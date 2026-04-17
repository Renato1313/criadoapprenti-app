#!/usr/bin/env python3
"""WhatsApp com automacao de conversa para capturar S500 e S10.

Fluxo:
1) Abre WhatsApp Web e envia saudacao + pedido inicial.
2) Aguarda respostas e so conclui quando capturar S500 e S10.
3) Se vier resposta sem os valores, envia follow-up educado automaticamente.
4) Se nao houver resposta por 40 minutos, envia lembrete educado.
5) Salva resultado no Excel com status claro.
"""

from __future__ import annotations

import json
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# ====== CONFIGURACAO ======
NUMERO = "5521996697733"  # sem +, sem espaco, sem traco
EMPRESA = "Empresa Exemplo"

SAUDACAO = "Bom dia"
SOLICITACAO_FALLBACK = "Poderia me informar os valores atuais de S500 e S10, por favor?"

USAR_IA_MENSAGENS = False
USAR_IA_EXTRACAO = True
OLLAMA_MODEL = "llama3.2:3b"
OLLAMA_ENDPOINT = "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT_SEGUNDOS = 12

ARQUIVO_XLSX = Path("saida/precos_combustivel.xlsx")
ABA_REGISTROS = "RegistrosWhatsApp"
CABECALHOS = [
    "DataHora",
    "Empresa",
    "Telefone",
    "SaudacaoEnviada",
    "SolicitacaoInicial",
    "UltimaSolicitacaoEnviada",
    "UltimaMensagemRecebida",
    "PrecoS500",
    "PrecoS10",
    "MetodoExtracao",
    "Status",
    "TentativasFollowup",
    "Observacoes",
]

TIMEOUT_RESPOSTA_SEGUNDOS = 7200  # 2h
POLL_SEGUNDOS = 2
INTERVALO_LEMBRETE_SEGUNDOS = 40 * 60  # 40 minutos
MAX_FOLLOWUPS = 6
PAUSAR_NO_FINAL = False
# ==========================


def to_float_br(raw: str) -> float:
    value = str(raw).strip()
    if "," in value and "." in value:
        value = value.replace(".", "").replace(",", ".")
    elif "," in value:
        value = value.replace(",", ".")
    return float(value)


def call_ollama_json(prompt: str) -> dict:
    response = requests.post(
        OLLAMA_ENDPOINT,
        json={
            "model": OLLAMA_MODEL,
            "prompt": prompt,
            "format": "json",
            "stream": False,
        },
        timeout=OLLAMA_TIMEOUT_SEGUNDOS,
    )
    response.raise_for_status()
    payload = response.json()
    return json.loads(payload["response"])


def mensagem_valida_para_envio(msg: str, empresa: str, exigir_s500_s10: bool) -> bool:
    msg_lower = msg.lower()
    bloqueadas = [
        "solicite",
        "entre em contato conosco",
        "estamos solicitando",
        empresa.lower().strip(),
    ]
    if any(p and p in msg_lower for p in bloqueadas):
        return False
    if exigir_s500_s10 and ("s500" not in msg_lower or "s10" not in msg_lower):
        return False
    return True


def gerar_solicitacao_ia(empresa: str, fallback: str) -> str:
    if not USAR_IA_MENSAGENS:
        return fallback

    prompt = (
        "Voce esta escrevendo UMA mensagem curta para WhatsApp no papel de COMPRADOR.\n"
        "Objetivo: pedir os valores atuais de combustivel S500 e S10 ao fornecedor.\n"
        "Regras obrigatorias:\n"
        "- portugues do Brasil\n"
        "- educada e objetiva\n"
        "- sem emoji\n"
        "- nao usar a palavra 'solicite'\n"
        "- nao usar 'entre em contato conosco'\n"
        "- nao mencionar o nome da empresa compradora\n"
        "- ate 160 caracteres\n"
        f"- nome da empresa compradora (NAO mencionar): {empresa}\n\n"
        "Retorne SOMENTE JSON valido com a chave 'mensagem'."
    )

    try:
        data = call_ollama_json(prompt)
        msg = re.sub(r"\s+", " ", str(data.get("mensagem", "")).strip())
        if msg and mensagem_valida_para_envio(msg=msg, empresa=empresa, exigir_s500_s10=True):
            return msg
        print("[aviso] IA gerou solicitacao fora do padrao, usando fallback.")
    except Exception as exc:  # noqa: BLE001
        print(f"[aviso] IA nao gerou solicitacao, usando fallback. Detalhe: {exc}")

    return fallback


def fallback_followup(faltantes: list[str], lembrete: bool) -> str:
    if faltantes == ["S500", "S10"]:
        return (
            "Quando puder, poderia me informar os valores atuais de S500 e S10, por favor?"
            if lembrete
            else "Poderia me informar os valores atuais de S500 e S10, por favor?"
        )
    if faltantes == ["S500"]:
        return (
            "Quando puder, poderia me informar tambem o valor de S500, por favor?"
            if lembrete
            else "Poderia me informar tambem o valor de S500, por favor?"
        )
    return (
        "Quando puder, poderia me informar tambem o valor de S10, por favor?"
        if lembrete
        else "Poderia me informar tambem o valor de S10, por favor?"
    )


def gerar_followup_ia(
    empresa: str,
    faltantes: list[str],
    ultima_resposta: str,
    lembrete: bool,
) -> str:
    fallback = fallback_followup(faltantes=faltantes, lembrete=lembrete)
    if not USAR_IA_MENSAGENS:
        return fallback

    faltantes_txt = " e ".join(faltantes)
    tipo = "lembrete educado" if lembrete else "pedido de complemento"
    prompt = (
        f"Escreva UMA mensagem curta de WhatsApp para {tipo}.\n"
        "Contexto: voce e comprador e quer os valores faltantes de combustivel.\n"
        f"Faltam estes valores: {faltantes_txt}.\n"
        "Regras obrigatorias:\n"
        "- portugues do Brasil\n"
        "- educada e objetiva\n"
        "- sem emoji\n"
        "- nao usar a palavra 'solicite'\n"
        "- nao usar 'entre em contato conosco'\n"
        "- nao mencionar empresa compradora\n"
        f"- ultima resposta recebida: {ultima_resposta[:200] or 'sem resposta'}\n"
        f"- deve citar explicitamente: {faltantes_txt}\n\n"
        "Retorne SOMENTE JSON valido com a chave 'mensagem'."
    )

    try:
        data = call_ollama_json(prompt)
        msg = re.sub(r"\s+", " ", str(data.get("mensagem", "")).strip())
        if not msg:
            return fallback
        msg_lower = msg.lower()
        for item in faltantes:
            if item.lower() not in msg_lower:
                return fallback
        if not mensagem_valida_para_envio(msg=msg, empresa=empresa, exigir_s500_s10=False):
            return fallback
        return msg
    except Exception as exc:  # noqa: BLE001
        print(f"[aviso] IA nao gerou follow-up, usando fallback. Detalhe: {exc}")
        return fallback


def extract_prices_regex(texto: str) -> dict[str, float]:
    s500 = re.search(r"(?i)\bS\s*500\b[^0-9]{0,25}([0-9]+(?:[.,][0-9]+)?)", texto)
    s10 = re.search(r"(?i)\bS\s*10\b[^0-9]{0,25}([0-9]+(?:[.,][0-9]+)?)", texto)

    out: dict[str, float] = {}
    if s500:
        out["S500"] = to_float_br(s500.group(1))
    if s10:
        out["S10"] = to_float_br(s10.group(1))
    return out


def extract_prices_ia(texto: str) -> dict[str, Optional[float]]:
    prompt = (
        "Extraia os valores de combustivel da mensagem abaixo.\n"
        "Retorne SOMENTE JSON no formato: "
        '{"S500": numero_ou_null, "S10": numero_ou_null}\n'
        "Nao invente valor.\n"
        f"Mensagem:\n{texto}"
    )
    data = call_ollama_json(prompt)

    result: dict[str, Optional[float]] = {"S500": None, "S10": None}
    for key in ("S500", "S10"):
        if key in data and data[key] is not None:
            try:
                result[key] = to_float_br(str(data[key]))
            except Exception:  # noqa: BLE001
                result[key] = None
    return result


def extract_prices(texto: str) -> tuple[Optional[float], Optional[float], str]:
    base = extract_prices_regex(texto)
    s500 = base.get("S500")
    s10 = base.get("S10")
    metodo = "regex"

    if s500 is not None and s10 is not None:
        return s500, s10, metodo

    # Evita IA em textos totalmente fora de contexto para nao "inventar" valor.
    tem_indicio = re.search(r"(?i)\bS\s*500\b|\bS\s*10\b|[0-9]+[.,][0-9]+", texto) is not None
    if not USAR_IA_EXTRACAO or not tem_indicio:
        return s500, s10, metodo

    try:
        ai_out = extract_prices_ia(texto)
        if s500 is None:
            s500 = ai_out.get("S500")
        if s10 is None:
            s10 = ai_out.get("S10")
        if s500 is not None or s10 is not None:
            metodo = "regex+ia"
    except Exception as exc:  # noqa: BLE001
        print(f"[aviso] IA falhou na extracao, mantendo regex. Detalhe: {exc}")

    return s500, s10, metodo


def get_last_inbound_message(page):
    msgs = page.locator("div.message-in")
    count = msgs.count()
    if count == 0:
        return None

    last = msgs.nth(count - 1)
    payload = last.evaluate(
        """(el) => {
            const clean = (s) => (s || "").replace(/\\u200e/g, "").trim();

            const copyable = el.querySelector("div.copyable-text");
            const stamp =
              el.getAttribute("data-id") ||
              (copyable ? copyable.getAttribute("data-pre-plain-text") : "") ||
              "";

            const nodes = el.querySelectorAll(
              "div.selectable-text.copyable-text, div.copyable-text, span[data-lexical-text='true'], span[dir='auto']"
            );

            const parts = [];
            nodes.forEach((n) => {
              const t = clean(n.textContent);
              if (t) parts.push(t);
            });

            let text = "";
            if (parts.length) {
              text = clean(parts.join("\\n"));
            } else {
              text = clean(el.innerText);
            }

            return { stamp, text };
        }"""
    )

    text = (payload.get("text") or "").strip()
    if not text:
        return None

    key = f"{payload.get('stamp', '')}|{text}"
    return {"key": key, "text": text}


def send_message(page, texto: str) -> None:
    caixa = page.locator('#main footer div[contenteditable="true"][role="textbox"]').first
    caixa.wait_for(timeout=20_000)
    caixa.click()
    caixa.fill("")
    page.keyboard.type(texto, delay=18)
    page.keyboard.press("Enter")


def faltantes_precos(s500: Optional[float], s10: Optional[float]) -> list[str]:
    faltantes = []
    if s500 is None:
        faltantes.append("S500")
    if s10 is None:
        faltantes.append("S10")
    return faltantes


def coletar_precos_em_conversa(
    page,
    baseline_key: Optional[str],
    solicitacao_inicial: str,
) -> dict:
    inicio = time.time()
    ultimo_evento = time.time()
    ultima_solicitacao = solicitacao_inicial
    ultimo_key = baseline_key
    ultima_msg_recebida = ""
    followups = 0
    historico_respostas: list[str] = []

    s500_final: Optional[float] = None
    s10_final: Optional[float] = None
    metodo_final = "regex"

    while time.time() - inicio < TIMEOUT_RESPOSTA_SEGUNDOS:
        atual = get_last_inbound_message(page)
        if atual and atual["key"] != ultimo_key:
            ultimo_key = atual["key"]
            ultima_msg_recebida = atual["text"]
            historico_respostas.append(ultima_msg_recebida)
            ultimo_evento = time.time()
            print(f"[recebido] {ultima_msg_recebida}")

            s500_msg, s10_msg, metodo_msg = extract_prices(ultima_msg_recebida)
            if metodo_msg == "regex+ia":
                metodo_final = "regex+ia"
            if s500_final is None and s500_msg is not None:
                s500_final = s500_msg
            if s10_final is None and s10_msg is not None:
                s10_final = s10_msg

            faltam = faltantes_precos(s500=s500_final, s10=s10_final)
            if not faltam:
                return {
                    "sucesso": True,
                    "s500": s500_final,
                    "s10": s10_final,
                    "metodo": metodo_final,
                    "ultima_msg": ultima_msg_recebida,
                    "ultima_solicitacao": ultima_solicitacao,
                    "status": "OK_COMPLETO",
                    "tentativas_followup": followups,
                    "obs": "Valores S500 e S10 capturados com sucesso.",
                    "historico": " || ".join(historico_respostas),
                }

            if followups < MAX_FOLLOWUPS:
                follow = gerar_followup_ia(
                    empresa=EMPRESA,
                    faltantes=faltam,
                    ultima_resposta=ultima_msg_recebida,
                    lembrete=False,
                )
                send_message(page, follow)
                ultima_solicitacao = follow
                followups += 1
                ultimo_evento = time.time()
                print(f"[follow-up] {follow}")

        if time.time() - ultimo_evento >= INTERVALO_LEMBRETE_SEGUNDOS and followups < MAX_FOLLOWUPS:
            faltam = faltantes_precos(s500=s500_final, s10=s10_final)
            lembrete = gerar_followup_ia(
                empresa=EMPRESA,
                faltantes=faltam if faltam else ["S500", "S10"],
                ultima_resposta=ultima_msg_recebida,
                lembrete=True,
            )
            send_message(page, lembrete)
            ultima_solicitacao = lembrete
            followups += 1
            ultimo_evento = time.time()
            print(f"[lembrete 40min] {lembrete}")

        page.wait_for_timeout(int(POLL_SEGUNDOS * 1000))

    faltam = faltantes_precos(s500=s500_final, s10=s10_final)
    if s500_final is None and s10_final is None and not historico_respostas:
        status = "TIMEOUT_SEM_RESPOSTA"
        obs = "Nenhuma resposta recebida no periodo."
    elif s500_final is None and s10_final is None:
        status = "TIMEOUT_SEM_VALORES"
        obs = "Respostas recebidas sem valores validos de S500/S10."
    else:
        status = "TIMEOUT_PARCIAL"
        obs = f"Timeout com valor parcial. Faltou: {', '.join(faltam)}."

    return {
        "sucesso": False,
        "s500": s500_final,
        "s10": s10_final,
        "metodo": metodo_final,
        "ultima_msg": ultima_msg_recebida,
        "ultima_solicitacao": ultima_solicitacao,
        "status": status,
        "tentativas_followup": followups,
        "obs": obs,
        "historico": " || ".join(historico_respostas),
    }


def obter_aba_registros(wb):
    if ABA_REGISTROS not in wb.sheetnames:
        ws = wb.create_sheet(ABA_REGISTROS)
    else:
        ws = wb[ABA_REGISTROS]
    return ws


def garantir_cabecalho(ws) -> None:
    primeira_linha = [ws.cell(row=1, column=i).value for i in range(1, len(CABECALHOS) + 1)]
    if primeira_linha == CABECALHOS:
        return

    if ws.max_row == 1 and all(v in (None, "") for v in primeira_linha):
        for col, header in enumerate(CABECALHOS, start=1):
            ws.cell(row=1, column=col, value=header)
        return

    ws.insert_rows(1)
    for col, header in enumerate(CABECALHOS, start=1):
        ws.cell(row=1, column=col, value=header)


def aplicar_estilo_planilha(ws) -> None:
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="D9D9D9")
    data_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, _ in enumerate(CABECALHOS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = data_border

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CABECALHOS))}{max(ws.max_row, 1)}"

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(CABECALHOS)):
        for cell in row:
            cell.border = data_border
            cell.alignment = Alignment(vertical="top", wrap_text=False)

    largura_por_cabecalho = {
        "SolicitacaoInicial": 40,
        "UltimaSolicitacaoEnviada": 42,
        "UltimaMensagemRecebida": 55,
        "Observacoes": 48,
    }
    for idx, header in enumerate(CABECALHOS, start=1):
        letter = get_column_letter(idx)
        if header in largura_por_cabecalho:
            ws.column_dimensions[letter].width = largura_por_cabecalho[header]
            continue

        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 28)

    idx_s500 = CABECALHOS.index("PrecoS500") + 1
    idx_s10 = CABECALHOS.index("PrecoS10") + 1
    idx_msg = CABECALHOS.index("UltimaMensagemRecebida") + 1
    idx_obs = CABECALHOS.index("Observacoes") + 1
    for row_idx in range(2, ws.max_row + 1):
        ws.cell(row=row_idx, column=idx_s500).number_format = "0.0000"
        ws.cell(row=row_idx, column=idx_s10).number_format = "0.0000"
        ws.cell(row=row_idx, column=idx_msg).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=row_idx, column=idx_obs).alignment = Alignment(vertical="top", wrap_text=True)


def salvar_excel(
    empresa: str,
    numero: str,
    saudacao_enviada: str,
    solicitacao_inicial: str,
    resultado: dict,
) -> None:
    ARQUIVO_XLSX.parent.mkdir(parents=True, exist_ok=True)

    if ARQUIVO_XLSX.exists():
        wb = load_workbook(ARQUIVO_XLSX)
    else:
        wb = Workbook()
        padrao = wb.active
        wb.remove(padrao)

    ws = obter_aba_registros(wb)
    garantir_cabecalho(ws)
    ws.append(
        [
            datetime.now().isoformat(timespec="seconds"),
            empresa,
            numero,
            saudacao_enviada,
            solicitacao_inicial,
            resultado["ultima_solicitacao"],
            resultado["ultima_msg"],
            resultado["s500"],
            resultado["s10"],
            resultado["metodo"],
            resultado["status"],
            resultado["tentativas_followup"],
            f'{resultado["obs"]} Historico: {resultado["historico"][:800]}',
        ]
    )
    aplicar_estilo_planilha(ws)
    wb.save(ARQUIVO_XLSX)


def main() -> None:
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir="perfil_whatsapp",
            headless=False,
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            print("Abrindo conversa pelo numero...")
            page.goto(
                f"https://web.whatsapp.com/send?phone={NUMERO}&app_absent=0",
                wait_until="domcontentloaded",
                timeout=120000,
            )
            page.wait_for_selector("#main footer", timeout=60_000)
            page.wait_for_timeout(1200)

            baseline = get_last_inbound_message(page)
            baseline_key = baseline["key"] if baseline else None

            solicitacao_inicial = gerar_solicitacao_ia(empresa=EMPRESA, fallback=SOLICITACAO_FALLBACK)
            print(f"Solicitacao inicial: {solicitacao_inicial}")

            send_message(page, SAUDACAO)
            page.wait_for_timeout(1500)
            send_message(page, solicitacao_inicial)
            print("Mensagens enviadas com sucesso! Monitorando conversa...")

            resultado = coletar_precos_em_conversa(
                page=page,
                baseline_key=baseline_key,
                solicitacao_inicial=solicitacao_inicial,
            )

            print(f"Status final: {resultado['status']}")
            print(
                f"Extraido -> S500: {resultado['s500']} | "
                f"S10: {resultado['s10']} | metodo: {resultado['metodo']}"
            )

            salvar_excel(
                empresa=EMPRESA,
                numero=NUMERO,
                saudacao_enviada=SAUDACAO,
                solicitacao_inicial=solicitacao_inicial,
                resultado=resultado,
            )
            print(f"Salvo em: {ARQUIVO_XLSX.resolve()}")

        except PlaywrightTimeoutError as exc:
            print(f"Erro de timeout no WhatsApp Web: {exc}")
        except Exception as exc:  # noqa: BLE001
            print(f"Erro: {exc}")
        finally:
            if PAUSAR_NO_FINAL:
                input("Pressione ENTER para fechar...")
            context.close()


if __name__ == "__main__":
    main()
