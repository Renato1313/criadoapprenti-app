#!/usr/bin/env python3
"""Script simples de WhatsApp com IA local (Ollama).

Fluxo:
1) Abre WhatsApp Web e conversa por numero.
2) Envia saudacao + solicitacao (gerada por IA ou fallback fixo).
3) Aguarda resposta.
4) Extrai S500/S10 por regex e usa IA como fallback.
5) Salva no Excel.
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
from playwright.sync_api import sync_playwright

# ====== CONFIGURACAO ======
NUMERO = "5521996697733"  # sem +, sem espaco, sem traco
EMPRESA = "Empresa Exemplo"

SAUDACAO = "Bom dia"
SOLICITACAO_FALLBACK = "16"

USAR_IA = True
OLLAMA_MODEL = "qwen2.5:3b"
OLLAMA_ENDPOINT = "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT_SEGUNDOS = 60

ARQUIVO_XLSX = Path("saida/precos_combustivel.xlsx")
ABA_REGISTROS = "RegistrosWhatsApp"
CABECALHOS = [
    "DataHora",
    "Empresa",
    "Telefone",
    "SaudacaoEnviada",
    "SolicitacaoEnviada",
    "MensagemRecebida",
    "PrecoS500",
    "PrecoS10",
    "MetodoExtracao",
    "Status",
    "Observacoes",
]
TIMEOUT_RESPOSTA_SEGUNDOS = 7200  # 2h
POLL_SEGUNDOS = 2
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


def gerar_solicitacao_ia(empresa: str, fallback: str) -> str:
    if not USAR_IA:
        return fallback

    prompt = (
        "Crie uma unica mensagem curta para WhatsApp em portugues.\n"
        "Objetivo: pedir preco atual de combustivel S500 e S10.\n"
        "Seja educado e objetivo, sem emoji.\n"
        f"Empresa: {empresa}\n"
        "Retorne SOMENTE JSON com chave 'mensagem'."
    )

    try:
        data = call_ollama_json(prompt)
        msg = str(data.get("mensagem", "")).strip()
        if msg:
            return msg
    except Exception as exc:  # noqa: BLE001
        print(f"[aviso] IA nao gerou solicitacao, usando fallback. Detalhe: {exc}")

    return fallback


def extract_prices_regex(texto: str) -> dict:
    s500 = re.search(r"(?i)\bS\s*500\b[^0-9]{0,20}([0-9]+(?:[.,][0-9]+)?)", texto)
    s10 = re.search(r"(?i)\bS\s*10\b[^0-9]{0,20}([0-9]+(?:[.,][0-9]+)?)", texto)

    out = {}
    if s500:
        out["S500"] = to_float_br(s500.group(1))
    if s10:
        out["S10"] = to_float_br(s10.group(1))
    return out


def extract_prices_ia(texto: str) -> dict:
    prompt = (
        "Extraia os valores de combustivel da mensagem abaixo.\n"
        "Retorne SOMENTE JSON no formato: "
        '{"S500": numero_ou_null, "S10": numero_ou_null}\n'
        "Mensagem:\n"
        f"{texto}"
    )
    data = call_ollama_json(prompt)

    result = {"S500": None, "S10": None}
    for key in ("S500", "S10"):
        if key in data and data[key] is not None:
            result[key] = to_float_br(str(data[key]))
    return result


def extract_prices(texto: str) -> tuple[Optional[float], Optional[float], str]:
    base = extract_prices_regex(texto)
    s500 = base.get("S500")
    s10 = base.get("S10")
    metodo = "regex"

    if s500 is not None and s10 is not None:
        return s500, s10, metodo

    if not USAR_IA:
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


def wait_new_inbound(page, baseline_key, timeout_seconds=7200, poll_seconds=2):
    limite = time.time() + timeout_seconds
    while time.time() < limite:
        atual = get_last_inbound_message(page)
        if atual and atual["key"] != baseline_key:
            return atual
        page.wait_for_timeout(int(poll_seconds * 1000))
    raise TimeoutError("Nenhuma nova mensagem recebida dentro do tempo limite.")


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

    # Colunas de texto longo com quebra de linha
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 55
    ws.column_dimensions["K"].width = 35
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 14

    # Ajuste automatico basico para as demais colunas
    for col_idx in range(1, len(CABECALHOS) + 1):
        letter = get_column_letter(col_idx)
        if letter in {"E", "F", "K", "G", "H"}:
            continue

        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 28)

    for row_idx in range(2, ws.max_row + 1):
        ws.cell(row=row_idx, column=7).number_format = "0.0000"
        ws.cell(row=row_idx, column=8).number_format = "0.0000"
        ws.cell(row=row_idx, column=6).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=row_idx, column=11).alignment = Alignment(vertical="top", wrap_text=True)


def montar_status_e_observacoes(
    s500: Optional[float],
    s10: Optional[float],
    metodo_extracao: str,
) -> tuple[str, str]:
    if s500 is not None and s10 is not None:
        return "OK_COMPLETO", f"Valores extraidos com {metodo_extracao}"
    if s500 is not None or s10 is not None:
        faltante = "S10" if s500 is not None else "S500"
        return "OK_PARCIAL", f"Faltou {faltante}; metodo {metodo_extracao}"
    return "SEM_VALORES", f"Nenhum valor encontrado; metodo {metodo_extracao}"


def salvar_excel(
    empresa: str,
    numero: str,
    saudacao_enviada: str,
    solicitacao_enviada: str,
    mensagem_recebida: str,
    s500: Optional[float],
    s10: Optional[float],
    metodo_extracao: str,
) -> None:
    ARQUIVO_XLSX.parent.mkdir(parents=True, exist_ok=True)

    if ARQUIVO_XLSX.exists():
        wb = load_workbook(ARQUIVO_XLSX)
    else:
        wb = Workbook()
        # Remove a aba padrao para deixar apenas o layout de registros.
        padrao = wb.active
        wb.remove(padrao)

    ws = obter_aba_registros(wb)
    garantir_cabecalho(ws)
    status, observacoes = montar_status_e_observacoes(s500=s500, s10=s10, metodo_extracao=metodo_extracao)

    ws.append(
        [
            datetime.now().isoformat(timespec="seconds"),
            empresa,
            numero,
            saudacao_enviada,
            solicitacao_enviada,
            mensagem_recebida,
            s500,
            s10,
            metodo_extracao,
            status,
            observacoes,
        ]
    )
    aplicar_estilo_planilha(ws)
    wb.save(ARQUIVO_XLSX)


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

        page.wait_for_timeout(12000)
        input("Se a conversa abriu, aperte ENTER aqui no prompt para continuar...")

        caixa = page.locator('footer div[contenteditable="true"]').first
        caixa.wait_for(timeout=20000)

        baseline = get_last_inbound_message(page)
        baseline_key = baseline["key"] if baseline else None

        solicitacao = gerar_solicitacao_ia(empresa=EMPRESA, fallback=SOLICITACAO_FALLBACK)
        print(f"Solicitacao escolhida: {solicitacao}")

        caixa.click()
        caixa.fill(SAUDACAO)
        caixa.press("Enter")

        page.wait_for_timeout(1500)
        caixa.fill(solicitacao)
        caixa.press("Enter")

        print("Mensagens enviadas com sucesso! Aguardando resposta...")

        resposta = wait_new_inbound(
            page,
            baseline_key=baseline_key,
            timeout_seconds=TIMEOUT_RESPOSTA_SEGUNDOS,
            poll_seconds=POLL_SEGUNDOS,
        )

        print("\nMensagem recebida:")
        print("-" * 40)
        print(resposta["text"])
        print("-" * 40)

        s500, s10, metodo = extract_prices(resposta["text"])
        print(f"Extraido -> S500: {s500} | S10: {s10} | metodo: {metodo}")

        salvar_excel(
            empresa=EMPRESA,
            numero=NUMERO,
            saudacao_enviada=SAUDACAO,
            solicitacao_enviada=solicitacao,
            mensagem_recebida=resposta["text"],
            s500=s500,
            s10=s10,
            metodo_extracao=metodo,
        )
        print(f"Salvo em: {ARQUIVO_XLSX.resolve()}")
    except Exception as e:  # noqa: BLE001
        print(f"Erro: {e}")
    finally:
        input("Pressione ENTER para fechar...")
        context.close()
