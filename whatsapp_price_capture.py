#!/usr/bin/env python3
"""Automacao WhatsApp Web para solicitacao e captura de precos.

Fluxo principal:
1) Le contatos da planilha (aba Contatos: empresa + telefone).
2) Abre WhatsApp Web com sessao persistente.
3) Para cada contato:
   - envia saudacao
   - gera mensagem de solicitacao via IA local (ou usa fallback configurado)
   - aguarda resposta
   - extrai S500 e S10 com regex + fallback IA
   - salva na mesma planilha (aba Registros)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import requests
from openpyxl import Workbook, load_workbook
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

INBOUND_SELECTOR = "div.message-in"
CONTACT_HEADERS = ["empresa", "telefone", "contexto", "mensagem_solicitacao"]
RECORD_HEADERS = [
    "timestamp",
    "empresa",
    "telefone",
    "saudacao_enviada",
    "solicitacao_enviada",
    "mensagem_recebida",
    "S500",
    "S10",
    "metodo_extracao",
    "status",
    "erro",
]


@dataclass
class InboundMessage:
    key: str
    text: str


@dataclass
class Contact:
    company: str
    phone: str
    context: str = ""
    request_message: str = ""


@dataclass
class ContactResult:
    company: str
    phone: str
    greeting: str
    request_sent: str
    message_received: str
    s500: Optional[float]
    s10: Optional[float]
    extraction_method: str
    status: str
    error: str


def _to_float_br(raw: str) -> float:
    value = str(raw).strip()
    if "," in value and "." in value:
        value = value.replace(".", "").replace(",", ".")
    elif "," in value:
        value = value.replace(",", ".")
    return float(value)


def normalize_phone(raw_phone: str) -> str:
    digits = re.sub(r"\D+", "", raw_phone or "")
    return digits


def normalize_header(name: str) -> str:
    clean = (name or "").strip().lower()
    return clean.replace(" ", "_")


def ensure_contacts_template(workbook_path: Path, contacts_sheet: str, records_sheet: str) -> None:
    workbook_path.parent.mkdir(parents=True, exist_ok=True)

    if workbook_path.exists():
        wb = load_workbook(workbook_path)
    else:
        wb = Workbook()
        wb.active.title = contacts_sheet

    if contacts_sheet not in wb.sheetnames:
        wb.create_sheet(contacts_sheet)
    ws_contacts = wb[contacts_sheet]
    if ws_contacts.max_row < 1 or not any(ws_contacts.iter_rows(min_row=1, max_row=1, values_only=True)):
        ws_contacts.append(CONTACT_HEADERS)

    if records_sheet not in wb.sheetnames:
        ws_records = wb.create_sheet(records_sheet)
        ws_records.append(RECORD_HEADERS)
    else:
        ws_records = wb[records_sheet]
        if ws_records.max_row < 1 or not any(ws_records.iter_rows(min_row=1, max_row=1, values_only=True)):
            ws_records.append(RECORD_HEADERS)

    wb.save(workbook_path)


def load_contacts(workbook_path: Path, contacts_sheet: str) -> list[Contact]:
    wb = load_workbook(workbook_path)
    ws = wb[contacts_sheet]

    header_cells = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    header_map: dict[str, int] = {}
    for idx, header in enumerate(header_cells):
        if header:
            header_map[normalize_header(str(header))] = idx

    if "empresa" not in header_map or "telefone" not in header_map:
        raise ValueError(
            "A aba Contatos precisa ter cabecalhos 'empresa' e 'telefone'. "
            f"Cabecalhos atuais: {list(header_map.keys())}"
        )

    contacts: list[Contact] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        company = str(row[header_map["empresa"]] or "").strip()
        phone = normalize_phone(str(row[header_map["telefone"]] or ""))
        context = str(row[header_map["contexto"]] or "").strip() if "contexto" in header_map else ""
        request_message = (
            str(row[header_map["mensagem_solicitacao"]] or "").strip()
            if "mensagem_solicitacao" in header_map
            else ""
        )

        if not company and not phone:
            continue
        if not company or not phone:
            print(f"[aviso] Linha ignorada por dados incompletos: empresa={company!r}, telefone={phone!r}")
            continue

        contacts.append(
            Contact(
                company=company,
                phone=phone,
                context=context,
                request_message=request_message,
            )
        )

    return contacts


def call_ollama_json(
    prompt: str,
    model: str,
    endpoint: str,
    timeout: int,
) -> dict:
    response = requests.post(
        endpoint,
        json={
            "model": model,
            "prompt": prompt,
            "format": "json",
            "stream": False,
        },
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    return json.loads(payload["response"])


def build_request_message(
    company: str,
    context: str,
    default_request: str,
    ai_enabled: bool,
    ai_model: str,
    ai_endpoint: str,
    ai_timeout: int,
) -> str:
    if not ai_enabled:
        return default_request

    prompt = (
        "Voce cria mensagens curtas para WhatsApp para pedir preco de combustivel.\n"
        "Retorne JSON com a chave 'mensagem'.\n"
        "Regras:\n"
        "- mensagem objetiva e educada\n"
        "- em portugues\n"
        "- pedir valores atuais de S500 e S10\n"
        "- sem emojis\n"
        "- uma unica mensagem curta\n"
        f"- empresa: {company}\n"
        f"- contexto adicional: {context or 'sem contexto adicional'}\n"
        f"- fallback permitido: {default_request}\n"
    )

    try:
        data = call_ollama_json(
            prompt=prompt,
            model=ai_model,
            endpoint=ai_endpoint,
            timeout=ai_timeout,
        )
        message = str(data.get("mensagem", "")).strip()
        if message:
            return message
    except Exception as exc:  # noqa: BLE001
        print(f"[aviso] IA nao gerou mensagem para {company}: {exc}")

    return default_request


def extract_prices_regex(text: str) -> dict[str, float]:
    s500_match = re.search(r"(?i)\bS\s*500\b[^0-9]{0,25}([0-9]+(?:[.,][0-9]+)?)", text)
    s10_match = re.search(r"(?i)\bS\s*10\b[^0-9]{0,25}([0-9]+(?:[.,][0-9]+)?)", text)

    prices: dict[str, float] = {}
    if s500_match:
        prices["S500"] = _to_float_br(s500_match.group(1))
    if s10_match:
        prices["S10"] = _to_float_br(s10_match.group(1))
    return prices


def extract_prices_ai(
    text: str,
    ai_model: str,
    ai_endpoint: str,
    ai_timeout: int,
) -> dict[str, Optional[float]]:
    prompt = (
        "Extraia os precos de combustivel da mensagem abaixo.\n"
        "Retorne somente JSON com chaves S500 e S10.\n"
        "Quando nao houver valor, use null.\n\n"
        f"MENSAGEM:\n{text}\n"
    )
    data = call_ollama_json(
        prompt=prompt,
        model=ai_model,
        endpoint=ai_endpoint,
        timeout=ai_timeout,
    )
    parsed: dict[str, Optional[float]] = {"S500": None, "S10": None}
    for key in ("S500", "S10"):
        if key not in data or data[key] is None:
            continue
        parsed[key] = _to_float_br(str(data[key]))
    return parsed


def extract_prices_with_fallback(
    text: str,
    ai_enabled: bool,
    ai_model: str,
    ai_endpoint: str,
    ai_timeout: int,
) -> tuple[Optional[float], Optional[float], str]:
    prices = extract_prices_regex(text)
    method = "regex"

    s500 = prices.get("S500")
    s10 = prices.get("S10")
    if s500 is not None and s10 is not None:
        return s500, s10, method

    if not ai_enabled:
        return s500, s10, method

    try:
        ai_prices = extract_prices_ai(
            text=text,
            ai_model=ai_model,
            ai_endpoint=ai_endpoint,
            ai_timeout=ai_timeout,
        )
        if s500 is None:
            s500 = ai_prices.get("S500")
        if s10 is None:
            s10 = ai_prices.get("S10")
        method = "regex+ia" if (s500 is not None or s10 is not None) else "regex"
    except Exception as exc:  # noqa: BLE001
        print(f"[aviso] IA nao conseguiu extrair precos: {exc}")

    return s500, s10, method


def get_last_inbound_message(page: Page) -> Optional[InboundMessage]:
    messages = page.locator(INBOUND_SELECTOR)
    count = messages.count()
    if count == 0:
        return None

    last_message = messages.nth(count - 1)
    payload = last_message.evaluate(
        """(el) => {
          const clean = (s) => (s || "").replace(/\\u200e/g, "").trim();

          const copyable = el.querySelector("div.copyable-text");
          const attrStamp =
            el.getAttribute("data-id") ||
            (copyable ? copyable.getAttribute("data-pre-plain-text") : "") ||
            "";

          const collected = [];
          const blockSelectors = [
            "div.selectable-text.copyable-text",
            "div.copyable-text",
            "span.selectable-text.copyable-text",
            "span[data-lexical-text='true']",
            "span[dir='auto']"
          ];

          for (const selector of blockSelectors) {
            const nodes = el.querySelectorAll(selector);
            for (const node of nodes) {
              const txt = clean(node.textContent);
              if (txt) collected.push(txt);
            }
            if (collected.length) break;
          }

          let text = "";
          if (collected.length) {
            text = clean(collected.join("\\n"));
          } else {
            text = clean(el.innerText);
          }

          return { stamp: attrStamp, text };
        }"""
    )

    text = (payload.get("text") or "").strip()
    if not text:
        return None

    stamp = payload.get("stamp") or ""
    key = f"{stamp}|{text}"
    return InboundMessage(key=key, text=text)


def wait_for_new_inbound_message(
    page: Page,
    baseline_key: Optional[str],
    timeout_seconds: int,
    poll_seconds: float,
) -> InboundMessage:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        current = get_last_inbound_message(page)
        if current and current.key != baseline_key:
            return current
        page.wait_for_timeout(int(poll_seconds * 1000))

    raise TimeoutError("Nenhuma nova mensagem recebida dentro do tempo limite.")


def wait_for_login(page: Page, login_timeout_seconds: int) -> None:
    page.goto("https://web.whatsapp.com/", wait_until="domcontentloaded")
    print("Aguardando WhatsApp Web carregar...")

    try:
        page.wait_for_selector("#pane-side", timeout=30_000)
        print("Sessao ativa detectada.")
        return
    except PlaywrightTimeoutError:
        print("Sessao nao encontrada imediatamente. Se necessario, escaneie o QR code.")

    page.wait_for_selector("#pane-side", timeout=login_timeout_seconds * 1000)
    print("Login concluido.")


def open_chat(page: Page, phone: str) -> None:
    url = f"https://web.whatsapp.com/send?phone={phone}&app_absent=0"
    page.goto(url, wait_until="domcontentloaded", timeout=120_000)

    try:
        page.wait_for_selector("#main footer", timeout=60_000)
    except PlaywrightTimeoutError as exc:
        invalid_msg = page.locator("text=Phone number shared via url is invalid")
        if invalid_msg.count() > 0:
            raise ValueError("Numero invalido para abertura direta no WhatsApp.") from exc
        raise RuntimeError("Nao foi possivel abrir a conversa no tempo esperado.") from exc


def send_message(page: Page, text: str) -> None:
    composer = page.locator("#main footer div[contenteditable='true'][role='textbox']").first
    composer.wait_for(timeout=20_000)
    composer.click()
    composer.fill("")
    page.keyboard.type(text, delay=18)
    page.keyboard.press("Enter")


def append_result(workbook_path: Path, records_sheet: str, result: ContactResult) -> None:
    wb = load_workbook(workbook_path)
    ws = wb[records_sheet]
    ws.append(
        [
            datetime.now().isoformat(timespec="seconds"),
            result.company,
            result.phone,
            result.greeting,
            result.request_sent,
            result.message_received,
            result.s500,
            result.s10,
            result.extraction_method,
            result.status,
            result.error,
        ]
    )
    wb.save(workbook_path)


def process_contact(
    page: Page,
    contact: Contact,
    greeting: str,
    default_request: str,
    timeout_seconds: int,
    poll_seconds: float,
    ai_enabled: bool,
    ai_model: str,
    ai_endpoint: str,
    ai_timeout: int,
) -> ContactResult:
    print(f"\n=== Processando {contact.company} ({contact.phone}) ===")
    request_message = (
        contact.request_message
        if contact.request_message
        else build_request_message(
            company=contact.company,
            context=contact.context,
            default_request=default_request,
            ai_enabled=ai_enabled,
            ai_model=ai_model,
            ai_endpoint=ai_endpoint,
            ai_timeout=ai_timeout,
        )
    )

    open_chat(page, contact.phone)
    baseline = get_last_inbound_message(page)
    baseline_key = baseline.key if baseline else None

    send_message(page, greeting)
    page.wait_for_timeout(1000)
    send_message(page, request_message)
    print(f"Enviado -> saudacao: {greeting!r} | solicitacao: {request_message!r}")

    inbound = wait_for_new_inbound_message(
        page=page,
        baseline_key=baseline_key,
        timeout_seconds=timeout_seconds,
        poll_seconds=poll_seconds,
    )
    print(f"Resposta recebida de {contact.company}.")

    s500, s10, method = extract_prices_with_fallback(
        text=inbound.text,
        ai_enabled=ai_enabled,
        ai_model=ai_model,
        ai_endpoint=ai_endpoint,
        ai_timeout=ai_timeout,
    )

    return ContactResult(
        company=contact.company,
        phone=contact.phone,
        greeting=greeting,
        request_sent=request_message,
        message_received=inbound.text,
        s500=s500,
        s10=s10,
        extraction_method=method,
        status="ok",
        error="",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Automacao WhatsApp com captura de precos em planilha.")
    parser.add_argument(
        "--workbook",
        default=os.getenv("WORKBOOK_XLSX", "saida/precos_combustivel.xlsx"),
        help="Planilha principal (.xlsx) com abas Contatos e Registros.",
    )
    parser.add_argument(
        "--contacts-sheet",
        default=os.getenv("CONTACTS_SHEET", "Contatos"),
        help="Nome da aba de contatos.",
    )
    parser.add_argument(
        "--records-sheet",
        default=os.getenv("RECORDS_SHEET", "Registros"),
        help="Nome da aba de registros.",
    )
    parser.add_argument(
        "--greeting",
        default=os.getenv("WA_GREETING", "Bom dia"),
        help='Mensagem inicial enviada para cada contato (default: "Bom dia").',
    )
    parser.add_argument(
        "--default-request",
        default=os.getenv("WA_DEFAULT_REQUEST", "16"),
        help='Fallback de solicitacao quando nao houver mensagem personalizada/IA (default: "16").',
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=int(os.getenv("WAIT_TIMEOUT_SECONDS", "7200")),
        help="Timeout em segundos para aguardar resposta por contato.",
    )
    parser.add_argument(
        "--poll",
        type=float,
        default=float(os.getenv("WAIT_POLL_SECONDS", "2")),
        help="Intervalo de polling em segundos.",
    )
    parser.add_argument(
        "--user-data-dir",
        default=os.getenv("WHATSAPP_USER_DATA_DIR", ".wweb_profile"),
        help="Diretorio para persistir sessao do WhatsApp.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Roda navegador em modo headless (na pratica, pode dificultar login no WhatsApp).",
    )
    parser.add_argument(
        "--login-timeout",
        type=int,
        default=int(os.getenv("LOGIN_TIMEOUT_SECONDS", "180")),
        help="Tempo maximo para concluir login no WhatsApp.",
    )
    parser.add_argument(
        "--ai-enabled",
        action="store_true",
        help="Ativa IA local (Ollama) para gerar solicitacao e fallback de extracao.",
    )
    parser.add_argument(
        "--ai-model",
        default=os.getenv("OLLAMA_MODEL", "qwen2.5:3b"),
        help="Modelo Ollama (default: qwen2.5:3b).",
    )
    parser.add_argument(
        "--ai-endpoint",
        default=os.getenv("OLLAMA_ENDPOINT", "http://localhost:11434/api/generate"),
        help="Endpoint de geracao do Ollama.",
    )
    parser.add_argument(
        "--ai-timeout",
        type=int,
        default=int(os.getenv("OLLAMA_TIMEOUT_SECONDS", "60")),
        help="Timeout de chamadas IA em segundos.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    workbook_path = Path(args.workbook)
    profile_path = Path(args.user_data_dir)

    ensure_contacts_template(
        workbook_path=workbook_path,
        contacts_sheet=args.contacts_sheet,
        records_sheet=args.records_sheet,
    )

    contacts = load_contacts(workbook_path=workbook_path, contacts_sheet=args.contacts_sheet)
    if not contacts:
        print(
            f"Nenhum contato encontrado na aba '{args.contacts_sheet}'. "
            f"Preencha a planilha em: {workbook_path.resolve()}"
        )
        return 1

    print(f"Contatos carregados: {len(contacts)}")
    print(f"Planilha ativa: {workbook_path.resolve()}")
    if args.ai_enabled:
        print(f"IA habilitada com modelo: {args.ai_model}")
    else:
        print("IA desabilitada. Usando solicitacao padrao/fixa.")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            headless=args.headless,
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            wait_for_login(page, login_timeout_seconds=args.login_timeout)

            ok_count = 0
            err_count = 0
            for contact in contacts:
                try:
                    result = process_contact(
                        page=page,
                        contact=contact,
                        greeting=args.greeting,
                        default_request=args.default_request,
                        timeout_seconds=args.timeout,
                        poll_seconds=args.poll,
                        ai_enabled=args.ai_enabled,
                        ai_model=args.ai_model,
                        ai_endpoint=args.ai_endpoint,
                        ai_timeout=args.ai_timeout,
                    )
                    ok_count += 1
                except Exception as exc:  # noqa: BLE001
                    err_count += 1
                    result = ContactResult(
                        company=contact.company,
                        phone=contact.phone,
                        greeting=args.greeting,
                        request_sent="",
                        message_received="",
                        s500=None,
                        s10=None,
                        extraction_method="erro",
                        status="erro",
                        error=str(exc),
                    )
                    print(f"[erro] Falha em {contact.company}: {exc}")

                append_result(
                    workbook_path=workbook_path,
                    records_sheet=args.records_sheet,
                    result=result,
                )

            print("\nProcessamento finalizado.")
            print(f"Sucesso: {ok_count} | Erros: {err_count}")
            print(f"Registros salvos em: {workbook_path.resolve()}")
            return 0 if err_count == 0 else 2
        finally:
            context.close()


if __name__ == "__main__":
    raise SystemExit(main())
