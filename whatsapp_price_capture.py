#!/usr/bin/env python3
"""Automacao WhatsApp Web para captura de precos de combustivel.

Fluxo:
1) Abre WhatsApp Web com sessao persistente.
2) Abre conversa por numero.
3) Envia saudacao e codigo de solicitacao (ex.: "16").
4) Aguarda uma nova mensagem recebida.
5) Extrai S500 e S10 via regex.
6) Salva o resultado em arquivo Excel (.xlsx).
"""

from __future__ import annotations

import argparse
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook, load_workbook
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

INBOUND_SELECTOR = "div.message-in"


@dataclass
class InboundMessage:
    key: str
    text: str


def _to_float_br(raw: str) -> float:
    value = raw.strip()
    if "," in value and "." in value:
        value = value.replace(".", "").replace(",", ".")
    elif "," in value:
        value = value.replace(",", ".")
    return float(value)


def extract_prices(text: str) -> dict[str, float]:
    s500_match = re.search(r"(?i)\bS\s*500\b[^0-9]{0,20}([0-9]+(?:[.,][0-9]+)?)", text)
    s10_match = re.search(r"(?i)\bS\s*10\b[^0-9]{0,20}([0-9]+(?:[.,][0-9]+)?)", text)

    prices: dict[str, float] = {}
    if s500_match:
        prices["S500"] = _to_float_br(s500_match.group(1))
    if s10_match:
        prices["S10"] = _to_float_br(s10_match.group(1))
    return prices


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
    page: Page, baseline_key: Optional[str], timeout_seconds: int, poll_seconds: float
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
    page.goto(url, wait_until="domcontentloaded")

    try:
        page.wait_for_selector("#main footer", timeout=60_000)
    except PlaywrightTimeoutError as exc:
        invalid_msg = page.locator("text=Phone number shared via url is invalid")
        if invalid_msg.count() > 0:
            raise ValueError("Numero invalido para abertura direta no WhatsApp.") from exc
        raise RuntimeError("Nao foi possivel abrir a conversa no tempo esperado.") from exc


def send_message(page: Page, text: str) -> None:
    composer = page.locator("#main footer div[contenteditable='true'][role='textbox']").first
    composer.wait_for(timeout=15_000)
    composer.click()
    page.keyboard.type(text, delay=20)
    page.keyboard.press("Enter")


def save_to_excel(
    output_path: Path,
    phone: str,
    greeting: str,
    request_code: str,
    message_text: str,
    s500: Optional[float],
    s10: Optional[float],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists():
        wb = load_workbook(output_path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Precos"
        ws.append(
            [
                "timestamp",
                "telefone",
                "saudacao",
                "codigo_solicitado",
                "mensagem_recebida",
                "S500",
                "S10",
            ]
        )

    ws.append(
        [
            datetime.now().isoformat(timespec="seconds"),
            phone,
            greeting,
            request_code,
            message_text,
            s500,
            s10,
        ]
    )
    wb.save(output_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Captura de precos no WhatsApp Web.")
    parser.add_argument("--phone", default=os.getenv("WHATSAPP_PHONE"), help="Telefone no formato internacional, ex: 5511999999999")
    parser.add_argument("--greeting", default=os.getenv("WA_GREETING", "Bom dia"), help='Mensagem inicial (default: "Bom dia")')
    parser.add_argument("--request-code", default=os.getenv("WA_CODE", "16"), help='Codigo de solicitacao (default: "16")')
    parser.add_argument("--timeout", type=int, default=int(os.getenv("WAIT_TIMEOUT_SECONDS", "7200")), help="Timeout em segundos para aguardar resposta")
    parser.add_argument("--poll", type=float, default=float(os.getenv("WAIT_POLL_SECONDS", "2")), help="Intervalo de polling em segundos")
    parser.add_argument("--output", default=os.getenv("OUTPUT_XLSX", "saida/precos_combustivel.xlsx"), help="Arquivo Excel de saida (.xlsx)")
    parser.add_argument("--user-data-dir", default=os.getenv("WHATSAPP_USER_DATA_DIR", ".wweb_profile"), help="Diretorio para persistir sessao do WhatsApp")
    parser.add_argument("--headless", action="store_true", help="Roda navegador em modo headless (na pratica, pode dificultar login no WhatsApp)")
    parser.add_argument("--login-timeout", type=int, default=int(os.getenv("LOGIN_TIMEOUT_SECONDS", "180")), help="Tempo maximo para concluir login no WhatsApp")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if not args.phone:
        parser.error("Informe --phone (ou variavel WHATSAPP_PHONE).")

    output_path = Path(args.output)
    profile_path = Path(args.user_data_dir)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(profile_path),
            headless=args.headless,
        )
        page = context.pages[0] if context.pages else context.new_page()

        try:
            wait_for_login(page, login_timeout_seconds=args.login_timeout)
            open_chat(page, args.phone)

            baseline = get_last_inbound_message(page)
            baseline_key = baseline.key if baseline else None
            print(f"Baseline de mensagem recebida: {baseline_key or 'sem historico'}")

            send_message(page, args.greeting)
            page.wait_for_timeout(1000)
            send_message(page, args.request_code)
            print(f'Mensagem enviada: "{args.greeting}" e "{args.request_code}"')
            print(f"Aguardando resposta por ate {args.timeout} segundos...")

            inbound = wait_for_new_inbound_message(
                page=page,
                baseline_key=baseline_key,
                timeout_seconds=args.timeout,
                poll_seconds=args.poll,
            )

            print("Mensagem recebida com sucesso:")
            print("-" * 40)
            print(inbound.text)
            print("-" * 40)

            prices = extract_prices(inbound.text)
            s500 = prices.get("S500")
            s10 = prices.get("S10")
            print(f"Extraido -> S500: {s500}, S10: {s10}")

            save_to_excel(
                output_path=output_path,
                phone=args.phone,
                greeting=args.greeting,
                request_code=args.request_code,
                message_text=inbound.text,
                s500=s500,
                s10=s10,
            )
            print(f"Registro salvo em: {output_path.resolve()}")
            return 0
        finally:
            context.close()


if __name__ == "__main__":
    raise SystemExit(main())
