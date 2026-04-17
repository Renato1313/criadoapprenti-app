# Automacao WhatsApp - Contatos em planilha + IA local opcional

Script em Python com Playwright para:

1. Ler contatos de uma planilha Excel (empresa + telefone)
2. Abrir WhatsApp Web com sessao persistente
3. Enviar saudacao e solicitacao por contato
4. Gerar solicitacao via IA local (Ollama), opcional
5. Aguardar resposta por contato
6. Capturar ultima mensagem recebida com estrategia robusta
7. Extrair S500 e S10 com regex + fallback IA
8. Salvar resultado na mesma planilha

## Requisitos

- Python 3.10+
- Linux/macOS/Windows
- (Opcional para IA) Ollama local em execucao

## Instalacao

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

Se for usar IA local:

```bash
ollama pull qwen2.5:3b
```

## Como funciona a planilha

Arquivo padrao:

```text
saida/precos_combustivel.xlsx
```

Abas usadas:

- `Contatos` (entrada)
- `Registros` (saida)

### Aba Contatos (cabecalho obrigatorio)

| empresa | telefone | contexto | mensagem_solicitacao |
| --- | --- | --- | --- |
| Posto XPTO | 5521999999999 | aceita codigo 16 | 16 |
| Distribuidora ABC | 5511998887777 | prefere texto completo | |

Campos:

- `empresa` (obrigatorio)
- `telefone` (obrigatorio)
- `contexto` (opcional, ajuda a IA a gerar melhor texto)
- `mensagem_solicitacao` (opcional, se preenchido tem prioridade sobre IA)

### Aba Registros (saida automatica)

Colunas salvas:

- `timestamp`
- `empresa`
- `telefone`
- `saudacao_enviada`
- `solicitacao_enviada`
- `mensagem_recebida`
- `S500`
- `S10`
- `metodo_extracao` (`regex` ou `regex+ia`)
- `status`
- `erro`

## Execucao

### Modo padrao (sem IA)

```bash
python whatsapp_price_capture.py \
  --workbook "saida/precos_combustivel.xlsx" \
  --greeting "Bom dia" \
  --default-request "16"
```

### Modo com IA local (Ollama)

```bash
python whatsapp_price_capture.py \
  --workbook "saida/precos_combustivel.xlsx" \
  --greeting "Bom dia" \
  --default-request "16" \
  --ai-enabled \
  --ai-model "qwen2.5:3b"
```

## Parametros uteis

```bash
python whatsapp_price_capture.py \
  --workbook "saida/precos_combustivel.xlsx" \
  --contacts-sheet "Contatos" \
  --records-sheet "Registros" \
  --greeting "Bom dia" \
  --default-request "16" \
  --timeout 7200 \
  --poll 2 \
  --user-data-dir ".wweb_profile" \
  --ai-enabled \
  --ai-model "qwen2.5:3b" \
  --ai-endpoint "http://localhost:11434/api/generate" \
  --ai-timeout 60
```

## Primeiro login no WhatsApp Web

No primeiro uso, o navegador pode abrir pedindo QR code:

1. Escaneie o QR code no celular
2. Aguarde carregar a lista de conversas
3. Nas proximas execucoes, a sessao sera reutilizada via `--user-data-dir`

## Observacoes

- O script cria automaticamente a planilha/abas se nao existirem.
- O WhatsApp Web pode alterar seletores ao longo do tempo; ha fallback de leitura de texto para aumentar resiliencia.
- Cada contato e processado individualmente; se um falhar, os demais continuam.
