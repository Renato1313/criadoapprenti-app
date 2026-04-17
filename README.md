# Automacao WhatsApp - Captura de Precos

Script em Python com Playwright para:

1. Abrir o WhatsApp Web com sessao persistente
2. Enviar mensagens para um numero (ex.: `Bom dia` e `16`)
3. Aguardar resposta por ate 2 horas
4. Capturar de forma robusta a ultima mensagem recebida
5. Extrair precos de `S500` e `S10`
6. Salvar resultado em Excel (`.xlsx`)

## Requisitos

- Python 3.10+
- Linux/macOS/Windows

## Instalacao

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Execucao

### Exemplo basico

```bash
python whatsapp_price_capture.py --phone 5511999999999
```

Por padrao, o script envia:

- saudacao: `Bom dia`
- codigo: `16`
- timeout de espera: `7200` segundos (2h)
- arquivo de saida: `saida/precos_combustivel.xlsx`

### Parametros uteis

```bash
python whatsapp_price_capture.py \
  --phone 5511999999999 \
  --greeting "Bom dia" \
  --request-code "16" \
  --timeout 7200 \
  --poll 2 \
  --output "saida/precos_combustivel.xlsx" \
  --user-data-dir ".wweb_profile"
```

## Primeiro login no WhatsApp Web

No primeiro uso, o navegador pode abrir pedindo QR code.

1. Escaneie o QR code no celular
2. Aguarde o carregamento da lista de conversas
3. Nas proximas execucoes, a sessao sera reutilizada via `--user-data-dir`

## Formatos de resposta aceitos

Exemplos capturados:

```text
S500: 6,45004
S10: 6,75516
```

ou

```text
S500 5,99
```

O parser tenta extrair `S500` e `S10` mesmo com variacoes de espaco, dois pontos e quebra de linha.

## Saida no Excel

Colunas geradas:

- `timestamp`
- `telefone`
- `saudacao`
- `codigo_solicitado`
- `mensagem_recebida`
- `S500`
- `S10`

Arquivo padrao:

```text
saida/precos_combustivel.xlsx
```

## Observacoes

- O WhatsApp Web pode alterar seletores ao longo do tempo; o script usa estrategia com fallback para melhorar resiliencia.
- Se nao encontrar novos dados dentro do timeout, o script encerra com erro de timeout.
