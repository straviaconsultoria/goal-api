# Excel Report API

Serviço único para gerar dezenas de relatórios XLSX para múltiplos tenants usando FastAPI + OpenPyXL.

## Regra principal

Para adicionar um relatório comum, NÃO altere Python, endpoint, Dockerfile ou serviço.

Copie:
`tenants/example/reports/operacional/`

para:
`tenants/<tenant>/reports/<relatorio>/`

Depois altere somente:
- `config.yaml`
- `templates/v1.xlsx`

A rota surge automaticamente:
`POST /reports/<tenant>/<relatorio>`

## Nova versão do layout

Adicione `templates/v2.xlsx` e altere no `config.yaml`:
`template: v2.xlsx`

## Endpoints

- `GET /health`
- `GET /catalog` (X-API-Key)
- `POST /reports/{tenant}/{report}` (X-API-Key)
- `/docs` para Swagger

## Payload genérico

```json
{
  "variables": {
    "titulo": "Relatório Linha 301",
    "data": "20/08/2026"
  },
  "sections": {
    "viagens": [
      {"horario": "05:30", "origem": "Garagem", "destino": "Centro", "bloco": "1"},
      {"horario": "06:10", "origem": "Centro", "destino": "Terminal", "bloco": "1"}
    ]
  }
}
```

## Como o template funciona

Campos simples: escreva no Excel `{{titulo}}`, `{{data}}`, etc.

Seção dinâmica: coloque no Excel o marcador configurado, por exemplo `{{VIAGENS}}`. A linha-modelo indicada no `config.yaml` deve ter a formatação visual desejada. O motor replica essa linha conforme a quantidade de registros e o conteúdo abaixo é deslocado automaticamente.

## Coolify

1. Suba este repositório no GitHub.
2. Crie uma Application no Coolify usando Dockerfile.
3. Configure a variável `API_KEY`.
4. Exponha a porta 8000.
5. Faça deploy.
6. Teste `/health` e `/docs`.
