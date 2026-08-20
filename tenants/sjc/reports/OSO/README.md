# Exemplo de relatório

Para criar um novo relatório, copie esta pasta inteira e altere apenas:
1. `config.yaml`
2. `templates/v1.xlsx`

Rota criada automaticamente:
`POST /reports/{tenant}/{nome-da-pasta-do-relatorio}`

Payload esperado neste exemplo:
```json
{
  "variables": {"titulo": "Linha 301", "data": "20/08/2026"},
  "sections": {
    "viagens": [
      {"horario": "05:30", "origem": "Garagem", "destino": "Centro", "bloco": "1"}
    ]
  }
}
```
