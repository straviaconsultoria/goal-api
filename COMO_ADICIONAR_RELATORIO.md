# Como adicionar um relatório em 2 arquivos

## 1. Crie/copiei a pasta

`tenants/<tenant>/reports/<novo_relatorio>/`

## 2. Coloque apenas

- `config.yaml`
- `templates/v1.xlsx`

## 3. Pronto

A rota será:
`POST /reports/<tenant>/<novo_relatorio>`

Não é necessário criar endpoint, editar `main.py`, criar container ou registrar o relatório.

## Versionar

Novo layout: adicione `v2.xlsx` e troque `template: v1.xlsx` por `template: v2.xlsx`.
