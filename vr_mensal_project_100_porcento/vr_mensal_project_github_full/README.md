# VR Mensal — 100% com Feriados + Validações

Gera `VR_MENSAL_YYYYMM.csv` consolidando `.xlsx` e considera **feriados** por UF (via `FERIADOS.xlsx`), além de produzir `VALIDACOES_YYYYMM.xlsx` (equivalente à *aba "validações"*).

## Como rodar (CLI)
```bash
pip install -r requirements.txt
python src/main.py --inicio 2025-04-15 --fim 2025-05-15 --competencia 2025-05
```

## Frontend (opcional)
```bash
python app.py
# http://localhost:5000
```

## Entradas em `data/`
- ATIVOS.xlsx, DESLIGADOS.xlsx, ESTÁGIO.xlsx/ESTAGIARIOS.xlsx, APRENDIZ.xlsx, EXTERIOR.xlsx, FÉRIAS.xlsx, AFASTAMENTOS.xlsx, ADMISSÃO ... .xlsx  
- Base dias uteis.xlsx, Base sindicato x valor.xlsx  
- **FERIADOS.xlsx** (novo, opcional porém recomendado) — colunas: `UF` (ou vazio p/ nacional), `DATA`, `DESCRICAO`. Datas em qualquer formato comum.

## Saídas em `out/`
- `VR_MENSAL_YYYYMM.csv`
- `VALIDACOES_YYYYMM.xlsx` — Resumo, Validacoes, Parametros

## Validações incluídas
- `Dias >= 0` e `Dias <= teto do período por UF` (com feriados)  
- `Desligado OK até dia 15` → `Dias = 0`  
- `Valor diário > 0`  
- Consistência: `TOTAL = Dias × Valor`, `80%/20%`
- Conferência de exclusões (Estagiário/Aprendiz/Afastado/Exterior não devem aparecer)
