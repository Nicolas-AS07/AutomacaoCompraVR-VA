# VR Mensal (.xlsx only) — Geração automatizada de VR por competência

Pipeline em Python que consolida planilhas (`.xlsx`) e gera o arquivo `VR_MENSAL_YYYYMM.csv` conforme regras operacionais (exclusões, férias, admissões no período, desligamentos com “OK até dia 15”, base de dias úteis por UF/sindicato e valores, 80/20). Inclui um frontend Flask simples para gerar e baixar o CSV via navegador.

## ✨ Recursos
- **Somente `.xlsx`** (engine `openpyxl`).
- **Normalização de cabeçalhos** (tolerante a acentos e sinônimos).
- **Exclusões automáticas**: Diretores, Estagiários, Aprendizes, Afastados (interseção no período), Exterior.
- **Regras de cálculo**: férias (dias úteis no período), admissões no meio, desligamentos com “OK até dia 15” (zera) / após 15 (proporcional), dias úteis por UF/sindicato (ou fallback seg–sex), valor diário por UF/sindicato (ou fallback SP/PR/RJ/RS).
- **Saída**: `Matricula, Admissão, Sindicato do Colaborador, Competência, Dias, VALOR DIÁRIO VR, TOTAL, Custo empresa (80%), Desconto profissional (20%), OBS GERAL`.
- **UI** (opcional): botão para gerar e baixar o CSV (Flask).

## 📁 Estrutura
```text
vr_mensal_project/
├─ src/
│  └─ main.py
├─ data/              # coloque aqui os .xlsx
├─ out/               # CSV gerado (VR_MENSAL_YYYYMM.csv)
├─ app.py             # frontend Flask (opcional)
├─ requirements.txt
└─ README.md
```

## 🧰 Requisitos
- Python 3.10+ recomendado

## 🚀 Instalação
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## 🧮 Execução (CLI)
1) Coloque os `.xlsx` em `data/`.  
2) Rode:
```bash
python src/main.py --inicio 2025-04-15 --fim 2025-05-15 --competencia 2025-05
```
Saída: `out/VR_MENSAL_202505.csv`.

## 🖥️ Frontend (Flask)
1) Garanta que `app.py` está na **raiz** (ao lado de `src/`, `data/`, `out/`).  
2) Inicie o servidor:
```bash
python app.py
```
3) Acesse http://localhost:5000 e clique em **Gerar & baixar CSV**.

### Alterar competência pela URL
```
http://localhost:5000/generate?inicio=2025-05-15&fim=2025-06-15&competencia=2025-06
```

### Variáveis de ambiente (opcionais)
- `VR_INICIO`, `VR_FIM`, `VR_COMPETENCIA` — valores padrão do app.  
- `VR_PROJECT_ROOT` — define manualmente a raiz do projeto se `app.py` estiver fora.

## 📊 Planilhas esperadas (nomes flexíveis)
- **ATIVOS.xlsx** (matrícula, admissão, sindicato, cargo opcional p/ “Diretor”).  
- **DESLIGADOS.xlsx** (matrícula, data demissão, `OK_COMUNICADO` e opcional `DATA_COMUNICADO`).  
- **ESTÁGIO.xlsx / ESTAGIARIOS.xlsx**, **APRENDIZ.xlsx**, **EXTERIOR.xlsx** (listas).  
- **FÉRIAS.xlsx** e **AFASTAMENTOS.xlsx** (matrícula, início, fim).  
- **ADMISSÃO ... .xlsx** (admitidos do período).  
- **Base dias uteis.xlsx** (UF/sindicato × mês).  
- **Base sindicato x valor.xlsx** (UF/sindicato × valor diário).

## 🧩 Notas
- **Feriados**: são considerados se a sua `Base dias uteis.xlsx` já trouxer dias úteis líquidos. No fallback (quando a coluna do mês não é encontrada), o cálculo usa seg–sex (sem feriados).  
- Evite deixar o CSV aberto no Excel durante a geração (bloqueio de escrita no Windows).

## 🤝 Contribuições
PRs são bem-vindos (relatório de excluídos, feriados por UF/município, rotinas automáticas de validação).
