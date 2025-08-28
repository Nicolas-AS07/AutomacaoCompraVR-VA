# VR Mensal — Web Cloud (Flask + Tailwind)

Frontend bonito, pronto para deploy em nuvem (Render/Railway).  
**Já inclui suas planilhas em `data/`**, mas também aceita upload (ZIP ou múltiplos .xlsx).

## Rodar local
```bash
pip install -r requirements.txt
python app.py
# abra http://localhost:5000
```

## Como usar
- **Modo incluído**: usa os `.xlsx` em `data/` (já copiados neste pacote).
- **Modo upload**: envie um `.zip` com `.xlsx` ou múltiplos `.xlsx`.
- Informe `Início`, `Fim`, `Competência (AAAA-MM)`, clique em **Gerar e baixar**.  
  O app roda `src/main.py` e devolve um ZIP com `VR_MENSAL_YYYYMM.csv`, `VALIDACOES_YYYYMM.xlsx` (quando habilitado) e logs.

## Deploy
### Render
- Web Service → conecte o repositório
- Build: `pip install -r requirements.txt`
- Start: `gunicorn -k gthread -w 2 -b 0.0.0.0:$PORT app:app`

### Railway
- Deploy do GitHub → defina `PORT` se necessário
- Start via `Procfile` já incluso

## Variáveis (opcionais)
- `VR_INICIO` / `VR_FIM` / `VR_COMPETENCIA` para defaults
- `MAX_CONTENT_LENGTH` (bytes) para limite de upload

## Estrutura
- `src/main.py` — gerador (o mesmo do projeto CLI)
- `data/` — **suas planilhas** já incluídas
- `out/` — saída padrão quando usar o modo incluído
- `templates/index.html` — UI (Tailwind)
- `Dockerfile` e `Procfile` — prontos para nuvem
