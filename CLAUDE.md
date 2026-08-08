# Instruções permanentes — Ferramenta NPD

Leia `PLANO.md` antes de qualquer trabalho neste repositório. Ele é o contexto completo do projeto: contexto de negócio, anatomia da planilha, anatomia das cotações, motor de custo, formato de escrita, ordem de construção, casos de teste e perguntas em aberto.

## Regras inegociáveis

1. **Nunca escrever no arquivo NPD original.** O original é `../Cotações Marchesoni/NPD_2026_04_08_26.xlsx` (fora deste repo). Todo teste e toda execução da ferramenta operam sobre cópias em `tests/fixtures/` ou `saida/`. Se algum código abrir o caminho do arquivo original em modo de escrita, pare.
2. **Construir na ordem da seção 11 do PLANO.md.** Cada etapa tem critério de aceite verificável. Não avançar para a próxima etapa sem passar na anterior — em especial a Etapa 1 (POC de escrita com foto embutida), que é bloqueante.
3. **A ferramenta nunca inventa dado.** Campo não extraído com confiança fica `None` e entra no relatório como pendente. Nunca zero, nunca string vazia, nunca chute.
4. **Regras de dependência entre módulos:** `ingest` não conhece `custo`. `custo` não conhece `escrita`. `ui` não contém regra de negócio. Só `escrita/ooxml.py` toca o arquivo da planilha. **`app/` é casca**: ele chama as mesmas funções que o `cli.py` chama e não decide nada — se uma regra de negócio aparecer dentro de `app/`, ela está no lugar errado. A tela (`app/web/`) nunca usa `innerHTML` com texto vindo de cotação, e nunca converte preço ou custo para número — eles atravessam o JSON como texto, porque `float` não representa 1,15.
5. **Perguntas em aberto (seção 13 do PLANO.md) não são técnicas** — travam partes do motor de custo (regime tributário, crédito de ICMS, custo do contêiner, origem do NCM). Não assumir respostas; parametrizar e marcar como pendente até confirmação do responsável pelo negócio.

## Ambiente

- Python do projeto: `npd-tool/.venv` (Python 3.9 do sistema — usar `from __future__ import annotations` em módulos com sintaxe `X | None` de tipos, pois 3.9 não resolve `|` em tipos em runtime).
- Ativar com `source .venv/bin/activate` a partir de `npd-tool/`.
- A janela do app precisa do extra: `pip install -e ".[app,dev]"`. Sem o `pywebview`, o app continua funcionando — a tela abre no navegador padrão.
