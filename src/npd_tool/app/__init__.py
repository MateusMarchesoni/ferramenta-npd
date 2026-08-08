"""A versão de janela da ferramenta — a mesma máquina, com tela em vez de terminal.

O CLI continua sendo o caminho canônico e o que a suíte testa. Este pacote é
uma **casca**: ele não sabe escolher preço, calcular m³ nem montar linha de
Funil. Tudo o que ele faz é perguntar ao `ui/candidatos.py` e ao
`escrita/ooxml.py` exatamente o que o `cli.py` já pergunta, na mesma ordem,
e desenhar o resultado.

Isso é a regra 4 do CLAUDE.md aplicada à tela: se um dia a conta de custo
mudar, ela muda num lugar só, e a janela acompanha sem saber que mudou.

    app/nucleo.py     a sessão: ler cotações, conferir, gravar
    app/servidor.py   um HTTP local que fala JSON com a tela
    app/janela.py     abre a janela nativa (ou o navegador, se não der)
    app/dialogos.py   os seletores de arquivo do próprio sistema
    app/config.py     lembra a última planilha e a última pasta
    app/web/          a interface (HTML/CSS/JS)
"""
