#!/bin/bash
# Um .command aberto por clique duplo começa na pasta pessoal do usuário, não
# na pasta dele mesmo — por isso o cd. Sem ele, a ferramenta procuraria a
# planilha no lugar errado.
cd "$(dirname "$0")" || exit 1
echo
"$(dirname "$0")/programa/npd-tool" abrir
echo
echo "============================================================"
echo " Agora abra a planilha, vá na aba Candidatos, marque com x os produtos e preencha NCM e Marca. Depois salve, feche, e clique em '2 - Conferir custo'."
echo "============================================================"
echo
read -n 1 -s -r -p "Pressione qualquer tecla para fechar."
echo
