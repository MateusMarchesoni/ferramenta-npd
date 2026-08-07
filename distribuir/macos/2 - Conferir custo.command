#!/bin/bash
# Um .command aberto por clique duplo começa na pasta pessoal do usuário, não
# na pasta dele mesmo — por isso o cd. Sem ele, a ferramenta procuraria a
# planilha no lugar errado.
cd "$(dirname "$0")" || exit 1
echo
"$(dirname "$0")/programa/npd-tool" conferir
echo
echo "============================================================"
echo " O custo também aparece na coluna M da aba Candidatos. Se estiver certo, clique em '3 - Gravar no funil'."
echo "============================================================"
echo
read -n 1 -s -r -p "Pressione qualquer tecla para fechar."
echo
