#!/bin/bash
# Monta o disco de instalação do Mac a partir do .app já construído.
#
#     distribuir/macos/montar-dmg.sh 0.1.0
#
# Sai `dist-instalador/Ferramenta-NPD-Mac-Apple-Silicon.dmg`: a pessoa abre o
# arquivo baixado e arrasta o ícone para a pasta Aplicativos. É a instalação
# que todo Mac usa há vinte anos, e por isso a que menos precisa de explicação.
#
# O que este script NÃO resolve — e não tem como resolver sem uma conta paga de
# desenvolvedor da Apple — é o aviso da primeira abertura. Um programa sem
# assinatura registrada na Apple faz o macOS parar e perguntar. A assinatura
# ad-hoc aqui embaixo evita o erro pior ("o aplicativo está danificado", que o
# macOS dá em Apple Silicon quando não há assinatura nenhuma), mas o aviso
# continua, e é ele que o LEIA-ME dentro do disco ensina a passar.
set -euo pipefail

VERSAO="${1:-0.0.0}"
NOME="Ferramenta NPD"
RAIZ="$(cd "$(dirname "$0")/../.." && pwd)"
APP="$RAIZ/dist-app/$NOME.app"
SAIDA="$RAIZ/dist-instalador"
DMG="$SAIDA/Ferramenta-NPD-Mac-Apple-Silicon.dmg"

[ -d "$APP" ] || { echo "não achei $APP — rode o pyinstaller antes"; exit 1; }

# Assinatura ad-hoc: não identifica ninguém, mas é o que o Apple Silicon exige
# para carregar o binário. `--deep` alcança as bibliotecas que vieram junto.
codesign --force --deep --sign - "$APP"
codesign --verify --deep --strict "$APP" && echo "assinatura ad-hoc conferida"

rm -rf "$SAIDA/monte" "$DMG"
mkdir -p "$SAIDA/monte"
cp -R "$APP" "$SAIDA/monte/"
ln -s /Applications "$SAIDA/monte/Aplicativos"
cp "$RAIZ/distribuir/macos/LEIA-ME-app.txt" "$SAIDA/monte/Leia antes de abrir.txt"

hdiutil create \
  -volname "$NOME $VERSAO" \
  -srcfolder "$SAIDA/monte" \
  -ov -format UDZO \
  "$DMG"

rm -rf "$SAIDA/monte"
echo "pronto: $DMG"
