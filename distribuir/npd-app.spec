# Receita do PyInstaller para a versão com janela.
#
#     pyinstaller distribuir/npd-app.spec
#
# Vale para Windows e para macOS, rodada na máquina do sistema alvo — o
# PyInstaller não faz build cruzado. Sai `dist/Ferramenta NPD.exe` no Windows e
# `dist/Ferramenta NPD.app` no Mac.
#
# É um arquivo separado do `npd-tool.spec` de propósito. O executável de
# terminal continua existindo, continua sendo o que a suíte exercita, e não
# deve engordar com pyobjc nem com WebView2 por causa de uma tela que ele não
# tem. Dois produtos, duas receitas.
import re
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

RAIZ = Path(SPECPATH).parent  # noqa: F821 — SPECPATH é injetado pelo PyInstaller

# A versão vem do código, lida como texto: importar `npd_tool` aqui obrigaria o
# pacote a estar instalado no Python que roda o PyInstaller. Escrever o número
# à mão neste arquivo criaria uma segunda fonte para ele, e a segunda é sempre
# a que fica para trás.
VERSAO = re.search(
    r'__version__ = "([^"]+)"',
    (RAIZ / "src" / "npd_tool" / "__init__.py").read_text(encoding="utf-8"),
).group(1)

# O pdfplumber carrega parte do pdfminer por nome, em tempo de execução; sem
# isto o programa monta, roda com xlsx e só quebra no primeiro PDF — que é o
# jeito mais caro de descobrir o problema.
ocultos = collect_submodules("pdfminer") + [
    "pdfplumber",
    "pypdfium2",
    "PIL",
    "PIL.Image",
]

# O `webview` escolhe o motor da janela por importação tardia: WebKit no Mac,
# WebView2 no Windows. O PyInstaller não enxerga nenhum dos dois olhando o
# código, então os dois entram na mão — e o do outro sistema fica de fora, para
# não arrastar meia biblioteca inútil para dentro do pacote.
ocultos += collect_submodules("webview")
if sys.platform == "darwin":
    ocultos += ["objc", "Foundation", "AppKit", "WebKit", "Quartz", "Security"]
elif sys.platform == "win32":
    ocultos += ["clr_loader", "pythonnet"]

# A interface são arquivos, não código: precisam ser copiados para dentro do
# pacote com o mesmo caminho relativo, ou o servidor não acha o index.html.
dados = [(str(RAIZ / "src" / "npd_tool" / "app" / "web"), "npd_tool/app/web")]

analise = Analysis(
    [str(RAIZ / "src" / "npd_tool" / "app" / "__main__.py")],
    pathex=[str(RAIZ / "src")],
    binaries=[],
    datas=dados,
    hiddenimports=ocultos,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "pytest", "IPython", "pandas"],
    noarchive=False,
)

pyz = PYZ(analise.pure)

NOME = "Ferramenta NPD"
icone = str(RAIZ / "distribuir" / "icone" / ("icone.icns" if sys.platform == "darwin" else "icone.ico"))

if sys.platform == "darwin":
    # No Mac o executável entra dentro do .app; ele não é o que a pessoa clica.
    exe = EXE(
        pyz,
        analise.scripts,
        [],
        exclude_binaries=True,
        name=NOME,
        debug=False,
        strip=False,
        upx=False,
        console=False,
        argv_emulation=True,  # arrastar arquivo para o ícone vira argumento
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=icone,
    )
    colecao = COLLECT(
        exe, analise.binaries, analise.datas, strip=False, upx=False, name=NOME
    )
    app = BUNDLE(
        colecao,
        name=f"{NOME}.app",
        icon=icone,
        bundle_identifier="br.com.marchesoni.ferramenta-npd",
        info_plist={
            "CFBundleName": NOME,
            "CFBundleDisplayName": NOME,
            "CFBundleShortVersionString": VERSAO,
            "CFBundleVersion": VERSAO,
            "NSHumanReadableCopyright": "Marchesoni",
            # sem isto a janela abre em resolução dobrada e borrada nas telas
            # Retina — o defeito mais visível que um app de Mac pode ter
            "NSHighResolutionCapable": True,
            # a tela acompanha o tema do sistema (claro e escuro): dizer que
            # não somos compatíveis faria o macOS forçar o modo claro
            "NSRequiresAquaSystemAppearance": False,
            "LSMinimumSystemVersion": "11.0",
            "LSApplicationCategoryType": "public.app-category.business",
        },
    )
else:
    # No Windows sai um .exe único: é o que o instalador copia e o que roda de
    # pen drive, sem pasta de apoio para alguém apagar por engano.
    exe = EXE(
        pyz,
        analise.scripts,
        analise.binaries,
        analise.datas,
        [],
        name=NOME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,  # é um programa de janela; console preta atrás é defeito
        disable_windowed_traceback=False,
        icon=icone,
    )
