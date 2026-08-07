# Receita do PyInstaller. Vale para Windows e para macOS — o mesmo arquivo,
# rodado na máquina do sistema alvo, porque o PyInstaller não faz build cruzado.
#
#     pyinstaller distribuir/npd-tool.spec
#
# Sai um executável único em `dist/`, com o Python e as bibliotecas dentro.
# Quem receber não instala nada.
from PyInstaller.utils.hooks import collect_submodules

# O pdfplumber carrega parte do pdfminer por nome, em tempo de execução; sem
# isto o executável monta, roda com xlsx e só quebra no primeiro PDF — que é o
# jeito mais caro de descobrir o problema.
ocultos = collect_submodules("pdfminer") + [
    "pdfplumber",
    "pypdfium2",
    "PIL",
]

analise = Analysis(
    ["../src/npd_tool/__main__.py"],
    pathex=["../src"],
    binaries=[],
    datas=[],
    hiddenimports=ocultos,
    hookspath=[],
    runtime_hooks=[],
    # nada de tela: a interface é a planilha (PLANO.md 13.5)
    excludes=["tkinter", "matplotlib", "numpy", "pytest", "IPython"],
    noarchive=False,
)

pyz = PYZ(analise.pure)

exe = EXE(
    pyz,
    analise.scripts,
    analise.binaries,
    analise.datas,
    [],
    name="npd-tool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # a saída é o relatório do que aconteceu; esconder não ajuda
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
