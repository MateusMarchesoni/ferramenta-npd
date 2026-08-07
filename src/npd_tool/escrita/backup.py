"""Cópia de segurança da planilha — o passo 1 da seção 7.2, antes de tudo.

A planilha NPD é o trabalho de meses de várias pessoas e não tem histórico de
versão. O risco listado como *crítico* na seção 14 do PLANO.md é corrompê-la.
Backup não é zelo: é a única coisa que separa um bug de uma perda definitiva.

O backup é feito por cópia binária do arquivo, sem abrir o zip — o que evita
que um erro de leitura do próprio pacote se propague para a cópia.
"""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

FORMATO_TIMESTAMP = "%Y%m%d-%H%M%S"
SUFIXO = "--backup-"


def caminho_do_backup(origem: Path, pasta: Path | None = None, momento: datetime | None = None) -> Path:
    origem = Path(origem)
    momento = momento or datetime.now()
    pasta = Path(pasta) if pasta else origem.parent / "backups"
    nome = f"{origem.stem}{SUFIXO}{momento.strftime(FORMATO_TIMESTAMP)}{origem.suffix}"
    return pasta / nome


LIMITE_DE_COLISOES = 100


def fazer_backup(origem: Path, pasta: Path | None = None, momento: datetime | None = None) -> Path:
    """Copia a planilha para `backups/` com timestamp e devolve o caminho.

    Dois comandos no mesmo segundo caem no mesmo nome. Nesse caso o backup ganha
    um sufixo em vez de sobrescrever o anterior ou abortar a operação: perder um
    backup é inaceitável, e recusar a gravação por causa do nome de um arquivo
    de segurança é desproporcional — o `abrir` seguido de `conferir` acontece em
    menos de um segundo com facilidade.
    """
    origem = Path(origem)
    if not origem.is_file():
        raise FileNotFoundError(f"planilha não encontrada para backup: {origem}")

    destino = caminho_do_backup(origem, pasta, momento)
    destino.parent.mkdir(parents=True, exist_ok=True)

    if destino.exists():
        for tentativa in range(2, LIMITE_DE_COLISOES + 1):
            alternativo = destino.with_name(
                f"{destino.stem}-{tentativa}{destino.suffix}"
            )
            if not alternativo.exists():
                destino = alternativo
                break
        else:
            raise FileExistsError(
                f"mais de {LIMITE_DE_COLISOES} backups no mesmo segundo em "
                f"{destino.parent} — alguma coisa está em laço"
            )

    shutil.copy2(origem, destino)
    return destino


def backups_existentes(origem: Path, pasta: Path | None = None) -> list[Path]:
    origem = Path(origem)
    pasta = Path(pasta) if pasta else origem.parent / "backups"
    if not pasta.is_dir():
        return []
    return sorted(pasta.glob(f"{origem.stem}{SUFIXO}*{origem.suffix}"))
