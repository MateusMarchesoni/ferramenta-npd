"""Gera o ícone do programa nos formatos que cada sistema exige.

    python distribuir/icone/gerar-icones.py

Sai daqui `icone.png` (1024, para referência), `icone.icns` (macOS) e
`icone.ico` (Windows). Os três ficam versionados: o build não precisa de
Pillow nem de máquina Mac para ter ícone, e o ícone não muda sozinho entre
uma entrega e outra.

O desenho é o mesmo do `app/web/icone.svg`, redesenhado em pixels porque não
há rasterizador de SVG entre as dependências — e acrescentar um só para fazer
um ícone que muda uma vez por ano seria caro pelo motivo errado.

Por que superelipse e não `rounded_rectangle`: o canto do ícone da Apple não é
um arco de círculo. É uma curva contínua, em que a curvatura cresce aos poucos
em vez de começar de uma vez. É a diferença entre um ícone que parece do
sistema e um que parece colado nele — e num ícone de 1024px ela é visível.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw

AQUI = Path(__file__).resolve().parent

LADO = 1024
SUPERAMOSTRA = 4  # desenha 4× maior e reduz: é o antisserrilhado dos pobres
EXPOENTE = 5.0    # 4 é quadrado demais, 6 é redondo demais; 5 é o da Apple

# a área útil do ícone do macOS: a arte não encosta na borda da tela
MARGEM = 96

TOPO = (0x4D, 0xA2, 0xFF)
MEIO = (0x0A, 0x84, 0xFF)
BASE = (0x3B, 0x34, 0xC9)


def _interpolar(a, b, t):
    return tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def _gradiente(largura: int, altura: int) -> Image.Image:
    """Azul do sistema no topo, índigo embaixo — a luz vem de cima, sempre."""
    imagem = Image.new("RGB", (1, altura))
    pixels = imagem.load()
    for y in range(altura):
        t = y / (altura - 1)
        if t < 0.48:
            cor = _interpolar(TOPO, MEIO, t / 0.48)
        else:
            cor = _interpolar(MEIO, BASE, (t - 0.48) / 0.52)
        pixels[0, y] = cor
    return imagem.resize((largura, altura))


def _mascara_superelipse(lado: int, margem: int) -> Image.Image:
    """|x|^n + |y|^n = 1 — a curva contínua, ponto a ponto."""
    mascara = Image.new("L", (lado, lado), 0)
    desenho = ImageDraw.Draw(mascara)
    raio = (lado - 2 * margem) / 2
    centro = lado / 2
    passos = 720
    pontos = []
    for indice in range(passos):
        angulo = 2 * 3.141592653589793 * indice / passos
        import math

        cosseno, seno = math.cos(angulo), math.sin(angulo)
        x = raio * (abs(cosseno) ** (2 / EXPOENTE)) * (1 if cosseno >= 0 else -1)
        y = raio * (abs(seno) ** (2 / EXPOENTE)) * (1 if seno >= 0 else -1)
        pontos.append((centro + x, centro + y))
    desenho.polygon(pontos, fill=255)
    return mascara


def _brilho(lado: int) -> Image.Image:
    """A luz de cima: branco a 34% que some antes da metade."""
    camada = Image.new("L", (1, lado))
    pixels = camada.load()
    for y in range(lado):
        t = y / (lado - 1)
        pixels[0, y] = round(max(0.0, 0.34 * (1 - t / 0.55)) * 255)
    return camada.resize((lado, lado))


def _funil(desenho: ImageDraw.ImageDraw, escala: float) -> None:
    """O símbolo: cotação entra em cima, produto cadastrado sai embaixo."""
    pontos = [(296, 300), (728, 300), (566, 528), (566, 708), (458, 760), (458, 528)]
    # o símbolo ocupa ~80% da área útil: cheio o bastante para ser reconhecido
    # em 16px, folgado o bastante para não parecer que estourou a moldura
    fator, centro = 0.82, 512.0
    pontos = [(centro + (x - centro) * fator, centro + (y - centro) * fator) for x, y in pontos]
    caminho = [(x * escala, y * escala) for x, y in pontos]
    desenho.polygon(caminho, fill=(255, 255, 255, 255))
    # o mesmo contorno em traço grosso arredonda as pontas do polígono
    desenho.line(caminho + [caminho[0]], fill=(255, 255, 255, 255),
                 width=round(30 * escala), joint="curve")


def desenhar(lado: int = LADO) -> Image.Image:
    grande = lado * SUPERAMOSTRA
    escala = grande / 1024

    arte = _gradiente(grande, grande).convert("RGBA")
    branco = Image.new("RGBA", (grande, grande), (255, 255, 255, 255))
    arte = Image.composite(branco, arte, _brilho(grande)).convert("RGBA")

    arte.putalpha(_mascara_superelipse(grande, round(MARGEM * escala)))

    simbolo = Image.new("RGBA", (grande, grande), (0, 0, 0, 0))
    _funil(ImageDraw.Draw(simbolo), escala)
    arte = Image.alpha_composite(arte, simbolo)

    return arte.resize((lado, lado), Image.LANCZOS)


def main() -> int:
    mestre = desenhar()
    mestre.save(AQUI / "icone.png")

    # Windows: um único .ico com todos os tamanhos que o Explorador pede
    mestre.save(
        AQUI / "icone.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )

    # macOS: o .icns é montado pelo `iconutil` a partir de um .iconset
    conjunto = AQUI / "icone.iconset"
    conjunto.mkdir(exist_ok=True)
    for tamanho in (16, 32, 128, 256, 512):
        mestre.resize((tamanho, tamanho), Image.LANCZOS).save(
            conjunto / f"icon_{tamanho}x{tamanho}.png"
        )
        mestre.resize((tamanho * 2, tamanho * 2), Image.LANCZOS).save(
            conjunto / f"icon_{tamanho}x{tamanho}@2x.png"
        )

    if sys.platform == "darwin":
        subprocess.run(
            ["iconutil", "-c", "icns", str(conjunto), "-o", str(AQUI / "icone.icns")],
            check=True,
        )
        print("icone.icns gerado")
    else:
        print("iconutil só existe no macOS — o .icns versionado continua valendo")

    print(f"ícones em {AQUI}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
