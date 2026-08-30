# -*- coding: utf-8 -*-
"""
O ICONE DO IRIS, DESENHADO POR CODIGO.

===========================================================================
POR QUE ISTO E UM SCRIPT E NAO UM .ICO SOLTO NO REPOSITORIO

Um binario versionado que ninguem sabe reproduzir e uma divida: no dia em
que a paleta mudar, ou em que alguem quiser outro tamanho, o arquivo vira
intocavel porque nao se sabe como ele foi feito. Aqui o desenho E o codigo,
e as cores vem da mesma paleta do Themes/Tokens.xaml.

===========================================================================
O DESENHO

"Iris" e a parte colorida do olho, e tambem uma flor. O olho ganha, porque
um anel concentrico e a unica coisa que ainda se le a 16 pixels -- que e o
tamanho que aparece na barra de tarefas, e o unico tamanho que importa de
verdade.

  - quadrado arredondado, no azul-escuro da superficie de acento
  - anel do acento (a iris)
  - pupila escura no meio
  - um brilho pequeno em cima a esquerda, que da profundidade sem virar ruido

Cada tamanho e desenhado em 8x e reduzido -- suavizacao de verdade, em vez
de reduzir um 256 e virar borrao no 16.

Uso:  python tools/gerar-icone.py
Saida: src/Iris.App/iris.ico
"""
import os

from PIL import Image, ImageDraw

# As mesmas cores de src/Iris.App/Themes/Tokens.xaml.
FUNDO = (41, 48, 74, 255)      # Color.Accent.Surface  #29304A
ACENTO = (130, 148, 255, 255)  # Color.Accent          #8294FF
PUPILA = (23, 26, 33, 255)     # Color.Surface.1       #171A21
BRILHO = (230, 234, 240, 255)  # Color.Text.Primary    #E6EAF0

TAMANHOS = (16, 24, 32, 48, 64, 128, 256)
ESCALA = 8


def desenhar(lado):
    n = lado * ESCALA
    img = Image.new("RGBA", (n, n), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Quadrado arredondado. O raio e proporcional para o canto ficar igual
    # em todos os tamanhos.
    d.rounded_rectangle([0, 0, n - 1, n - 1], radius=n * 0.22, fill=FUNDO)

    centro = n / 2.0

    # A IRIS: anel do acento. Largura generosa -- anel fino some no 16.
    raio_ext = n * 0.34
    d.ellipse([centro - raio_ext, centro - raio_ext,
               centro + raio_ext, centro + raio_ext], fill=ACENTO)

    # A PUPILA.
    raio_pup = n * 0.145
    d.ellipse([centro - raio_pup, centro - raio_pup,
               centro + raio_pup, centro + raio_pup], fill=PUPILA)

    # O BRILHO, em cima a esquerda. Pequeno de proposito: a 16 pixels ele
    # vira meio pixel claro, que e exatamente a impressao de volume que se
    # quer -- e nao um ponto branco competindo com a pupila.
    raio_bri = n * 0.075
    bx = centro - raio_ext * 0.42
    by = centro - raio_ext * 0.42
    d.ellipse([bx - raio_bri, by - raio_bri, bx + raio_bri, by + raio_bri],
              fill=BRILHO)

    return img.resize((lado, lado), Image.LANCZOS)


def main():
    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    destino = os.path.join(raiz, "src", "Iris.App", "iris.ico")

    imagens = [desenhar(t) for t in TAMANHOS]

    # O primeiro e o maior; o Pillow guarda os demais como quadros do .ico.
    imagens[-1].save(destino, format="ICO",
                     sizes=[(t, t) for t in TAMANHOS])

    print("gerado %s" % destino)
    print("tamanhos: %s" % ", ".join(str(t) for t in TAMANHOS))
    print("%d bytes" % os.path.getsize(destino))


if __name__ == "__main__":
    main()
