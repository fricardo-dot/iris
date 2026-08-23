# -*- coding: utf-8 -*-
"""Garante BOM UTF-8 em todo tools/*.ps1.

O PowerShell 5.1 le .ps1 sem BOM como ANSI. Um caractere multibyte (um
travessao, um acento) vira lixo, e se o lixo contiver aspas o parser
quebra em cascata com erros que apontam para linhas erradas.

Ja aconteceu tres vezes neste projeto. Rode isto depois de gerar ou editar
qualquer script com ferramenta que grave sem BOM.
"""
import glob, io, os, sys

BOM = u'﻿'
mudou = []
for p in glob.glob(os.path.join(os.path.dirname(__file__), '*.ps1')):
    with io.open(p, 'rb') as f:
        bruto = f.read()
    if bruto.startswith(b'\xef\xbb\xbf'):
        continue
    texto = bruto.decode('utf-8')
    with io.open(p, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(texto)
    mudou.append(os.path.basename(p))

print("BOM adicionado em %d arquivo(s): %s" % (len(mudou), ", ".join(mudou) or "-"))
