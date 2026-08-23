# Mede o custo REAL de uma pagina, do jeito que o broker faz.
#
# SOMENTE LEITURA: nao cria, nao move, nao apaga nada.
#
# A primeira versao desta medicao ordenava UMA vez e cronometrava so
# Items.Item(i). Estava errada como modelo do broker: ReadPage refaz
# folder.Items, Sort e Count a CADA pagina, inclusive no "Carregar mais".
# Ou seja, a versao anterior media a parte barata e deixava de fora
# justamente a que pode degradar com o tamanho da pasta.
#
# Aqui as tres fases sao cronometradas em separado, e os offsets vao em
# ordem ALEATORIA — em ordem crescente, profundidade e aquecimento de cache
# ficam confundidos um com o outro.
#
# DIFERENCAS que sobram em relacao ao ReadPage, e que este script NAO
# reproduz:
#   - o broker filtra o que nao e MailItem; aqui tudo e tocado igual;
#   - o Summarize tolera excecao por propriedade, com Texto()/Numero();
#     aqui uma propriedade que estoure derruba a medicao;
#   - o broker passa pela fila da STA e pelo message filter.
# Ou seja: serve para COMPARAR fases e offsets, nao como latencia ponta a
# ponta da aplicacao.

param(
    [int]$PastaId = 6,          # 6 = Caixa de Entrada
    [int]$TamanhoDaPagina = 50,
    [int]$Execucoes = 3
)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$pasta = $ns.GetDefaultFolder($PastaId)

# Nao encadear: $pasta.Items.Count cria um RCW intermediario que ninguem
# libera. Mesma regra do codigo de producao (R7).
$itensDaPasta = $pasta.Items
$total = $itensDaPasta.Count
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($itensDaPasta)

Write-Output "pasta: $($pasta.Name)  |  itens: $total  |  pagina: $TamanhoDaPagina"
Write-Output ""
Write-Output "exec | offset | Items+Sort | Count | leitura | ms/item | total"
Write-Output "-----|--------|-----------|-------|---------|---------|------"

$offsets = @(0, 100, 300, 600, 900) | Where-Object { $_ + $TamanhoDaPagina -le $total }

for ($e = 1; $e -le $Execucoes; $e++) {
    # Ordem aleatoria a cada execucao.
    foreach ($offset in ($offsets | Sort-Object { Get-Random })) {

        # --- Fase 1: obter a colecao e ordenar, como ReadPage faz ---
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $items = $pasta.Items
        $items.Sort("[ReceivedTime]", $true)
        $msOrdenar = $sw.Elapsed.TotalMilliseconds

        # --- Fase 2: contar ---
        $sw.Restart()
        $null = $items.Count
        $msContar = $sw.Elapsed.TotalMilliseconds

        # --- Fase 3: ler a pagina, com TODAS as propriedades do DTO ---
        $sw.Restart()
        $lidos = 0
        for ($i = $offset + 1; $i -le $offset + $TamanhoDaPagina; $i++) {
            $m = $items.Item($i)
            $null = $m.EntryID
            $null = $m.Subject
            $null = $m.SenderName
            $null = $m.ReceivedTime
            $null = $m.Size
            $null = $m.UnRead
            $null = $m.Permission
            $null = $m.MessageClass
            $anexos = $m.Attachments
            $null = $anexos.Count
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($anexos)
            $lidos++
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m)
        }
        $msLer = $sw.Elapsed.TotalMilliseconds

        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($items)

        $porItem = [math]::Round($msLer / $lidos, 2)
        $soma = [int]($msOrdenar + $msContar + $msLer)
        "{0,4} | {1,6} | {2,9} | {3,5} | {4,7} | {5,7} | {6} ms" -f `
            $e, $offset, [int]$msOrdenar, [int]$msContar, [int]$msLer, $porItem, $soma
    }
}

Write-Output ""
Write-Output "Items+Sort e Count rodam a CADA pagina. Se eles crescerem com o"
Write-Output "tamanho da pasta, o custo fixo por pagina e que limita, nao o"
Write-Output "acesso por indice."

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
