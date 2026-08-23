# Mede o custo de acesso por INDICE em offsets crescentes.
#
# SOMENTE LEITURA: nao cria, nao move, nao apaga nada. A pergunta e se
# Items.Item(i) e O(1) no Object Model ou se degrada com a profundidade —
# se degradar, "Carregar mais" fica progressivamente mais lento e a
# paginacao precisa de outra estrategia.

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$pasta = $ns.GetDefaultFolder(6)   # Caixa de Entrada
$itens = $pasta.Items

# Mesma ordenacao que o Iris usa na lista.
$itens.Sort("[ReceivedTime]", $true)

$total = $itens.Count
Write-Output "pasta: $($pasta.Name)  |  itens: $total"
Write-Output ""
Write-Output "offset | ms/item | amostra"
Write-Output "-------|---------|--------"

$tamanhoDaPagina = 50

foreach ($offset in 0, 100, 300, 600, 900) {
  if ($offset + $tamanhoDaPagina -gt $total) { continue }

  # Aquecimento fora da medicao.
  $null = $itens.Item($offset + 1)

  $sw = [Diagnostics.Stopwatch]::StartNew()
  $lidos = 0
  for ($i = $offset + 1; $i -le $offset + $tamanhoDaPagina; $i++) {
    $m = $itens.Item($i)
    # Toca as MESMAS propriedades que o DTO da lista usa.
    $null = $m.Subject
    $null = $m.SenderName
    $null = $m.ReceivedTime
    $null = $m.UnRead
    $lidos++
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m)
  }
  $sw.Stop()

  $porItem = [math]::Round($sw.Elapsed.TotalMilliseconds / $lidos, 2)
  "{0,6} | {1,7} | {2} itens em {3} ms" -f $offset, $porItem, $lidos, [int]$sw.Elapsed.TotalMilliseconds
}

Write-Output ""
Write-Output "Se ms/item ficar estavel entre offsets, o acesso e O(1)."
Write-Output "Se crescer com o offset, a paginacao por indice nao escala."

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
