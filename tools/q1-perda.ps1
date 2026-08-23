# Q1, parte 5: POR QUE a paginacao por cursor perdeu 20% dos itens.
#
# SOMENTE LEITURA.
#
# A paginacao por ReceivedTime devolveu 803 de 1003. Rapido e errado e pior
# que lento e certo, e descobrir a causa importa mais que o ganho de 20x.
#
# Duas hipoteses:
#   H1 — EMPATE. O filtro usa "< fronteira" estrito. Itens com ReceivedTime
#        IGUAL ao ultimo da pagina anterior sao pulados.
#   H2 — AUSENCIA. Itens sem ReceivedTime nao satisfazem NENHUM filtro de
#        comparacao, e somem da paginacao inteira.

param([int]$PastaId = 6)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$pasta = $ol.GetNamespace("MAPI").GetDefaultFolder($PastaId)

$t = $pasta.GetTable()
$t.Columns.RemoveAll()
[void]$t.Columns.Add("EntryID")
[void]$t.Columns.Add("ReceivedTime")
[void]$t.Columns.Add("MessageClass")
$t.Sort("ReceivedTime", $true)

$todos = @()
while (-not $t.EndOfTable) {
    $a = $t.GetArray(200)
    for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
        $todos += [pscustomobject]@{
            Id     = "$($a.GetValue($r,0))"
            Quando = $a.GetValue($r,1)
            Classe = "$($a.GetValue($r,2))"
        }
    }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)

Write-Output "total na tabela: $($todos.Count)"
Write-Output ""

# H2 — quantos nao tem ReceivedTime?
$semData = $todos | Where-Object { $null -eq $_.Quando }
Write-Output "H2 — sem ReceivedTime: $($semData.Count)"
if ($semData.Count -gt 0) {
    $semData | Group-Object Classe | ForEach-Object { "     {0,4}x {1}" -f $_.Count, $_.Name }
}
Write-Output ""

# H1 — quantos empatam no segundo?
$comData = $todos | Where-Object { $null -ne $_.Quando }
$grupos = $comData | Group-Object { ([datetime]$_.Quando).ToString("yyyy-MM-dd HH:mm:ss") }
$empatados = $grupos | Where-Object { $_.Count -gt 1 }
$itensEmpatados = ($empatados | Measure-Object Count -Sum).Sum
$excedente = $itensEmpatados - $empatados.Count

Write-Output "H1 — grupos com o MESMO segundo: $($empatados.Count)"
Write-Output "     itens nesses grupos: $itensEmpatados"
Write-Output "     itens que um filtro estrito PULARIA: $excedente"
Write-Output ""
$empatados | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
    "     {0,3} itens em {1}" -f $_.Count, $_.Name
}

Write-Output ""
Write-Output "CONCLUSAO"
$perdaPrevista = $semData.Count + $excedente
Write-Output "  perda prevista pelas duas hipoteses: $perdaPrevista"
$observada = $todos.Count - 803
Write-Output "  perda observada na paginacao:        $observada"

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
