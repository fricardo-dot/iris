# Q1, parte 6: o filtro DASL de data interpreta a string como LOCAL ou UTC?
#
# SOMENTE LEITURA.
#
# A paginacao por cursor perdeu 200 de 1003 itens, e empate + ausencia de
# ReceivedTime explicam so 6. A perda e sistematica por pagina, o que aponta
# para deslocamento de fuso: se o filtro le a string como UTC e o
# ReceivedTime da tabela vem local, cada fronteira pula uma janela do
# tamanho do offset do fuso.
#
# O teste: pegar um instante conhecido do meio da caixa e contar quantos
# itens o filtro devolve com a data em LOCAL e em UTC. O correto e o que
# bater com a contagem feita na mao.

param([int]$PastaId = 6)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$pasta = $ol.GetNamespace("MAPI").GetDefaultFolder($PastaId)

function Instantes() {
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    [void]$t.Columns.Add("ReceivedTime")
    $t.Sort("ReceivedTime", $true)
    $lista = @()
    while (-not $t.EndOfTable) {
        $a = $t.GetArray(200)
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $lista += [datetime]$a.GetValue($r, 0)
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    return $lista
}

function Contar($filtro) {
    $t = $pasta.GetTable($filtro)
    $n = 0
    while (-not $t.EndOfTable) { $a = $t.GetArray(200); $n += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1) }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    return $n
}

$todos = Instantes
Write-Output "itens na pasta: $($todos.Count)"

# Fronteira: o 50o item, que e onde a pagina 1 terminaria.
$fronteira = $todos[49]
$esperado = ($todos | Where-Object { $_ -lt $fronteira }).Count

Write-Output "fronteira (50o item, hora LOCAL): $($fronteira.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Output "itens estritamente anteriores, contados na mao: $esperado"
Write-Output ""
Write-Output "offset do fuso desta maquina: $([TimeZoneInfo]::Local.GetUtcOffset($fronteira))"
Write-Output ""

$prop = "urn:schemas:httpmail:datereceived"
$comoLocal = $fronteira.ToString("yyyy-MM-dd HH:mm:ss")
$comoUtc   = $fronteira.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")

Write-Output "filtro                                   | devolve | diferenca"
Write-Output "-----------------------------------------|---------|----------"

foreach ($par in @(@("string LOCAL ", $comoLocal), @("string UTC   ", $comoUtc))) {
    $filtro = "@SQL=" + '"' + $prop + '"' + " < '" + $par[1] + "'"
    try {
        $n = Contar $filtro
        "{0} {1,-24} | {2,7} | {3}" -f $par[0], $par[1], $n, ($n - $esperado)
    } catch {
        "{0} {1,-24} | ERRO    | {2}" -f $par[0], $par[1], $_.Exception.Message.Substring(0, 40)
    }
}

Write-Output ""
Write-Output "A linha com diferenca ZERO e a interpretacao correta."

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
