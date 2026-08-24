# Q8 - Qual e o ALCANCE MINIMO OBSERVADO deste ambiente?
#
# ALCANCE, nao FRONTEIRA - e a distincao e o motivo de este cabecalho ter
# sido reescrito. A mensagem mais antiga que o OOM alcanca e onde os dados por
# acaso acabam: pode ser mais nova que o limite da janela (nao existe correio
# mais velho) ou mais velha que ele (itens que chegaram por outro caminho).
#
# A versao anterior deste script chamava isso de "corte de observabilidade" e
# de "abaixo do qual nada e observavel" - exatamente a inferencia que o modelo
# em EnvironmentPolicy.AlcanceMedido proibe. Serve para DESCREVER o que foi
# visto, nunca para concluir "mais antigo que isto esta fora da janela".
$ErrorActionPreference = "Stop"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$PT = "http://schemas.microsoft.com/mapi/proptag/"
$TAG_RECEIVED = "urn:schemas:httpmail:datereceived"

$global:minData = $null
$global:maxData = $null
$global:total = 0
$global:pastas = 0
$global:vazias = 0
$global:erros = 0
$global:comItens = 0

function Varrer($pasta, $prof) {
    if ($prof -gt 8) { return }
    $global:pastas++
    $n = 0
    try { $n = $pasta.Items.Count } catch { $n = -1 }
    if ($n -eq 0) { $global:vazias++ }

    if ($n -gt 0) {
        $global:comItens++
        Write-Host ("  pasta '{0}' classe={1} n={2}" -f $pasta.Name, $pasta.DefaultItemType, $n)
        $tb = $null
        try {
            $tb = $pasta.GetTable()
            $tb.Columns.RemoveAll()
            $tb.Columns.Add($TAG_RECEIVED) | Out-Null
            # Sem Sort: min/max nao dependem de ordem, e Table.Sort recebe
            # OlSortOrder, nao booleano - passar $true da "Value does not fall
            # within the expected range", mensagem que nao nomeia o parametro.
            while (-not $tb.EndOfTable) {
                # GetArray, nao GetRows: GetRows e do ADO. Pelo late binding
                # o erro sai como "does not contain a method named 'GetRows'",
                # que ao menos nomeia o problema - raro neste projeto.
                $arr = $tb.GetArray(500)
                for ($i = $arr.GetLowerBound(0); $i -le $arr.GetUpperBound(0); $i++) {
                    $d = $arr[$i, 0]
                    if ($d -ne $null -and $d -is [DateTime]) {
                        $global:total++
                        if ($global:minData -eq $null -or $d -lt $global:minData) { $global:minData = $d }
                        if ($global:maxData -eq $null -or $d -gt $global:maxData) { $global:maxData = $d }
                    }
                }
            }
        } catch {
            $global:erros += 1
            if ($global:erros -le 5) {
                Write-Host ("  ERRO em '{0}' (n={1}): {2}" -f $pasta.Name, $n, $_.Exception.Message)
            }
        }
    }

    $filhas = $null
    try {
        $filhas = $pasta.Folders
        $c = $filhas.Count
        for ($i = 1; $i -le $c; $i++) {
            $f = $null
            try { $f = $filhas.Item($i); Varrer $f ($prof + 1) }
            finally { if ($f) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) } }
        }
    } catch {
    } finally {
        if ($filhas) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
    }
}

foreach ($st in $ns.Stores) {
    Write-Host "store: $($st.DisplayName)  cached=$($st.IsCachedExchange)"
    $raiz = $null
    try { $raiz = $st.GetRootFolder(); Varrer $raiz 0 }
    catch { Write-Host "  erro: $($_.Exception.Message)" }
    finally { if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) } }
}

Write-Host ""
Write-Host "pastas varridas : $global:pastas"
Write-Host "pastas com 0    : $global:vazias"
Write-Host "pastas com itens: $global:comItens"
Write-Host "erros de tabela : $global:erros"
Write-Host "mensagens datadas: $global:total"
Write-Host "mais ANTIGA     : $global:minData"
Write-Host "mais NOVA       : $global:maxData"
if ($global:minData -and $global:maxData) {
    $dias = [int]($global:maxData - $global:minData).TotalDays
    Write-Host "amplitude       : $dias dias"
    Write-Host "alcance minimo  : $($global:minData.ToString('yyyy-MM-dd'))  (o que se enxergou, NAO a fronteira da janela)"
}
