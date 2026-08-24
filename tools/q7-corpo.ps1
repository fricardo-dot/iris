# Q7: quanto custa ler o CORPO, e o que isso significa para a busca.
#
# SOMENTE LEITURA, e com uma regra a mais: NADA DE CONTEUDO SAI DAQUI.
#
# Nenhum corpo, trecho de corpo, assunto ou endereco e impresso ou gravado.
# So COMPRIMENTO, TEMPO e CONTAGEM. Isso nao e escrupulo decorativo: o
# proprio plano proibe persistir corpo antes de criptografia e retencao
# estarem decididas, porque seria criar o risco R2-I durante o experimento
# que deveria informa-lo.
#
# ------------------------------------------------------------------
# O QUE A Q7 VIROU, depois da Q2
#
# A pergunta original era custo unitario: "quanto custa indexar um corpo?".
# Depois da Q2 ela ganhou uma segunda metade, que e a que decide o schema:
#
#   A que entidade o indice se liga — encarnacao, item logico, ou
#   associacao item-pasta?
#
# Se ligar a ENCARNACAO, um Move cria encarnacao nova (medido na §11.1) e o
# corpo inteiro precisa ser reindexado. Se ligar ao ITEM LOGICO, o Move nao
# custa nada. A diferenca e de ordens de grandeza, e depende de uma
# decisao de schema — nao de otimizacao depois.
#
# ------------------------------------------------------------------
# CUIDADO COM DOWNLOAD
#
# Ler .Body de um item que esta no cache so como cabecalho DISPARA o
# download. Este script confere DownloadState antes e PULA o que nao
# estiver completo — medir e uma coisa, provocar trafego e outra.

$ErrorActionPreference = "Stop"
$AMOSTRA = 200

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
function Solta($o) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }

$in = $ns.GetDefaultFolder(6)

# ---- referencia: quanto custa SO o metadado, pela Table ----
$sw = [Diagnostics.Stopwatch]::StartNew()
$t = $in.GetTable()
$cols = $t.Columns
$cols.RemoveAll()
foreach ($c in @("Subject","SenderName","ReceivedTime","Size","UnRead","MessageClass")) {
    $x = $cols.Add($c); Solta $x
}
Solta $cols
$nMeta = 0
while (-not $t.EndOfTable) {
    $a = $t.GetArray(200)
    $nMeta += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1)
}
Solta $t
$sw.Stop()
$msMeta = $sw.Elapsed.TotalMilliseconds
Write-Host ("METADADO pela Table: {0} itens em {1:N0} ms  ({2:N2} ms/item)" -f `
    $nMeta, $msMeta, ($msMeta / [Math]::Max($nMeta,1)))
Write-Host ""

# ---- amostra para o corpo ----
$itens = $in.Items
$itens.Sort("[ReceivedTime]", $true)
$total = $itens.Count
$passo = [Math]::Max(1, [int]($total / $AMOSTRA))

$lidos = 0; $pulados = 0; $erros = 0
$msTexto = 0.0; $msHtml = 0.0
$charsTexto = [int64]0; $charsHtml = [int64]0
$maiorTexto = 0; $maiorHtml = 0
$semCorpo = 0
$faixas = @{ "0-1k" = 0; "1-5k" = 0; "5-20k" = 0; "20-100k" = 0; "100k+" = 0 }

for ($i = 1; $i -le $total -and $lidos -lt $AMOSTRA; $i += $passo) {
    $it = $null
    try {
        $it = $itens.Item($i)

        # NAO usar [Microsoft.Office.Interop.Outlook.MailItem] aqui: no
        # PowerShell o objeto vem como System.__ComObject e o cast devolve
        # NULO, sem erro. Filtrar por classe e o que funciona.
        $cls = ""
        try { $cls = "$($it.MessageClass)" } catch { }
        if (-not $cls.StartsWith("IPM.Note", [StringComparison]::OrdinalIgnoreCase)) {
            $pulados++; continue
        }

        # olHeaderOnly = 0, olFullItem = 1. Eu tinha invertido, e o filtro
        # pulava exatamente os itens completos. Ler .Body de um item que so
        # tem cabecalho DISPARARIA download.
        $ds = -1
        try { $ds = [int]$it.DownloadState } catch { }
        if ($ds -ne 1) { $pulados++; continue }

        $mail = $it

        $sw = [Diagnostics.Stopwatch]::StartNew()
        $corpo = $mail.Body
        $sw.Stop()
        $msTexto += $sw.Elapsed.TotalMilliseconds
        $n = if ($null -eq $corpo) { 0 } else { $corpo.Length }
        $charsTexto += $n
        if ($n -gt $maiorTexto) { $maiorTexto = $n }
        if ($n -eq 0) { $semCorpo++ }
        $corpo = $null

        $sw = [Diagnostics.Stopwatch]::StartNew()
        $html = $mail.HTMLBody
        $sw.Stop()
        $msHtml += $sw.Elapsed.TotalMilliseconds
        $h = if ($null -eq $html) { 0 } else { $html.Length }
        $charsHtml += $h
        if ($h -gt $maiorHtml) { $maiorHtml = $h }
        $html = $null

        $faixa = if ($n -lt 1000) { "0-1k" }
                 elseif ($n -lt 5000) { "1-5k" }
                 elseif ($n -lt 20000) { "5-20k" }
                 elseif ($n -lt 100000) { "20-100k" }
                 else { "100k+" }
        $faixas[$faixa]++
        $lidos++
    } catch {
        $erros++
    } finally { Solta $it }
}
Solta $itens

if ($lidos -eq 0) {
    Write-Host ""
    Write-Host ("FALHA: a amostra ficou VAZIA (pulados {0}, erros {1})." -f $pulados, $erros)
    Write-Host "Zero itens lidos nao e medicao, e filtro quebrado. Nao publique"
    Write-Host "nenhum numero desta execucao."
    exit 1
}
Write-Host ("CORPO, amostra de {0} itens (pulados {1}, erros {2})" -f $lidos, $pulados, $erros)
Write-Host ("-" * 62)
$mt = $msTexto / [Math]::Max($lidos,1)
$mh = $msHtml / [Math]::Max($lidos,1)
Write-Host ("  .Body     : {0,8:N0} ms total | {1,6:N2} ms/item | {2,8:N0} chars/item" -f `
    $msTexto, $mt, ($charsTexto / [Math]::Max($lidos,1)))
Write-Host ("  .HTMLBody : {0,8:N0} ms total | {1,6:N2} ms/item | {2,8:N0} chars/item" -f `
    $msHtml, $mh, ($charsHtml / [Math]::Max($lidos,1)))
Write-Host ("  maior texto: {0:N0} chars | maior HTML: {1:N0} chars" -f $maiorTexto, $maiorHtml)
Write-Host ("  itens com corpo VAZIO: {0}" -f $semCorpo)
Write-Host ""
Write-Host "  distribuicao do corpo em texto:"
foreach ($k in @("0-1k","1-5k","5-20k","20-100k","100k+")) {
    $pct = 100.0 * $faixas[$k] / [Math]::Max($lidos,1)
    Write-Host ("     {0,-9} {1,4} itens  {2,5:N1}%" -f $k, $faixas[$k], $pct)
}

Write-Host ""
Write-Host ("=" * 62)
Write-Host "O QUE ISSO CUSTA EM ESCALA"
Write-Host ("=" * 62)
$msMetaItem = $msMeta / [Math]::Max($nMeta,1)
Write-Host ("  metadado : {0,8:N2} ms/item" -f $msMetaItem)
Write-Host ("  + corpo  : {0,8:N2} ms/item   ({1:N0}x mais caro)" -f `
    $mt, ($mt / [Math]::Max($msMetaItem, 0.001)))
Write-Host ""
foreach ($q in @(1004, 3000, 17668)) {
    $segMeta = $q * $msMetaItem / 1000
    $segCorpo = $q * $mt / 1000
    $mbTexto = $q * ($charsTexto / [Math]::Max($lidos,1)) * 2 / 1MB
    Write-Host ("  {0,6} msgs -> metadado {1,7:N1} s | corpo {2,7:N1} s | texto ocuparia {3,7:N1} MB" -f `
        $q, $segMeta, $segCorpo, $mbTexto)
}

Write-Host ""
Write-Host "AMPLIFICACAO (a metade que a Q2 acrescentou):"
Write-Host "  Se o indice se ligar a ENCARNACAO, cada Move reindexa o corpo"
Write-Host "  inteiro, porque a §11.1 mediu que o Move cria encarnacao nova."
Write-Host ("  Custo de reindexar UMA mensagem: {0:N2} ms." -f $mt)
Write-Host "  Se se ligar ao ITEM LOGICO, o Move custa ZERO de reindexacao."
Write-Host "  E decisao de SCHEMA, nao de otimizacao depois."

Solta $in
Write-Host ""
Write-Host "Nenhum conteudo foi impresso ou gravado: so comprimento e tempo."
