# Q1, parte 2: a PROTECAO (IRM/rotulo) da para saber em lote?
#
# SOMENTE LEITURA.
#
# Esta pergunta ficou de fora do teste de colunas porque eu tinha marcado
# "Permission = SIM" usando o proptag 0x0E01000B, que e PR_DELETE_AFTER_SUBMIT
# e nao permissao. A coluna foi ACEITA e devolveu nulo — o falso positivo
# classico: coluna ausente volta vazia em vez de dar erro.
#
# MailItem.Permission e propriedade de nivel OOM. A pergunta certa e: existe
# algum sinal EM TABELA que diga se a mensagem e protegida?
#
# Hipotese principal: MessageClass. Mensagem protegida por RMS costuma ser
# IPM.Note.rpmsg.Message; S/MIME e IPM.Note.SMIME*.

param([int]$PastaId = 6, [int]$Amostra = 400)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$pasta = $ns.GetDefaultFolder($PastaId)

Write-Output "pasta: $($pasta.Name)"
Write-Output ""

# --- 1. Que MessageClass existem nesta caixa? ---
$t = $pasta.GetTable()
$classes = @{}
$n = 0
while (-not $t.EndOfTable -and $n -lt $Amostra) {
    $linha = $t.GetNextRow()
    $c = "$($linha.Item('MessageClass'))"
    if ($classes.ContainsKey($c)) { $classes[$c]++ } else { $classes[$c] = 1 }
    $n++
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)

Write-Output "MessageClass em $n itens:"
$classes.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    "  {0,5}x  {1}" -f $_.Value, $_.Key
}

# --- 2. MailItem.Permission concorda com MessageClass? ---
# Abre alguns itens e compara. Se Permission for sempre 0 e nao houver
# classe protegida, esta caixa nao permite concluir nada — e dizer isso e o
# resultado.
Write-Output ""
Write-Output "Conferindo MailItem.Permission em 30 itens (abrindo o item):"

$itens = $pasta.Items
$itens.Sort("[ReceivedTime]", $true)
$comPermissao = 0
$abertos = 0
for ($i = 1; $i -le [Math]::Min(30, $itens.Count); $i++) {
    $m = $itens.Item($i)
    if ($m.MessageClass -like "IPM.Note*") {
        try {
            if ([int]$m.Permission -ne 0) {
                $comPermissao++
                Write-Output "  protegido: classe=$($m.MessageClass) permission=$($m.Permission)"
            }
            $abertos++
        } catch {
            Write-Output "  falhou ao ler Permission: $($_.Exception.Message.Substring(0,40))"
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m)
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens)

Write-Output "  $abertos itens abertos, $comPermissao com Permission <> 0"
Write-Output ""
if ($comPermissao -eq 0) {
    Write-Output "CONCLUSAO PARCIAL: nenhuma mensagem protegida nesta amostra."
    Write-Output "Nao da para provar que MessageClass detecta protecao usando"
    Write-Output "uma caixa que nao tem mensagem protegida. Fica NAO VALIDADO."
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
