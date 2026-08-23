# Q1, parte 1: QUAIS colunas a Table entrega.
#
# SOMENTE LEITURA.
#
# O numero agregado nao interessa antes desta matriz. Se sete colunas vierem
# em lote mas protecao e anexos ainda exigirem abrir cada item, o ganho
# desaparece — e e por isso que a matriz vem primeiro.
#
# Cada coluna e testada SOZINHA: adicionar varias e ver a tabela falhar nao
# diz qual delas e a culpada.

param([int]$PastaId = 6)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$pasta = $ns.GetDefaultFolder($PastaId)

Write-Output "pasta: $($pasta.Name)"
Write-Output ""

# --- Colunas que a tabela traz sem pedir ---
$t = $pasta.GetTable()
Write-Output "COLUNAS PADRAO ($($t.Columns.Count)):"
for ($i = 1; $i -le $t.Columns.Count; $i++) {
    Write-Output "  - $($t.Columns.Item($i).Name)"
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)

Write-Output ""
Write-Output "CADA COLUNA DO MailSummary, TESTADA SOZINHA:"
Write-Output ""
Write-Output "coluna                          | add | valor lido"
Write-Output "--------------------------------|-----|------------"

# Nome do DTO -> candidatos de nome de coluna, do mais simples ao DASL.
$alvos = [ordered]@{
    "EntryID"       = @("EntryID")
    "Subject"       = @("Subject")
    "SenderName"    = @("SenderName", "urn:schemas:httpmail:fromname")
    "ReceivedTime"  = @("ReceivedTime", "urn:schemas:httpmail:datereceived")
    "Size"          = @("Size", "http://schemas.microsoft.com/mapi/proptag/0x0E080003")
    "UnRead"        = @("UnRead", "urn:schemas:httpmail:read")
    "MessageClass"  = @("MessageClass")
    "HasAttachment" = @("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B",
                        "urn:schemas:httpmail:hasattachment")
    "Permission"    = @("Permission",
                        "http://schemas.microsoft.com/mapi/proptag/0x0E01000B")
    "LastModified"  = @("LastModificationTime",
                        "urn:schemas:httpmail:datereceived")
    "SearchKey"     = @("http://schemas.microsoft.com/mapi/proptag/0x300B0102")
    "InternetMsgId" = @("http://schemas.microsoft.com/mapi/proptag/0x1035001E",
                        "urn:schemas:httpmail:message-id")
}

foreach ($nome in $alvos.Keys) {
    $ok = $false
    $usado = ""
    $amostra = ""

    foreach ($candidato in $alvos[$nome]) {
        $tab = $null
        try {
            $tab = $pasta.GetTable()
            [void]$tab.Columns.Add($candidato)

            # Adicionar pode aceitar e a leitura falhar. So conta se ler.
            if (-not $tab.EndOfTable) {
                $linha = $tab.GetNextRow()
                $v = $linha.Item($candidato)
                $amostra = if ($null -eq $v) { "(nulo)" } else {
                    $texto = "$v"
                    if ($texto.Length -gt 22) { $texto.Substring(0, 22) + "…" } else { $texto }
                }
            } else {
                $amostra = "(pasta vazia)"
            }
            $ok = $true
            $usado = $candidato
        } catch {
            $amostra = ""
        } finally {
            if ($tab) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($tab) }
        }
        if ($ok) { break }
    }

    $marca = if ($ok) { "SIM" } else { "NAO" }
    $detalhe = if ($ok) { "$amostra" } else { "nenhum candidato serviu" }
    "{0,-31} | {1,-3} | {2}" -f $nome, $marca, $detalhe
    if ($ok -and $usado -ne $nome) { "{0,-31} |     | via: {1}" -f "", $usado }
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
