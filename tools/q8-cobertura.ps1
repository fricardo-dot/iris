# A pasta tem N, e eu enxergo M. Qual propriedade devolve N?
#
# A UI do Outlook reporta 17.728 itens na Caixa de Entrada e 145 em
# "1. Backup". O OOM, via Items.Count, devolve 1.018 e 35. Se alguma
# propriedade MAPI der o numero da UI, o Iris consegue MEDIR a propria
# cobertura em vez de supor - e cobertura medida e exatamente o que falta
# para a FolderCoverage sair de "Desconhecida".
#
# Candidatos:
#   PR_CONTENT_COUNT        0x36020003  itens na pasta
#   PR_CONTENT_UNREAD       0x36030003  nao lidos
#   PR_ASSOC_CONTENT_COUNT  0x36170003  itens associados (ocultos)
#   PR_DELETED_MSG_COUNT    0x66430003
$ErrorActionPreference = "Stop"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$tags = [ordered]@{
    "PR_CONTENT_COUNT"       = ($PT + "0x36020003")
    "PR_CONTENT_UNREAD"      = ($PT + "0x36030003")
    "PR_ASSOC_CONTENT_COUNT" = ($PT + "0x36170003")
}

function Ler($pasta, $tag) {
    $pa = $null
    try { $pa = $pasta.PropertyAccessor; return $pa.GetProperty($tag) }
    catch { return "erro" }
    finally { if ($pa) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) } }
}

$alvos = @("Caixa de Entrada", "1. Backup", "Itens Enviados", "Lixo Eletronico",
           "Lixo Eletrônico", "Rascunhos", "Itens Excluidos", "Itens Excluídos",
           "Arquivo Morto", "0. E-mails Lidos", "Spam")

$linhas = @()
function Varrer($pasta, $prof) {
    if ($prof -gt 6) { return }
    if ($alvos -contains $pasta.Name) {
        $r = [ordered]@{ Pasta = $pasta.Name }
        try { $r["Items.Count"] = $pasta.Items.Count } catch { $r["Items.Count"] = "erro" }
        foreach ($k in $tags.Keys) { $r[$k] = (Ler $pasta $tags[$k]) }
        # Table: a contagem que a paginacao do Iris realmente veria
        try {
            $tb = $pasta.GetTable(); $c = $tb.Columns; $c.RemoveAll()
            $x = $c.Add("EntryID"); $n = 0
            while (-not $tb.EndOfTable) {
                $a = $tb.GetArray(500); $n += ($a.GetUpperBound(0) - $a.GetLowerBound(0) + 1)
            }
            $r["Table"] = $n
        } catch { $r["Table"] = "erro" }
        $script:linhas += ,(New-Object PSObject -Property $r)
    }
    $filhas = $null
    try {
        $filhas = $pasta.Folders; $c = $filhas.Count
        for ($i = 1; $i -le $c; $i++) {
            $f = $null
            try { $f = $filhas.Item($i); Varrer $f ($prof + 1) }
            finally { if ($f) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) } }
        }
    } catch { } finally { if ($filhas) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) } }
}

foreach ($st in $ns.Stores) {
    $raiz = $null
    try { $raiz = $st.GetRootFolder(); Varrer $raiz 0 }
    finally { if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) } }
}

$linhas | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
