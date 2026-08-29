# Q2: TODO par que alguma evidencia une, caracterizado propriedade a
# propriedade.
#
# SOMENTE LEITURA.
#
# Substitui q2-evidencias.ps1, q2-colisoes.ps1 e q2-par.ps1, que juntos
# tinham quatro defeitos:
#
#   1. Olhavam so as 4 pastas padrao, e eu chamei aquilo de "a caixa". A
#      arvore tem 127 pastas e 2281 itens.
#   2. $grupos.Count sem @() devolvia o tamanho do GRUPO quando havia um so.
#   3. q2-par.ps1 selecionava por PREFIXO DE ASSUNTO, achou 13 itens, ficou
#      com 2 e sobrescreveu os outros 11 em silencio. Os dois que sobraram
#      por acaso eram os certos. Agora a selecao e pelo EntryID do grupo de
#      colisao de fato encontrado.
#
#      ATENCAO: este script imprime os DOIS PRIMEIROS itens de cada grupo.
#      Uma versao anterior deste comentario prometia que ele "PARA se o
#      grupo nao tiver exatamente os itens esperados". Nao para, e nunca
#      parou. Grupo com 3+ itens e AVISADO abaixo, mas nao interrompe.
#   4. O comentario dizia que o discriminador era "enviado tem SubmitTime e
#      nao tem DeliveryTime". Os dados desmentem: as duas propriedades
#      aparecem nos DOIS itens do par. Quem discrimina e
#      PR_TRANSPORT_MESSAGE_HEADERS (so o transporte escreve) junto com
#      PR_MESSAGE_FLAGS.
#
# PRIVACIDADE: assunto truncado em 30 chars e NOME DE EXIBICAO do remetente
# truncado em 18. E nome, nao dominio — o comentario anterior prometia
# dominio e imprimia nome.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$detalhe = [ordered]@{
    "SearchKey"     = "0x300B0102"
    "MessageID"     = "0x1035001E"
    "RecordKey"     = "0x0FF90102"
    "ChangeKey"     = "0x65E20102"
    "SubmitTime"    = "0x00390040"
    "DeliveryTime"  = "0x0E060040"
    "MsgFlags"      = "0x0E070003"
    "TransportHdrs" = "0x007D001E"
    "ConvIndex"     = "0x00710102"
}

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Hex($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [byte[]]) {
        if ($v.Length -eq 0) { return $null }
        return (($v | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    if ($v -is [string]) { $s = $v.Trim(); if ($s -eq "") { return $null }; return $s }
    return "$v"
}

# ---------- 1. varrer a arvore inteira ----------
$itens = New-Object System.Collections.ArrayList
$falhas = New-Object System.Collections.ArrayList

# CORTE E RAMO CEGO. O laco para na profundidade 12 e engole falha ao abrir
# as filhas; sem contar os dois, "nenhuma colisao" seria afirmacao sobre o que
# nao foi percorrido.
$script:cortados = 0
$script:ramosCegos = 0

function Varrer($pasta, [string]$caminho, [int]$prof) {
    if ($prof -gt 12) { $script:cortados++; return }
    $t = $null
    try {
        $t = $pasta.GetTable()
        $cols = $t.Columns
        try {
            $cols.RemoveAll()
            # Os parenteses NAO sao opcionais. Dentro de @(...) o PowerShell
            # le $PT + "0x.." em modo argumento e devolve TRES elementos
            # ("$PT", "+", "0x..") em vez de um. O Add recebia a URL do
            # prefixo sozinha e as 127 pastas falhavam com "Value does not
            # fall within the expected range".
            foreach ($c in @("EntryID", "Subject", "SenderName",
                             ($PT + "0x300B0102"), ($PT + "0x1035001E"))) {
                [void]$cols.Add($c)
            }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cols) }

        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                [void]$script:itens.Add([pscustomobject]@{
                    Pasta   = $caminho
                    Id      = "$($a.GetValue($r,0))"
                    Assunto = "$($a.GetValue($r,1))"
                    Remet   = "$($a.GetValue($r,2))"
                    Sk      = Hex $a.GetValue($r,3)
                    Mid     = Hex $a.GetValue($r,4)
                })
            }
        }
    } catch {
        # Catch vazio ja me custou uma varredura inteira devolvendo zero em
        # silencio. Se uma pasta falhar, eu quero saber qual e por que.
        [void]$script:falhas.Add("$caminho : $($_.Exception.Message)")
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    }

    $filhas = $null
    try { $filhas = $pasta.Folders } catch { $script:ramosCegos++; return }
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try { Varrer $f "$caminho/$($f.Name)" ($prof + 1) }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
}

$stores = $ns.Stores
for ($s = 1; $s -le $stores.Count; $s++) {
    $store = $stores.Item($s)
    $raiz = $null
    try { $raiz = $store.GetRootFolder(); Varrer $raiz $store.DisplayName 0 }
    catch { Write-Host "STORE INACESSIVEL: $($_.Exception.Message)" }
    finally {
        if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($store)
    }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)

Write-Host "corpus: $($itens.Count) itens"
if ($falhas.Count -gt 0) {
    Write-Host "pastas que falharam: $($falhas.Count)"
    $falhas | Select-Object -First 3 | ForEach-Object { Write-Host "   $_" }
}
Write-Host ""

# ---------- 2. achar os grupos ----------
# O @() e obrigatorio: sem ele, um grupo unico devolve o tamanho do grupo.
$grupos = @()
foreach ($campo in @("Sk", "Mid")) {
    $g = @($itens | Where-Object { $null -ne $_.$campo } | Group-Object $campo |
           Where-Object { $_.Count -gt 1 })
    foreach ($grupo in $g) {
        $grupos += [pscustomobject]@{
            Por = $(if ($campo -eq "Sk") { "SearchKey" } else { "Message-ID" })
            Itens = @($grupo.Group)
        }
    }
}
Write-Host "grupos de colisao: $($grupos.Count)"
if ($grupos.Count -eq 0) {
    Write-Host "nenhuma colisao ENTRE OS ITENS QUE ESTE ROTEIRO LEU."
    if ($script:cortados -gt 0 -or $script:ramosCegos -gt 0 -or $falhas.Count -gt 0) {
        Write-Host ("  RESSALVA: {0} ramo(s) cortados na profundidade 12, {1} ramo(s) que nao" -f $script:cortados, $script:ramosCegos)
        Write-Host ("  consegui abrir, {0} pasta(s) que falharam. Nao e 'nao existe colisao'." -f $falhas.Count)
    }
    exit 0
}

# ---------- 3. caracterizar cada um ----------
function Curto([string]$s, [int]$n) {
    if ([string]::IsNullOrEmpty($s)) { return "(vazio)" }
    if ($s.Length -le $n) { return $s }
    return $s.Substring(0, $n) + "..."
}

$vistos = @{}
foreach ($g in $grupos) {
    # Um mesmo par colide por SearchKey E por Message-ID; caracterizo uma vez.
    $assinatura = (($g.Itens | Select-Object -ExpandProperty Id | Sort-Object) -join "|")
    if ($vistos.ContainsKey($assinatura)) { continue }
    $vistos[$assinatura] = $true

    Write-Host ("=" * 96)
    Write-Host ("PAR unido por {0} — {1} itens" -f $g.Por, $g.Itens.Count)
    Write-Host ("=" * 96)

    $colunas = @()
    foreach ($it in $g.Itens) {
        $item = $null
        try { $item = $ns.GetItemFromID($it.Id) } catch {
            Write-Host "  (item inacessivel: $($it.Pasta))"; continue
        }
        try {
            $col = [ordered]@{}
            $col["Pasta"]   = $it.Pasta.Split("/")[-1]
            $col["Assunto"] = Curto $it.Assunto 30
            $col["Remet"]   = Curto $it.Remet 18
            $col["Classe"]  = $item.MessageClass
            $col["Tamanho"] = $item.Size
            $pa = $item.PropertyAccessor
            try {
                foreach ($k in $detalhe.Keys) {
                    try {
                        $v = $pa.GetProperty($PT + $detalhe[$k])
                        $col[$k] = if ($k -eq "TransportHdrs") {
                            if ($null -eq $v -or "$v" -eq "") { "vazio" } else { "$("$v".Length) chars" }
                        } else { Hex $v }
                    } catch { $col[$k] = "(ausente)" }
                }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pa) }
            $colunas += $col
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }

    if ($colunas.Count -lt 2) { Write-Host "  (menos de 2 itens legiveis)"; continue }
    if ($colunas.Count -gt 2) {
        Write-Host ("  AVISO: grupo tem {0} itens; a tabela mostra so os 2 primeiros." -f $colunas.Count)
    }

    $chaves = @("Pasta","Assunto","Remet","Classe","Tamanho") + @($detalhe.Keys)
    Write-Host ("{0,-14} | {1,-34} | {2,-34} | igual?" -f "propriedade", "A", "B")
    Write-Host ("-" * 96)
    foreach ($k in $chaves) {
        $a = "$($colunas[0][$k])"; $b = "$($colunas[1][$k])"
        $ma = if ($a.Length -gt 34) { $a.Substring(0,31) + "..." } else { $a }
        $mb = if ($b.Length -gt 34) { $b.Substring(0,31) + "..." } else { $b }
        Write-Host ("{0,-14} | {1,-34} | {2,-34} | {3}" -f $k, $ma, $mb, $(if ($a -eq $b) { "SIM" } else { "nao" }))
    }
    Write-Host ""
}
