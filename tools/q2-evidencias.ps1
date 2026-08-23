# Q2, parte 1: como as evidencias de correlacao se comportam na caixa real.
#
# SOMENTE LEITURA. Nao cria, nao move, nao apaga.
#
# PRIVACIDADE: nada de assunto, endereco ou corpo sai daqui. Message-ID e
# SearchKey aparecem so como HASH e como contagem — o que interessa e se
# eles COLIDEM e se estao PRESENTES, nao o que dizem.
#
# A pergunta da Q2 nao e "qual chave vence". E:
#
#   Quais evidencias permitem correlacionar duas manifestacoes do mesmo
#   item, e ONDE elas erram?
#
# O erro que mais importa nao e deixar de unir um item movido. E UNIR DOIS
# ITENS DISTINTOS: ai o resumo da IA e o estado do usuario vao parar na
# mensagem errada. E o risco R2-G.

$ErrorActionPreference = "Stop"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$PR_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"
$PR_MSG_ID     = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"

function Hash([string]$texto) {
    if ([string]::IsNullOrEmpty($texto)) { return $null }
    $sha = [Security.Cryptography.SHA256]::Create()
    $b = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($texto))
    $sha.Dispose()
    return ([BitConverter]::ToString($b) -replace '-','').Substring(0, 12)
}

function Levantar($pasta, [string]$rotulo) {
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    [void]$t.Columns.Add("EntryID")
    [void]$t.Columns.Add("MessageClass")
    [void]$t.Columns.Add($PR_SEARCH_KEY)
    [void]$t.Columns.Add($PR_MSG_ID)

    $itens = @()
    while (-not $t.EndOfTable) {
        $a = $t.GetArray(200)
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $sk = $a.GetValue($r, 2)
            $skTexto = if ($null -eq $sk) { $null } else {
                if ($sk -is [array]) { ($sk | ForEach-Object { $_.ToString("x2") }) -join "" } else { "$sk" }
            }
            $itens += [pscustomobject]@{
                Pasta   = $rotulo
                Id      = "$($a.GetValue($r,0))"
                Classe  = "$($a.GetValue($r,1))"
                SkHash  = Hash $skTexto
                MidHash = Hash ("$($a.GetValue($r,3))".Trim())
                TemMid  = -not [string]::IsNullOrWhiteSpace("$($a.GetValue($r,3))")
                TemSk   = -not [string]::IsNullOrWhiteSpace($skTexto)
            }
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    return ,$itens
}

# 6=Entrada, 5=Enviados, 16=Rascunhos, 3=Excluidos
$pastas = @(
    @{ Id = 6;  Rotulo = "Entrada"   },
    @{ Id = 5;  Rotulo = "Enviados"  },
    @{ Id = 16; Rotulo = "Rascunhos" },
    @{ Id = 3;  Rotulo = "Excluidos" }
)

$todos = @()
Write-Host "pasta      | itens | com Message-ID | com SearchKey"
Write-Host "-----------|-------|----------------|--------------"

foreach ($p in $pastas) {
    try {
        $pasta = $ns.GetDefaultFolder($p.Id)
        $itens = Levantar $pasta $p.Rotulo
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
    } catch {
        Write-Host ("{0,-10} | ERRO: {1}" -f $p.Rotulo, $_.Exception.Message.Substring(0,40))
        continue
    }

    $todos += $itens
    $comMid = ($itens | Where-Object { $_.TemMid }).Count
    $comSk  = ($itens | Where-Object { $_.TemSk }).Count
    Write-Host ("{0,-10} | {1,5} | {2,5} ({3,3}%) | {4,5} ({5,3}%)" -f `
        $p.Rotulo, $itens.Count,
        $comMid, [int](100 * $comMid / [Math]::Max($itens.Count,1)),
        $comSk,  [int](100 * $comSk  / [Math]::Max($itens.Count,1)))
}

Write-Host ""
Write-Host "TOTAL: $($todos.Count) itens"
Write-Host ""

# --- Onde falta Message-ID, e de que tipo sao esses itens ---
$semMid = $todos | Where-Object { -not $_.TemMid }
Write-Host "SEM Message-ID: $($semMid.Count)"
if ($semMid.Count -gt 0) {
    $semMid | Group-Object Pasta, Classe | Sort-Object Count -Descending |
        Select-Object -First 10 | ForEach-Object {
            Write-Host ("   {0,4}x  {1}" -f $_.Count, $_.Name)
        }
}
Write-Host ""

# --- COLISAO: o mesmo Message-ID em itens DIFERENTES ---
$porMid = $todos | Where-Object { $_.TemMid } | Group-Object MidHash |
          Where-Object { $_.Count -gt 1 }

Write-Host "Message-ID compartilhado por mais de um item: $($porMid.Count) grupos"
if ($porMid.Count -gt 0) {
    $mesmaPasta = 0; $entrePastas = 0
    foreach ($g in $porMid) {
        $pastasDoGrupo = ($g.Group | Select-Object -ExpandProperty Pasta -Unique)
        if ($pastasDoGrupo.Count -gt 1) { $entrePastas++ } else { $mesmaPasta++ }
    }
    Write-Host "   na MESMA pasta : $mesmaPasta"
    Write-Host "   entre PASTAS   : $entrePastas"
    Write-Host ""
    Write-Host "   exemplos (so a combinacao de pastas e o tamanho do grupo):"
    $porMid | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object {
        $combo = (($_.Group | Select-Object -ExpandProperty Pasta) | Sort-Object) -join "+"
        Write-Host ("     {0} itens em [{1}]" -f $_.Count, $combo)
    }
}
Write-Host ""

# --- COLISAO de SearchKey ---
$porSk = $todos | Where-Object { $_.TemSk } | Group-Object SkHash |
         Where-Object { $_.Count -gt 1 }
Write-Host "SearchKey compartilhado por mais de um item: $($porSk.Count) grupos"
if ($porSk.Count -gt 0) {
    $porSk | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object {
        $combo = (($_.Group | Select-Object -ExpandProperty Pasta) | Sort-Object) -join "+"
        Write-Host ("     {0} itens em [{1}]" -f $_.Count, $combo)
    }
}
Write-Host ""
Write-Host "LEITURA: grupo com mais de um item significa que aquela evidencia"
Write-Host "NAO distingue os itens do grupo. Se algum grupo tiver itens que sao"
Write-Host "mensagens DIFERENTES, usar aquela evidencia sozinha funde os dois."
