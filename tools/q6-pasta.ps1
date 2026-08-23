# Q6a: a identidade de PASTA sobrevive a renomear, mover e recriar?
#
# ESCREVE. So em pastas que este script cria, na raiz da caixa, e remove no
# fim (Delete de pasta e soft: vai para Itens Excluidos).
#
# Nenhuma pasta do usuario e tocada. Nenhum item e criado, movido ou
# apagado.
#
# POR QUE ISTO IMPORTA: a secao 11.1 mediu que o EntryID de ITEM NAO
# sobrevive a um Move. O FolderKey do Iris e EntryID + StoreID. Se o
# EntryID de PASTA tambem mudar, a chave de pasta do cache tem exatamente o
# mesmo defeito, e ninguem tinha verificado.
#
# CENARIOS:
#   A) RENOMEAR       - o nome faz parte da identidade?
#   B) MOVER          - a hierarquia faz parte?
#   C) APAGAR+RECRIAR - pasta recriada com o MESMO NOME e a mesma pasta?
#                       Se o EntryID for reaproveitado, o cache atribui a
#                       uma pasta nova o historico da antiga. Pior caso.
#
# COMO B E TESTADO, e o detalhe importa: MAPIFolder.MoveTo e um Sub e NAO
# devolve a pasta — diferente de MailItem.Move, que devolve. Entao em vez
# de pegar o retorno, o teste guarda o EntryID ANTES e tenta RESOLVE-LO
# depois com GetFolderFromID. Isso responde a pergunta de forma direta:
# "o identificador antigo ainda aponta para a pasta?".
#
# SEM RETRY. Mutacao que falha no meio nao se repete as cegas.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Solta($o) {
    if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
}

function Hex($v) {
    if ($null -eq $v) { return "(nulo)" }
    if ($v -is [byte[]]) {
        if ($v.Length -eq 0) { return "(vazio)" }
        return (($v | ForEach-Object { $_.ToString("X2") }) -join "")
    }
    return "$v"
}

function Foto($pasta) {
    if ($null -eq $pasta) { throw "Foto recebeu pasta nula" }
    $r = [ordered]@{
        Nome    = $pasta.Name
        EntryID = $pasta.EntryID
        StoreID = $pasta.StoreID
    }
    $pa = $null
    try {
        $pa = $pasta.PropertyAccessor
        foreach ($p in @(
            @{ N = "RecordKey"; T = "0x0FF90102" },
            @{ N = "ChangeKey"; T = "0x65E20102" })) {
            try { $r[$p.N] = Hex $pa.GetProperty($PT + $p.T) }
            catch { $r[$p.N] = "(ERRO AO LER)" }
        }
    } finally { Solta $pa }
    return $r
}

function Comparar($antes, $depois, [string]$oQue) {
    Write-Host ""
    Write-Host ("=" * 78)
    Write-Host $oQue
    Write-Host ("=" * 78)
    $mudou = @()
    foreach ($k in $antes.Keys) {
        $a = "$($antes[$k])"; $b = "$($depois[$k])"
        $igual = ($a -eq $b)
        if (-not $igual) { $mudou += $k }
        $mostra = if ($igual) {
            if ($a.Length -gt 44) { $a.Substring(0, 41) + "..." } else { $a }
        } else {
            $n = [Math]::Min($a.Length, $b.Length)
            $c = 0
            while ($c -lt $n -and $a[$c] -eq $b[$c]) { $c++ }
            "prefixo comum de $c de $($a.Length)"
        }
        Write-Host ("  {0,-9} | {1,-6} | {2}" -f $k, $(if ($igual) { "IGUAL" } else { "MUDOU" }), $mostra)
    }
    return ,$mudou
}

function Resolve($entryId, $storeId) {
    # O identificador antigo ainda aponta para alguma pasta?
    try { return $ns.GetFolderFromID($entryId, $storeId) } catch { return $null }
}

function Achar($pai, [string]$nome) {
    foreach ($f in $pai.Folders) {
        if ($f.Name -eq $nome) { return $f }
        Solta $f
    }
    return $null
}

$raiz = $ns.GetDefaultFolder(6).Parent
$resumo = @()

try {
    # ---------------------------------------------------------------
    # A) RENOMEAR
    # ---------------------------------------------------------------
    $a = $raiz.Folders.Add("Iris Q6 alfa")
    $antes = Foto $a
    $a.Name = "Iris Q6 alfa renomeada"
    $depois = Foto $a
    $m = Comparar $antes $depois "A) RENOMEAR uma pasta"
    $resumo += @{ Caso = "renomear"; Mudou = @($m | Where-Object { $_ -ne "Nome" }) }

    # ---------------------------------------------------------------
    # B) MOVER para dentro de outra pasta
    # ---------------------------------------------------------------
    $b = $raiz.Folders.Add("Iris Q6 beta")
    $antesMove = Foto $a
    $idAntigo = $antesMove.EntryID
    $storeId = $antesMove.StoreID

    $a.MoveTo($b)      # Sub: nao devolve nada
    Solta $a

    # 1) o identificador ANTIGO ainda resolve?
    $porIdAntigo = Resolve $idAntigo $storeId
    $resolveu = ($null -ne $porIdAntigo)

    # 2) e onde a pasta esta de fato?
    $movida = Achar $b "Iris Q6 alfa renomeada"
    if ($null -eq $movida) { throw "a pasta sumiu depois do MoveTo" }
    $depoisMove = Foto $movida

    $m = Comparar $antesMove $depoisMove "B) MOVER a pasta para dentro de outra"
    Write-Host ""
    Write-Host ("  GetFolderFromID com o EntryID ANTIGO: {0}" -f `
        $(if ($resolveu) { "AINDA RESOLVE" } else { "NAO resolve mais" }))
    if ($resolveu) {
        $mesmoNome = ($porIdAntigo.Name -eq "Iris Q6 alfa renomeada")
        Write-Host ("     e aponta para a pasta certa? {0}" -f $(if ($mesmoNome) { "SIM" } else { "NAO" }))
        Solta $porIdAntigo
    }
    $resumo += @{ Caso = "mover"; Mudou = @($m); Resolve = $resolveu }
    Solta $movida

    # ---------------------------------------------------------------
    # C) APAGAR e RECRIAR com o MESMO NOME
    # ---------------------------------------------------------------
    $gama1 = $raiz.Folders.Add("Iris Q6 gama")
    $antesGama = Foto $gama1
    $gama1.Delete()
    Solta $gama1

    $gama2 = $raiz.Folders.Add("Iris Q6 gama")
    $depoisGama = Foto $gama2
    $m = Comparar $antesGama $depoisGama "C) APAGAR e RECRIAR com o MESMO NOME"

    $reaproveitou = ($antesGama.EntryID -eq $depoisGama.EntryID)
    Write-Host ""
    if ($reaproveitou) {
        Write-Host "  !!! O EntryID FOI REAPROVEITADO. Uma pasta NOVA herdaria o"
        Write-Host "      historico da antiga no cache. E o pior caso possivel."
    } else {
        Write-Host "  O EntryID NAO foi reaproveitado: pasta recriada e outra pasta."
    }
    $resumo += @{ Caso = "recriar"; Mudou = @($m | Where-Object { $_ -ne "Nome" }); Reuso = $reaproveitou }
    Solta $gama2
    Solta $b

} catch {
    Write-Host ""
    Write-Host "!!! FALHA NO MEIO DA MUTACAO !!!"
    Write-Host $_.Exception.Message
    Write-Host "Confira as pastas 'Iris Q6 *' na raiz."
} finally {
    Write-Host ""
    Write-Host ("=" * 78)
    Write-Host "LIMPEZA"
    Write-Host ("=" * 78)
    # Recursiva: a pasta movida vive dentro de outra.
    function Limpar($pai) {
        foreach ($f in @($pai.Folders)) {
            try {
                if ($f.Name -like "Iris Q6*") {
                    Limpar $f
                    $n = $f.Items.Count
                    if ($n -eq 0) {
                        Write-Host ("  removendo: {0}" -f $f.Name)
                        $f.Delete()
                    } else {
                        Write-Host ("  MANTIDA ({0} itens): {1}" -f $n, $f.Name)
                    }
                }
            } catch {
                Write-Host ("  nao consegui limpar: {0}" -f $_.Exception.Message)
            } finally { Solta $f }
        }
    }
    Limpar $raiz
    Solta $raiz
}

Write-Host ""
Write-Host ("=" * 78)
Write-Host "VEREDITO"
Write-Host ("=" * 78)
foreach ($r in $resumo) {
    $lista = if ($r.Mudou.Count -eq 0) { "nada" } else { ($r.Mudou -join ", ") }
    Write-Host ("  {0,-10} -> mudou: {1}" -f $r.Caso, $lista)
}
Write-Host ""
Write-Host "LEITURA: se o EntryID de pasta mudar em algum destes casos, o"
Write-Host "FolderKey = EntryID + StoreID tem o MESMO defeito que a 11.1 achou"
Write-Host "na chave de item, e a pasta tambem precisa de identidade interna."
