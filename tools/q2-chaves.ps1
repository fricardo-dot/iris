# Q2, rodada 2: as pastas ALCANCADAS, todas as chaves candidatas.
#
# O cabecalho dizia "TODAS as pastas" e "a arvore inteira e percorrida", e
# nao e verdade: a travessia corta na profundidade 12 e podia perder um ramo
# em silencio. Os zeros da matriz apareciam sob promessa de cobertura
# completa, que e a forma mais cara de afirmar ausencia. Agora os dois casos
# sao contados e saem no rodape.
#
# SOMENTE LEITURA.
#
# Substitui q2-evidencias.ps1, que tinha tres defeitos:
#
#   1. Olhava so as 4 pastas padrao e eu chamei o resultado de "a caixa".
#      Subpasta, Lixo Eletronico, Caixa de Saida e pasta do usuario ficaram
#      de fora. Agora a travessia cobre todos os stores, ATE a profundidade 12.
#   2. $grupos.Count sem @() em volta le a propriedade do GroupInfo quando
#      ha UM grupo so. Eu ja tinha DOCUMENTADO esse erro e deixei ele no
#      codigo que gerou a tabela publicada.
#   3. Aceitava como "chave presente" qualquer coisa que virasse texto. Um
#      valor de erro ou um tipo inesperado contava como presenca. Agora o
#      tipo e verificado e o que nao for byte[] ou string vira ANOMALIA
#      contada, nao presenca.
#
# E mede o que a rodada 1 nao mediu: PR_RECORD_KEY e PR_SOURCE_KEY, com
# presenca e UNICIDADE no corpus inteiro. Eu tinha descartado a RecordKey
# olhando UM par, porque os bytes finais dela batem com o EntryID. Bytes em
# comum nao provam equivalencia semantica, e a RecordKey e justamente a
# candidata que pode sobreviver a um Move.

$ErrorActionPreference = "Stop"
$PT = "http://schemas.microsoft.com/mapi/proptag/"

# Cada coluna e tentada uma a uma: se Add falhar e eu usar indice fixo,
# todas as colunas seguintes passam a apontar para outra propriedade em
# silencio. Por isso o mapa nome -> indice e montado do resultado real.
$candidatas = @(
    @{ Nome = "EntryID";   Dasl = "EntryID" },
    @{ Nome = "Classe";    Dasl = "MessageClass" },
    @{ Nome = "SearchKey"; Dasl = $PT + "0x300B0102" },
    @{ Nome = "MessageID"; Dasl = $PT + "0x1035001E" },
    @{ Nome = "RecordKey"; Dasl = $PT + "0x0FF90102" },
    @{ Nome = "SourceKey"; Dasl = $PT + "0x65E00102" },
    @{ Nome = "ChangeKey"; Dasl = $PT + "0x65E20102" }
)

$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

$tipos = @{}
function Contar([string]$chave) {
    $script:tipos[$chave] = 1 + [int]$script:tipos[$chave]
}

function Normalizar($v, [string]$nome) {
    if ($null -eq $v) { Contar "$nome/nulo"; return $null }
    $t = $v.GetType()
    if ($t.FullName -eq "System.Byte[]") {
        Contar "$nome/byte[]"
        if ($v.Length -eq 0) { Contar "$nome/VAZIO"; return $null }
        return (($v | ForEach-Object { $_.ToString("x2") }) -join "")
    }
    if ($t.FullName -eq "System.String") {
        Contar "$nome/string"
        $s = $v.Trim()
        if ($s -eq "") { Contar "$nome/VAZIO"; return $null }
        return $s
    }
    # Nao converto em texto: se veio de um tipo que eu nao previ, eu nao sei
    # o que estou comparando.
    Contar "$nome/ANOMALIA:$($t.Name)"
    return $null
}

$itens = New-Object System.Collections.ArrayList
$pastasVistas = 0
$pastasComErro = New-Object System.Collections.ArrayList
$script:cortados = 0
$script:ramosCegos = 0
$script:semColunaChave = 0   # a pasta nao expoe EntryID: ela E o ramo perdido
$script:semStore = 0         # nem cheguei a abrir o store

function Varrer($pasta, [string]$caminho, [int]$profundidade) {
    if ($profundidade -gt 12) { $script:cortados++; return }
    $script:pastasVistas++

    $t = $null
    try {
        $t = $pasta.GetTable()
        $colunas = $t.Columns
        try {
            $colunas.RemoveAll()
            $mapa = @{}
            $i = 0
            foreach ($c in $candidatas) {
                try { [void]$colunas.Add($c.Dasl); $mapa[$c.Nome] = $i; $i++ }
                catch { }   # recusada por este provider: fica FORA do mapa
            }
        } finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($colunas)
        }
        # RETURN AQUI PULA AS FILHAS TAMBEM. Pasta sem coluna EntryID nao
        # entra no corpus E leva o ramo inteiro junto, sem contar.
        # CAUSA DIFERENTE, CONTADOR DIFERENTE. Isto e "a pasta nao expoe a
        # coluna", e nao "o Folders falhou" -- somar os dois num contador so
        # faz a ressalva explicar a causa errada.
        if (-not $mapa.ContainsKey("EntryID")) { $script:semColunaChave++; return }

        while (-not $t.EndOfTable) {
            $a = $t.GetArray(200)
            for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
                $linha = [ordered]@{ Pasta = $caminho }
                foreach ($c in $candidatas) {
                    $linha[$c.Nome] = if ($mapa.ContainsKey($c.Nome)) {
                        Normalizar $a.GetValue($r, $mapa[$c.Nome]) $c.Nome
                    } else { $null }
                }
                [void]$script:itens.Add([pscustomobject]$linha)
            }
        }
    } catch {
        [void]$script:pastasComErro.Add("$caminho : $($_.Exception.Message)")
    } finally {
        if ($t) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t) }
    }

    $filhas = $null
    try { $filhas = $pasta.Folders } catch { $script:ramosCegos++; return }
    try {
        for ($k = 1; $k -le $filhas.Count; $k++) {
            $f = $filhas.Item($k)
            try { Varrer $f "$caminho/$($f.Name)" ($profundidade + 1) }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f) }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
}

$stores = $ns.Stores
Write-Host "stores: $($stores.Count)"
for ($s = 1; $s -le $stores.Count; $s++) {
    # PEGAR O STORE E LER O NOME TAMBEM PODEM FALHAR, e estavam fora do try:
    # o roteiro abortava sem contar nada, e a matriz saia sobre um universo
    # menor do que ela dizia.
    $store = $null
    try { $store = $stores.Item($s) } catch { $script:semStore++; continue }
    $nome = "(store $s)"
    try { $nome = [string]$store.DisplayName } catch { }
    Write-Host ("  [{0}] {1}" -f $s, $nome)
    $raiz = $null
    try {
        $raiz = $store.GetRootFolder()
        Varrer $raiz $nome 0
    } catch {
        $script:semStore++
        Write-Host "      (sem acesso: $($_.Exception.Message))"
    } finally {
        if ($raiz) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raiz) }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($store)
    }
}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)

Write-Host ""
Write-Host "pastas percorridas: $pastasVistas"
$script:cego = $script:cortados + $script:ramosCegos + $script:semColunaChave + $script:semStore
if ($script:cego -gt 0) {
    Write-Host "O QUE NAO FOI PERCORRIDO:" -ForegroundColor DarkYellow
    if ($script:cortados -gt 0)      { Write-Host ("  {0} ramo(s) cortados na profundidade 12" -f $script:cortados) }
    if ($script:ramosCegos -gt 0)    { Write-Host ("  {0} ramo(s) cujo Folders falhou" -f $script:ramosCegos) }
    if ($script:semColunaChave -gt 0) { Write-Host ("  {0} pasta(s) sem a coluna EntryID -- e o ramo abaixo delas" -f $script:semColunaChave) }
    if ($script:semStore -gt 0)      { Write-Host ("  {0} store(s) que nao consegui abrir" -f $script:semStore) }
    Write-Host "  Os zeros da matriz sao sobre o que foi lido, e nao sobre a caixa." -ForegroundColor DarkYellow
}
Write-Host "itens: $($itens.Count)"
if ($pastasComErro.Count -gt 0) {
    Write-Host "pastas com erro: $($pastasComErro.Count)"
    $pastasComErro | Select-Object -First 5 | ForEach-Object { Write-Host "   $_" }
}
Write-Host ""

Write-Host "TIPOS devolvidos por GetArray (so conta presenca no tipo certo):"
$tipos.Keys | Sort-Object | ForEach-Object { Write-Host ("   {0,-30} {1}" -f $_, $tipos[$_]) }
Write-Host ""

Write-Host ("{0,-12} | {1,8} | {2,5} | {3,9} | {4,6} | {5}" -f `
    "chave", "presente", "%", "distintos", "grupos", "maior")
Write-Host ("-" * 72)

$total = $itens.Count
foreach ($c in @("SearchKey","MessageID","RecordKey","SourceKey","ChangeKey","EntryID")) {
    $com = @($itens | Where-Object { $null -ne $_.$c })
    $todosGrupos = @($com | Group-Object $c)
    $g = @($todosGrupos | Where-Object { $_.Count -gt 1 })
    $maior = 0
    foreach ($grupo in $g) { if ($grupo.Count -gt $maior) { $maior = $grupo.Count } }
    Write-Host ("{0,-12} | {1,8} | {2,4}% | {3,9} | {4,6} | {5}" -f `
        $c, $com.Count, [int](100*$com.Count/[Math]::Max($total,1)),
        $todosGrupos.Count, $g.Count, $maior)
}

Write-Host ""
Write-Host "AUSENCIA de Message-ID, por pasta e classe:"
$itens | Where-Object { $null -eq $_.MessageID } | Group-Object Pasta, Classe |
    Sort-Object Count -Descending | Select-Object -First 12 | ForEach-Object {
        Write-Host ("   {0,5}x  {1}" -f $_.Count, $_.Name)
    }

Write-Host ""
Write-Host "MATRIZ DE COLISAO"
Write-Host "grupos de SearchKey repetida -> o Message-ID dentro bate?"
$gsk = @($itens | Where-Object { $null -ne $_.SearchKey } | Group-Object SearchKey |
        Where-Object { $_.Count -gt 1 })
$mesmoMid = 0; $midDif = 0; $midFalta = 0
# -ExpandProperty DESCARTA $null DO PIPELINE, e isso mentia na matriz.
#
# Um grupo com um Message-ID presente e outro AUSENTE saia do ExpandProperty
# com um unico valor -- e caia em "igual" em vez de "falta". O numero que
# decide o desenho da correlacao estava contando ausencia como acordo. O laco
# abaixo le a propriedade item a item, que preserva o $null.
foreach ($g in $gsk) {
    $mids = @()
    foreach ($it in $g.Group) { $mids += ,$it.MessageID }
    if ($mids -contains $null) { $midFalta++ }
    elseif (@($mids | Sort-Object -Unique).Count -eq 1) { $mesmoMid++ }
    else { $midDif++ }
}
Write-Host ("   grupos                     : {0}" -f $gsk.Count)
Write-Host ("      Message-ID igual        : {0}" -f $mesmoMid)
Write-Host ("      Message-ID DIFERENTE    : {0}  <- SearchKey unindo msg distintas" -f $midDif)
Write-Host ("      Message-ID ausente      : {0}" -f $midFalta)

Write-Host ""
Write-Host "grupos de Message-ID repetido -> a SearchKey dentro bate?"
$gmid = @($itens | Where-Object { $null -ne $_.MessageID } | Group-Object MessageID |
         Where-Object { $_.Count -gt 1 })
$skIgual = 0; $skDif = 0; $skFalta = 0
foreach ($g in $gmid) {
    # Mesmo motivo do bloco acima: o ExpandProperty engolia o $null e um
    # grupo com uma SearchKey ausente virava "igual".
    $sks = @()
    foreach ($it in $g.Group) { $sks += ,$it.SearchKey }
    if ($sks -contains $null) { $skFalta++ }
    elseif (@($sks | Sort-Object -Unique).Count -eq 1) { $skIgual++ }
    else { $skDif++ }
}
Write-Host ("   grupos                     : {0}" -f $gmid.Count)
Write-Host ("      SearchKey igual         : {0}" -f $skIgual)
Write-Host ("      SearchKey DIFERENTE     : {0}  <- SearchKey SEPARA o que o MID une" -f $skDif)
Write-Host ("      SearchKey ausente       : {0}  <- nao e acordo nem divergencia" -f $skFalta)

$saida = Join-Path $env:TEMP "q2-corpus.csv"
$itens | Export-Csv -Path $saida -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "corpus salvo em $saida (fora do repositorio)"
