<#
.SYNOPSIS
    Mede o ALCANCE do calendario local, e compara com o do correio.

.DESCRIPTION
    SOMENTE LEITURA. Nao cria, nao move, nao apaga, nao responde convite.

    POR QUE ESTE ROTEIRO EXISTE

    Em 28/08/2026 a Fase 6 entregou a leitura do calendario e a agenda na
    janela. A agenda diz quantos compromissos leu e se a leitura truncou --
    e NAO diz nada sobre o que existe alem do que o Outlook expoe.

    Isso ficou registrado como divida com esse nome: "a cobertura do
    calendario nunca foi medida". Este roteiro paga a divida.

    A PERGUNTA, E POR QUE ELA NAO E OBVIA

    O medir-janela.ps1 mediu que as pastas de CORREIO tem horizonte comum de
    ~31 dias. A pergunta aqui e outra: o calendario tem o mesmo horizonte?

    Se tiver, a janela de sincronizacao vale para tudo, e a agenda herda a
    mesma ressalva do acervo -- "pode faltar, e o Iris nao conclui ausencia".

    Se NAO tiver -- se o calendario for anos e o correio for um mes --,
    entao a janela e do CORREIO, e a agenda pode dizer bem mais sobre si do
    que diz hoje. As duas respostas mudam o que a tela deve afirmar, e
    nenhuma delas era conhecida.

    O QUE ELE MEDE

    1. O span do calendario: primeiro e ultimo compromisso, sem expansao de
       serie. Mestres so, que e o que existe como item.
    2. A distribuicao por ano, para ver se ha corte ou se e continuo.
    3. Quantos compromissos ha ANTES do horizonte do correio. Se houver
       muitos, esta provado que a janela nao alcanca o calendario.

    O QUE ELE NAO MEDE

    Quantos compromissos existem no SERVIDOR. Isso continua inalcancavel
    pelo OOM, exatamente como no correio -- e por isso nem este roteiro nem
    a agenda podem concluir ausencia.

.NOTES
    ASCII puro e sem BOM, como os outros roteiros deste diretorio.
#>
[CmdletBinding()]
param(
    # O horizonte do correio, para a comparacao. O padrao e o medido em
    # 28/08/2026 pelo medir-janela.ps1.
    [string]$HorizonteDoCorreio = '2026-07-28'
)

$ErrorActionPreference = 'Stop'

Write-Host "Medindo o alcance do calendario local (somente leitura)..." -ForegroundColor Cyan
Write-Host ""

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
    Write-Host "O Outlook nao esta aberto. Abra-o e rode de novo." -ForegroundColor Red
    exit 1
}
$ns = $ol.GetNamespace('MAPI')

$pasta = $null
$itens = $null
try {
    $pasta = $ns.GetDefaultFolder(9)   # olFolderCalendar
    $itens = $pasta.Items

    # SEM IncludeRecurrences, de proposito.
    #
    # A pergunta e "o que EXISTE localmente", e ocorrencia de serie nao
    # existe como item: ela e calculada. Ligar a expansao aqui daria um
    # numero inflado que nao corresponde a nada guardado.
    $itens.Sort("[Start]", $false)
    $n = $itens.Count

    Write-Host ("pasta:  {0}" -f $pasta.Name)
    Write-Host ("itens:  {0}   (mestres, sem expansao de serie)" -f $n)
    Write-Host ""

    if ($n -eq 0) {
        Write-Host "Calendario vazio. Nada a medir." -ForegroundColor Yellow
        exit 0
    }

    $porAno = @{}
    $antesDoHorizonte = 0
    $limite = [datetime]::Parse($HorizonteDoCorreio)
    $maisAntigo = $null
    $maisNovo = $null
    $lidos = 0
    $recusados = 0

    for ($i = 1; $i -le $n; $i++) {
        $item = $null
        try {
            $item = $itens.Item($i)
            $inicio = $null
            try { $inicio = [datetime]$item.Start } catch { }
            if ($null -eq $inicio) { $recusados++; continue }

            $lidos++
            if ($null -eq $maisAntigo -or $inicio -lt $maisAntigo) { $maisAntigo = $inicio }
            if ($null -eq $maisNovo   -or $inicio -gt $maisNovo)   { $maisNovo = $inicio }

            $ano = $inicio.Year
            if (-not $porAno.ContainsKey($ano)) { $porAno[$ano] = 0 }
            $porAno[$ano]++

            if ($inicio -lt $limite) { $antesDoHorizonte++ }
        } catch {
            $recusados++
        } finally {
            if ($item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
        }
    }

    Write-Host ("mais antigo:  {0:yyyy-MM-dd}" -f $maisAntigo)
    Write-Host ("mais novo:    {0:yyyy-MM-dd}" -f $maisNovo)
    Write-Host ("span:         {0} dias" -f [int]($maisNovo - $maisAntigo).TotalDays)
    if ($recusados -gt 0) {
        Write-Host ("recusados:    {0} (nao consegui ler a data)" -f $recusados) -ForegroundColor DarkYellow
    }
    Write-Host ""

    Write-Host "POR ANO:" -ForegroundColor Cyan
    foreach ($ano in ($porAno.Keys | Sort-Object)) {
        $q = $porAno[$ano]
        Write-Host ("  {0}  {1,5}  {2}" -f $ano, $q, ('#' * [Math]::Min(50, [int]($q / 5))))
    }

    Write-Host ""
    Write-Host ("COMPARACAO COM O CORREIO (horizonte {0}):" -f $HorizonteDoCorreio) -ForegroundColor Cyan
    Write-Host ("  compromissos ANTERIORES ao horizonte do correio: {0} de {1}" -f `
        $antesDoHorizonte, $lidos)

    Write-Host ""
    if ($antesDoHorizonte -eq 0) {
        Write-Host "MESMO HORIZONTE. O calendario tambem corta onde o correio corta." -ForegroundColor Yellow
        Write-Host "  A janela de sincronizacao alcanca os dois, e a agenda herda a"
        Write-Host "  mesma ressalva do acervo: pode faltar, e nao se conclui ausencia."
    } elseif ($antesDoHorizonte -lt ($lidos / 20)) {
        Write-Host "QUASE O MESMO HORIZONTE, com poucos itens antes." -ForegroundColor Yellow
        Write-Host "  Poucos itens antigos podem ser excecao, e nao alcance. Olhe a"
        Write-Host "  distribuicao por ano antes de concluir."
    } else {
        Write-Host "HORIZONTES DIFERENTES." -ForegroundColor Green
        Write-Host ("  Ha {0} compromissos anteriores ao corte do correio. A janela de" -f $antesDoHorizonte)
        Write-Host "  sincronizacao NAO alcanca o calendario da mesma forma -- ele guarda"
        Write-Host "  muito mais historico localmente que o correio."
        Write-Host ""
        Write-Host "  CONSEQUENCIA: a agenda pode afirmar mais sobre si do que o acervo"
        Write-Host "  afirma. O que ela continua NAO podendo e concluir ausencia, porque"
        Write-Host "  a contagem do servidor segue inalcancavel pelo OOM."
    }

    Write-Host ""
    Write-Host "O QUE ISTO NAO DIZ:" -ForegroundColor DarkGray
    Write-Host "  Quantos compromissos existem no servidor. Inalcancavel pelo OOM," -ForegroundColor DarkGray
    Write-Host "  igual ao correio -- e por isso nem este roteiro nem a agenda" -ForegroundColor DarkGray
    Write-Host "  concluem ausencia." -ForegroundColor DarkGray
    Write-Host "  Tambem nao diz nada sobre calendarios secundarios ou compartilhados:" -ForegroundColor DarkGray
    Write-Host "  mede a pasta padrao desta conta." -ForegroundColor DarkGray

} catch {
    Write-Host ("nao consegui ler o calendario: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    # Ordem inversa a aquisicao, R7.
    if ($itens) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens) }
    if ($pasta) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta) }
    if ($ns)    { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns) }
    if ($ol)    { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol) }
}
