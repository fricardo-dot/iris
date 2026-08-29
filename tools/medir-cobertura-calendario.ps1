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
    entao o corte observado no correio nao esta sendo aplicado a este
    calendario, e a agenda pode dizer mais sobre si do que diz hoje.

    O QUE A RESPOSTA NAO AUTORIZA A CONCLUIR

    Que o calendario esteja INTEIRO localmente. Ausencia de corte nao e
    presenca de tudo: pode haver outro corte mais antigo que o roteiro nao
    alcanca, e a contagem do servidor continua inalcancavel pelo OOM.

    E a comparacao com o correio e por DATA, nao por store: este roteiro le
    a pasta padrao de calendario, e o medir-janela.ps1 percorre todos os
    stores sem separar por StoreID. Dizer "o mesmo store" seria afirmar uma
    correlacao que nenhum dos dois instrumentos estabelece.

    O QUE ELE MEDE

    1. O span do calendario: primeiro e ultimo compromisso, sem expansao de
       serie. Mestres so, que e o que existe como item.
    2. A distribuicao por ano, para ver se ha corte ou se e continuo.
    3. Quantos compromissos ha ANTES do horizonte do correio. Muitos deles
       mostram que o corte de ~31 dias NAO APARECE neste calendario -- e so
       isso. Nao prova "a janela nao alcanca o calendario": este roteiro nao
       correlaciona StoreID com o do medir-janela, entao a comparacao e por
       data e nao por caixa, e ele nao procura corte mais antigo que o item
       mais velho que achou.

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
    # O StoreID sai impresso porque a comparacao com o correio depende dele,
    # e o medir-janela nao o imprime. Sem os dois lados, "mesmo store" e
    # suposicao.
    try { Write-Host ("store:  {0}..." -f $pasta.StoreID.Substring(0, 32)) } catch { }
    Write-Host ("itens:  {0}   (mestres, sem expansao de serie)" -f $n)
    Write-Host ""

    if ($n -eq 0) {
        # "VAZIO" SERIA AFIRMACAO DE AUSENCIA, e este roteiro nao pode fazer
        # essa -- pelo mesmo motivo que ele repete no cabecalho: a contagem do
        # servidor e inalcancavel pelo OOM. Zero item EXPOSTO LOCALMENTE nao e
        # zero compromisso.
        Write-Host "NENHUM COMPROMISSO EXPOSTO LOCALMENTE nesta pasta." -ForegroundColor Yellow
        Write-Host "  Isso NAO quer dizer que o calendario esteja vazio: a contagem do"
        Write-Host "  servidor continua inalcancavel pelo OOM. Nao ha o que medir aqui,"
        Write-Host "  e tambem nao ha o que concluir."
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
        # O SEGUNDO ZERO DESTE ROTEIRO, e ele escapou da correcao do primeiro.
        # Item cuja data nao foi lida entra em $recusados e NAO e classificado
        # -- qualquer um deles pode ser anterior ao horizonte. Com recusados,
        # nem "nenhum antes" se sustenta.
        if ($recusados -gt 0) {
            Write-Host ("NENHUM COMPROMISSO LEGIVEL ANTES DO HORIZONTE -- e {0} item(ns) nao" -f $recusados) -ForegroundColor Yellow
            Write-Host "  tiveram a data lida. Qualquer um deles pode ser anterior, entao"
            Write-Host "  nem 'nenhum antes' se sustenta aqui."
        } else {
            Write-Host "NENHUM COMPROMISSO ANTES DO HORIZONTE DO CORREIO." -ForegroundColor Yellow
        }
        Write-Host "  Isso NAO prova que o calendario corta ali: um calendario que so"
        Write-Host "  tenha compromissos recentes da exatamente este retrato. E os dois"
        Write-Host "  roteiros nao correlacionam StoreID, entao a comparacao e por data"
        Write-Host "  e nao por caixa."
        Write-Host "  O que se sustenta: nada aqui contradiz a ressalva do acervo, e"
        Write-Host "  ausencia continua nao sendo concluivel."
    } elseif ($antesDoHorizonte -lt ($lidos / 20)) {
        Write-Host "QUASE O MESMO HORIZONTE, com poucos itens antes." -ForegroundColor Yellow
        Write-Host "  Poucos itens antigos podem ser excecao, e nao alcance. Olhe a"
        Write-Host "  distribuicao por ano antes de concluir."
    } else {
        Write-Host "HORIZONTES DIFERENTES." -ForegroundColor Green
        Write-Host ("  Ha {0} compromissos anteriores ao corte do correio." -f $antesDoHorizonte)
        Write-Host ""
        Write-Host "  O QUE ISTO SUSTENTA, na formulacao mais estreita que serve:"
        Write-Host "  o corte de ~31 dias observado nas pastas de correio NAO aparece"
        Write-Host "  neste calendario padrao local."
        Write-Host ""
        Write-Host "  O QUE NAO SUSTENTA: que o calendario esteja inteiro; que nao"
        Write-Host "  exista outro corte mais antigo; que nenhuma politica de cache"
        Write-Host "  alcance calendario; nem a comparacao 'mesmo store', porque este"
        Write-Host "  roteiro le a pasta padrao e o medir-janela percorre todos os"
        Write-Host "  stores sem separar por StoreID."
        Write-Host ""
        Write-Host "  CONSEQUENCIA para a tela: repetir na agenda a ressalva de janela"
        Write-Host "  do acervo seria ressalva emprestada. O que ela continua NAO"
        Write-Host "  podendo e concluir ausencia -- por falta de prova, e nao por"
        Write-Host "  janela."
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
