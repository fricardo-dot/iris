<#
.SYNOPSIS
    Mede o EFEITO da janela de sincronizacao do Exchange em cache.

.DESCRIPTION
    SOMENTE LEITURA. Nao cria, nao move, nao apaga, nao envia. Nao abre corpo
    de mensagem nem anexo -- le tres propriedades por pasta e duas datas.

    POR QUE ESTE ROTEIRO EXISTE

    A janela de sincronizacao nao e legivel: nem pelo objeto Store, nem pelos
    294 valores do registro do perfil, nem pelas 65.536 tags MAPI que a Fase 2
    varreu. O ESCOPO.md registrou a saida: "a saida nao e achar a configuracao;
    e medir o EFEITO dela".

    O efeito e um horizonte. Se a janela for de N meses, nenhuma pasta tera
    item local mais antigo que hoje menos N -- e o corte aparece em TODAS as
    pastas ao mesmo tempo, porque a janela e do store e nao da pasta.

    O QUE UMA EXECUCAO NAO PROVA

    Uma execucao e um RETRATO. Ela mostra que existe um horizonte comum hoje;
    nao mostra que ele anda com o tempo, que e o que "janela deslizante"
    afirma. Migracao de caixa, recriacao do OST, retencao corporativa ou uma
    operacao em massa naquela data produziriam o mesmo retrato.

    Duas execucoes separadas por alguns dias distinguem: se o corte anda
    junto, e deslizante; se fica parado, foi um evento.

    O QUE DISTINGUE JANELA DE HABITO DO USUARIO

    Uma pasta so, cortando em 31 dias, nao prova nada: pode ser so uma caixa
    onde o dono arquiva tudo todo mes. O sinal de janela e a COINCIDENCIA --
    muitas pastas, com volumes e usos diferentes, cortando no MESMO dia.

    Este roteiro imprime o item mais antigo de cada pasta justamente para que
    essa coincidencia seja visivel ou desminta a hipotese.

.NOTES
    ASCII puro e sem BOM, de proposito: o PowerShell 5.1 do ambiente le
    arquivo com BOM de forma inconsistente quando o roteiro e chamado por
    -File.
#>
[CmdletBinding()]
param(
    # Pastas com menos itens que isto nao entram na conta: um horizonte
    # medido sobre tres mensagens nao e medida, e ainda por cima poluiria a
    # busca pela coincidencia.
    [int]$MinimoDeItens = 20,

    # Salva o resultado como CSV, para quem quiser cruzar depois.
    [string]$Csv
)

$ErrorActionPreference = 'Stop'

Write-Host "Medindo o efeito da janela de sincronizacao (somente leitura)..." -ForegroundColor Cyan
Write-Host ""

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
    Write-Host "O Outlook nao esta aberto. Abra-o e rode de novo." -ForegroundColor Red
    exit 1
}

$ns = $ol.GetNamespace('MAPI')

$linhas = New-Object System.Collections.ArrayList

# Percorre a arvore. Nao encadeia expressao COM: cada colecao intermediaria
# recebe nome, pela R7 do ESCOPO.
function Percorrer($pastas, $trilha) {
    foreach ($f in $pastas) {
        $caminho = if ($trilha) { "$trilha\$($f.Name)" } else { $f.Name }

        $ehCorreio = $false
        try { $ehCorreio = ($f.DefaultItemType -eq 0) } catch { }

        if ($ehCorreio) {
            $itens = $null
            try {
                $itens = $f.Items
                $n = $itens.Count
                if ($n -ge $MinimoDeItens) {
                    # Sort do proprio Outlook: a Fase 0 mediu 2-5 ms em 770
                    # itens, contra ~16 ms POR ITEM se montassemos DTO.
                    $itens.Sort("[ReceivedTime]", $false)

                    $primeiro = $null
                    $ultimo = $null
                    $maisAntigo = $null
                    $maisNovo = $null
                    try {
                        $primeiro = $itens.GetFirst()
                        if ($primeiro) { $maisAntigo = $primeiro.ReceivedTime }
                    } catch { }
                    try {
                        $ultimo = $itens.GetLast()
                        if ($ultimo) { $maisNovo = $ultimo.ReceivedTime }
                    } catch { }

                    if ($maisAntigo -and $maisNovo) {
                        $dias = [int]([datetime]$maisNovo - [datetime]$maisAntigo).TotalDays
                        $null = $linhas.Add([pscustomobject]@{
                            Pasta       = $caminho
                            Itens       = $n
                            MaisAntigo  = ([datetime]$maisAntigo).ToString('yyyy-MM-dd')
                            MaisNovo    = ([datetime]$maisNovo).ToString('yyyy-MM-dd')
                            DiasDeSpan  = $dias
                        })
                    }

                    if ($primeiro) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($primeiro) }
                    if ($ultimo)   { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ultimo) }
                }
            } catch {
                Write-Host ("  nao consegui ler '{0}': {1}" -f $caminho, $_.Exception.Message) -ForegroundColor DarkYellow
            } finally {
                if ($itens) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens) }
            }
        }

        $filhas = $null
        try {
            $filhas = $f.Folders
            Percorrer $filhas $caminho
        } catch {
        } finally {
            if ($filhas) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
        }

        # A PROPRIA PASTA. Ela e um RCW como qualquer outro, e nao liberar
        # prende o Outlook vivo ate o host do PowerShell morrer. Foi apontado
        # na revisao de 28/08: o roteiro pregava a R7 no cabecalho e a violava
        # na travessia.
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($f)
    }
}

$raizes = $null
try {
    $raizes = $ns.Folders
    Percorrer $raizes ""
} finally {
    # ORDEM INVERSA A AQUISICAO, que e a R7 escrita no ESCOPO: colecao,
    # depois namespace, depois a aplicacao.
    if ($raizes) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($raizes) }
    if ($ns)     { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns) }
    if ($ol)     { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol) }
}

if ($linhas.Count -eq 0) {
    Write-Host "Nenhuma pasta de correio com $MinimoDeItens itens ou mais." -ForegroundColor Yellow
    exit 0
}

$ordenadas = $linhas | Sort-Object MaisAntigo
$ordenadas | Format-Table -AutoSize

Write-Host ""
Write-Host "A COINCIDENCIA, que e o que interessa:" -ForegroundColor Cyan

$grupos = $ordenadas | Group-Object MaisAntigo | Sort-Object Count -Descending
foreach ($g in $grupos | Select-Object -First 5) {
    Write-Host ("  {0}  ->  {1} pasta(s)" -f $g.Name, $g.Count)
}

$maior = $grupos | Select-Object -First 1
$total = $ordenadas.Count
Write-Host ""
if ($maior.Count -ge 3 -and $maior.Count -ge ($total / 2)) {
    $h = [datetime]::Parse($maior.Name)
    $dias = [int]([datetime]::Today - $h).TotalDays
    Write-Host ("HORIZONTE COMUM em {0}: {1} de {2} pastas cortam no mesmo dia, {3} dias atras." -f `
        $maior.Name, $maior.Count, $total, $dias) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "O QUE UMA MEDICAO SO SUSTENTA:" -ForegroundColor DarkGray
    Write-Host "  Que HOJE o conjunto exposto pelo OOM tem esse horizonte local." -ForegroundColor DarkGray
    Write-Host "  NAO sustenta que ele e DESLIZANTE, nem que a causa e a janela." -ForegroundColor DarkGray
    Write-Host "  Idade da caixa, migracao, recriacao do cache, retencao corporativa" -ForegroundColor DarkGray
    Write-Host "  ou uma operacao em massa naquela data explicariam o mesmo retrato." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  O QUE SEPARA: rodar de novo daqui a alguns dias. Se o corte andar" -ForegroundColor DarkGray
    Write-Host "  junto, e deslizante; se ficar parado em $($maior.Name), foi um evento." -ForegroundColor DarkGray
} else {
    Write-Host "SEM horizonte comum: as pastas comecam em datas espalhadas." -ForegroundColor Green
    Write-Host "O corte de uma pasta isolada e habito de arquivamento, e nao janela."
}

Write-Host ""
Write-Host "O QUE ISTO NAO DIZ, EM NENHUMA HIPOTESE:" -ForegroundColor DarkGray
Write-Host "  Quantos itens existem no servidor alem do horizonte. Isso continua" -ForegroundColor DarkGray
Write-Host "  inalcancavel pelo OOM, e por isso o Iris nao conclui ausencia." -ForegroundColor DarkGray
Write-Host "  Tambem nao diz nada sobre pastas com menos de $MinimoDeItens itens," -ForegroundColor DarkGray
Write-Host "  nem sobre pastas cuja leitura falhou -- essas foram avisadas acima." -ForegroundColor DarkGray
Write-Host "  E mistura stores, se houver mais de um: a janela e do store, entao" -ForegroundColor DarkGray
Write-Host "  duas caixas com janelas diferentes apareceriam como ruido." -ForegroundColor DarkGray

if ($Csv) {
    $ordenadas | Export-Csv -Path $Csv -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "CSV em $Csv"
}
