<#
.SYNOPSIS
    Conta quantas linhas da Table devolvem NULO nas colunas que o Iris le.

.DESCRIPTION
    SOMENTE LEITURA.

    POR QUE ESTE ROTEIRO EXISTE

    A revisao externa de 28/08/2026 apontou que o caminho de paginacao
    transforma AUSENCIA em FATO: quando a conversao falha, Size vira 0,
    UnRead vira False, anexo vira False e texto vira vazio. E a mesma
    familia do defeito do message_class, que era constante fabricada.

    Corrigir isso e mexer em DTO, esquema e tela -- campos anulaveis,
    migracao, e uma forma de a UI mostrar "nao sei" sem parecer "nao".
    Decisao de tamanho.

    ENTAO MEDE PRIMEIRO. Se a Table nunca devolve nulo nestas colunas, o
    defeito e teorico e a migracao seria trabalho grande por risco que nao
    se materializa. Se devolve, e real e o tamanho se justifica.

    E foi assim que o message_class foi achado: 1.123 linhas com UMA classe
    distinta era limpo demais para ser medida.

    O QUE ELE MEDE

    Le a mesma Table que o Iris le, com as mesmas colunas, e conta por
    coluna quantas celulas voltam NULO ou nao convertem. Nao interpreta:
    imprime a contagem.

.NOTES
    ASCII puro e sem BOM.
#>
[CmdletBinding()]
param(
    # Pasta a medir, pelo nome. Vazio = Caixa de Entrada padrao.
    [string]$Pasta = ''
)

$ErrorActionPreference = 'Stop'

# As MESMAS colunas do MessagePaging.vb. Copiar a lista e o ponto: medir
# outras colunas nao diria nada sobre o que o Iris le.
$Colunas = @(
    @{ Nome = 'EntryID';       Dasl = 'http://schemas.microsoft.com/mapi/proptag/0x66700102' },
    @{ Nome = 'Subject';       Dasl = 'Subject' },
    @{ Nome = 'SenderName';    Dasl = 'SenderName' },
    @{ Nome = 'ReceivedTime';  Dasl = 'ReceivedTime' },
    @{ Nome = 'Size';          Dasl = 'Size' },
    @{ Nome = 'UnRead';        Dasl = 'UnRead' },
    @{ Nome = 'MessageClass';  Dasl = 'MessageClass' },
    @{ Nome = 'HasAttachment'; Dasl = 'http://schemas.microsoft.com/mapi/proptag/0x0E1B000B' }
)

Write-Host "Medindo nulos nas colunas que o Iris le (somente leitura)..." -ForegroundColor Cyan
Write-Host ""

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
    Write-Host "O Outlook nao esta aberto." -ForegroundColor Red
    exit 1
}
$ns = $ol.GetNamespace('MAPI')

$pastaAlvo = $null
$table = $null
$colecao = $null
try {
    if ($Pasta) {
        # Busca simples por nome no primeiro nivel de cada store.
        foreach ($raiz in $ns.Folders) {
            foreach ($f in $raiz.Folders) {
                if ($f.Name -eq $Pasta) { $pastaAlvo = $f; break }
            }
            if ($pastaAlvo) { break }
        }
        if (-not $pastaAlvo) { Write-Host "Nao achei a pasta '$Pasta'." -ForegroundColor Red; exit 1 }
    } else {
        $pastaAlvo = $ns.GetDefaultFolder(6)   # olFolderInbox
    }

    Write-Host ("pasta:  {0}" -f $pastaAlvo.Name)

    $table = $pastaAlvo.GetTable()
    $colecao = $table.Columns
    $colecao.RemoveAll()
    foreach ($c in $Colunas) {
        try { [void]$colecao.Add($c.Dasl) }
        catch { Write-Host ("  coluna RECUSADA: {0}" -f $c.Nome) -ForegroundColor DarkYellow }
    }

    $nulos = @{}
    $vazios = @{}
    foreach ($c in $Colunas) { $nulos[$c.Nome] = 0; $vazios[$c.Nome] = 0 }
    $linhas = 0

    while (-not $table.EndOfTable) {
        $linha = $table.GetNextRow()
        $linhas++
        for ($i = 0; $i -lt $Colunas.Count; $i++) {
            $nome = $Colunas[$i].Nome
            $v = $null
            try { $v = $linha.Item($i + 1) } catch { }

            if ($null -eq $v -or $v -is [DBNull]) {
                $nulos[$nome]++
            } elseif ($v -is [string] -and $v.Length -eq 0) {
                $vazios[$nome]++
            }
        }
        if ($linha) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($linha) }
    }

    Write-Host ("linhas: {0}" -f $linhas)
    Write-Host ""
    Write-Host ("{0,-16} {1,8} {2,8}" -f 'COLUNA', 'NULOS', 'VAZIOS')
    Write-Host ("{0,-16} {1,8} {2,8}" -f '------', '-----', '------')
    $achouNulo = $false
    foreach ($c in $Colunas) {
        $n = $nulos[$c.Nome]; $z = $vazios[$c.Nome]
        if ($n -gt 0) { $achouNulo = $true }
        $cor = if ($n -gt 0) { 'Yellow' } else { 'Gray' }
        Write-Host ("{0,-16} {1,8} {2,8}" -f $c.Nome, $n, $z) -ForegroundColor $cor
    }

    Write-Host ""
    if ($achouNulo) {
        Write-Host "HA NULOS. A ausencia existe de verdade nesta caixa, e hoje ela" -ForegroundColor Yellow
        Write-Host "  vira fato: 0, False, ou texto vazio. A correcao se justifica."
    } else {
        Write-Host "NENHUM NULO em nenhuma coluna, nesta pasta, agora." -ForegroundColor Green
        Write-Host "  O defeito continua REAL como contrato -- a conversao ainda"
        Write-Host "  fabrica em caso de falha -- e nao se manifesta nesta amostra."
        Write-Host "  Isso NAO prova que nunca acontece: prova que nao acontece aqui."
    }

    Write-Host ""
    Write-Host "O QUE ISTO NAO DIZ:" -ForegroundColor DarkGray
    Write-Host "  Nada sobre outras pastas, outras contas, ou o mesmo item depois" -ForegroundColor DarkGray
    Write-Host "  de uma mudanca. E nada sobre a conversao FALHAR com valor" -ForegroundColor DarkGray
    Write-Host "  presente -- so conta ausencia." -ForegroundColor DarkGray

} catch {
    Write-Host ("falhou: {0}" -f $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    if ($colecao)   { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($colecao) }
    if ($table)     { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($table) }
    if ($pastaAlvo) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pastaAlvo) }
    if ($ns)        { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns) }
    if ($ol)        { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol) }
}
