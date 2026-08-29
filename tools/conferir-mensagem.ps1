<#
.SYNOPSIS
    Diz se o Iris vai ACEITAR uma mensagem, antes de voce clicar no botao.

.DESCRIPTION
    O canario de 26/08/2026 parou tres vezes seguidas, e cada parada custou
    uma ida e volta: a pasta nao era a autorizada, depois o anexo da
    assinatura, depois a referencia cid: do mesmo anexo. Cada uma dessas
    recusas e correta -- e nenhuma delas era visivel antes da tentativa.

    Este roteiro le o Outlook SOMENTE PARA LER e aplica as mesmas regras do
    ContentPipeline do Iris: anexo, cid:/data:, corpo vazio, campo longo
    demais. Nada e enviado, nada e modificado, e o corpo NAO e impresso --
    so o veredito e o motivo.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File "C:\Users\Ricardo\Documents\Iris\tools\conferir-mensagem.ps1" -Pasta "Iris-Teste"
#>
[CmdletBinding()]
param(
    # A pasta onde procurar. O nome, e nao o caminho todo.
    [Parameter(Mandatory = $true)]
    [string] $Pasta,

    # Confere so as mensagens cujo assunto contenha isto. Sem o filtro,
    # confere todas as da pasta.
    [string] $Assunto = ""
)

$ErrorActionPreference = 'Stop'

# Os mesmos limites do ContentPipeline. Se eles mudarem la e nao aqui, este
# roteiro passa a mentir -- entao estao com o nome do lugar de onde vieram.
$MaxAssunto    = 300
$MaxRemetente  = 200
$MaxTexto      = 20000

# cid: e data:<tipo>/ -- a mesma expressao do ContentPipeline.Embutido.
$Embutido = [regex]::new('(cid:|data:[a-z]+/)', 'IgnoreCase')

# RAMO QUE NAO FOI VISTO -- mesma correcao do preparar-ativacao.
$script:ramosCegos = 0

function Achar-Pasta($raiz, $nome, $caminho) {
    $achadas = @()
    $filhas = $null
    try {
        $filhas = $raiz.Folders
        for ($i = 1; $i -le $filhas.Count; $i++) {
            $f = $null
            try {
                $f = $filhas.Item($i)
                $meu = "$caminho\$($f.Name)"
                if ($f.Name -eq $nome) {
                    $achadas += [pscustomobject]@{ Pasta = $f; Caminho = $meu }
                } else {
                    $achadas += Achar-Pasta $f $nome $meu
                }
            } catch {
                # UM RAMO INTEIRO nao foi visto, e isto era silencioso.
                $script:ramosCegos++
            }
        }
    } finally {
        if ($filhas) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($filhas) }
    }
    return $achadas
}

Write-Host ""
Write-Host "Abrindo o Outlook somente para LER..." -ForegroundColor Cyan

$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")

$alvos = @()
for ($s = 1; $s -le $ns.Folders.Count; $s++) {
    $raiz = $ns.Folders.Item($s)
    $alvos += Achar-Pasta $raiz $Pasta $raiz.Name
}

if ($alvos.Count -eq 0) {
    if ($script:ramosCegos -gt 0) {
        Write-Host ("Nao achei pasta chamada '{0}' -- e {1} ramo(s) da arvore nao" -f $Pasta, $script:ramosCegos) -ForegroundColor Red
        Write-Host "  pude percorrer. Ela pode estar num deles." -ForegroundColor Red
    } else {
        Write-Host "Nao achei pasta chamada '$Pasta'." -ForegroundColor Red
    }
    exit 1
}

$alvo = $alvos[0]
Write-Host "Pasta: $($alvo.Caminho)" -ForegroundColor Green
Write-Host ""

$itens = $alvo.Pasta.Items
$total = $itens.Count
$conferidas = 0
$aprovadas = 0

for ($i = 1; $i -le $total; $i++) {
    $m = $null
    try {
        $m = $itens.Item($i)
        $tema = ""
        try { $tema = [string]$m.Subject } catch { }

        if ($Assunto -and ($tema -notlike "*$Assunto*")) { continue }
        $conferidas++

        Write-Host ("-" * 68)
        Write-Host $tema

        $problemas = @()

        # 1. ANEXO. O ContentPipeline recusa QUALQUER mensagem com anexo --
        #    inclusive imagem de assinatura, que e o caso comum e o que mais
        #    surpreende.
        $nAnexos = 0
        $nomes = @()
        try {
            $anexos = $m.Attachments
            $nAnexos = $anexos.Count
            for ($a = 1; $a -le $nAnexos; $a++) {
                $nomes += [string]$anexos.Item($a).FileName
            }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($anexos)
        } catch { }

        if ($nAnexos -gt 0) {
            $problemas += "tem $nAnexos anexo(s): $($nomes -join ', ')"
        }

        # 2. FORMATO. 1 = texto sem formatacao, 2 = RTF, 3 = HTML.
        $formato = 0
        try { $formato = [int]$m.BodyFormat } catch { }
        $nomeFormato = switch ($formato) {
            1 { "texto sem formatacao" }
            2 { "RTF" }
            3 { "HTML" }
            default { "desconhecido ($formato)" }
        }

        # 3. REFERENCIA EMBUTIDA, no corpo e nos campos.
        $corpo = ""
        try { if ($formato -eq 3) { $corpo = [string]$m.HTMLBody } else { $corpo = [string]$m.Body } } catch { }

        $de = ""
        try { $de = [string]$m.SenderEmailAddress } catch { }
        if (-not $de) { try { $de = [string]$m.SentOnBehalfOfName } catch { } }

        if ($Embutido.IsMatch($corpo)) { $problemas += "o corpo tem cid: ou data:" }
        if ($Embutido.IsMatch($tema))  { $problemas += "o assunto tem cid: ou data:" }
        if ($Embutido.IsMatch($de))    { $problemas += "o remetente tem cid: ou data:" }

        # 4. TEXTO, e os limites.
        $texto = ""
        try { $texto = [string]$m.Body } catch { }
        $texto = $texto.Trim()
        if ($texto.Length -eq 0)          { $problemas += "corpo vazio" }
        if ($tema.Length -gt $MaxAssunto) { $problemas += "assunto passa de $MaxAssunto" }
        if ($de.Length -gt $MaxRemetente) { $problemas += "remetente passa de $MaxRemetente" }
        if ($texto.Length -gt $MaxTexto)  { $problemas += "corpo passa de $MaxTexto" }

        Write-Host ("  formato .... {0}" -f $nomeFormato)
        Write-Host ("  anexos ..... {0}" -f $nAnexos)
        Write-Host ("  corpo ...... {0} caracteres" -f $texto.Length)

        if ($problemas.Count -eq 0) {
            $aprovadas++
            Write-Host "  O IRIS ACEITA." -ForegroundColor Green
        } else {
            Write-Host "  O IRIS RECUSA:" -ForegroundColor Red
            foreach ($p in $problemas) { Write-Host "    * $p" -ForegroundColor Yellow }
        }
    } catch {
        Write-Host ("  nao consegui ler este item: {0}" -f $_.Exception.Message) -ForegroundColor Red
    } finally {
        if ($m) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($m) }
    }
}

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens)

Write-Host ("-" * 68)
Write-Host ""
if ($conferidas -eq 0) {
    Write-Host "Nenhuma mensagem CONFERIDA (a pasta expoe $total item(ns) localmente)." -ForegroundColor Yellow
    Write-Host "  Zero conferidas nao diz nada sobre o conteudo da pasta: diz que este"
    Write-Host "  roteiro nao chegou a conferir nada."
} else {
    Write-Host ("$aprovadas de $conferidas passariam pelo ContentPipeline.")
}
Write-Host ""
Write-Host "Isto NAO confere a pasta autorizada nem a ativacao -- so o conteudo." -ForegroundColor Cyan
Write-Host "Para a pasta, o que manda e o EntryID gravado em ativacao.json." -ForegroundColor Cyan
