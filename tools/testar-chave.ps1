<#
.SYNOPSIS
    Testa a credencial do OpenRouter SEM mandar conteudo de e-mail nenhum.

.DESCRIPTION
    O canario deu "ProvedorRecusou": alguma coisa devolveu erro HTTP. As duas
    hipoteses levam a acoes opostas -- chave invalida (401) ou nenhum endpoint
    atendendo a politica de dados -- e o diario nao guardava o codigo, entao
    nao dava para distinguir.

    Este roteiro le a MESMA credencial que o Iris le, do Gerenciador de
    Credenciais, e chama /api/v1/key -- que so devolve informacao sobre a
    propria chave. Nao ha completion, nao ha modelo, e NENHUM byte de e-mail
    sai daqui.

    A chave nao e impressa em lugar nenhum. So o codigo HTTP e os campos
    nao-secretos da resposta.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File "C:\Users\Ricardo\Documents\Iris\tools\testar-chave.ps1"
#>
[CmdletBinding()]
param(
    [string] $Alvo = "Iris/OpenRouter",

    # Manda uma completion MINIMA com o mesmo bloco `provider` da ativacao.
    #
    # O conteudo e a palavra "ping": nenhum byte de e-mail. Serve para saber se
    # a restricao de roteamento -- zdr, data_collection, only -- tem endpoint
    # que a atenda, sem descobrir isso com uma mensagem de verdade.
    [switch] $Sondar,

    [string] $Ativacao = (Join-Path ([Environment]::GetFolderPath(
        [Environment+SpecialFolder]::CommonApplicationData)) "Iris\ativacao.json")
)

$ErrorActionPreference = 'Stop'

# Le do Gerenciador de Credenciais pelo mesmo caminho do Iris: CredReadW,
# blob em UTF-16. Se este roteiro achar a chave e o Iris nao, o problema e do
# Iris; se nenhum dos dois achar, e do cadastro.
if (-not ("Cred" -as [type])) {
    Add-Type -Namespace Win32 -Name Cred -MemberDefinition @'
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct CREDENTIAL {
    public uint Flags; public uint Type;
    public IntPtr TargetName; public IntPtr Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public uint CredentialBlobSize; public IntPtr CredentialBlob;
    public uint Persist; public uint AttributeCount;
    public IntPtr Attributes; public IntPtr TargetAlias; public IntPtr UserName;
}
[DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "CredReadW")]
public static extern bool CredRead(string target, uint type, uint reserved, out IntPtr cred);
[DllImport("advapi32.dll", EntryPoint = "CredFree")]
public static extern void CredFree(IntPtr buffer);
'@
}

$ponteiro = [IntPtr]::Zero
$chave = $null
try {
    if (-not [Win32.Cred]::CredRead($Alvo, 1, 0, [ref]$ponteiro)) {
        Write-Host ""
        Write-Host "Nao achei credencial no alvo '$Alvo'." -ForegroundColor Red
        Write-Host "Cadastre com:  cmdkey /generic:$Alvo /user:iris /pass" -ForegroundColor Yellow
        exit 1
    }
    $c = [System.Runtime.InteropServices.Marshal]::PtrToStructure(
        $ponteiro, [type][Win32.Cred+CREDENTIAL])
    $chave = [System.Runtime.InteropServices.Marshal]::PtrToStringUni(
        $c.CredentialBlob, [int]($c.CredentialBlobSize / 2)).TrimEnd([char]0).Trim()
} finally {
    if ($ponteiro -ne [IntPtr]::Zero) { [Win32.Cred]::CredFree($ponteiro) }
}

Write-Host ""
Write-Host "Credencial encontrada no alvo '$Alvo'."
# COMPRIMENTO E PREFIXO, e nunca a chave. O prefixo distingue "colei a chave
# certa" de "colei outra coisa" sem revelar nada util a quem ler a tela.
Write-Host ("  comprimento .. {0} caracteres" -f $chave.Length)
Write-Host ("  comeca com ... {0}..." -f $chave.Substring(0, [Math]::Min(7, $chave.Length)))
Write-Host ""
Write-Host "Perguntando ao OpenRouter sobre a propria chave (sem conteudo)..." -ForegroundColor Cyan

try {
    $r = Invoke-WebRequest -Uri "https://openrouter.ai/api/v1/key" `
                           -Headers @{ Authorization = "Bearer $chave" } `
                           -Method Get -UseBasicParsing -TimeoutSec 30
    Write-Host ""
    Write-Host ("HTTP {0} -- a chave VALE." -f [int]$r.StatusCode) -ForegroundColor Green
    Write-Host ""
    Write-Host "O que o OpenRouter diz dela (util para o limite que o Codex pediu):"
    Write-Host $r.Content
    Write-Host ""
    Write-Host "Entao o erro do canario NAO foi a chave. A hipotese que sobra e a" -ForegroundColor Yellow
    Write-Host "restricao de roteamento: zdr + only=google pode nao ter endpoint." -ForegroundColor Yellow
} catch {
    $resp = $_.Exception.Response
    Write-Host ""
    if ($resp) {
        $codigo = [int]$resp.StatusCode
        Write-Host ("HTTP {0}" -f $codigo) -ForegroundColor Red
        if ($codigo -eq 401) {
            Write-Host "A CHAVE nao vale. Recadastre com cmdkey e tente de novo." -ForegroundColor Yellow
        } elseif ($codigo -eq 407) {
            Write-Host "Proxy pedindo autenticacao: ha um intermediario na rede." -ForegroundColor Yellow
        }
        try {
            $leitor = New-Object IO.StreamReader($resp.GetResponseStream())
            Write-Host $leitor.ReadToEnd()
        } catch { }
    } else {
        Write-Host ("Nem chegou a haver resposta HTTP: {0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-Host "Isso aponta para rede, DNS ou TLS -- e nao para a chave." -ForegroundColor Yellow
    }
    # Sem chave que valha, sondar roteamento nao diz nada.
    exit 1
}

# ---------------------------------------------------------------------------
if (-not $Sondar) {
    Write-Host ""
    Write-Host "Para sondar tambem a restricao de roteamento (sem conteudo de" -ForegroundColor Cyan
    Write-Host "e-mail), repita com -Sondar." -ForegroundColor Cyan
    exit 0
}

if (-not (Test-Path -LiteralPath $Ativacao)) {
    Write-Host ""
    Write-Host "Nao achei a ativacao em: $Ativacao" -ForegroundColor Red
    exit 1
}

$a = Get-Content -Raw -LiteralPath $Ativacao | ConvertFrom-Json

Write-Host ""
Write-Host "Sondando o roteamento com o MESMO bloco provider da ativacao..." -ForegroundColor Cyan
Write-Host ("  modelo ..... {0}" -f $a.modelo)
Write-Host ("  only ....... {0}" -f ($a.provedoresPermitidos -join ', '))
Write-Host ("  zdr ........ {0}" -f $a.exigirRetencaoZero)
Write-Host ""

# O MESMO bloco que o OpenRouterAssistantProvider monta -- e o conteudo e
# "ping". Se isto passar e o Iris falhar, o problema nao e o roteamento.
$provider = @{ allow_fallbacks = $false }
if ($a.exigirRetencaoZero) {
    $provider.zdr = $true
    $provider.data_collection = "deny"
}
if ($a.provedoresPermitidos.Count -gt 0) { $provider.only = @($a.provedoresPermitidos) }

$corpo = @{
    model       = $a.modelo
    messages    = @(@{ role = "user"; content = "ping" })
    temperature = 0
    provider    = $provider
} | ConvertTo-Json -Depth 6

try {
    $r = Invoke-WebRequest -Uri "https://openrouter.ai/api/v1/chat/completions" `
                           -Headers @{ Authorization = "Bearer $chave" } `
                           -Method Post -ContentType "application/json" `
                           -Body ([Text.Encoding]::UTF8.GetBytes($corpo)) `
                           -UseBasicParsing -TimeoutSec 60
    Write-Host ("HTTP {0} -- o roteamento ACEITA a restricao." -f [int]$r.StatusCode) -ForegroundColor Green
    $j = $r.Content | ConvertFrom-Json
    Write-Host ("  atendeu .... {0}" -f $j.provider)
    Write-Host ("  modelo ..... {0}" -f $j.model)
    Write-Host ""
    Write-Host "Entao a ativacao esta roteavel. Pode repetir o canario." -ForegroundColor Green
} catch {
    $resp = $_.Exception.Response
    Write-Host ""
    if ($resp) {
        Write-Host ("HTTP {0} -- o roteamento RECUSA." -f [int]$resp.StatusCode) -ForegroundColor Red
        try {
            $leitor = New-Object IO.StreamReader($resp.GetResponseStream())
            Write-Host $leitor.ReadToEnd()
        } catch { }
        Write-Host ""
        Write-Host "E este e o mesmo erro que o Iris levou. A saida e afrouxar a" -ForegroundColor Yellow
        Write-Host "restricao na cerimonia, ou trocar de modelo -- e as duas sao" -ForegroundColor Yellow
        Write-Host "decisao sua, nao ajuste de codigo." -ForegroundColor Yellow
    } else {
        Write-Host ("Nem houve resposta HTTP: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
    exit 1
}
