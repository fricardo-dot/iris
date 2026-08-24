# A tecnica da JANELA funciona para o Outlook?
#
# SOMENTE LEITURA. Nao muda nada; so tenta obter o modelo de objetos por um
# caminho alternativo ao ROT.
#
# Antes de escrever VB, medir. A §21.2 mediu que GetActiveObject pode
# falhar por minutos com o Outlook rodando e visivel. O Codex propos usar
# AccessibleObjectFromWindow com OBJID_NATIVEOM, que NAO PODE iniciar o
# Outlook porque depende de uma janela que ja existe.
#
# Mas a tecnica e conhecida sobretudo para Word e Excel. Para o Outlook
# quero ver com meus olhos:
#   - a janela existe e com que classe;
#   - o que AccessibleObjectFromWindow devolve;
#   - se da para chegar em Application a partir dali;
#   - e se o objeto obtido REALMENTE FUNCIONA (uma leitura de verdade,
#     nao so "nao lancou" — a licao da coluna Permission da Q1).

$ErrorActionPreference = "Stop"

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Janelas {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr FindWindowEx(IntPtr parent, IntPtr child, string cls, string title);
    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hwnd, StringBuilder buf, int max);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("oleacc.dll", PreserveSig=false)]
    [return: MarshalAs(UnmanagedType.Interface)]
    public static extern object AccessibleObjectFromWindow(
        IntPtr hwnd, uint objectId, ref Guid riid);
}
"@

$OBJID_NATIVEOM = [uint32]"0xFFFFFFF0"
$IID_IDispatch = [Guid]"00020400-0000-0000-C000-000000000046"

Write-Host "=== procurando janelas do Outlook ==="
$classes = @("rctrl_renwnd32", "OutlookGrid", "Framework::CFrame")
foreach ($cls in $classes) {
    $h = [IntPtr]::Zero
    $n = 0
    do {
        $h = [Janelas]::FindWindowEx([IntPtr]::Zero, $h, $cls, $null)
        if ($h -ne [IntPtr]::Zero) {
            $n++
            $idProc = 0
            [void][Janelas]::GetWindowThreadProcessId($h, [ref]$idProc)
            $vis = [Janelas]::IsWindowVisible($h)
            $nome = "?"
            try { $nome = (Get-Process -Id $idProc -ErrorAction Stop).ProcessName } catch { }
            Write-Host ("  classe={0,-18} hwnd={1}  pid={2} ({3})  visivel={4}" -f `
                $cls, $h, $idProc, $nome, $vis)
        }
    } while ($h -ne [IntPtr]::Zero -and $n -lt 5)
    if ($n -eq 0) { Write-Host ("  classe={0,-18} (nenhuma)" -f $cls) }
}

Write-Host ""
Write-Host "=== tentando obter o modelo de objetos pela janela ==="
$h = [Janelas]::FindWindowEx([IntPtr]::Zero, [IntPtr]::Zero, "rctrl_renwnd32", $null)
if ($h -eq [IntPtr]::Zero) {
    Write-Host "  nenhuma janela rctrl_renwnd32. A tecnica nao se aplica aqui."
    exit 1
}

$pidJanela = 0
[void][Janelas]::GetWindowThreadProcessId($h, [ref]$pidJanela)
Write-Host ("  janela {0}, do PID {1}" -f $h, $pidJanela)

$obj = $null
try {
    $riid = $IID_IDispatch
    $obj = [Janelas]::AccessibleObjectFromWindow($h, $OBJID_NATIVEOM, [ref]$riid)
} catch {
    Write-Host ("  AccessibleObjectFromWindow LANCOU: {0}" -f $_.Exception.Message)
    exit 1
}

if ($null -eq $obj) {
    Write-Host "  devolveu NULO."
    exit 1
}

Write-Host ("  devolveu um objeto: {0}" -f $obj.GetType().FullName)

# O que e esse objeto? Para o Outlook, espera-se algo do modelo nativo.
foreach ($prop in @("Name", "Class", "Application", "CurrentFolder", "Session")) {
    try {
        $v = $obj.$prop
        $desc = if ($null -eq $v) { "(nulo)" }
                elseif ($v -is [string] -or $v -is [int]) { "$v" }
                else { $v.GetType().FullName }
        Write-Host ("     .{0,-14} -> {1}" -f $prop, $desc)
    } catch {
        Write-Host ("     .{0,-14} -> (erro)" -f $prop)
    }
}

Write-Host ""
Write-Host "=== chegando em Application e USANDO de verdade ==="
$app = $null
try { $app = $obj.Application } catch { }
if ($null -eq $app) {
    Write-Host "  nao consegui .Application a partir do objeto da janela."
    exit 1
}
Write-Host ("  Application obtido: {0}" -f $app.GetType().FullName)

# NAO basta "nao lancou". Le de verdade — a licao da coluna Permission.
try {
    Write-Host ("     .Name         = {0}" -f $app.Name)
    Write-Host ("     .Version      = {0}" -f $app.Version)
    $ns = $app.GetNamespace("MAPI")
    Write-Host ("     GetNamespace  = ok")
    $stores = $ns.Stores
    Write-Host ("     Stores.Count  = {0}" -f $stores.Count)
    $in = $ns.GetDefaultFolder(6)
    Write-Host ("     Entrada       = {0} itens" -f $in.Items.Count)
    Write-Host ""
    Write-Host "  *** A TECNICA FUNCIONA: leitura real bem-sucedida. ***"
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($in)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($stores)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
} catch {
    Write-Host ("  FALHOU numa leitura real: {0}" -f $_.Exception.Message)
    Write-Host "  Objeto obtido mas inutil — seria 'Connected mentiroso'."
}

# o PID mudou? Ninguem pode ter iniciado Outlook novo.
Write-Host ""
$procs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
Write-Host ("  processos OUTLOOK.EXE agora: {0} (PIDs {1})" -f $procs.Count, ($procs.Id -join ","))
Write-Host ("  a janela usada era do PID {0}" -f $pidJanela)
