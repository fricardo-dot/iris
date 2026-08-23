# Q2, parte 4: quao perto dois itens DISTINTOS chegam?
#
# SOMENTE LEITURA.
#
# A parte 3 mostrou onde a evidencia FORTE erra (enviado x recebido). Esta
# parte olha o outro lado: se alguem for tentado a correlacionar por
# heuristica — assunto, remetente, tamanho, instante — quantos itens
# comprovadamente diferentes ela funde?
#
# ORACULO, e ele so vale numa direcao:
#
#   Message-ID DIFERENTE (os dois nao vazios) => mensagens diferentes.
#   Isso e definicao do RFC 5322, nao inferencia minha. Qualquer regra que
#   una um par desses esta ERRADA, e da para contar.
#
#   Message-ID IGUAL NAO autoriza unir. A parte 3 mostrou o porque: o
#   enviado e o recebido da mesma mensagem compartilham Message-ID E
#   SearchKey e ainda assim sao itens distintos, com estado proprio cada
#   um. "Mesma mensagem" e mais grosso que "mesmo item".
#
# Por isso a coluna abaixo se chama "mesma msg" e nao "certos": ela conta
# unioes que o oraculo NAO condena, nao unioes que ele aprova.
#
# PRIVACIDADE: assunto so agregado; nada de corpo nem endereco.

$PT_MSG_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
$ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$ns = $ol.GetNamespace("MAPI")

function Levantar($id, [string]$rotulo) {
    $pasta = $ns.GetDefaultFolder($id)
    $t = $pasta.GetTable()
    $t.Columns.RemoveAll()
    foreach ($c in @("EntryID","Subject","SenderName","ReceivedTime","Size",$PT_MSG_ID)) {
        [void]$t.Columns.Add($c)
    }
    $itens = @()
    while (-not $t.EndOfTable) {
        $a = $t.GetArray(200)
        for ($r = $a.GetLowerBound(0); $r -le $a.GetUpperBound(0); $r++) {
            $itens += [pscustomobject]@{
                Pasta   = $rotulo
                Id      = "$($a.GetValue($r,0))"
                Assunto = "$($a.GetValue($r,1))"
                Remet   = "$($a.GetValue($r,2))"
                Quando  = $a.GetValue($r,3)
                Tam     = [int]("0" + "$($a.GetValue($r,4))")
                Mid     = "$($a.GetValue($r,5))".Trim()
            }
        }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($t)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta)
    return ,$itens
}

$todos = @()
foreach ($p in @(@(6,"Entrada"), @(5,"Enviados"), @(16,"Rascunhos"), @(3,"Excluidos"))) {
    $todos += Levantar $p[0] $p[1]
}
Write-Host "corpus: $($todos.Count) itens"

# So itens com Message-ID: sao os que o oraculo consegue julgar.
$julgaveis = @($todos | Where-Object { $_.Mid -ne "" })
Write-Host "julgaveis pelo oraculo (tem Message-ID): $($julgaveis.Count)"
Write-Host ""

function Avaliar {
    param([string]$Nome, [scriptblock]$Chave)

    $g = $julgaveis | Group-Object { & $Chave $_ } | Where-Object { $_.Count -gt 1 }
    $g = @($g)

    $paresErrados = 0     # par unido pela regra, com Message-ID diferente
    $mesmaMsg     = 0     # par unido pela regra, com Message-ID igual
    $piorGrupo    = 0

    foreach ($grupo in $g) {
        $itens = @($grupo.Group)
        if ($itens.Count -gt $piorGrupo) { $piorGrupo = $itens.Count }
        for ($i = 0; $i -lt $itens.Count; $i++) {
            for ($j = $i + 1; $j -lt $itens.Count; $j++) {
                if ($itens[$i].Mid -eq $itens[$j].Mid) { $mesmaMsg++ } else { $paresErrados++ }
            }
        }
    }

    Write-Host ("{0,-46} | {1,6} | {2,7} | {3,6} | {4}" -f `
        $Nome, $g.Count, $paresErrados, $mesmaMsg, $piorGrupo)
    return $paresErrados
}

Write-Host ("{0,-46} | {1,6} | {2,7} | {3,6} | {4}" -f `
    "regra de correlacao", "grupos", "ERRADOS", "mesma msg", "maior grupo")
Write-Host ("-" * 96)

[void](Avaliar "assunto"                        { param($x) $x.Assunto })
[void](Avaliar "assunto + remetente"            { param($x) "$($x.Assunto)|$($x.Remet)" })
[void](Avaliar "assunto + remetente + tamanho"  { param($x) "$($x.Assunto)|$($x.Remet)|$($x.Tam)" })
[void](Avaliar "assunto + remetente + instante" { param($x) "$($x.Assunto)|$($x.Remet)|$($x.Quando)" })
[void](Avaliar "assunto+remet+tamanho+instante" { param($x) "$($x.Assunto)|$($x.Remet)|$($x.Tam)|$($x.Quando)" })
Write-Host ("-" * 96)
$e1 = Avaliar "Message-ID (o proprio oraculo)"  { param($x) $x.Mid }

Write-Host ""
Write-Host "ERRADOS   = pares que a regra une e que tem Message-ID DIFERENTE:"
Write-Host "            mensagens comprovadamente distintas sendo fundidas."
Write-Host "mesma msg = pares da mesma mensagem. NAO conte como acerto: o unico"
Write-Host "            par assim nesta caixa e o enviado x recebido da parte 3,"
Write-Host "            que sao itens distintos e nao devem ser unidos."
Write-Host ""
Write-Host "Ou seja: nesta caixa NAO EXISTE par que devesse ser correlacionado."
Write-Host "Toda uniao disponivel aqui e erro. Um par legitimo so aparece com"
Write-Host "um item MOVIDO — e mover exige autorizacao do usuario."
Write-Host ""
Write-Host "A ultima linha e o oraculo contra si mesmo: tem que dar 0 errados."
if ($e1 -ne 0) { Write-Host "FALHA: o oraculo se contradisse."; exit 1 }
