# D5 - a latencia por lote, medida em CONDICAO DECLARADA.
#
# Este script existe por causa de um erro que quase entrou no relatorio final.
#
# A primeira medicao deu "max 62 ms, todos abaixo do orcamento". Depois o Codex
# leu o arquivo de medicoes inteiro e achou cinco execucoes, tres delas acima
# de 100 ms - 107, 120, 115. Eu tinha reportado a favoravel e ignorado as
# outras quatro.
#
# So que a explicacao nao era "o produto e lento". Era que o MSTest roda com 12
# workers em paralelo, e as execucoes ruins aconteceram com a SUITE INTEIRA no
# ar: outros testes disputando a mesma STA do Outlook, mais os testes de crash
# gerando processos. Medido:
#
#   isolado                      max  57- 58 ms
#   dois testes COM em paralelo  max  74- 76 ms
#   suite inteira, 12 workers    max      184 ms
#
# Nenhum dos dois numeros que eu quase publiquei era a medicao do produto. O
# AMBIENTE DE MEDICAO ERA PARTE DA MEDICAO, e eu nao tinha declarado qual.
#
# A condicao normativa e a ISOLADA: o produto nao roda 317 testes em paralelo
# com ele mesmo. E dizer isso e obrigatorio - numero sem condicao declarada nao
# e medicao, e uma das duas leituras que ele admite.
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)

$repeticoes = if ($args.Count -gt 0) { [int]$args[0] } else { 3 }

Remove-Item -Recurse -Force medicoes -EA SilentlyContinue

Write-Host "Medindo a D5 em condicao ISOLADA, $repeticoes execucoes."
Write-Host "Requer o Outlook classico aberto."
Write-Host ""

for ($i = 1; $i -le $repeticoes; $i++) {
    Write-Host "  execucao $i..."
    dotnet test Iris.slnx --nologo -v q `
        --filter "Importa_uma_pasta_real_e_mede_a_latencia_por_lote" | Out-Null
}

Write-Host ""
Write-Host "=== resultado (condicao: ISOLADA) ==="
Get-Content medicoes/2.2b-medicoes.txt | Select-String -Pattern '^  min'

Write-Host ""
Write-Host "Para comparar com a condicao contaminada, rode a suite inteira:"
Write-Host "  dotnet test Iris.slnx"
Write-Host "e olhe o mesmo arquivo. A diferenca e o custo do paralelismo do"
Write-Host "TESTE, nao do produto."
