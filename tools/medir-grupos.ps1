<#
.SYNOPSIS
    Mede o que o OOM expoe para TAREFAS, CALENDARIO e CONTATOS.

.DESCRIPTION
    SOMENTE LEITURA. Nao cria, nao move, nao apaga, nao envia, nao aceita
    convite. Le propriedades de itens que ja existem.

    POR QUE ESTE ROTEIRO EXISTE

    As Fases 5, 6 e 7 do ESCOPO -- Tarefas, Calendario e Contatos -- estao
    escritas em uma frase cada. Planeja-las sem medir seria repetir o erro que
    a Fase 0 existiu para evitar: a versao 1 do ESCOPO afirmava que WPF em VB
    nao era suportado no .NET moderno, e a premissa era falsa.

    O QUE ELE RESPONDE, E POR QUE CADA UMA IMPORTA

    1. As pastas existem, e quantos itens tem? Um modulo para uma pasta vazia
       e trabalho sem consumidor.

    2. As propriedades que o modulo precisaria SAO LEGIVEIS? A Fase 0 mediu
       que Permission nao vem por coluna de Table, e isso mudou o desenho
       inteiro do gate. Descobrir o equivalente DEPOIS de escrever o modulo
       custa o modulo.

    3. Quanto custa montar um DTO? A Fase 0 mediu ~16 ms por MENSAGEM, e foi
       esse numero que tornou o cache obrigatorio. Se tarefa custar o mesmo,
       a Fase 5 herda a mesma conclusao antes de comecar.

    4. A recorrencia do calendario aparece? Compromisso recorrente e a
       diferenca entre "listar eventos" e "expandir uma serie" -- e as duas
       coisas tem tamanhos muito diferentes.

    O QUE ELE NAO FAZ

    Nao cria item nenhum. Saber se DA para criar uma tarefa exige criar uma
    tarefa, e isso e mutacao na caixa do dono. Fica para quando ele estiver
    na maquina.

.NOTES
    ASCII puro e sem BOM, como os outros roteiros deste diretorio.
#>
[CmdletBinding()]
param(
    # Quantos itens medir por grupo ao cronometrar o DTO.
    [int]$Amostra = 30
)

$ErrorActionPreference = 'Stop'

function Ler($bloco) {
    # Leitura defensiva: propriedade que o Outlook recusa nao pode derrubar a
    # medicao -- "nao consegui ler" e o proprio resultado.
    try { & $bloco } catch { $null }
}

Write-Host "Medindo o que o OOM expoe para Tarefas, Calendario e Contatos..." -ForegroundColor Cyan
Write-Host "SOMENTE LEITURA. Nada e criado, movido ou apagado." -ForegroundColor DarkGray
Write-Host ""

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
    Write-Host "O Outlook nao esta aberto. Abra-o e rode de novo." -ForegroundColor Red
    exit 1
}
$ns = $ol.GetNamespace('MAPI')

# olFolderCalendar=9, olFolderContacts=10, olFolderTasks=13
$grupos = @(
    @{ Nome = 'CALENDARIO'; Id = 9;  Fase = 6 },
    @{ Nome = 'CONTATOS';   Id = 10; Fase = 7 },
    @{ Nome = 'TAREFAS';    Id = 13; Fase = 5 }
)

foreach ($g in $grupos) {
    Write-Host ("=" * 66) -ForegroundColor DarkGray
    Write-Host ("{0}  (Fase {1})" -f $g.Nome, $g.Fase) -ForegroundColor Cyan

    $pasta = $null
    $itens = $null
    try {
        $pasta = $ns.GetDefaultFolder($g.Id)
    } catch {
        Write-Host "  A pasta padrao nao existe nesta conta." -ForegroundColor Yellow
        continue
    }

    try {
        $itens = $pasta.Items
        $n = $itens.Count
        Write-Host ("  pasta:  {0}" -f $pasta.Name)
        Write-Host ("  itens:  {0}" -f $n)

        if ($n -eq 0) {
            # "VAZIA" E AFIRMACAO DE AUSENCIA. O Count do OOM e o que esta
            # EXPOSTO LOCALMENTE, e a contagem do servidor continua
            # inalcancavel -- e essa e a ressalva que o projeto inteiro repete.
            Write-Host "  NENHUM ITEM EXPOSTO LOCALMENTE nesta pasta." -ForegroundColor Yellow
            Write-Host "    Nao da para concluir que ela esteja vazia no servidor; da para"
            Write-Host "    dizer que nao ha o que agrupar AQUI, que e o que decide o modulo."
            continue
        }

        # ORDENAR ANTES DE AMOSTRAR.
        #
        # Items.Item(i) sem Sort devolve a ordem interna do provedor, que nao e
        # ordem nenhuma que signifique alguma coisa. Uma amostra dos "30
        # primeiros" nessa ordem nao e amostra: e um pedaco arbitrario, e
        # relatar "0 recorrentes em 30" a partir dele seria dar aparencia de
        # medida a um acaso.
        #
        # Ordenado por data, os 30 primeiros sao os 30 mais recentes -- que e
        # uma populacao que da para nomear.
        try {
            switch ($g.Id) {
                9  { $itens.Sort("[Start]", $true) }
                13 { $itens.Sort("[DueDate]", $true) }
                10 { $itens.Sort("[FullName]", $false) }
            }
        } catch {
            Write-Host "  (nao consegui ordenar: a amostra e a ordem interna do provedor)" -ForegroundColor DarkYellow
        }

        # ---- as propriedades, medidas item a item ----
        $quantos = [Math]::Min($Amostra, $n)
        $campos = @{}
        $relogio = [Diagnostics.Stopwatch]::StartNew()
        $lidos = 0
        $recorrentes = 0

        for ($i = 1; $i -le $quantos; $i++) {
            $item = $null
            try {
                $item = $itens.Item($i)
                $lidos++

                switch ($g.Id) {
                    9 {
                        $campos['Subject']        += [int](Ler { $null -ne $item.Subject })
                        $campos['Start']          += [int](Ler { $null -ne $item.Start })
                        $campos['End']            += [int](Ler { $null -ne $item.End })
                        $campos['Location']       += [int](Ler { $null -ne $item.Location })
                        $campos['AllDayEvent']    += [int](Ler { $null -ne $item.AllDayEvent })
                        $campos['Organizer']      += [int](Ler { $null -ne $item.Organizer })
                        $campos['ResponseStatus'] += [int](Ler { $null -ne $item.ResponseStatus })
                        $campos['IsRecurring']    += [int](Ler { $null -ne $item.IsRecurring })
                        $campos['BusyStatus']     += [int](Ler { $null -ne $item.BusyStatus })
                        $campos['Recipients.Cnt'] += [int](Ler { $null -ne $item.Recipients.Count })
                        if (Ler { $item.IsRecurring }) { $recorrentes++ }
                    }
                    10 {
                        $campos['FullName']       += [int](Ler { $null -ne $item.FullName })
                        $campos['CompanyName']    += [int](Ler { $null -ne $item.CompanyName })
                        $campos['Email1Address']  += [int](Ler { -not [string]::IsNullOrEmpty($item.Email1Address) })
                        $campos['MobileTel']      += [int](Ler { -not [string]::IsNullOrEmpty($item.MobileTelephoneNumber) })
                        $campos['JobTitle']       += [int](Ler { $null -ne $item.JobTitle })
                        $campos['LastModified']   += [int](Ler { $null -ne $item.LastModificationTime })
                    }
                    13 {
                        $campos['Subject']        += [int](Ler { $null -ne $item.Subject })
                        $campos['DueDate']        += [int](Ler { $null -ne $item.DueDate })
                        $campos['StartDate']      += [int](Ler { $null -ne $item.StartDate })
                        $campos['Status']         += [int](Ler { $null -ne $item.Status })
                        $campos['Importance']     += [int](Ler { $null -ne $item.Importance })
                        $campos['PercentComplete']+= [int](Ler { $null -ne $item.PercentComplete })
                        $campos['Owner']          += [int](Ler { $null -ne $item.Owner })
                        $campos['Complete']       += [int](Ler { $null -ne $item.Complete })
                        $campos['ReminderSet']    += [int](Ler { $null -ne $item.ReminderSet })
                    }
                }
            } catch {
                Write-Host ("  item {0}: {1}" -f $i, $_.Exception.Message) -ForegroundColor DarkYellow
            } finally {
                if ($item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
            }
        }
        $relogio.Stop()

        if ($lidos -gt 0) {
            $porItem = $relogio.Elapsed.TotalMilliseconds / $lidos
            Write-Host ("  custo:  {0:N1} ms por item  ({1} itens medidos)" -f $porItem, $lidos)
            Write-Host "          referencia: a Fase 0 mediu ~16 ms por MENSAGEM," -ForegroundColor DarkGray
            Write-Host "          e foi esse numero que tornou o cache obrigatorio." -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  propriedades legiveis (de $lidos):"
            foreach ($k in ($campos.Keys | Sort-Object)) {
                $v = $campos[$k]
                $marca = if ($v -eq $lidos) { "ok " } elseif ($v -eq 0) { "NAO" } else { "par" }
                Write-Host ("    [{0}] {1,-16} {2}/{3}" -f $marca, $k, $v, $lidos)
            }
            if ($g.Id -eq 9) {
                Write-Host ""
                Write-Host ("  recorrentes na amostra: {0} de {1} (os mais recentes por [Start])" -f $recorrentes, $lidos)
                Write-Host "          recorrencia e a diferenca entre LISTAR eventos e" -ForegroundColor DarkGray
                Write-Host "          EXPANDIR uma serie, e as duas tem tamanhos diferentes." -ForegroundColor DarkGray
            }
        }
    } catch {
        Write-Host ("  nao consegui ler: {0}" -f $_.Exception.Message) -ForegroundColor Red
    } finally {
        if ($itens) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($itens) }
        if ($pasta) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pasta) }
    }
    Write-Host ""
}

Write-Host ("=" * 66) -ForegroundColor DarkGray
Write-Host "O QUE ISTO NAO MEDIU:" -ForegroundColor DarkGray
Write-Host "  Se DA para criar. Saber isso exige criar, e criar e mutacao na" -ForegroundColor DarkGray
Write-Host "  caixa do dono -- fica para quando ele estiver na maquina." -ForegroundColor DarkGray

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
