Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Iris.Model

Namespace Global.Iris.Integration

    ''' <summary>
    ''' <b>As identidades do dono, num arquivo que ele pode abrir e corrigir.</b>
    '''
    ''' ------------------------------------------------------------------
    ''' <b>POR QUE UM ARQUIVO, E NÃO SÓ O QUE O OUTLOOK DIZ</b>
    '''
    ''' O Outlook sabe as contas configuradas. Ele não sabe, com a mesma
    ''' segurança, os <i>aliases</i> pelos quais a organização entrega mensagem
    ''' em nome do dono, nem as caixas compartilhadas em que ele responde como
    ''' se fosse outra pessoa. Errar isso não dá erro: dá uma fila de respostas
    ''' pendentes cobrando do dono mensagens que ele mesmo escreveu.
    '''
    ''' Então o Outlook <b>semeia</b> e o dono <b>corrige</b>. A semeadura só
    ''' acontece enquanto não há <b>identidade nenhuma</b> escrita:
    ''' reescrever a cada abertura apagaria a correção, que é o único motivo
    ''' de o arquivo existir — e não reparar um arquivo sem endereço nenhum
    ''' congelaria o conjunto vazio para sempre.
    '''
    ''' ------------------------------------------------------------------
    ''' <b>FALHA VALE COMO CONJUNTO VAZIO</b>
    '''
    ''' Não conseguir ler devolve <see cref="MinhasIdentidades"/> vazio, e vazio
    ''' responde <c>Desconhecida</c> para tudo. A fila mostra linhas incertas em
    ''' vez de linhas erradas — e uma linha que se declara incerta é conferível,
    ''' enquanto um palpite com cara de fato não é.
    ''' </summary>
    Public NotInheritable Class IdentidadesEmArquivo

        Private ReadOnly _caminho As String

        Public Sub New(Optional caminho As String = Nothing)
            _caminho = If(caminho, CaminhoPadrao())
        End Sub

        Public Shared Function CaminhoPadrao() As String
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Iris", "identidades.txt")
        End Function

        Public ReadOnly Property Caminho As String
            Get
                Return _caminho
            End Get
        End Property

        ''' <summary>Existe arquivo? Se não, ninguém semeou ainda.</summary>
        Public ReadOnly Property Existe As Boolean
            Get
                Try
                    Return File.Exists(_caminho)
                Catch
                    Return False
                End Try
            End Get
        End Property

        ''' <summary>
        ''' Lê o conjunto. Linha vazia e linha começando com <c>#</c> são
        ''' ignoradas — o cabeçalho semeado é todo comentário, e um leitor que
        ''' engolisse o cabeçalho como endereço acharia que o dono se chama
        ''' "uma por linha".
        ''' </summary>
        Public Function Ler() As MinhasIdentidades
            Try
                If Not File.Exists(_caminho) Then Return New MinhasIdentidades({})

                Dim uteis = File.ReadAllLines(_caminho, Encoding.UTF8).
                            Select(Function(l) l.Trim()).
                            Where(Function(l) l.Length > 0 AndAlso Not l.StartsWith("#"))
                Return New MinhasIdentidades(uteis)
            Catch
                Return New MinhasIdentidades({})
            End Try
        End Function

        ''' <summary>
        ''' <b>Semeia — e só quando não há nada a desfazer.</b>
        '''
        ''' Devolve <c>True</c> quando escreveu. Arquivo com <b>pelo menos uma</b>
        ''' identidade devolve <c>False</c> e não é tocado: o dono pode ter
        ''' apagado uma linha de propósito, e semeadura que insiste desfaz
        ''' correção.
        '''
        ''' ------------------------------------------------------------------
        ''' <b>MAS ARQUIVO SEM ENDEREÇO NENHUM É REPARADO</b>
        '''
        ''' A regra era <i>existe, não toca</i>, e ela congelava o pior estado
        ''' possível: um arquivo só com o cabeçalho de comentários — meia escrita
        ''' interrompida, disco cheio, cópia de perfil pela metade — produzia
        ''' conjunto vazio <b>para sempre</b>, e a fila respondia "não sei" a
        ''' toda mensagem sem nada explicando por quê. Era exatamente o defeito
        ''' que o comentário da semeadura dizia estar evitando.
        '''
        ''' A troca de <c>File.Exists</c> por <i>tem identidade</i> preserva o que
        ''' importa — nada que o dono escreveu é desfeito — e repara o único caso
        ''' em que não há o que preservar.
        '''
        ''' Achado por revisão externa em 31/08/2026.
        '''
        ''' Sem nenhum endereço para semear, também não escreve: um arquivo só
        ''' com comentários pareceria semeado e não estaria.
        ''' </summary>
        Public Function Semear(enderecos As IEnumerable(Of String)) As Boolean
            Try
                If Ler().Quantas > 0 Then Return False

                Dim limpos = New MinhasIdentidades(enderecos).Listar()
                If limpos.Count = 0 Then Return False

                Dim linhas As New List(Of String) From {
                    "# As identidades do dono desta caixa, uma por linha.",
                    "#",
                    "# O Iris usa esta lista para saber QUEM escreveu cada mensagem, e",
                    "# assim separar 'estou esperando alguem' de 'alguem esta me",
                    "# esperando'. A pasta nao serve para isso: alias, conta adicional",
                    "# e caixa compartilhada quebram a regra 'esta em Itens Enviados,",
                    "# logo fui eu'.",
                    "#",
                    "# Acrescente os seus aliases. Numa organizacao Exchange, o",
                    "# remetente interno pode aparecer como endereco X.500 comecando",
                    "# com /O= -- se voce vir mensagens suas na fila de respostas",
                    "# pendentes, e essa forma que falta aqui.",
                    "#",
                    "# Linha vazia e linha com # sao ignoradas."}
                linhas.AddRange(limpos)

                Directory.CreateDirectory(Path.GetDirectoryName(_caminho))
                File.WriteAllLines(_caminho, linhas, Encoding.UTF8)
                Return True
            Catch
                Return False
            End Try
        End Function

    End Class

End Namespace
