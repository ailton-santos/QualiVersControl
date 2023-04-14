Option Explicit

' Declarando as variáveis
Dim versaoAtual As Integer ' Versão atual do documento
Dim dataAtualizacao As Date ' Data da última atualização do documento
Dim descricaoAtualizacao As String ' Descrição da última atualização do documento

' Função para adicionar uma nova versão ao documento
Sub adicionarVersao()
    ' Atualiza as variáveis
    versaoAtual = versaoAtual + 1
    dataAtualizacao = Date
    descricaoAtualizacao = InputBox("Digite uma descrição para a atualização:")
    
    ' Adiciona a nova versão ao histórico
    With Sheets("Controle de Versões")
        .Range("A1").EntireRow.Insert ' Insere uma nova linha no início da tabela
        .Range("A2").Value = versaoAtual
        .Range("B2").Value = dataAtualizacao
        .Range("C2").Value = descricaoAtualizacao
    End With
    
    ' Mostra a mensagem de confirmação
    MsgBox "Nova versão adicionada com sucesso!"
End Sub

' Função para mostrar o histórico de versões
Sub mostrarHistorico()
    ' Limpa a planilha
    Sheets("Controle de Versões").Cells.ClearContents
    
    ' Configura a planilha
    With Sheets("Controle de Versões")
        .Range("A1").Value = "Controle de Versões"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A3").Value = "Versão"
        .Range("B3").Value = "Data de Atualização"
        .Range("C3").Value = "Descrição da Atualização"
        
        ' Preenche o histórico com as versões anteriores
        Dim ultimaLinha As Integer
        ultimaLinha = .Range("A" & .Rows.Count).End(xlUp).Row
        For i = 2 To ultimaLinha
            .Cells(i, 1).Value = versaoAtual - (ultimaLinha - i)
            .Cells(i, 2).Value = dataAtualizacao
            .Cells(i, 3).Value = descricaoAtualizacao
        Next i
    End With
End Sub

' Função principal
Sub main()
    ' Inicializa as variáveis
    versaoAtual = 1
    dataAtualizacao = Date
    descricaoAtualizacao = "Primeira versão do documento"
    
    ' Mostra o menu de opções
    Dim opcao As Integer
    opcao = MsgBox("O que deseja fazer?" & vbNewLine & _
                    "1. Adicionar nova versão" & vbNewLine & _
                    "2. Mostrar histórico de versões", _
                    vbQuestion + vbYesNo)
    
    If opcao = vbYes Then ' Se escolheu adicionar nova versão
        adicionarVersao()
    Else ' Se escolheu mostrar histórico de versões
        mostrarHistorico()
    End If
End Sub
