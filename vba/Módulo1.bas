Attribute VB_Name = "Módulo1"
Sub AdicionarGasto()
    Dim wsForm As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim nextID As Long
    
    Dim dataGasto As Date
    Dim descricao As String
    Dim categoria As String
    Dim valorTotal As Double
    Dim formaPag As String
    Dim numParcelas As Integer
    Dim valorParcela As Double
    Dim cartao As String
    
    Dim i As Integer
    Dim dataParcela As Date
    Dim descParcela As String
    
    ' Conectar com as abas da planilha
    On Error Resume Next
    Set wsForm = ThisWorkbook.Sheets("ENTRADA E GASTOS")
    Set wsData = ThisWorkbook.Sheets("PLAN GASTOS")
    On Error GoTo 0
    
    If wsForm Is Nothing Or wsData Is Nothing Then
        MsgBox "Verifique se as abas 'ENTRADA E GASTOS' e 'PLAN GASTOS' existem no ficheiro.", vbCritical, "Erro de Aba"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 1. Ler dados da linha 3 da aba ENTRADA E GASTOS
    If IsDate(wsForm.Range("B3").Value) Then
        dataGasto = CDate(wsForm.Range("B3").Value)
    Else
        dataGasto = Date
    End If
    
    descricao = Trim(wsForm.Range("C3").Value)
    categoria = wsForm.Range("D3").Value
    valorTotal = CDbl(wsForm.Range("E3").Value)
    formaPag = wsForm.Range("F3").Value
    numParcelas = CInt(wsForm.Range("G3").Value)
    cartao = wsForm.Range("H3").Value
    
    ' Validações de segurança
    If descricao = "" Then
        MsgBox "Por favor, preencha a Descrição do gasto.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If valorTotal <= 0 Then
        MsgBox "O Valor Total deve ser maior que zero.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If numParcelas < 1 Then numParcelas = 1
    valorParcela = valorTotal / numParcelas
    
    ' 2. Encontrar a última linha preenchida na Coluna C (ID) da aba PLAN GASTOS
    lastRow = wsData.Cells(wsData.Rows.Count, "C").End(xlUp).Row
    If lastRow < 3 Then lastRow = 3
    
    ' Gerar o próximo ID automático
    If lastRow = 3 Then
        nextID = 1
    Else
        nextID = Val(wsData.Cells(lastRow, 3).Value) + 1
    End If
    
    ' 3. Escrever os dados na aba PLAN GASTOS (a partir da Coluna C / Coluna 3)
    For i = 1 To numParcelas
        lastRow = lastRow + 1
        
        ' Calcular data da parcela (avança 1 mês a cada parcela)
        dataParcela = DateAdd("m", i - 1, dataGasto)
        
        If numParcelas > 1 Then
            descParcela = descricao & " (" & i & "/" & numParcelas & ")"
        Else
            descParcela = descricao
        End If
        
        ' Inserir valores nas colunas C até N
        wsData.Cells(lastRow, 3).Value = nextID                      ' Coluna C - ID
        wsData.Cells(lastRow, 4).Value = dataParcela                 ' Coluna D - Data
        wsData.Cells(lastRow, 5).Value = descParcela                 ' Coluna E - Descrição
        wsData.Cells(lastRow, 6).Value = categoria                   ' Coluna F - Categoria
        wsData.Cells(lastRow, 7).Value = valorTotal                  ' Coluna G - Valor Total
        wsData.Cells(lastRow, 8).Value = formaPag                    ' Coluna H - Forma de Pagamento
        wsData.Cells(lastRow, 9).Value = numParcelas                 ' Coluna I - Nº Parcelas
        wsData.Cells(lastRow, 10).Value = valorParcela               ' Coluna J - Valor da Parcela
        wsData.Cells(lastRow, 11).Value = cartao                     ' Coluna K - Cartão
        
        ' Fórmulas automáticas para Mês, Ano e Mês/Ano apontando para a Coluna D (Data)
        wsData.Cells(lastRow, 12).Formula = "=MONTH(D" & lastRow & ")"        ' Coluna L - Mês
        wsData.Cells(lastRow, 13).Formula = "=YEAR(D" & lastRow & ")"         ' Coluna M - Ano
        'wsData.Cells(lastRow, 14).Formula = "=TEXT(D" & lastRow & ", ""mm/yyyy"")" ' Coluna N - Mês/Ano'
        
        ' Formatação visual das células gravadas
        wsData.Cells(lastRow, 4).NumberFormat = "dd/mm/yyyy"
        wsData.Cells(lastRow, 7).NumberFormat = """R$""#,##0.00"
        wsData.Cells(lastRow, 10).NumberFormat = """R$""#,##0.00"
        
        nextID = nextID + 1
    Next i
    
    ' 4. Limpar a linha de entrada na aba ENTRADA E GASTOS
    wsForm.Range("B3").Formula = "=TODAY()" ' Restaura a data atual em B3
    wsForm.Range("C3").Value = ""           ' Limpa Descrição
    wsForm.Range("E3").Value = 0            ' Zera Valor Total
    wsForm.Range("F3").Value = "À vista"    ' Reseta Forma de Pagamento
    wsForm.Range("G3").Value = 1            ' Reseta Parcelas para 1
    
    Application.ScreenUpdating = True
    
    MsgBox "Gasto adicionado com sucesso!", vbInformation, "Sucesso"
End Sub
