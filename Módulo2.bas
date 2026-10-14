Attribute VB_Name = "Módulo2"
Sub AdicionarReceita()
    Dim wsForm As Worksheet
    Dim wsRec As Worksheet
    Dim lastRow As Long
    Dim nextID As Long
    
    Dim dataReg As Date
    Dim tipoRec As String
    Dim valorRec As Double
    Dim mesRef As String
    
    ' Conectar com as abas do seu arquivo
    On Error Resume Next
    Set wsForm = ThisWorkbook.Sheets("ENTRADA E GASTOS")
    Set wsRec = ThisWorkbook.Sheets("RECEITA")
    On Error GoTo 0
    
    If wsForm Is Nothing Or wsRec Is Nothing Then
        MsgBox "Verifique se as abas 'ENTRADA E GASTOS' e 'RECEITA' existem na sua planilha.", vbCritical, "Erro de Aba"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 1. Ler dados do formulário de entrada (Linha 3 - Colunas J até M)
    If IsDate(wsForm.Range("J3").Value) Then
        dataReg = CDate(wsForm.Range("J3").Value)
    Else
        dataReg = Date
    End If
    
    tipoRec = Trim(wsForm.Range("K3").Value)
    valorRec = CDbl(wsForm.Range("L3").Value)
    mesRef = Trim(wsForm.Range("M3").Value)
    
    ' Validações de segurança
    If tipoRec = "" Then
        MsgBox "Por favor, selecione o Tipo de receita (Adiantamento ou Salário).", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If valorRec <= 0 Then
        MsgBox "O valor da receita deve ser maior que zero.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    ' 2. Localizar a última linha preenchida na Coluna B (ID) da aba RECEITA
    lastRow = wsRec.Cells(wsRec.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    
    ' Gerar o ID automático
    If lastRow = 2 Then
        nextID = 1
    Else
        nextID = Val(wsRec.Cells(lastRow, 2).Value) + 1
    End If
    
    lastRow = lastRow + 1
    
    ' 3. Escrever os dados na aba RECEITA (Colunas B até H)
    wsRec.Cells(lastRow, 2).Value = nextID            ' Coluna B - ID
    wsRec.Cells(lastRow, 3).Value = dataReg           ' Coluna C - DATA DE REGISTRO
    wsRec.Cells(lastRow, 4).Value = tipoRec           ' Coluna D - TIPO
    wsRec.Cells(lastRow, 5).Value = valorRec          ' Coluna E - VALOR
    wsRec.Cells(lastRow, 6).Value = mesRef            ' Coluna F - MÊS REFERENTE
    
    ' Fórmulas para Mês e Ano baseados na Data de Registro (Coluna C)
    wsRec.Cells(lastRow, 7).Formula = "=MONTH(C" & lastRow & ")" ' Coluna G - MÊS
    wsRec.Cells(lastRow, 8).Formula = "=YEAR(C" & lastRow & ")"  ' Coluna H - ANO
    
    ' Formatação das Células
    wsRec.Cells(lastRow, 3).NumberFormat = "dd/mm/yyyy"
    wsRec.Cells(lastRow, 5).NumberFormat = """R$""#,##0.00"
    
    ' 4. Limpar o Formulário de Entrada
    wsForm.Range("J3").Formula = "=TODAY()"
    wsForm.Range("K3").Value = ""
    wsForm.Range("L3").Value = 0
    wsForm.Range("M3").Value = ""
    
    Application.ScreenUpdating = True
    
    MsgBox tipoRec & " adicionado com sucesso na aba RECEITA!", vbInformation, "Sucesso"
End Sub
