Attribute VB_Name = "Módulo3"
Sub AdicionarInvestimento()
    Dim wsForm As Worksheet
    Dim wsInv As Worksheet
    Dim dataReg As Date
    Dim mesRef As String
    
    ' Conectar com as abas do arquivo
    On Error Resume Next
    Set wsForm = ThisWorkbook.Sheets("ENTRADA E GASTOS")
    Set wsInv = ThisWorkbook.Sheets("Investimentos")
    On Error GoTo 0
    
    If wsForm Is Nothing Or wsInv Is Nothing Then
        MsgBox "Verifique se as abas 'ENTRADA E GASTOS' e 'Investimentos' existem na sua planilha.", vbCritical, "Erro de Aba"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 1. Ler Data de Registro (J3) e Mês de Referência (K9)
    If IsDate(wsForm.Range("J3").Value) Then
        dataReg = CDate(wsForm.Range("J3").Value)
    Else
        dataReg = Date
    End If
    
    mesRef = Trim(wsForm.Range("K9").Value)
    
    If mesRef = "" Then
        MsgBox "Por favor, preencha o Mês de Referência em K9.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    ' 2. Processar e gravar as 4 caixinhas
    ProcessarEGravar wsForm, wsInv, dataReg, mesRef, "Investimentos", "L14", "K14"
    ProcessarEGravar wsForm, wsInv, dataReg, mesRef, "Apartamento", "L15", "K15"
    ProcessarEGravar wsForm, wsInv, dataReg, mesRef, "Lazer", "L16", "K16"
    ProcessarEGravar wsForm, wsInv, dataReg, mesRef, "Outros", "L17", "K17"
    
    ' 3. LIMPAR OS VALORES RESERVADOS (Voltam para 0 ou Vazio)
    wsForm.Range("L14:L17").ClearContents ' Limpa o conteúdo das células L14 a L17
    
    Application.ScreenUpdating = True
    
    MsgBox "Valores salvos na aba 'Investimentos' e campos limpos com sucesso!", vbInformation, "Sucesso"
End Sub

' Sub-rotina para verificar valor e gravar na aba Investimentos
Private Sub ProcessarEGravar(wsForm As Worksheet, wsInv As Worksheet, dataReg As Date, mesRef As String, nomeCaixinha As String, celReal As String, celSug As String)
    Dim valor As Double
    
    If wsForm.Range(celReal).Value > 0 Then
        valor = CDbl(wsForm.Range(celReal).Value)
    ElseIf wsForm.Range(celSug).Value > 0 Then
        valor = CDbl(wsForm.Range(celSug).Value)
    End If
    
    If valor > 0 Then
        RegistrarCaixinha wsInv, dataReg, nomeCaixinha, valor, mesRef
    End If
End Sub

' Sub-rotina auxiliar para inserir a linha
Private Sub RegistrarCaixinha(wsInv As Worksheet, dataReg As Date, caixinha As String, valor As Double, mesRef As String)
    Dim lastRow As Long
    Dim nextID As Long
    
    lastRow = wsInv.Cells(wsInv.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then lastRow = 3
    
    If lastRow = 3 Then
        nextID = 1
    Else
        nextID = Val(wsInv.Cells(lastRow, 2).Value) + 1
    End If
    
    lastRow = lastRow + 1
    
    wsInv.Cells(lastRow, 2).Value = nextID            ' Coluna B - ID
    wsInv.Cells(lastRow, 3).Value = dataReg           ' Coluna C - DATA
    wsInv.Cells(lastRow, 4).Value = caixinha          ' Coluna D - DESTINO
    wsInv.Cells(lastRow, 5).Value = valor             ' Coluna E - VALOR
    wsInv.Cells(lastRow, 6).Value = mesRef            ' Coluna F - MÊS REFERENTE
    
    wsInv.Cells(lastRow, 7).Formula = "=MONTH(C" & lastRow & ")" ' Coluna G - MÊS
    wsInv.Cells(lastRow, 8).Formula = "=YEAR(C" & lastRow & ")"  ' Coluna H - ANO
    
    wsInv.Cells(lastRow, 3).NumberFormat = "dd/mm/yyyy"
    wsInv.Cells(lastRow, 5).NumberFormat = """R$""#,##0.00"
End Sub
