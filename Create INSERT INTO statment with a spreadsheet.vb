Sub CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub


Public Sub Insert_Into()
Dim l, c, l_fim, c_fim As Long
Dim query, tabela As String

tabela = InputBox("Digite o nome da tabela igual está no banco de dados")
query = "INSERT INTO " & tabela & " ("

c_fim = Cells(1, Columns.Count).End(xlToLeft).Column

'Loop para o caso da última linha da primeira coluna estar vazia
l_fim = Range("A" & Rows.Count).End(xlUp).Row
    For c = 1 To c_fim
        
        If (Cells(Rows.Count, c).End(xlUp).Row) > l_fim Then
            l_fim = Cells(Rows.Count, c).End(xlUp).Row
        End If
        
    Next

'Loop para escrever a primeira linha do INSERT INTO
For c = 1 To c_fim
    
    If c = c_fim Then
        query = query & Cells(1, c) & ") VALUES"
    Else
        query = query & Cells(1, c) & ", "
    End If
    
Next

'Loop para preencher os valores
For l = 2 To l_fim
    query = query & vbNewLine & "("
    
    For c = 1 To c_fim
        If c = c_fim Then
          If Cells(l, c) = 0 Then 'Se o valor da célula for zero manter null
            query = query & "NULL" & ""
          Else
            query = query & "'" & Replace(Cells(l, c), ",", ".") & "'"
          End If
        Else
            If Cells(l, c) = 0 Then 'Se o valor da célula for zero manter null
                query = query & "NULL" & ", "
            Else
                query = query & "'" & Replace(Cells(l, c), ",", ".") & "', "
            End If
        End If
    Next
    
    If l = l_fim Then
        query = query & ")"
    Else
        query = query & "),"
    End If
Next

CopyText (query)

End Sub