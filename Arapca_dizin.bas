Attribute VB_Name = "Arapca_dizin"
    Dim Latin() As String
    Dim Arap() As String
    
Sub arapcadizin()
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    'Arapcadizinform.Show
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    ReDim Latin(pars + 100) As String
    
    ReDim Arap(pars + 100) As String
    'MsgBox pars
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        'Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Loop While Selection.Font.bold = True
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Latin(ptr) = Selection.text
        strip Latin(ptr)
        'MsgBox ">>>" + Latin + "<<<<"
        'Arapcadizinform.Latin.Caption = Latin
        
        
        For i = 1 To 50
            Selection.MoveRight Unit:=wdCharacter
            If Selection.Font.Name = "Arapca (TDK-3)" Then
                found = True
                Exit For
            End If
        Next i
        
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdMove
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Loop While Selection.Font.Name = "Arapca (TDK-3)"
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
        
        'Arap(ptr) = normalizearapword(Selection.text)
        Arap(ptr) = Selection.text
        strip Arap(ptr)
        'Arapcadizinform.Arap.Caption = Arap
        'MsgBox Arap
        
        'Documents("dizin.doc").Activate
        'Selection.InsertAfter Arap
        'Selection.Font.Name = "Arapca (TDK-3)"
        'Selection.Collapse Direction:=wdCollapseEnd
        'Selection.InsertAfter Chr(9)
        'If Len(Arap) < 10 Then Selection.InsertAfter Chr(9)
        'Selection.InsertAfter Latin
        'Selection.Font.Name = "Times New Roman"
        'Selection.Collapse Direction:=wdCollapseEnd
        'Selection.InsertAfter Chr(13)
        'Documents(Realdoc).Activate
        
        'Selection.StartOf Unit:=wdParagraph
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
    writedizin (ptr)
End Sub

Static Sub strip(ByRef str As String)

    Dim i As Integer
    
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) <> " " Then
            Exit For
        End If
        str = Left(str, i - 1)
    Next i
End Sub

Private Static Sub writedizin(count As Integer)

    elifbainit
    'sortdizin (count)
    
    For i = 1 To count + 1
        Documents("dizin.doc").Activate
        Selection.InsertAfter Arap(i)
        Selection.Font.Name = "Arapca (TDK-3)"
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.InsertAfter Chr(9)
        'If Len(Arap) < 10 Then Selection.InsertAfter Chr(9)
        Selection.InsertAfter Latin(i)
        Selection.Font.Name = "Times New Roman"
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.InsertAfter Chr(13)
    Next i
    

End Sub

Private Static Sub sortdizin(count As Integer)

    Dim changed As Boolean
    Dim i As Long
    Dim result As Boolean
    Dim buf As String
    
    Do
        changed = False
        For i = 0 To count
            result = arcmp(Arap(i), Arap(i + 1))
        
            If result = False Then
                buf = Arap(i)
                Arap(i) = Arap(i + 1)
                Arap(i + 1) = buf
            
                buf = Latin(i)
                Latin(i) = Latin(i + 1)
                Latin(i + 1) = buf
            
                changed = True
                i = 0
            End If
        Next i
    Loop While changed = False
End Sub
Sub Cudi_arapcadizin()
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    'Arapcadizinform.Show
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    ReDim Latin(pars + 100) As String
    
    ReDim Arap(pars + 100) As String
    'MsgBox pars
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        'Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
        
        Do
            Selection.MoveRight Unit:=wdCharacter
        Loop While Selection.text <> "|"
        firstword = ""
        Do
                Selection.MoveRight Unit:=wdCharacter, Extend:=wdMove
                If Selection.text <> "|" Then
                    firstword = firstword + Selection.text
                End If
        Loop While Selection.text <> "|"
        MsgBox firstword

        
        Arap(ptr) = firstword
        strip Arap(ptr)
        Selection.StartOf Unit:=wdParagraph
        Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
        Latin(ptr) = Selection.text
        
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
    cudiwritedizin (ptr)
End Sub

Private Static Sub cudiwritedizin(count As Integer)

    elifbainit
    'sortdizin (count)
    
    For i = 1 To count + 1
        Documents("dizin.doc").Activate
        Selection.InsertAfter Latin(i)
        Selection.Collapse Direction:=wdCollapseEnd
        'Selection.InsertAfter Chr(13)
    Next i
    

End Sub

