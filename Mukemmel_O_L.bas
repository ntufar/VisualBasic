Attribute VB_Name = "Mukemmel_O_L"
    Dim Latin() As String
    Dim Arap() As String
    Dim GercekArap() As String
    
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
    ReDim GercekArap(pars + 100) As String
    'MsgBox pars
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        'Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Loop While Selection.Font.Name = "Arapca (TDK-3)"
        Selection.MoveLeft Unit:=wdCharacter, count:=2, Extend:=wdExtend
        GercekArap(ptr) = Selection.text
        Arap(ptr) = normalizearapword(Selection.text)
        strip Arap(ptr)
        'MsgBox ">>>" + Arap(ptr) + "<<<<"
        'Latin(ptr) = Selection.text
            'MsgBox ">>>" + Arap(ptr) + "<<<<"
            'Arapcadizinform.Latin.Caption = Latin
        
        
        For i = 1 To 50
            Selection.MoveRight Unit:=wdCharacter, Extend:=wdMove
            'If Selection.Font.Name = "Times New Roman" Then
            '    found = True
            '    Exit For
            'End If
            If Selection.text = Chr(9) Then
                found = True
                Exit For
            End If
        Next i
        
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        'Do
        '    Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        'Loop While Selection.Font.Name = "Arapca (TDK-3)"
        'Selection.MoveLeft Unit:=wdCharacter, count:=2, Extend:=wdExtend
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
        
        Latin(ptr) = Selection.text
        'GercekArap(ptr) = Selection.text
        'Arap(ptr) = Selection.text
        'strip Arap(ptr)
        'Arap(ptr) = normalizearapword(Arap(ptr))
        'Arapcadizinform.Arap.Caption = Arap
        'MsgBox Latin(ptr)
        
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
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) = "[" Or Mid(str, i, 1) = "[" Then
            Mid(str, i, 1) = " "
        End If
    Next i
'    For i = 1 To Len(str)
'        If Mid(str, i, 1) = Chr(&H20) Then
'            Mid(str, i, Len(str) - i) = Mid(str, i + 1, Len(str) - i - 1)
'        End If
'    Next i
End Sub

Private Static Sub writedizin(count As Integer)

    elifbainit
    sortdizin (count)
    
    For i = 1 To count
        Documents("dizin.doc").Activate
        Selection.InsertAfter GercekArap(i)
        Selection.Font.Name = "Arapca (TDK-3)"
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.InsertAfter Chr(9)
        'If Len(Arap) < 10 Then Selection.InsertAfter Chr(9)
        Selection.InsertAfter Latin(i)
        Selection.Font.Name = "Arial" '"Times New Roman"
        Selection.Font.Size = 8
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.InsertAfter Chr(13)
    Next i
    Documents("dizin.doc").Save
    

End Sub

Private Static Sub sortdizin(count As Integer)

    Dim changed As Boolean
    Dim i As Long
    Dim result As Boolean
    Dim buf As String
    
    Do
        changed = False
        i = 2
        Do
            If i < 2 Then i = 2
            result = arcmp(Arap(i - 1), Arap(i))
        
            If result = False Then
                buf = Arap(i - 1)
                Arap(i - 1) = Arap(i)
                Arap(i) = buf
                
                buf = GercekArap(i - 1)
                GercekArap(i - 1) = GercekArap(i)
                GercekArap(i) = buf
            
            
                buf = Latin(i - 1)
                Latin(i - 1) = Latin(i)
                Latin(i) = buf
            
                changed = True
                i = i - 20
                
            End If
            If i = count Then Exit Do
            i = i + 1
        Loop While True
    Loop While changed = True
End Sub

