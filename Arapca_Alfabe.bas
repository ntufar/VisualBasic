Attribute VB_Name = "Arapca_Alfabe"
Dim Words() As String
Dim Para() As Integer

Sub arapcaalfabe()
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    'Arapcadizinform.Show
    ptr = 0
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    ReDim Words(pars + 100) As String
    ReDim Para(pars + 100) As Integer
    MsgBox pars
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        'Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Loop While Selection.Font.bold = True
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
        'Latin(ptr) = Selection.Text
        'strip Latin(ptr)
        
        
        '----------------------------
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
        
        Words(ptr) = Selection.text
        strip Words(ptr)
        Para(ptr) = ptr

        '----------------------------
        
        
        
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
        If ptr = pars - 1 Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
    writerezult (ptr)
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

Private Static Sub writerezult(count As Integer)

    elifbainit
    sortdizin (count)
    
    Realdoc = ActiveWindow.Document
    Documents("dizin.doc").Activate
    For i = 0 To count
        
        Set myRange = Documents(Realdoc).Paragraphs(Para(i)).Range
        myRange.Copy
        
        Selection.Paste
        Selection.Collapse Direction:=wdCollapseEnd
        
        


        'Selection.InsertAfter Arap(i)
        'Selection.Font.Name = "Arapca (TDK-3)"
        'Selection.Collapse Direction:=wdCollapseEnd
        'Selection.InsertAfter Chr(9)
        'If Len(Arap) < 10 Then Selection.InsertAfter Chr(9)
        'Selection.InsertAfter Latin(i)
        'Selection.Font.Name = "Times New Roman"
        'Selection.Collapse Direction:=wdCollapseEnd
        'Selection.InsertAfter Chr(13)
    Next i
    

End Sub

Private Static Sub sortdizin(count As Integer)

    Dim changed As Boolean
    Dim i As Long
    Dim result As Boolean
    Dim buf As String
    Dim x As Integer
    
    Do
        changed = False
        For i = 0 To count
            result = arcmp(Words(i), Words(i + 1))
        
            If result = False Then
                buf = Words(i)
                Words(i) = Words(i + 1)
                Words(i + 1) = buf
            
                x = Para(i)
                Para(i) = Para(i + 1)
                Para(i + 1) = x
                
                
            
                changed = True
            End If
        Next i
    Loop While changed = True
End Sub

