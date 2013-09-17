Attribute VB_Name = "arapcalar"
Sub arapcaparantez()
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    
    
    Selection.HomeKey Unit:=wdStory
    'Arapcadizinform.Show
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        'Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        Loop While Selection.Text <> "("
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Loop While Mid(Selection.Text, Len(Selection.Text), 1) <> ")"
        
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
        
        Selection.Font.Name = "Arapca (TDK-3)"
        
        
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
End Sub

