Attribute VB_Name = "IP"
Sub IP1()
    
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    Realdoc = ActiveWindow.Document
    
    
    'Selection.HomeKey Unit:=wdStory
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    'MsgBox pars
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        Loop While Selection.text <> ")"
        Selection.MoveRight Unit:=wdCharacter, count:=2, Extend:=wdMove
        'MsgBox "aaa"
        Selection.MoveEnd Unit:=wdParagraph
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Selection.Delete
        Selection.TypeText text:=Chr(11) & Chr(11) & Chr(11) & Chr(11)
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
End Sub

