Attribute VB_Name = "Dizin_araclari"
Option Compare Text
Dim Words() As String
Dim Para() As Integer


Sub Maddebasilari_cikart()
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    Dim fword As String
    Dim sword As String
    Dim cmp As Boolean
    Dim count As Integer
    
    
    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    
    ReDim Words(pars + 10)
    ReDim Para(pars + 10)
    'MsgBox pars
    

    Do
        Selection.StartOf Unit:=wdParagraph
        'Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
        
        count = 1
        Do
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
            count = count + 1
        Loop While (Selection.Font.bold = True Or _
                Mid(Selection.text, Len(Selection.text), 1) = " ") Or _
                count < 3
        Selection.MoveLeft Unit:=wdCharacter, count:=3, Extend:=wdExtend
        Words(ptr) = Selection.text
        Para(ptr) = ptr
        'MsgBox ">>>" + Words(ptr) + "<<<<"
        'Arapcadizinform.Latin.Caption = Latin
        
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        If ptr = pars Then Exit Do
        ptr = ptr + 1
    Loop While True
    'If changed = True Then GoTo doitagain
    writeit (ptr)
    MsgBox "bitti"
End Sub

Private Sub writeit(count As Integer)

    
    Realdoc = ActiveWindow.Document
    
    Documents("dizin.doc").Activate
    For i = 1 To count
        Selection.TypeText Words(i) + "   "
        Selection.TypeText Chr(13)
        'Selection.Collapse Direction:=wdCollapseEnd
        'Documents(Realdoc).Activate
        'Selection.HomeKey unit:=wdStory
        'Selection.MoveDown unit:=wdParagraph, count:=para(i) - 1, Extend:=wdMove
        'Set myRange = Documents(Realdoc).Paragraphs(Para(i)).Range
        'myRange.Copy
        'Selection.Expand unit:=wdParagraph
        'Selection.Copy
        'ActiveDocument.Paragraphs(para(i) + 1).Range.Copy
        'Documents("dizin.doc").Activate
        'Selection.Paste
        'Selection.Collapse Direction:=wdCollapseEnd
    Next i
    

End Sub

