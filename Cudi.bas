Attribute VB_Name = "Cudi"
Sub Cudi_maddebasi_normal()
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    'MsgBox pars
    
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        Do
            Selection.Font.bold = False
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        Loop While Selection.text <> Chr(9)
        
        
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
End Sub

Sub Cudi_bold_italik()
    Dim ptr As Long
    Dim chars As Long
    Dim bold As Boolean
    Dim italic As Boolean
        
    
    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    ptr = 1
    chars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces)
    'MsgBox pars
    
    bold = False
    italic = False
    Do
        If Selection.Font.bold = True And bold = False Then
            Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdMove
            Selection.InsertBefore ("%b")
            Selection.MoveRight Unit:=wdCharacter, count:=2, Extend:=wdMove
            bold = True
        End If
        If Selection.Font.bold = False And bold = True Then
            Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdMove
            Selection.InsertBefore ("%0b")
            Selection.MoveRight Unit:=wdCharacter, count:=2, Extend:=wdMove
            bold = False
        End If
        
        If Selection.Font.italic = True And italic = False Then
            Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdMove
            Selection.InsertBefore ("%i")
            Selection.MoveRight Unit:=wdCharacter, count:=2, Extend:=wdMove
            italic = True
        End If
        If Selection.Font.italic = False And italic = True Then
            Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdMove
            Selection.InsertBefore ("%0i")
            Selection.MoveRight Unit:=wdCharacter, count:=2, Extend:=wdMove
            italic = False
        End If
        
        
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        
        If ptr = chars Then Exit Do
        ptr = ptr + 1
    Loop While True
End Sub


