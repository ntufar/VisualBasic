Attribute VB_Name = "Resimli"
Sub Resimli_maddebasi_normal()
    
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
            'Selection.Font.bold = False
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
        Loop While Selection.Font.bold = True
        Selection.MoveLeft Unit:=wdCharacter, count:=2, Extend:=wdExtend
        Selection.Font.bold = False
        'MsgBox (Selection.text)
        Selection.Collapse Direction:=wdCollapseEnd
        If Selection.text = " " Then
            Selection.Delete Unit:=wdCharacter, count:=1
        End If
        Selection.InsertAfter ("|")
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        Selection.Delete
        
        'MsgBox ("Holario")
        
        Do
            Selection.Font.bold = False
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        Loop While Selection.text <> ")"
        Selection.Font.bold = False
        
        Selection.text = "|"
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        Selection.Delete
        If Selection.text = " " Then
            Selection.Delete Unit:=wdCharacter, count:=1
        End If
        'MsgBox ("Holario1")
        
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
End Sub

Sub Resimli_bold_italik()
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



Sub Resimli_maddebasi_normal_arapca()
    
    Dim i As Integer
    Dim found As Boolean
    Dim ptr As Integer
    Dim pras As Long
    
    Selection.HomeKey Unit:=wdStory
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    'MsgBox pars
    
    Do
        'found = False
        Selection.StartOf Unit:=wdParagraph
        Selection.HomeKey Unit:=wdLine, Extend:=wdMove
        'MsgBox ("H")
        Do
            'Selection.Font.bold = False
            'MsgBox (Selection.text)
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
            'Selection.Collapse Direction:=wdCollapseEnd
            'MsgBox Selection.text
            
            'MsgBox Mid(Selection.text, Len(Selection.text), 1)
        Loop While Selection.Font.bold = True And Selection.text <> "1" _
            And Selection.text <> Chr(9)
        
        
        
        Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdMove
        'MsgBox "holario"
        If Selection.text = " " Then
            Selection.Delete Unit:=wdCharacter, count:=1
        Else
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        End If
        'Selection.InsertAfter ("|" + Chr(&H86) + "|") '"Re"
        'Selection.InsertAfter ("|" + Chr(&H91) + "|") '"Þin"
        'Selection.InsertAfter ("|" + Chr(&H92) + "|") '"Sat"
        'Selection.InsertAfter ("|" + Chr(&HF7) + "|") '"Dat"
        'Selection.InsertAfter ("|" + Chr(&HB9) + "|") '"Tý"
        'Selection.InsertAfter ("|" + Chr(&HB4) + "|") '"Zý"
        'Selection.InsertAfter ("|" + Chr(&H9F) + "|") '"Ayin"
        'Selection.InsertAfter ("|" + Chr(&H8D) + "|") '"Gayin"
        'Selection.InsertAfter ("|" + Chr(&H2D) + "|") '"Fe"
        Selection.InsertAfter ("|" + Chr(&H82) + "|") '"Kaf"
        'Selection.InsertAfter ("|" + Chr(&H84) + "|") '"Kef"
        'Selection.InsertAfter ("|" + Chr(&H89) + "|") '"Lâm"
        'Selection.InsertAfter ("|" + Chr(&HC2) + "|") '"Mim"
        'Selection.InsertAfter ("|" + Chr(&HCA) + "|") '"Nun"
        'Selection.InsertAfter ("|" + Chr(&HCB) + "|") '"Vav"
        'Selection.InsertAfter ("|" + Chr(&H88) + "|") '"He"
        Selection.Collapse Direction:=wdCollapseEnd
        If Selection.text = " " Then
            Selection.Delete Unit:=wdCharacter, count:=1
        End If
        Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
        Selection.Font.bold = False
        'MsgBox Selection.text
        
        'Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        'Selection.Delete
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        'If i <> 1 Then Exit Do
        If ptr = pars Then Exit Do
        ptr = ptr + 1
        'MsgBox "holario"
    Loop While True
End Sub

