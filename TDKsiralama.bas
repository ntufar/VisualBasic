Attribute VB_Name = "TDKsiralama"
Dim Words() As String
Dim Para() As Integer

Sub TDKSirala()

    Realdoc = ActiveWindow.Document
    
    
    Selection.HomeKey Unit:=wdStory
    'Arapcadizinform.Show
    ptr = 1
    pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)
    
    ReDim Words(pars + 10)
    ReDim Para(pars + 10)


    Do
        Selection.StartOf Unit:=wdParagraph
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        Para(ptr) = ptr
        Words(ptr) = Selection.text
        strip Words(ptr)
        'MsgBox Words(ptr)
        
        i = Selection.MoveDown(Unit:=wdParagraph, count:=1, Extend:=wdMove)
        
        If ptr = pars Then Exit Do
        ptr = ptr + 1
    Loop While True
    writeit (ptr)
    'If changed = True Then GoTo doitagain
    MsgBox "bitti"
End Sub

'---------------------- STRIP -----------------

Private Static Sub strip(ByRef str As String)

    Dim i As Integer
    
    For i = 1 To Len(str)
        If Mid(str, i, 1) <> " " And Mid(str, i, 1) <> Chr(9) Then
            Exit For
        End If
        str = Right(str, Len(str) - 1)
        'MsgBox "Holario >>>" + str + "<<<"
        
        i = i - 1
    Next i
    
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) <> " " And Mid(str, i, 1) <> Chr(9) Then
            Exit For
        End If
        str = Left(str, i - 1)
        'MsgBox "Holario1 >>>" + str + "<<<"
        Exit For
    Next i
    
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) = "-" Then
            Mid(str, i, Len(str) - i) = Right(str, Len(str) - i)
            str = Left(str, Len(str) - 1)
            'MsgBox str
        End If
    Next i
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) = " " Then
            Mid(str, i, Len(str) - i) = Right(str, Len(str) - i)
            str = Left(str, Len(str) - 1)
            MsgBox str
        End If
    Next i
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) = "’" Or Mid(str, i, 1) = "'" Then
            Mid(str, i, Len(str) - i) = Right(str, Len(str) - i)
            str = Left(str, Len(str) - 1)
            'MsgBox str
        End If
    Next i
    '-----------
    For i = Len(str) To 1 Step -1
        If Asc(Mid(str, i, 1)) = &H49 Then    'I
            Mid(str, i, 1) = "ý"
        Else
            If Asc(Mid(str, i, 1)) = &HDD Then    'Ý
                Mid(str, i, 1) = "i"
            Else
                If Mid(str, i, 1) = "â" Or Mid(str, i, 1) = "Â" Then
                    Mid(str, i, 1) = "a"
                Else
                    If Mid(str, i, 1) = "Î" Or Mid(str, i, 1) = "î" Then
                        Mid(str, i, 1) = "i"
                    Else
                        If Mid(str, i, 1) = "Û" Or Mid(str, i, 1) = "û" Then
                            Mid(str, i, 1) = "u"
                        Else
                            '---------
                             Mid(str, i, 1) = LCase(Mid(str, i, 1))
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub writeit(count As Integer)

    sortdizin (count)
    
    Realdoc = ActiveWindow.Document
    
    Documents("dizin.doc").Activate
    For i = 1 To count
        Set myRange = Documents(Realdoc).Paragraphs(Para(i)).Range
        myRange.Copy
        Selection.Paste
        Selection.Collapse Direction:=wdCollapseEnd
    Next i
    

End Sub
Private Static Sub sortdizin(count As Integer)

    Dim changed As Boolean
    Dim i As Long
    Dim result As Boolean
    Dim buf As String
    Dim ibuf As Integer
    
    
    Do
        changed = False
        For i = 1 To count - 1
            result = latcmp2(Words(i), Words(i + 1))
        
            'If result = True Then
             '   MsgBox words(i) + "  " + words(i + 1) + "  true"
            'Else
             '   MsgBox words(i) + "  " + words(i + 1) + "  false"
            'End If
            If result = False Then
                'MsgBox "holario"
                buf = Words(i)
                Words(i) = Words(i + 1)
                Words(i + 1) = buf
            
                ibuf = Para(i)
                Para(i) = Para(i + 1)
                Para(i + 1) = ibuf
            
                changed = True
                i = i - 10
                If i <= 0 Then i = 1
                'i = 0
            End If
        Next i
    Loop While changed = True
End Sub

Function latcmp2(ByVal first As String, ByVal second As String) As Boolean
    ' return true if first is >= second

    Dim fp As Integer
    Dim sp As Integer
    Dim lmt As Integer
    Dim hvf As String
    Dim hvs As String
    Dim p As Integer
    Dim i As Integer
    
    latcmp2 = True
    
    If Len(first) < Len(second) Then
        lmt = Len(first)
    Else
        lmt = Len(second)
    End If
        
    For p = 1 To lmt
        hvf = Mid(first, p, 1)
        hvs = Mid(second, p, 1)
        If (Asc(hvf) = &H69 And Asc(hvs) = &HEE) Or _
                (Asc(hvf) = &HDD And Asc(hvs) = &HCE) Or _
                (Asc(hvf) = &HEE And Asc(hvs) = &H69) Or _
                (Asc(hvf) = &HCE And Asc(hvs) = &HDD) Then
                    dummy = dummy 'nothing
        Else
            If hvf = "i" And hvs = "i" Then
                If (Asc(hvf) = &H69) And _
                    (Asc(hvs) = &HFD) Then
                        latcmp2 = False
                        Exit Function
                Else
                    If (Asc(hvs) = &H69) And _
                        (Asc(hvf) = &HFD) Then
                            latcmp2 = True
                            Exit Function
                    End If
                End If
            Else
                If (Asc(hvf) = &H61 And Asc(hvs) = &HE2) Or _
                    (Asc(hvf) = &H41 And Asc(hvs) = &HC2) Or _
                    (Asc(hvf) = &HE2 And Asc(hvs) = &H61) Or _
                    (Asc(hvf) = &HC2 And Asc(hvs) = &H41) Then
                        dummy = dummy 'nothing
                Else
                    i = StrComp(hvf, hvs, vbTextCompare) 'vbBinaryCompare)
                    If i = 1 Then
                        latcmp2 = False
                        Exit Function
                    Else
                        If i = -1 Then
                            latcmp2 = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next p
    
    If Len(first) > Len(second) Then
        latcmp2 = False
    Else
        latcmp2 = True
    End If

End Function



