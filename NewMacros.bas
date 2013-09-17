Attribute VB_Name = "NewMacros"
Sub TDK_Arapça()
'
' TDK_Arapça Makro
' Makro, Nikolay Tufar tarafýndan 07.08.98 tarihinde kaydedildi
'
   ' If Selection.Font.Name = "Arapca (TDK-3)" Then
   '     MsgBox "Arapca!!!!"
'    TDK_Arapca.Arab = "Holario!!"
    TDK_Arapca.Show
  '  End If
    
End Sub
Sub toTimes()
Attribute toTimes.VB_Description = "Makro, Nikolay Tufar tarafýndan 26.10.98 tarihinde kaydedildi"
Attribute toTimes.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.toTimes"
'
' toTimes Makro
' Makro, Nikolay Tufar tarafýndan 26.10.98 tarihinde kaydedildi
'
  '  Selection.Start = 1
    While 1 = 1
        If Selection.Font.Name <> "Arapca (TDK-3)" Then
            With Selection.Font
                .Name = "Times New Roman"
                .Size = 10
            End With
        End If
        Selection.MoveRight Unit:=wdCharacter, count:=1
    Wend
End Sub
Sub FontArapca_BOLD()
Attribute FontArapca_BOLD.VB_Description = "Makro, SOZLUK1 tarafýndan 29.07.99 tarihinde kaydedildi"
Attribute FontArapca_BOLD.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro1"
'
' Makro1 Makro
' Makro, SOZLUK1 tarafýndan 29.07.99 tarihinde kaydedildi
'
    Selection.Font.Name = "Arapca (TDK-3)"
    Selection.Font.bold = True
End Sub
Sub addparanthesis()

    Dim i As Integer
    
    'Selection.HomeKey unit:=wdStory
    Do While True
        Selection.StartOf Unit:=wdParagraph
        Do
            Selection.MoveRight Unit:=wdCharacter
        Loop Until Selection.Font.Name = "Arapca (TDK-3)"
        
        Selection.MoveLeft Unit:=wdCharacter
        Selection.InsertAfter " ("
        Selection.Font.bold = False
        Selection.MoveRight Unit:=wdCharacter
        
        Do
            Selection.MoveRight Unit:=wdCharacter
        Loop While Selection.Font.Name = "Arapca (TDK-3)"
        
        'Do
        '    Selection.MoveLeft unit:=wdCharacter
        'Loop While Selection.Text = " "
        
        Selection.InsertBefore ")"
        Selection.Font.bold = False
        
        Selection.MoveLeft Unit:=wdCharacter, count:=3
        For i = 1 To 5
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
            If Selection.text = ":" Then
                    'MsgBox "holario"
                    Selection.Delete
                    Exit For
            End If
        Next i
        
        i = Selection.MoveDown(Unit:=wdParagraph)
        If i = 0 Then Exit Do
    Loop 'while true
End Sub
Sub removespacesfromarab()
    Dim i As Integer
    Dim found As Boolean
    
    Selection.HomeKey Unit:=wdStory
    Do
        found = False
        Selection.StartOf Unit:=wdParagraph
        For i = 1 To 50
            Selection.MoveRight Unit:=wdCharacter
            If Selection.text = "(" Then
                found = True
                Exit For
            End If
        Next i
        
        Do
            If Selection.text = ")" Then Exit Do
            If Selection.text = " " And _
                    Selection.Font.Name = "Times New Roman" Then
                Selection.Delete
            End If
            
            Selection.MoveRight Unit:=wdCharacter
        Loop While True
        
        i = Selection.MoveDown(Unit:=wdParagraph)
        If i = 0 Then Exit Do
    Loop While True
End Sub

Sub pageview()
ActiveWindow.ActivePane.View.Type = wdPageView
End Sub

Sub linktoprev()
Selection.HeaderFooter.LinkToPrevious = False

End Sub

Sub bkdizin()
Dim pars As Integer
Dim i As Integer
Dim n As Integer

pars = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs)

Selection.HomeKey Unit:=wdStory

For i = 1 To pars
    Selection.StartOf Unit:=wdParagraph
    n = 0
    If Selection.Font.Name <> "Arapca (TDK-3)" Then GoTo nt123
    While Selection.Font.Name = "Arapca (TDK-3)"
        Selection.MoveRight Unit:=wdCharacter
        n = n + 1
        If n = 100 Then GoTo nt123
    Wend

    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, count:=1, _
            Extend:=wdExtend
    Selection.Cut
    Selection.StartOf Unit:=wdParagraph
    Selection.Paste
    Selection.InsertAfter Chr(9)
nt123:
    'Selection.EndOf unit:=wdParagraph
    Selection.MoveDown Unit:=wdParagraph, count:=1
    
Next i

End Sub
Sub Makro2()
Attribute Makro2.VB_Description = "Makro, SOZLUK1 tarafýndan 01.11.99 tarihinde kaydedildi"
Attribute Makro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro2"
'
' Makro2 Makro
' Makro, SOZLUK1 tarafýndan 01.11.99 tarihinde kaydedildi
'
    Selection.Sections(1).Headers(1).PageNumbers.Add PageNumberAlignment:= _
        wdAlignPageNumberOutside, FirstPage:=True
End Sub
Sub Makro3()
Attribute Makro3.VB_Description = "Makro, SOZLUK1 tarafýndan 15.06.00 tarihinde kaydedildi"
Attribute Makro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro3"
'
' Makro3 Makro
' Makro, SOZLUK1 tarafýndan 15.06.00 tarihinde kaydedildi
'
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.Font.Size = 12
End Sub
Sub Once0nk()
Attribute Once0nk.VB_Description = "Makro, SOZLUK1 tarafýndan 11.07.00 tarihinde kaydedildi"
Attribute Once0nk.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Once0nk"
'
' Once0nk Makro
' Makro, SOZLUK1 tarafýndan 11.07.00 tarihinde kaydedildi
'
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End Sub
Sub Siir1()
Attribute Siir1.VB_Description = "Makro, SOZLUK1 tarafýndan 11.07.00 tarihinde kaydedildi"
Attribute Siir1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Siir1"
'
' Siir1 Makro
' Makro, SOZLUK1 tarafýndan 11.07.00 tarihinde kaydedildi
'
    ActiveWindow.ActivePane.SmallScroll Down:=3
    Selection.TypeBackspace
    Selection.TypeParagraph
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.MoveUp Unit:=wdLine, count:=3
    Selection.MoveDown Unit:=wdLine, count:=8
    Selection.MoveUp Unit:=wdLine, count:=3
    Selection.TypeBackspace
    Selection.TypeParagraph
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.TypeBackspace
    Selection.TypeParagraph
End Sub

Sub oncekigibi()
    Selection.HeaderFooter.LinkToPrevious = False
End Sub
Sub Makro4()
Attribute Makro4.VB_Description = "Makro, SOZLUK1 tarafýndan 29.08.00 tarihinde kaydedildi"
Attribute Makro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro4"
'
' Makro4 Makro
' Makro, SOZLUK1 tarafýndan 29.08.00 tarihinde kaydedildi
'
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Selection.Font.Name = "Arapca (TDK-3)"
    Selection.Font.Size = 24
End Sub
Sub DIZIN_isaretle()
Attribute DIZIN_isaretle.VB_Description = "Makro, SOZLUK1 tarafýndan 19.12.00 tarihinde kaydedildi"
Attribute DIZIN_isaretle.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.DIZIN_isaretle"
'
' DIZIN_isaretle Makro
' Makro, SOZLUK1 tarafýndan 19.12.00 tarihinde kaydedildi
'
    Selection.Copy
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
        "XE ", PreserveFormatting:=False
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.TypeText text:=""""
    Selection.Paste
    Selection.TypeText text:=""""
    'Selection.MoveDown Unit:=wdLine, count:=1
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
        "TC ", PreserveFormatting:=False
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.TypeText text:=""""
    Selection.Paste
    Selection.TypeText text:=""" \f m "
    Selection.MoveDown Unit:=wdLine, count:=1

End Sub
Sub Makro1()
Attribute Makro1.VB_Description = "Makro, SOZLUK1 tarafýndan 03.01.01 tarihinde kaydedildi"
Attribute Makro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro1"
'
' Makro1 Makro
' Makro, SOZLUK1 tarafýndan 03.01.01 tarihinde kaydedildi
'
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
        "TC ", PreserveFormatting:=False
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.TypeText text:=""""
    Selection.Paste
    Selection.TypeText text:=""" \f m "
    Selection.MoveDown Unit:=wdLine, count:=1
End Sub
Sub Makro5()
Attribute Makro5.VB_Description = "Makro, SOZLUK1 tarafýndan 27.02.01 tarihinde kaydedildi"
Attribute Makro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro5"
'
' Makro5 Makro
' Makro, SOZLUK1 tarafýndan 27.02.01 tarihinde kaydedildi
'
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdCharacter, count:=4, Extend:=wdExtend
    Selection.TypeText text:=vbTab
    Selection.Font.Name = "Arapca (TDK-3)"
    Selection.TypeText text:="s"
    'Selection.MoveUp Unit:=wdLine, count:=1
    'Selection.MoveDown Unit:=wdLine, count:=2
    'Selection.HomeKey Unit:=wdLine
    'Selection.MoveDown Unit:=wdLine, count:=1
    'Selection.MoveUp Unit:=wdLine, count:=1
    'Selection.MoveDown Unit:=wdLine, count:=1
    'Selection.MoveUp Unit:=wdLine, count:=3
    'Selection.MoveDown Unit:=wdLine, count:=2
End Sub
