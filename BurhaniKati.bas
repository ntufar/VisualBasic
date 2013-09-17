Attribute VB_Name = "BurhaniKati"
Dim CurPage As Integer

Private Sub DoHeading(ByVal IsOdd As Integer, ByRef text As String, ByVal pageoffset As Integer)

    Dim dummy As String
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Or ActiveWindow.ActivePane.View.Type _
         = wdMasterView Then
        ActiveWindow.ActivePane.View.Type = wdPageView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HeaderFooter.LinkToPrevious = False
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Or ActiveWindow.ActivePane.View.Type _
         = wdMasterView Then
        ActiveWindow.ActivePane.View.Type = wdPageView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.WholeStory
    dummy = CurPage + pageoffset - 1
    'If IsOdd Then
    '    Selection.text = text + "  " + dummy
    'Else
    '    Selection.text = dummy + "  " + text
    'End If
    Selection.text = text
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 9.5
    Selection.Font.bold = True
    Selection.HeaderFooter.LinkToPrevious = False
    'Okul sozlugu icin
    'With Selection.ParagraphFormat
    '    .SpaceBefore = 120
    'End With

    Selection.HeaderFooter.LinkToPrevious = False
    If IsOdd Then
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Else
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub


Sub BurhaniKatiTDKBaslik()
'
' TDKHeading Macro
' Macro created 02.03.98 by CA
' Modified for Okul Sozlugu by NT
'
    Dim NumPages As Integer
    Dim dummy As Integer
    Dim bigentry As Boolean
    Dim text As String
    Dim pageoffset As Integer
    
    'ilksayfaform.Show
    
    'pageoffset = ilksayfaform.ilksayfa.text
    pageoffset = 1
    'MsgBox pageoffset
    
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Repaginate
    ActiveWindow.ActivePane.View.ShowAll = True
    CurPage = 2
    NumPages = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticPages)
    Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, count:=1, Name:=""
    
    While CurPage <= NumPages
            CurPage = Selection.Information(wdActiveEndPageNumber)
            'Selection.HomeKey Unit:=wdLine
            'oldsel = Selection.Start
            'Selection.StartOf Unit:=wdParagraph, Extend:=wdMove
            'If Selection.StartOf <> oldsel Then
             '   Selection.InsertBreak Type:=wdSectionBreakContinuous
              '  Selection.MoveDown Unit:=wdParagraph
            'Else
             '   Selection.InsertBreak Type:=wdSectionBreakContinuous
            'End If
            'Selection.InsertBreak Type:=wdSectionBreakOddPage
            dobreak

            bigentry = False
            If (CurPage Mod 2) <> 0 Then
                ' tek
                Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, count:=1, Name:=""
                Selection.StartOf Unit:=wdParagraph
                dummy = Selection.Information(wdActiveEndPageNumber)
                If dummy <> CurPage Then
                        Selection.MoveUp Unit:=wdParagraph
                        Selection.StartOf Unit:=wdParagraph
                End If
            Else
                ' cift
                Selection.StartOf Unit:=wdParagraph, Extend:=wdMove
                If Selection.Font.bold = False Then
                    Selection.MoveDown Unit:=wdParagraph
                End If
                dummy = Selection.Information(wdActiveEndPageNumber)
                If dummy <> CurPage Then bigentry = True
            End If
            
            If bigentry = True Then
                Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, Name:=CurPage
                Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
                Selection.Copy
            Else
                Selection.StartOf Unit:=wdParagraph, Extend:=wdMove
            
                Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
                While Selection.Font.bold = True
                    If Mid(Selection.text, Len(Selection.text), 1) = "," _
                        Or Mid(Selection.text, Len(Selection.text), 1) = "(" _
                        Or Mid(Selection.text, Len(Selection.text), 1) = "1" _
                    Then
                        GoTo n723
                    End If
                    Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
                Wend
n723:
                Selection.MoveLeft Unit:=wdCharacter, Extend:=wdExtend
                TextToCopy = Selection.text
                text = Selection.text
            End If
            

            
    '        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
     '       Selection.HeaderFooter.LinkToPrevious = False
            
      '      Selection.InsertAfter Text:=TextToCopy
       '     ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        '    Selection.WholeStory
         '   Selection.Paste
            
            If (CurPage Mod 2) = 0 Then
                Call DoHeading(0, text, pageoffset)
            Else
                Call DoHeading(1, text, pageoffset)
            End If
            
            ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
            ActiveDocument.Repaginate
            Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, count:=1, Name:=""
            'Selection.InsertBreak Type:=wdSectionBreakContinuous
            'CurPage = Selection.Information(wdActiveEndPageNumber)
            CurPage = CurPage + 1
    Wend

End Sub

Private Sub dobreak()
'
' dobreak Makro
' Makro, SOZLUK1 tarafýndan 05.10.99 tarihinde kaydedildi
'
    'MsgBox wdSectionBreakContinuous
    'MsgBox wdSectionOddPage
    'MsgBox wdSectionContinuous
    'MsgBox ActiveDocument.PageSetup.SectionStart
    'dodoc
    Selection.InsertBreak Type:=3 'wdSectionBreakContinuous
    If Selection.Font.bold = False Then
        With Selection.ParagraphFormat
            .FirstLineIndent = CentimetersToPoints(0)
        End With
    End If
    'MsgBox wdSectionContinuous
    'MsgBox ActiveDocument.PageSetup.SectionStart

End Sub


Private Sub dodoc()
'
' beisil Makro
' Makro, SOZLUK1 tarafýndan 05.10.99 tarihinde kaydedildi
'
    With ActiveDocument.PageSetup
        '.LineNumbering.Active = False
        '.Orientation = wdOrientPortrait
        '.TopMargin = CentimetersToPoints(6.15)
        '.BottomMargin = CentimetersToPoints(6.15)
        '.LeftMargin = CentimetersToPoints(4.7)
        '.RightMargin = CentimetersToPoints(4.7)
        '.Gutter = CentimetersToPoints(0)
        '.HeaderDistance = CentimetersToPoints(1.25)
        '.FooterDistance = CentimetersToPoints(5.15)
        '.PageWidth = CentimetersToPoints(21.59)
        '.PageHeight = CentimetersToPoints(27.94)
        '.FirstPageTray = wdPrinterDefaultBin
        '.OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionContinuous
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = wdUndefined
    End With
End Sub



