Attribute VB_Name = "OkulSozlugu"
Dim CurPage As Integer

Sub TDKHeading()
'
' TDKHeading Macro
' Macro created 02.03.98 by CA
' Modified for Okul Sozlugu by NT
'
    
    Dim TempCurPage As Integer
    Dim TempCurPage2 As Integer
    Dim NumPages As Integer
    Dim SelStartOdd As Long
    Dim SelEndOdd As Long
    Dim SelStartEven As Long
    Dim SelEndEven As Long
    Dim DoAllign As Integer
    
    Selection.HomeKey Unit:=wdStory 'Pass 1
    ActiveDocument.Repaginate
    ActiveWindow.ActivePane.View.ShowAll = True
    CurPage = 0
    NumPages = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticPages)
    While CurPage <= NumPages
        Selection.Find.ClearFormatting
        With Selection.Find
            .text = "�"
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        If Selection.Find.found Then
            Selection.Delete Unit:=wdCharacter, count:=1
            Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
            Selection.Font.Hidden = True
        Else
            CurPage = NumPages + 1
        End If
    Wend
    
    Selection.HomeKey Unit:=wdStory 'Pass 2
    ActiveDocument.Repaginate
    For i = 2 To NumPages
        Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, count:=1, Name:=""
        Selection.InsertBreak Type:=wdSectionBreakContinuous
        DoAllign = False
        Selection.MoveDown Unit:=wdParagraph, count:=1, Extend:=wdExtend
        SelStartEven = Selection.Start
        SelEndEven = Selection.End
        If InStr(Selection.text, ("�")) <= 0 Then
            Selection.Find.ClearFormatting
            Selection.Find.Font.Hidden = True
            With Selection.Find
                .text = ""
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute
            If Selection.Find.found Then
                If Not (Selection.Start >= SelStartEven And Selection.Start <= SelEndEven) Then
                    DoAllign = True
                End If
            Else
                DoAllign = True
            End If
        End If
        Selection.Start = SelStartEven
        Selection.End = SelStartEven
        If DoAllign Then
            Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
        End If
    Next i
    
    Selection.HomeKey Unit:=wdStory 'Pass 3
    ActiveDocument.Repaginate
    CurPage = 0
    SelStartOdd = -1
    NumPages = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticPages)
    While CurPage <= NumPages
        Selection.Find.ClearFormatting
        Selection.Find.Font.Hidden = True
        With Selection.Find
            .text = ""
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        
        If Selection.Find.found Then
            Selection.Font.Hidden = False
            Selection.MoveRight Unit:=wdCharacter, count:=1
            TempCurPage = Selection.Information(wdActiveEndPageNumber)
            If TempCurPage <> CurPage And TempCurPage <> 1 Then
                If (TempCurPage Mod 2) = 0 Then
                    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                    SelStartEven = Selection.Start
                    SelEndEven = Selection.End
                    If TempCurPage <> 2 Then
                        Selection.Start = SelStartOdd
                        Selection.End = SelEndOdd
                        Selection.Copy
                        If (Selection.Information(wdActiveEndPageNumber) Mod 2) = 0 Then
                            Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, count:=1, Name:=""
                            Selection.MoveRight Unit:=wdCharacter, count:=1
                        End If
                        Call DoHeading(True)
                        SelStartEven = SelStartEven
                        SelEndEven = SelEndEven
                    End If
                    Selection.Start = SelStartEven
                    Selection.End = SelEndEven
                    Selection.Copy
                    Selection.MoveRight Unit:=wdCharacter, count:=1
                    Call DoHeading(False)
                Else
                    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                    SelStartOdd = Selection.Start
                    SelEndOdd = Selection.End
                End If
                CurPage = TempCurPage
            Else
                Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                SelStartOdd = Selection.Start
                SelEndOdd = Selection.End
            End If
        Else
            CurPage = NumPages + 1
        End If
    Wend
    
    If (Selection.Information(wdActiveEndPageNumber) Mod 2) <> 0 Then
        Selection.Start = SelStartOdd
        Selection.End = SelEndOdd
        'Selection.Copy
        Call DoHeading(True)
    End If
  
    Selection.HomeKey Unit:=wdStory
    ActiveWindow.ActivePane.View.ShowAll = False
End Sub
Sub DoHeading(ByVal IsOdd As Integer, ByRef text As String, ByVal pageoffset As Integer)

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
    If IsOdd Then
        Selection.text = text + "  " + dummy
    Else
        Selection.text = dummy + "  " + text
    End If
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


Sub TDKBaslik()
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
    
    ilksayfaform.Show
    
    pageoffset = ilksayfaform.ilksayfa.text
    
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
' Makro, SOZLUK1 taraf�ndan 05.10.99 tarihinde kaydedildi
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
' Makro, SOZLUK1 taraf�ndan 05.10.99 tarihinde kaydedildi
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

