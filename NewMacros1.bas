Attribute VB_Name = "NewMacros1"

Sub dizin1()
Attribute dizin1.VB_Description = "Makro, SOZLUK1 tarafýndan 24.06.99 tarihinde kaydedildi"
Attribute dizin1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.dizin1"
'
' dizin1 Makro
' Makro, SOZLUK1 tarafýndan 24.06.99 tarihinde kaydedildi
'
    'Selection.Start = 0
    'While Selection.Start <> Document.End
    Realdoc = ActiveWindow.Document
    'MsgBox Realdoc
    Documents("Dizin.doc").Activate
    Selection.HomeKey Unit:=wdStory
    Ext = False
    While Ext = False
        Documents("Dizin.doc").Activate
        Selection.MoveRight Unit:=wdSentence
        If Selection.MoveRight = 1 Then
            Selection.Expand Unit:=wdSentence
            TextToFind = Selection.text
            Documents(Realdoc).Activate
            Selection.HomeKey Unit:=wdStory
            'MsgBox TextToFind
            Selection.Find.ClearFormatting
            With Selection.Find
                .text = TextToFind
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Do While Selection.Find.Execute = True
                Selection.Copy
                Selection.MoveRight Unit:=wdCharacter, count:=2
                Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:= _
                    "XE ", PreserveFormatting:=False
                Selection.MoveLeft Unit:=wdCharacter, count:=2
                Selection.TypeText text:=""""
                Selection.Paste
                Selection.TypeText text:=""""
            Loop
        Else: Ext = True
        End If
    Wend
    Documents(Realdoc).Activate
End Sub

Sub tab1()

    Realdoc = ActiveWindow.Document
    Selection.HomeKey Unit:=wdStory
    Documents("yk.doc").Activate
    Selection.HomeKey Unit:=wdStory
    doexit = False
    
    i = 2
    While 1

        Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
        WordToLookFor = Selection.text
        
        Range
        ' in document
        Documents(Realdoc).Activate
        Selection.HomeKey Unit:=wdStory
        Do
            Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
            If Selection.text = WordToLookFor Then
                MsgBox "Found!! " + WordToLookFor
            End If
            Selection.MoveRight Unit:=wdSentence
            
            If Selection.MoveRight = 0 Then Exit Do
            
        Loop Until Frue
        ' in document (END)
        
        
        'If WordToLookFor = "zoka" Then
         '   MsgBox "Bitti"
          '  GoTo eltable
        'End If
        Documents("yk.doc").Activate
        Selection.MoveRight Unit:=wdSentence
        If Selection.MoveRight = 0 Then GoTo eltable
    Wend
eltable:
    
    
End Sub
Sub tab2()

    Dim appAcc As Access.Application
    
    ' Get a reference to the Access Application object.
    Set appAcc = New Access.Application
    
    Realdoc = ActiveWindow.Document
    Selection.HomeKey Unit:=wdStory
    Documents("yk.doc").Activate
    Selection.HomeKey Unit:=wdStory
    Documents(Realdoc).Activate
    doexit = False
    
    i = 2
    While True

        Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
        WordToLookFor = Selection.text
        
        
        ' in document
        Documents("yk.doc").Activate
        Selection.HomeKey Unit:=wdStory
        Do
            Selection.HomeKey Unit:=wdLine
            Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
            If Selection.text = WordToLookFor Then
                MsgBox "Found!! " + Selection.text + "," + WordToLookFor
            End If
            Selection.MoveDown Unit:=wdParagraph
            
            If Selection.MoveRight = 0 Then Exit Do
            
        Loop Until Frue
        ' in document (END)
        
        
        
        'If WordToLookFor = "zoka" Then
         '   MsgBox "Bitti"
          '  GoTo eltable
        'End If
        Documents(Realdoc).Activate
        Selection.MoveRight Unit:=wdSentence
        If Selection.MoveRight = 0 Then GoTo eltable
    Wend
eltable:
    
    
End Sub

Sub tab3()

    Dim appAcc As Access.Application
    
    ' Get a reference to the Access Application object.
    Set appAcc = New Access.Application
    
    With appAcc
        ' Open the Northwind database.
        ' Modify the path as needed.
        .OpenCurrentDatabase "\\BELGIN\C\IBM\REGISTER\NT\sozlukac.mdb"
        
        
      ' Open a form.
      '.DoCmd.OpenForm "Employees", acNormal
      
      '.DoCmd.OpenQuery QueryName:="Yabanci Kelimeler"
      
      .DoCmd.RunSQL "SELECT WORDMULT.HEAD_MULT, ABBREVIATIONS.ABBREVIATION, WORDMULT.LANGUAGE FROM WORDMULT, ABBREVIATIONS where WORDMULT.LANGUAGE1 = ABBREVIATIONS.ABBR_ID and WORDMULT.HEAD_MULT = 'abajur'; "
      '.DoCmd.RunSQL "drop TABLE aaa;"

      
      
      ' Close the database.
      .CloseCurrentDatabase
      
      ' Quit the application.
      .quit
     End With
     
     
     ' Close the object reference.
     Set appAcc = Nothing


        
    
    Realdoc = ActiveWindow.Document
    Selection.HomeKey Unit:=wdStory
    Documents("yk.doc").Activate
    Selection.HomeKey Unit:=wdStory
    Documents(Realdoc).Activate
    doexit = False
    
    i = 2
    While True

        Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
        WordToLookFor = Selection.text
        
        
        ' in document
        Documents("yk.doc").Activate
        Selection.HomeKey Unit:=wdStory
        Do
            Selection.HomeKey Unit:=wdLine
            Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
            If Selection.text = WordToLookFor Then
                MsgBox "Found!! " + Selection.text + "," + WordToLookFor
            End If
            Selection.MoveDown Unit:=wdParagraph
            
            If Selection.MoveRight = 0 Then Exit Do
            
        Loop Until Frue
        ' in document (END)
        
        
        
        'If WordToLookFor = "zoka" Then
         '   MsgBox "Bitti"
          '  GoTo eltable
        'End If
        Documents(Realdoc).Activate
        Selection.MoveRight Unit:=wdSentence
        If Selection.MoveRight = 0 Then GoTo eltable
    Wend
eltable:
    
    

End Sub

Sub tab4()
'Dim appAcc As Access.Application
    
    Dim strDb As String
    Dim oDataBase  As Database
    Dim oWorkSpace As Workspace
    Dim oRecordSet As Recordset
    Dim qdf As QueryDef
    Dim iRowNum    As Integer


    
    Set oWorkSpace = _
        CreateWorkspace(Name:="JW", _
        UserName:="877068", Password:="877068", UseType:=dbUseJet)
        
    
    strDb = _
        "\\BELGIN\C\IBM\REGISTER\NT\sozlukac.mdb"
    Set oDataBase = OpenDatabase(strDb)
    
    
        
    
    Realdoc = ActiveWindow.Document
    Selection.HomeKey Unit:=wdStory
    Documents("log.doc").Activate
    Selection.HomeKey Unit:=wdStory
    Documents(Realdoc).Activate
    doexit = False
    
    i = 2
    While True

        Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
        While Selection.Font.bold = True
            If Mid(Selection.text, Len(Selection.text), 1) = "," _
                Or Mid(Selection.text, Len(Selection.text), 1) = "(" _
            Then
                GoTo n723
            End If
            Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
        Wend
n723:
        Selection.MoveLeft Unit:=wdCharacter, Extend:=wdExtend
        WordToLookFor = Selection.text
        
    
        For p = 1 To Len(WordToLookFor)
            'MsgBox Mid(WordToLookFor, p, 1)
            'MsgBox Asc(Mid(WordToLookFor, p, 1))
            If Asc(Mid(WordToLookFor, p, 1)) = 39 Then
                'MsgBox "holario"
                b1 = Mid(WordToLookFor, 1, p)
                b2 = Mid(WordToLookFor, p + 1, Len(WordToLookFor) - p - 1)
                WordToLookFor = b1 + Chr(39) + b2
                p = p + 1
            End If
        Next p
        For p = 1 To Len(WordToLookFor)
            'MsgBox Mid(WordToLookFor, p, 1)
            'MsgBox Asc(Mid(WordToLookFor, p, 1))
            If Asc(Mid(WordToLookFor, p, 1)) = 146 Then
                'MsgBox "holario"
                b1 = Mid(WordToLookFor, 1, p)
                b2 = Mid(WordToLookFor, p + 1, Len(WordToLookFor) - p - 1)
                WordToLookFor = b1 + Chr(146) + b2
                p = p + 1
            End If
        Next p
        
        'MsgBox WordToLookFor
        Set qdf = oDataBase.CreateQueryDef("nt10", _
            "SELECT WORDMULT.HEAD_MULT, ABBREVIATIONS.ABBREVIATION," + _
            "WORDMULT.LANGUAGE FROM WORDMULT, ABBREVIATIONS " + _
            "where WORDMULT.LANGUAGE1 = ABBREVIATIONS.ABBR_ID " + _
            "and wordmult.head_mult = '" + WordToLookFor + "';" _
        )
        
        Set oRecordSet = qdf.OpenRecordset()
    
        bulundu = False
        With oRecordSet
            Do While Not .EOF
                bulundu = True
                aaa = !HEAD_MULT
                bbb = !ABBREVIATION
                ccc = !Language
                'MsgBox WordToFind + "," + aaa + "," + bbb + "," + ccc
                Documents("log.doc").Activate
                'Selection.InsertAfter ""
                Selection.InsertAfter aaa
                Selection.InsertAfter ", "
                Selection.InsertAfter bbb
                Selection.InsertAfter ", "
                Selection.InsertAfter ccc
                Selection.InsertAfter Chr(13)
                
                Documents(Realdoc).Activate
                '******************************
                For n = 1 To 10
                    Selection.MoveRight Unit:=wdWord
                    Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
                    'MsgBox Selection.Text
                    eklendi = False
                    If Selection.text + "." = bbb Then
                        Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
                        Selection.Delete
                        Selection.InsertAfter ccc
                        Selection.InsertAfter " "
                        Selection.Font.Name = "TIMESCVR"
                        Selection.Font.italic = True
                        'MsgBox Selection.Text
                        'Selection.InsertAfter "-nt-"
                        eklendi = True
                        Exit For
                    End If
                Next n
                GoTo nt123
                '******************************
                .MoveNext
            Loop
nt123:
            'If bunundu = True Then
             '   Documents("log.doc").Activate
              '  If eklendi = True Then
               '     Selection.InsertAfter " Eklendi "
'                Else
 '                   Selection.InsertAfter " Eklenmedi!!! "
  '              End If
   '             Selection.InsertAfter Chr(13)
    '            Documents(Realdoc).Activate
     '       End If
        End With
        
        
        oRecordSet.Close
        qdf.Close
        oDataBase.QueryDefs.Delete qdf.Name
        
        
        Documents(Realdoc).Activate
        Selection.StartOf Unit:=wdParagraph
        Selection.MoveDown Unit:=wdParagraph
        If Selection.MoveDown = 0 Then GoTo eltable
        Selection.StartOf Unit:=wdParagraph
    Wend
eltable:
    
    oDataBase.Close
    oWorkSpace.Close
    Documents("log.doc").Activate
    Selection.Font.Name = "TIMESCVR"
    ali2sanity
    Documents(Realdoc).Activate
    ali2sanity

End Sub
Sub ali2sanity()
Attribute ali2sanity.VB_Description = "Makro, SOZLUK1 tarafýndan 02.07.99 tarihinde kaydedildi"
Attribute ali2sanity.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Makro2"
'
' Makro2 Makro
' Makro, SOZLUK1 tarafýndan 02.07.99 tarihinde kaydedildi
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\12"
        .Replacement.text = "Far."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\11"
        .Replacement.text = "Ar."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\13"
        .Replacement.text = "Fr."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\14"
        .Replacement.text = "Ýt."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\15"
        .Replacement.text = "Yun."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\19"
        .Replacement.text = "T."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\16"
        .Replacement.text = "Lât."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\18"
        .Replacement.text = "Ýng."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\17"
        .Replacement.text = "O.T."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\20"
        .Replacement.text = "Ýsp."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\25"
        .Replacement.text = "Ýbr."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\23"
        .Replacement.text = "Alm."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\21"
        .Replacement.text = "Erm."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\24"
        .Replacement.text = "Sl."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\22"
        .Replacement.text = "Rus."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\26"
        .Replacement.text = "Mac."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\27"
        .Replacement.text = "Bulg."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\28"
        .Replacement.text = "Port."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\346"
        .Replacement.text = "Jap."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\348"
        .Replacement.text = "Arn."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '----------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\352"
        .Replacement.text = "Norv."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
