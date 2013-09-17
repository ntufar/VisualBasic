Attribute VB_Name = "Kisaltmalar"
Option Compare Text
Dim SQLquotebuf  As String

'Dim Words(100000) As String
'Dim Para() As Integer


Sub kisaltmalarekle()
    Dim strDb As String
    Dim oDataBase  As Database
    Dim oWorkSpace As Workspace
    Dim oRecordSet As Recordset
    Dim qdf As QueryDef
    Dim udf As QueryDef
    Dim iRowNum    As Integer
    Dim WordToLookFor As String
    Dim i As Integer


    
    Set oWorkSpace = _
        CreateWorkspace(Name:="JW", _
        UserName:="877068", Password:="877068", UseType:=dbUseJet)
        
    
    strDb = _
        "d:\nt\hatice\ksay.mdb"
    Set oDataBase = OpenDatabase(strDb)

    
    
    Realdoc = ActiveWindow.Document
    'Selection.HomeKey Unit:=wdStory
    'Documents("kl.doc").Activate
    'Selection.HomeKey Unit:=wdStory
    'Documents(Realdoc).Activate
    Selection.HomeKey Unit:=wdStory
  
 
    
    Dim qwerty As Integer
    i = 1
    
    While True

        'Selection.StartOf unit:=wdParagraph
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
        If Selection.text = "(" Then
            Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdMove
            While Mid(Selection.text, Len(Selection.text), 1) <> ")" And _
                Selection.Font.italic = True
                    Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
            Wend
            If Len(Selection.text) <> 1 Then
                WordToLookFor = Selection.text
                WordToLookFor = Mid(WordToLookFor, 1, Len(WordToLookFor) - 1)
                'Words(i) = WordToLookFor
                
                Set qdf = oDataBase.CreateQueryDef("nt10", _
                    "insert into ksay(kelime) values('" + _
                    SQLquote(WordToLookFor) + "');" _
                )
        
                qdf.Execute
                qdf.Close
                oDataBase.QueryDefs.Delete ("nt10")

                
                i = i + 1
                
                'MsgBox WordToLookFor
                Selection.Collapse Direction:=wdCollapseEnd
                If WordToLookFor = Chr(13) Or WordToLookFor = Chr(9) _
                    Or WordToLookFor = "." Or WordToLookFor = "," _
                    Or WordToLookFor = "(" Or WordToLookFor = ")" Then
                    Selection.MoveRight Unit:=wdWord, Extend:=wdMove
                    GoTo doit_again
                End If
            End If
        End If
        
        'Selection.Copy
        
        

doit_again:
        qwerty = Selection.MoveRight(Unit:=wdCharacter, count:=2, Extend:=wdMove)
        If (qwerty <> 2) Then GoTo eltable
        qwerty = Selection.MoveLeft(Unit:=wdCharacter, count:=2, Extend:=wdMove)
    Wend
eltable:

    oDataBase.Close
    oWorkSpace.Close
    '    Documents("dizin.doc").Activate
    '    For k = 1 To i
    '        Selection.InsertAfter Words(k)
    '        Selection.InsertAfter Chr(13)
    '    Next k

End Sub
Sub kelime_sil()
    Dim strDb As String
    Dim oDataBase  As Database
    Dim oWorkSpace As Workspace
    Dim oRecordSet As Recordset
    Dim qdf As QueryDef
    Dim udf As QueryDef
    Dim iRowNum    As Integer


    
    Set oWorkSpace = _
        CreateWorkspace(Name:="JW", _
        UserName:="877068", Password:="877068", UseType:=dbUseJet)
        
    
    strDb = _
        "d:\nt\hatice\ksay.mdb"
    Set oDataBase = OpenDatabase(strDb)

    
    
        
        Set qdf = oDataBase.CreateQueryDef("nt10", _
            "delete from ksay;")
        
        qdf.Execute
        qdf.Close
        oDataBase.QueryDefs.Delete ("nt10")
        

    oDataBase.Close
    oWorkSpace.Close

End Sub

Sub kelime_alfabetik()
    Dim strDb As String
    Dim oDataBase  As Database
    Dim oWorkSpace As Workspace
    Dim oRecordSet As Recordset
    Dim qdf As QueryDef
    Dim udf As QueryDef
    Dim iRowNum    As Integer
    Dim sayac As Integer


    
    Set oWorkSpace = _
        CreateWorkspace(Name:="JW", _
        UserName:="877068", Password:="877068", UseType:=dbUseJet)
        
    
    strDb = _
        "d:\nt\hatice\ksay.mdb"
    Set oDataBase = OpenDatabase(strDb)

    

        
        Set qdf = oDataBase.CreateQueryDef("nt10", _
            "select kelime, count(kelime) as sayi" + _
            " from ksay group by kelime" + _
            " order by kelime;")
        
        Set oRecordSet = qdf.OpenRecordset()
        sayac = 0
        With oRecordSet
            Do While Not .EOF
                aaa = !kelime
                bbb = !sayi
                Selection.InsertAfter aaa
                Selection.InsertAfter Chr(9)
                Selection.InsertAfter bbb
                Selection.InsertAfter Chr(13)
                .MoveNext
                sayac = sayac + 1
            Loop
        End With
        
        
    qdf.Close
    oDataBase.QueryDefs.Delete ("nt10")
        
                Selection.InsertAfter "--------"
                Selection.InsertAfter Chr(13)
                Selection.InsertAfter "Toplam ayrý kelime:"
                Selection.InsertAfter Chr(9)
                Selection.InsertAfter sayac
                Selection.InsertAfter Chr(13)

   
   Set qdf = oDataBase.CreateQueryDef("nt10", _
            "select count(kelime)as toplam" + _
            " from ksay " + _
            " ;")
   Set oRecordSet = qdf.OpenRecordset()
   With oRecordSet
            Do While Not .EOF
                bbb = !toplam
                'Selection.InsertAfter "--------"
                Selection.InsertAfter Chr(13)
                Selection.InsertAfter "Toplam kelime:"
                Selection.InsertAfter Chr(9)
                Selection.InsertAfter bbb
                Selection.InsertAfter Chr(13)
                .MoveNext
            Loop
    End With
                
    
    qdf.Close
    oDataBase.QueryDefs.Delete ("nt10")
        
   
        

    oDataBase.Close
    oWorkSpace.Close
    
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.Font.Size = 12


End Sub

Sub kelime_sayi()
    Dim strDb As String
    Dim oDataBase  As Database
    Dim oWorkSpace As Workspace
    Dim oRecordSet As Recordset
    Dim qdf As QueryDef
    Dim udf As QueryDef
    Dim iRowNum    As Integer
    Dim sayac As Integer


    
    Set oWorkSpace = _
        CreateWorkspace(Name:="JW", _
        UserName:="877068", Password:="877068", UseType:=dbUseJet)
        
    
    strDb = _
        "d:\nt\hatice\ksay.mdb"
    Set oDataBase = OpenDatabase(strDb)

    
    
        
        Set qdf = oDataBase.CreateQueryDef("nt10", _
            "select kelime, count(kelime) as sayi" + _
            " from ksay group by kelime" + _
            " order by count(kelime) desc;")
        
        Set oRecordSet = qdf.OpenRecordset()
        sayac = 0
        With oRecordSet
            Do While Not .EOF
                aaa = !kelime
                bbb = !sayi
                Selection.InsertAfter aaa
                Selection.InsertAfter Chr(9)
                Selection.InsertAfter bbb
                Selection.InsertAfter Chr(13)
                .MoveNext
                sayac = sayac + 1
            Loop
        End With
        
        
        qdf.Close
        oDataBase.QueryDefs.Delete ("nt10")
                Selection.InsertAfter "--------"
                Selection.InsertAfter Chr(13)
                Selection.InsertAfter "Toplam ayrý kelime:"
                Selection.InsertAfter Chr(9)
                Selection.InsertAfter sayac
                Selection.InsertAfter Chr(13)
        
        
   Set qdf = oDataBase.CreateQueryDef("nt10", _
            "select count(kelime)as toplam" + _
            " from ksay " + _
            " ;")
   Set oRecordSet = qdf.OpenRecordset()
   With oRecordSet
            Do While Not .EOF
                bbb = !toplam
                'Selection.InsertAfter "--------"
                Selection.InsertAfter Chr(13)
                Selection.InsertAfter "Toplam kelime:"
                Selection.InsertAfter Chr(9)
                Selection.InsertAfter bbb
                Selection.InsertAfter Chr(13)
                .MoveNext
            Loop
    End With
    
    qdf.Close
    oDataBase.QueryDefs.Delete ("nt10")
        
        

    oDataBase.Close
    oWorkSpace.Close

    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(1.25)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5), _
        Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    Selection.Font.Size = 12


End Sub



Function SQLquote(str As String) As String

    Dim pb As Integer
    Dim sp As Integer
    
    
    SQLquotebuf = Space(Len(str) + 100)
    

    bp = 1
    sp = 1
    
    For sp = 1 To Len(str)
        Mid(SQLquotebuf, bp, 1) = Mid(str, sp, 1)
        If Mid(str, sp, 1) = "'" Then
            bp = bp + 1
            Mid(SQLquotebuf, bp, 1) = "'"
            'MsgBox str + buf
        End If
        bp = bp + 1
    Next sp
    SQLquote = SQLquotebuf
End Function


