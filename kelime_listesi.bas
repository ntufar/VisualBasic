Attribute VB_Name = "kelime_listesi"
Sub kelime_ekle()
    Dim strDb As String
    Dim oDataBase  As Database
    Dim oWorkSpace As Workspace
    Dim oRecordSet As Recordset
    Dim qdf As QueryDef
    Dim udf As QueryDef
    Dim iRowNum    As Integer
    Dim WordToLookFor As String


    
    Set oWorkSpace = _
        CreateWorkspace(Name:="JW", _
        UserName:="877068", Password:="877068", UseType:=dbUseJet)
        
    
    strDb = _
        "d:\nt\hatice\ksay.mdb"
    Set oDataBase = OpenDatabase(strDb)

    
    
    Realdoc = ActiveWindow.Document
    'Selection.HomeKey Unit:=wdStory
    Documents("kl.doc").Activate
    Selection.HomeKey Unit:=wdStory
    Documents(Realdoc).Activate
    Selection.HomeKey Unit:=wdStory
  
 
    
    Dim qwerty As Integer
    
    While True

        'Selection.StartOf unit:=wdParagraph
        Selection.MoveRight Unit:=wdWord, Extend:=wdExtend
        WordToLookFor = Selection.text
        Selection.Collapse Direction:=wdCollapseEnd
        If WordToLookFor = Chr(13) Or WordToLookFor = Chr(9) _
            Or WordToLookFor = "." Or WordToLookFor = "," _
            Or WordToLookFor = "(" Or WordToLookFor = ")" Then
            Selection.MoveRight Unit:=wdWord, Extend:=wdMove
            GoTo doit_again
        End If
        
        'Selection.Copy
        
        Set qdf = oDataBase.CreateQueryDef("nt10", _
            "insert into ksay(kelime) values('" + _
            SQLquote(WordToLookFor) + "');" _
        )
        
        qdf.Execute
        qdf.Close
        oDataBase.QueryDefs.Delete ("nt10")
        

doit_again:
        qwerty = Selection.MoveRight(Unit:=wdWord, count:=2, Extend:=wdMove)
        If (qwerty <> 2) Then GoTo eltable
        qwerty = Selection.MoveLeft(Unit:=wdWord, count:=2, Extend:=wdMove)
    Wend
eltable:

    oDataBase.Close
    oWorkSpace.Close

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
    Dim buf As String
    
    buf = "                                                  "
    

    bp = 1
    sp = 1
    
    For sp = 1 To Len(str)
        Mid(buf, bp, 1) = Mid(str, sp, 1)
        If Mid(str, sp, 1) = "'" Then
            bp = bp + 1
            Mid(buf, bp, 1) = "'"
            'MsgBox str + buf
        End If
        bp = bp + 1
    Next sp
    SQLquote = buf
End Function
