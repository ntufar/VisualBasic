Attribute VB_Name = "Elifba"
Dim elifbayalin(50) As Integer
Dim elifbabasta(50) As Integer
Dim elifbaortada(50) As Integer
Dim elifbasonda(50) As Integer
Dim elifbahareke(15) As Integer



Const aelif As Integer = &HAC
Const eelif As Integer = &HB3
Const ielif As Integer = &H26
Const elif As Integer = &HAB
Const be As Integer = &HBB
Const pe As Integer = &HDB
Const te As Integer = &HAE
Const se As Integer = &HC0
Const cim As Integer = &HC3
Const chim As Integer = &H80
Const ha As Integer = &HD5
Const hi As Integer = &H8C
Const dal As Integer = &H9C
Const zel As Integer = &H8B
Const re As Integer = &H97
Const ze As Integer = &H93
Const je As Integer = &H8A
Const ti As Integer = &HB9
Const zi As Integer = &HB4
Const ain As Integer = &H9F
Const gain As Integer = &H8D
Const ar_sin As Integer = &H94
Const shin As Integer = &H91
Const sat As Integer = &H92
Const dat As Integer = &HF7
Const fe As Integer = &H2D
Const ffe As Integer = &H98
Const kaf As Integer = &H82
Const kef As Integer = &H84
Const lam As Integer = &H89
Const mim As Integer = &HC2
Const nun As Integer = &HCA
Const vav As Integer = &HCB
Const hvav As Integer = &H83
Const uvav As Integer = &H99
Const ovav As Integer = &H9B
Const yuvav As Integer = &H34
Const yovav As Integer = &H35
Const pppvav As Integer = &H28
Const he As Integer = &H88
Const the As Integer = &H85
Const hhe As Integer = &HB6
Const hye As Integer = &H54
Const ye As Integer = &HC8
Rem -----------------------------------
Const b_aelif As Integer = aelif
Const b_eelif As Integer = eelif
Const b_ielif As Integer = ielif
Const b_elif As Integer = elif
Const b_be As Integer = &H8F
Const b_pe As Integer = &HC4
Const b_te As Integer = &HD4
Const b_se As Integer = &HDD
Const b_cim As Integer = &H32
Const b_chim As Integer = &H81
Const b_ha As Integer = &H30
Const b_hi As Integer = &H95
Const b_dal As Integer = dal
Const b_zel As Integer = zel
Const b_re As Integer = re
Const b_ze As Integer = ze
Const b_je As Integer = je
Const b_ti As Integer = &HB9
Const b_zi As Integer = &HBA
Const b_ain As Integer = &H90
Const b_gain As Integer = &H9E
Const b_sin As Integer = &H8E
Const b_shin As Integer = &H96
Const b_sat As Integer = &H23
Const b_dat As Integer = &H24
Const b_fe As Integer = &HA7
Const b_ffe As Integer = &HD1
Const b_kaf As Integer = &HAD
Const b_kef As Integer = &HBD
Const b_lam As Integer = &HBC
Const b_mim As Integer = &HA6
Const b_nun As Integer = &HEC
Const b_vav As Integer = vav
Const b_hvav As Integer = hvav
Const b_uvav As Integer = uvav
Const b_ovav As Integer = ovav
Const b_yuvav As Integer = yuvav
Const b_yovav As Integer = yovav
Const b_pppvav As Integer = pppvav
Const b_he As Integer = &HA3
Const b_the As Integer = the
Const b_hhe As Integer = hhe
Const b_ye As Integer = &HB2
Const b_hye As Integer = &H7A
Rem ----------------------------------
Const e_aelif As Integer = aelif
Const e_eelif As Integer = eelif
Const e_ielif As Integer = &H53
Const e_elif As Integer = &H55
Const e_be As Integer = &H56
Const e_pe As Integer = &HE9
Const e_te As Integer = &H58
Const e_se As Integer = &H59
Const e_cim As Integer = &H5A
Const e_chim As Integer = &HEA
Const e_ha As Integer = &H60
Const e_hi As Integer = &H61
Const e_dal As Integer = &H62
Const e_zel As Integer = &H63
Const e_re As Integer = &H64
Const e_ze As Integer = &H65
Const e_je As Integer = &HF3
Const e_ti As Integer = &H6A
Const e_zi As Integer = &H6B
Const e_ain As Integer = &H6C
Const e_gain As Integer = &H6D
Const e_sin As Integer = &H66
Const e_shin As Integer = &H67
Const e_sat As Integer = &H68
Const e_dat As Integer = &H69
Const e_fe As Integer = &H6E
Const e_ffe As Integer = &HED
Const e_kaf As Integer = &H6F
Const e_kef As Integer = &H70
Const e_lam As Integer = &H71
Const e_mim As Integer = &H72
Const e_nun As Integer = &H73
Const e_vav As Integer = &H75
Const e_hvav As Integer = hvav
Const e_uvav As Integer = &HA4
Const e_ovav As Integer = &HA9
Const e_yuvav As Integer = &H36
Const e_yovav As Integer = &H37
Const e_pppvav As Integer = &H29
Const e_he As Integer = &HEB
Const e_the As Integer = &H57
Const e_hhe As Integer = hhe
Const e_ye As Integer = &H76
Const e_hye As Integer = &H54
Rem //------------------------------
Const m_aelif As Integer = aelif
Const m_eelif As Integer = eelif
Const m_ielif As Integer = e_ielif
Const m_elif As Integer = e_elif
Const m_be As Integer = &HBE
Const m_pe As Integer = &HE1
Const m_te As Integer = &HD7
Const m_se As Integer = &HDE
Const m_cim As Integer = &H9D
Const m_chim As Integer = &HC7
Const m_ha As Integer = &H31
Const m_hi As Integer = &HA5
Const m_dal As Integer = e_dal
Const m_zel As Integer = e_zel
Const m_re As Integer = e_re
Const m_ze As Integer = e_ze
Const m_je As Integer = e_je
Const m_ti As Integer = &H44
Const m_zi As Integer = &H45
Const m_ain As Integer = &H46
Const m_gain As Integer = &H47
Const m_sin As Integer = &H8E
Const m_shin As Integer = &H41
Const m_sat As Integer = &H42
Const m_dat As Integer = &H43
Const m_fe As Integer = &H48
Const m_ffe As Integer = &HE7
Const m_kaf As Integer = &H49
Const m_kef As Integer = &H4A
Const m_lam As Integer = &H2A
Const m_mim As Integer = &H4C
Const m_nun As Integer = &H4D
Const m_vav As Integer = vav
Const m_hvav As Integer = hvav
Const m_uvav As Integer = uvav
Const m_ovav As Integer = ovav
Const m_yuvav As Integer = yuvav
Const m_yovav As Integer = yovav
Const m_pppvav As Integer = pppvav
Const m_he As Integer = &HE4
Const m_the As Integer = the
Const m_hhe As Integer = hhe
Const m_ye As Integer = &H9A
Const m_hye As Integer = &H18
Rem //----------------------------------
Const lamelif As Integer = &HF4
Const b_lamelif As Integer = lamelif
Const e_lamelif As Integer = &HF6
Const m_lamelif As Integer = e_lamelif

Const alamelif As Integer = &HFC
Const b_alamelif As Integer = lamelif
Const e_alamelif As Integer = &H78
Const m_alamelif As Integer = e_lamelif

Const gef As Integer = &HAF
Const b_gef As Integer = &HD6
Const m_gef As Integer = &H7E
Const e_gef As Integer = &H40

Const nazal As Integer = &H21
Const e_nazal As Integer = &H22
Const m_nazal As Integer = &H25
Const b_nazal As Integer = &H27

Const hemze As Integer = &HA1
Rem //------------------------------ hareke
Const sukun As Integer = &HDA
Const ustun As Integer = &HD3
Const kesre As Integer = &HC5
Const sedde As Integer = &HD2
Const tenvin As Integer = &HCE
Const ustsedde As Integer = &H5D
Const asedde As Integer = &HE6
Const likevav As Integer = &H33
Const likevavn As Integer = &HCF

Const likevavsedde As Integer = &H2B
Const aa As Integer = &HD8
Const ii As Integer = &HF5
Const iii As Integer = &HCC
Const ppp As Integer = &H2C

Public Sub elifbainit()
Rem ------------------ Yalin
elifbayalin(1) = aelif
elifbayalin(2) = elif
elifbayalin(3) = eelif
elifbayalin(4) = hemze
elifbayalin(5) = be
elifbayalin(6) = pe
elifbayalin(7) = te
elifbayalin(8) = se
elifbayalin(9) = cim
elifbayalin(10) = chim
elifbayalin(11) = ha
elifbayalin(12) = hi
elifbayalin(13) = dal
elifbayalin(14) = zel
elifbayalin(15) = re
elifbayalin(16) = ze
elifbayalin(17) = je

elifbayalin(18) = ar_sin
elifbayalin(19) = shin
elifbayalin(20) = sat
elifbayalin(21) = dat

elifbayalin(22) = ti
elifbayalin(23) = zi
elifbayalin(24) = ain
elifbayalin(25) = gain

elifbayalin(26) = fe
elifbayalin(27) = ffe
elifbayalin(28) = kaf
elifbayalin(29) = kef
elifbayalin(30) = lam
elifbayalin(31) = mim
elifbayalin(32) = nun
elifbayalin(33) = vav
elifbayalin(34) = hvav
elifbayalin(35) = uvav
elifbayalin(36) = ovav
elifbayalin(37) = yuvav
elifbayalin(38) = yovav
elifbayalin(39) = pppvav
elifbayalin(40) = he
elifbayalin(41) = the
elifbayalin(42) = hhe
elifbayalin(43) = hye
elifbayalin(44) = ye
Rem ----------
Rem ------------------ Basta
elifbabasta(1) = b_aelif
elifbabasta(2) = b_elif
elifbabasta(3) = b_eelif
elifbabasta(4) = hemze
elifbabasta(5) = b_be
elifbabasta(6) = b_pe
elifbabasta(7) = b_te
elifbabasta(8) = b_se
elifbabasta(9) = b_cim
elifbabasta(10) = b_chim
elifbabasta(11) = b_ha
elifbabasta(12) = b_hi
elifbabasta(13) = b_dal
elifbabasta(14) = b_zel
elifbabasta(15) = b_re
elifbabasta(16) = b_ze
elifbabasta(17) = b_je

elifbabasta(18) = b_sin
elifbabasta(19) = b_shin
elifbabasta(20) = b_sat
elifbabasta(21) = b_dat

elifbabasta(22) = b_ti
elifbabasta(23) = b_zi
elifbabasta(24) = b_ain
elifbabasta(25) = b_gain

elifbabasta(26) = b_fe
elifbabasta(27) = b_ffe
elifbabasta(28) = b_kaf
elifbabasta(29) = b_kef
elifbabasta(30) = b_lam
elifbabasta(31) = b_mim
elifbabasta(32) = b_nun
elifbabasta(33) = b_vav
elifbabasta(34) = b_hvav
elifbabasta(35) = b_uvav
elifbabasta(36) = b_ovav
elifbabasta(37) = b_yuvav
elifbabasta(38) = b_yovav
elifbabasta(39) = b_pppvav
elifbabasta(40) = b_he
elifbabasta(41) = b_the
elifbabasta(42) = b_hhe
elifbabasta(43) = b_hye
elifbabasta(44) = b_ye
Rem ----------
Rem ------------------ Sonda
elifbasonda(1) = e_aelif
elifbasonda(2) = e_elif
elifbasonda(3) = e_eelif
elifbasonda(4) = hemze
elifbasonda(5) = e_be
elifbasonda(6) = e_pe
elifbasonda(7) = e_te
elifbasonda(8) = e_se
elifbasonda(9) = e_cim
elifbasonda(10) = e_chim
elifbasonda(11) = e_ha
elifbasonda(12) = e_hi
elifbasonda(13) = e_dal
elifbasonda(14) = e_zel
elifbasonda(15) = e_re
elifbasonda(16) = e_ze
elifbasonda(17) = e_je

elifbasonda(18) = e_sin
elifbasonda(19) = e_shin
elifbasonda(20) = e_sat
elifbasonda(21) = e_dat

elifbasonda(22) = e_ti
elifbasonda(23) = e_zi
elifbasonda(24) = e_ain
elifbasonda(25) = e_gain

elifbasonda(26) = e_fe
elifbasonda(27) = e_ffe
elifbasonda(28) = e_kaf
elifbasonda(29) = e_kef
elifbasonda(30) = e_lam
elifbasonda(31) = e_mim
elifbasonda(32) = e_nun
elifbasonda(33) = e_vav
elifbasonda(34) = e_hvav
elifbasonda(35) = e_uvav
elifbasonda(36) = e_ovav
elifbasonda(37) = e_yuvav
elifbasonda(38) = e_yovav
elifbasonda(39) = e_pppvav
elifbasonda(40) = e_he
elifbasonda(41) = e_the
elifbasonda(42) = e_hhe
elifbasonda(43) = e_hye
elifbasonda(44) = e_ye
Rem ----------
Rem ------------------ Ortada
elifbaortada(1) = m_aelif
elifbaortada(2) = m_elif
elifbaortada(3) = m_eelif
elifbaortada(4) = hemze
elifbaortada(5) = m_be
elifbaortada(6) = m_pe
elifbaortada(7) = m_te
elifbaortada(8) = m_se
elifbaortada(9) = m_cim
elifbaortada(10) = m_chim
elifbaortada(11) = m_ha
elifbaortada(12) = m_hi
elifbaortada(13) = m_dal
elifbaortada(14) = m_zel
elifbaortada(15) = m_re
elifbaortada(16) = m_ze
elifbaortada(17) = m_je

elifbaortada(18) = m_sin
elifbaortada(19) = m_shin
elifbaortada(20) = m_sat
elifbaortada(21) = m_dat


elifbaortada(22) = m_ti
elifbaortada(23) = m_zi
elifbaortada(24) = m_ain
elifbaortada(25) = m_gain

elifbaortada(26) = m_fe
elifbaortada(27) = m_ffe
elifbaortada(28) = m_kaf
elifbaortada(29) = m_kef
elifbaortada(30) = m_lam
elifbaortada(31) = m_mim
elifbaortada(32) = m_nun
elifbaortada(33) = m_vav
elifbaortada(34) = m_hvav
elifbaortada(35) = m_uvav
elifbaortada(36) = m_ovav
elifbaortada(37) = m_yuvav
elifbaortada(38) = m_yovav
elifbaortada(39) = m_pppvav
elifbaortada(40) = m_he
elifbaortada(41) = m_the
elifbaortada(42) = m_hhe
elifbaortada(43) = m_hye
elifbaortada(44) = m_ye

Rem --------- Hareke
elifbahareke(1) = sukun
elifbahareke(2) = ustun
elifbahareke(3) = kesre
elifbahareke(4) = sedde
elifbahareke(5) = tenvin
elifbahareke(6) = ustsedde
elifbahareke(7) = asedde
elifbahareke(8) = likevav
elifbahareke(9) = likevavn
elifbahareke(10) = likevavsedde
elifbahareke(11) = aa
elifbahareke(12) = ii
elifbahareke(13) = iii
elifbahareke(14) = ppp

Rem ---------
End Sub

Sub sirala()
    
    Dim count As Integer
    Dim firstword As String
    Dim secondword As String
    Dim wordfound As Boolean
    Dim sawbrace As Boolean
    Dim changed As Boolean
    Dim even As Boolean
    Dim qwerty As Integer
    
    elifbainit   'elifba tablolari doldurur
    firstword = ""
    secondword = ""
    even = False
    

doitagain:
    Selection.HomeKey Unit:=wdStory
doitagainwithouthome:
    'If even = True Then
    '    even = False
    '    Selection.MoveDown Unit:=wdParagraph
    'Else
    '    even = False
    'End If
    changed = False
    Do
        ' birinci kelime bul
        sawbrace = False
        wordfound = False
        Selection.StartOf Unit:=wdParagraph
        
        count = 0
        Do
            If Selection.text = "(" Then sawbrace = True
            If Selection.text = ")" Then sawbrace = False
            If sawbrace = True And _
                Selection.Font.Name = "Arapca (TDK-3)" Then
                    wordfound = True
                    'MsgBox "found!!"
                    Exit Do
            End If
            Selection.MoveRight Unit:=wdCharacter
            count = count + 1
            If count > 40 Then
                Exit Do
            End If
        Loop While True 'Selection.MoveRight <> 0
        'Selection.StartOf Unit:=wdWord
        Selection.MoveLeft Unit:=wdCharacter
        firstword = ""
        If wordfound = True Then
            Do
                Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
                If Selection.Font.Name <> "Arapca (TDK-3)" Or _
                    Selection.text = ")" Then
                    'firstword = Selection.Text
                    'MsgBox "holario"
                    GoTo ed002 'Exit Do
                End If
                'firstword = firstword + Selection.Text
                'MsgBox firstword
            Loop While True 'Selection.MoveRight <> 0
        End If
ed002:
       Selection.MoveLeft Unit:=wdCharacter, Extend:=wdExtend
       firstword = Selection.text
        
        
        If wordfound = True Then
            'MsgBox "first word: " + firstword
            'Documents("x.doc").Activate
            'Selection.InsertAfter firstword
            'Selection.InsertAfter Chr(13)
            'Documents("O.doc").Activate
        End If
        
        
        ' -------- ikinci kelime bul ----------
        qwerty = Selection.MoveDown(Unit:=wdParagraph)
        If qwerty = 0 Then
            'MsgBox "holario"
            GoTo eeol1
        End If
        sawbrace = False
        wordfound = False
        Selection.StartOf Unit:=wdParagraph
        count = 0
        Do
            If Selection.text = "(" Then sawbrace = True
            If Selection.text = ")" Then sawbrace = False
            If sawbrace = True And _
                Selection.Font.Name = "Arapca (TDK-3)" Then
                    wordfound = True
                    Exit Do
            End If
            Selection.MoveRight Unit:=wdCharacter
            count = count + 1
            If count > 40 Then
                Exit Do
            End If
        Loop While True
        Selection.MoveLeft Unit:=wdCharacter
        
        If wordfound = False Then
            secondword = ""
        Else
            Do
                Selection.MoveRight Unit:=wdCharacter, Extend:=wdExtend
                If Selection.Font.Name <> "Arapca (TDK-3)" Then
                    Selection.MoveLeft Unit:=wdCharacter, Extend:=wdExtend
                    secondword = Selection.text
                    Exit Do
                End If
            Loop While True
        End If
        
        If wordfound = True Then
            'MsgBox "first word: " + firstword
            'Documents("x.doc").Activate
            'Selection.InsertAfter secondword
            'Selection.InsertAfter Chr(13)
            'Documents("O.doc").Activate
        End If

        
        '---------- firstword and secondword are ready
        '---------- cursor is in the second paragraph
        'MsgBox firstword
        'MsgBox secondword
        'Documents("x.doc").Activate
        If firstword <> "" And secondword = "" Then
                Selection.StartOf Unit:=wdParagraph
                Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                Selection.Cut
            
                Selection.MoveUp Unit:=wdParagraph
                Selection.Paste
                changed = True
                GoTo endofloop1
        End If
        If firstword <> "" And secondword <> "" Then
            If arcmp(firstword, secondword) = False Then
                Selection.StartOf Unit:=wdParagraph
                Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                Selection.Cut
            
                Selection.MoveUp Unit:=wdParagraph
                Selection.Paste
                changed = True
                Selection.MoveUp Unit:=wdParagraph, count:=3
                GoTo doitagainwithouthome
            
                'Selection.InsertAfter ">>>>>>>>>>>>>>>>"
                'Selection.InsertAfter Chr(13)
            Else
                ' nothing
                'Selection.InsertAfter "<<<<<<<<<<<<<<<<"
                'Selection.InsertAfter Chr(13)
            End If
        End If
        'Documents("O.doc").Activate
endofloop1:
        'qwerty = Selection.MoveDown(Unit:=wdParagraph)
        'If qwerty = 0 Then GoTo eeol1
        'Selection.MoveUp Unit:=wdParagraph
    Loop While True 'qwerty <> 0
eeol1:
    If changed = True Then GoTo doitagain
End Sub

Function arcmp(ByVal first As String, ByVal second As String) As Boolean
    ' return true if first is >= second

    Dim fp As Integer
    Dim sp As Integer
    Dim lmt As Integer
    Dim fb As String
    Dim sb As String
    Dim hvf As Integer
    Dim hvs As Integer
    
    'first = normalizearapword(first)
    'second = normalizearapword(second)
    fp = Len(first)
    sp = Len(second)
    If fp < sp Then
        lmt = fp
    Else
        lmt = sp
    End If
    
    arcmp = False
    For i = 1 To lmt
        fb = Mid(first, i, 1)
        sb = Mid(second, i, 1)
        hvf = harfval(fb)
        hvs = harfval(sb)
        If hvf < hvs Then
             arcmp = True
             Exit Function
        End If
        If hvf = hvs Then
             arcmp = True
        Else
             arcmp = False
             Exit Function
        End If
    Next i
    'If arcmp = True Then Exit Function
    'MsgBox "same length"
    If fp <= sp Then
        arcmp = True
    Else
        arcmp = False
    End If
End Function
Private Function harfval(ByVal harf As String) As Integer
    'Dim elifbayalin(50) As byte
    'Dim elifbabasta(50) As byte
    'Dim elifbaortada(50) As byte
    'Dim elifbasonda(50) As byte
    Dim i As Integer
    Dim h As Integer
    
    h = Asc(Left(harf, 1))
    For i = 1 To 44
        If elifbayalin(i) = h Or _
            elifbabasta(i) = h Or _
            elifbaortada(i) = h Or _
            elifbasonda(i) = h Then
                harfval = i
                Exit Function
        End If
        'If h = 32 Then
        '    harfval = 99
        '    Exit Function
        'End If
    Next i
    'Documents("x.doc").Activate
    'Selection.InsertAfter h
    'MsgBox "not found" + h
    harfval = 0
End Function
Public Function normalizearapword(ByVal word As String) As String
    Dim nw As String
    Dim wp As Integer
    Dim nwp As Integer
    Dim wl As Integer
    Dim nwl As Integer
    Dim i As Integer
    Dim xbuf As String
    'Dim i As Integer
    
    'normalizearapword = word
    'Exit Function
    elifbainit
doitagain:
    nw = "                                                  "
    wl = Len(word)
    nwl = Len(nw)
    
    
    nwp = 1
    For wp = wl To 1 Step -1
        'MsgBox Asc(Mid(word, wp, 1))
        If Asc(Mid(word, wp, 1)) = gef Or _
             Asc(Mid(word, wp, 1)) = b_gef Or _
             Asc(Mid(word, wp, 1)) = m_gef Or _
             Asc(Mid(word, wp, 1)) = e_gef Then
                  Mid(word, wp, 1) = Chr(kef)
        End If
        If Asc(Mid(word, wp, 1)) = eelif Or _
             Asc(Mid(word, wp, 1)) = b_eelif Or _
             Asc(Mid(word, wp, 1)) = m_eelif Or _
             Asc(Mid(word, wp, 1)) = e_eelif Then
                  Mid(word, wp, 1) = Chr(elif)
        End If
        If Asc(Mid(word, wp, 1)) = eilif Or _
             Asc(Mid(word, wp, 1)) = b_ielif Or _
             Asc(Mid(word, wp, 1)) = m_ielif Or _
             Asc(Mid(word, wp, 1)) = e_ielif Then
                  Mid(word, wp, 1) = Chr(elif)
        End If
        If Asc(Mid(word, wp, 1)) = &H3C Then
                  Mid(word, wp, 1) = Chr(chim)
        End If
        If Asc(Mid(word, wp, 1)) = &H8E Then
                  Mid(word, wp, 1) = Chr(ar_sin)
        End If
        If Asc(Mid(word, wp, 1)) = &H38 Then
                  Mid(word, wp, 1) = Chr(se)
        End If
        If Asc(Mid(word, wp, 1)) = &H4E Then
                  Mid(word, wp, 1) = Chr(he)
        End If
        If Asc(Mid(word, wp, 1)) = &H39 Then
                  Mid(word, wp, 1) = Chr(se)
        End If
        If Asc(Mid(word, wp, 1)) = alamelif Or _
             Asc(Mid(word, wp, 1)) = b_alamelif Or _
             Asc(Mid(word, wp, 1)) = m_alamelif Or _
             Asc(Mid(word, wp, 1)) = e_alamelif Then
                  'Mid(word, wp, 1) = Chr(lam)
                  xbuf = Mid(word, 1, wp - 1)
                  xbuf = xbuf + Chr(elif)
                  xbuf = xbuf + Chr(lam)
                  xbuf = xbuf + Mid(word, wp + 1, wl - wp)
                  word = xbuf
                  GoTo doitagain
        End If
        If Asc(Mid(word, wp, 1)) = lamelif Or _
             Asc(Mid(word, wp, 1)) = b_lamelif Or _
             Asc(Mid(word, wp, 1)) = m_lamelif Or _
             Asc(Mid(word, wp, 1)) = e_lamelif Then
                  'Mid(word, wp, 1) = Chr(lam)
                  xbuf = Mid(word, 1, wp - 1)
                  xbuf = xbuf + Chr(elif)
                  xbuf = xbuf + Chr(lam)
                  xbuf = xbuf + Mid(word, wp + 1, wl - wp)
                  word = xbuf
                  GoTo doitagain
        End If
        
        
        If ishareke(Mid(word, wp, 1)) = True Then
          'nothing Mid(nw, nwp, 1) = Mid(word, wp, 1)
            nwp = nwp - 1
          Else
              Mid(nw, nwp, 1) = Mid(word, wp, 1)
        End If
        nwp = nwp + 1
    Next wp
    
spagain:
    nwl = Len(nw)
    For i = 1 To Len(nw)
        If Mid(nw, i, 1) = " " Then
            xbuf = Left(nw, i - 1)
            xbuf = xbuf + Right(nw, Len(nw) - i)
            'Mid(nw, i, Len(nw) - i) = Mid(nw, i + 1, Len(nw) - i - 1)
            nw = xbuf
            GoTo spagain
        End If
    Next i
    'normalizearapword = Right(nw, nwl - nwp)
    normalizearapword = nw
    'MsgBox ">>>" + normalizearapword + "<<<" + ">>>>>" + word + "<<<<<"
End Function

Private Function ishareke(ByVal c As String) As Boolean
    Dim i As Integer
    For i = 1 To 14
        If Asc(Left(c, 1)) = elifbahareke(i) Then
            ishareke = True
            Exit Function
        End If
    Next i
    ishareke = False
    Exit Function
End Function
