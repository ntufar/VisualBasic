Attribute VB_Name = "EskiElifba"
Dim elifbayalin(50) As String
Dim elifbabasta(50) As String
Dim elifbaortada(50) As String
Dim elifbasonda(50) As String
Dim elifbahareke(15) As String

Private Sub elifbainit()

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
Const e_aelif As Integer = elif
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
Const m_aelif As Integer = elif
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

Rem ------------------ Yalin
elifbayalin(1) = Chr(aelif)
elifbayalin(2) = Chr(eelif)
elifbayalin(3) = Chr(ielif)
elifbayalin(4) = Chr(elif)
elifbayalin(5) = Chr(be)
elifbayalin(6) = Chr(pe)
elifbayalin(7) = Chr(te)
elifbayalin(8) = Chr(se)
elifbayalin(9) = Chr(cim)
elifbayalin(10) = Chr(chim)
elifbayalin(11) = Chr(ha)
elifbayalin(12) = Chr(hi)
elifbayalin(13) = Chr(dal)
elifbayalin(14) = Chr(zel)
elifbayalin(15) = Chr(re)
elifbayalin(16) = Chr(ze)
elifbayalin(17) = Chr(je)

elifbayalin(18) = Chr(ar_sin)
elifbayalin(19) = Chr(shin)
elifbayalin(20) = Chr(sat)
elifbayalin(21) = Chr(dat)

elifbayalin(22) = Chr(ti)
elifbayalin(23) = Chr(zi)
elifbayalin(24) = Chr(ain)
elifbayalin(25) = Chr(gain)

elifbayalin(26) = Chr(fe)
elifbayalin(27) = Chr(ffe)
elifbayalin(28) = Chr(kaf)
elifbayalin(29) = Chr(kef)
elifbayalin(30) = Chr(lam)
elifbayalin(31) = Chr(mim)
elifbayalin(32) = Chr(nun)
elifbayalin(33) = Chr(vav)
elifbayalin(34) = Chr(hvav)
elifbayalin(35) = Chr(uvav)
elifbayalin(36) = Chr(ovav)
elifbayalin(37) = Chr(yuvav)
elifbayalin(38) = Chr(yovav)
elifbayalin(39) = Chr(pppvav)
elifbayalin(40) = Chr(he)
elifbayalin(41) = Chr(the)
elifbayalin(42) = Chr(hhe)
elifbayalin(43) = Chr(hye)
elifbayalin(44) = Chr(ye)
Rem ----------
Rem ------------------ Basta
elifbabasta(1) = Chr(b_aelif)
elifbabasta(2) = Chr(b_eelif)
elifbabasta(3) = Chr(b_ielif)
elifbabasta(4) = Chr(b_elif)
elifbabasta(5) = Chr(b_be)
elifbabasta(6) = Chr(b_pe)
elifbabasta(7) = Chr(b_te)
elifbabasta(8) = Chr(b_se)
elifbabasta(9) = Chr(b_cim)
elifbabasta(10) = Chr(b_chim)
elifbabasta(11) = Chr(b_ha)
elifbabasta(12) = Chr(b_hi)
elifbabasta(13) = Chr(b_dal)
elifbabasta(14) = Chr(b_zel)
elifbabasta(15) = Chr(b_re)
elifbabasta(16) = Chr(b_ze)
elifbabasta(17) = Chr(b_je)

elifbabasta(18) = Chr(b_ar_sin)
elifbabasta(19) = Chr(b_shin)
elifbabasta(20) = Chr(b_sat)
elifbabasta(21) = Chr(b_dat)

elifbabasta(22) = Chr(b_ti)
elifbabasta(23) = Chr(b_zi)
elifbabasta(24) = Chr(b_ain)
elifbabasta(25) = Chr(b_gain)

elifbabasta(26) = Chr(b_fe)
elifbabasta(27) = Chr(b_ffe)
elifbabasta(28) = Chr(b_kaf)
elifbabasta(29) = Chr(b_kef)
elifbabasta(30) = Chr(b_lam)
elifbabasta(31) = Chr(b_mim)
elifbabasta(32) = Chr(b_nun)
elifbabasta(33) = Chr(b_vav)
elifbabasta(34) = Chr(b_hvav)
elifbabasta(35) = Chr(b_uvav)
elifbabasta(36) = Chr(b_ovav)
elifbabasta(37) = Chr(b_yuvav)
elifbabasta(38) = Chr(b_yovav)
elifbabasta(39) = Chr(b_pppvav)
elifbabasta(40) = Chr(b_he)
elifbabasta(41) = Chr(b_the)
elifbabasta(42) = Chr(b_hhe)
elifbabasta(43) = Chr(b_hye)
elifbabasta(44) = Chr(b_ye)
Rem ----------
Rem ------------------ Sonda
elifbasonda(1) = Chr(e_aelif)
elifbasonda(2) = Chr(e_eelif)
elifbasonda(3) = Chr(e_ielif)
elifbasonda(4) = Chr(e_elif)
elifbasonda(5) = Chr(e_be)
elifbasonda(6) = Chr(e_pe)
elifbasonda(7) = Chr(e_te)
elifbasonda(8) = Chr(e_se)
elifbasonda(9) = Chr(e_cim)
elifbasonda(10) = Chr(e_chim)
elifbasonda(11) = Chr(e_ha)
elifbasonda(12) = Chr(e_hi)
elifbasonda(13) = Chr(e_dal)
elifbasonda(14) = Chr(e_zel)
elifbasonda(15) = Chr(e_re)
elifbasonda(16) = Chr(e_ze)
elifbasonda(17) = Chr(e_je)

elifbasonda(18) = Chr(e_ar_sin)
elifbasonda(19) = Chr(e_shin)
elifbasonda(20) = Chr(e_sat)
elifbasonda(21) = Chr(e_dat)

elifbasonda(22) = Chr(e_ti)
elifbasonda(23) = Chr(e_zi)
elifbasonda(24) = Chr(e_ain)
elifbasonda(25) = Chr(e_gain)

elifbasonda(26) = Chr(e_fe)
elifbasonda(27) = Chr(e_ffe)
elifbasonda(28) = Chr(e_kaf)
elifbasonda(29) = Chr(e_kef)
elifbasonda(30) = Chr(e_lam)
elifbasonda(31) = Chr(e_mim)
elifbasonda(32) = Chr(e_nun)
elifbasonda(33) = Chr(e_vav)
elifbasonda(34) = Chr(e_hvav)
elifbasonda(35) = Chr(e_uvav)
elifbasonda(36) = Chr(e_ovav)
elifbasonda(37) = Chr(e_yuvav)
elifbasonda(38) = Chr(e_yovav)
elifbasonda(39) = Chr(e_pppvav)
elifbasonda(40) = Chr(e_he)
elifbasonda(41) = Chr(e_the)
elifbasonda(42) = Chr(e_hhe)
elifbasonda(43) = Chr(e_hye)
elifbasonda(44) = Chr(e_ye)
Rem ----------
Rem ------------------ Ortada
elifbaortada(1) = Chr(m_aelif)
elifbaortada(2) = Chr(m_eelif)
elifbaortada(3) = Chr(m_ielif)
elifbaortada(4) = Chr(m_elif)
elifbaortada(5) = Chr(m_be)
elifbaortada(6) = Chr(m_pe)
elifbaortada(7) = Chr(m_te)
elifbaortada(8) = Chr(m_se)
elifbaortada(9) = Chr(m_cim)
elifbaortada(10) = Chr(m_chim)
elifbaortada(11) = Chr(m_ha)
elifbaortada(12) = Chr(m_hi)
elifbaortada(13) = Chr(m_dal)
elifbaortada(14) = Chr(m_zel)
elifbaortada(15) = Chr(m_re)
elifbaortada(16) = Chr(m_ze)
elifbaortada(17) = Chr(m_je)

elifbaortada(28) = Chr(m_ar_sin)
elifbaortada(29) = Chr(m_shin)
elifbaortada(20) = Chr(m_sat)
elifbaortada(21) = Chr(m_dat)


elifbaortada(22) = Chr(m_ti)
elifbaortada(23) = Chr(m_zi)
elifbaortada(24) = Chr(m_ain)
elifbaortada(25) = Chr(m_gain)

elifbaortada(26) = Chr(m_fe)
elifbaortada(27) = Chr(m_ffe)
elifbaortada(28) = Chr(m_kaf)
elifbaortada(29) = Chr(m_kef)
elifbaortada(30) = Chr(m_lam)
elifbaortada(31) = Chr(m_mim)
elifbaortada(32) = Chr(m_nun)
elifbaortada(33) = Chr(m_vav)
elifbaortada(34) = Chr(m_hvav)
elifbaortada(35) = Chr(m_uvav)
elifbaortada(36) = Chr(m_ovav)
elifbaortada(37) = Chr(m_yuvav)
elifbaortada(38) = Chr(m_yovav)
elifbaortada(39) = Chr(m_pppvav)
elifbaortada(40) = Chr(m_he)
elifbaortada(41) = Chr(m_the)
elifbaortada(42) = Chr(m_hhe)
elifbaortada(43) = Chr(m_hye)
elifbaortada(44) = Chr(m_ye)

Rem --------- Hareke
elifbahareke(1) = Chr(sukun)
elifbahareke(2) = Chr(ustun)
elifbahareke(3) = Chr(kesre)
elifbahareke(4) = Chr(sedde)
elifbahareke(5) = Chr(tenvin)
elifbahareke(6) = Chr(ustsedde)
elifbahareke(7) = Chr(asedde)
elifbahareke(8) = Chr(likevav)
elifbahareke(9) = Chr(likevavn)
elifbahareke(10) = Chr(likevavsedde)
elifbahareke(11) = Chr(aa)
elifbahareke(12) = Chr(ii)
elifbahareke(13) = Chr(iii)
elifbahareke(14) = Chr(ppp)

Rem ---------
End Sub

Private Sub sirala()
    
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

Private Function arcmp(ByVal first As String, ByVal second As String) As Boolean
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
        lmt = fp - 1
    Else
        lmt = sp - 1
    End If
    
    For i = 0 To lmt
        fb = Mid(first, fp - i, 1)
        sb = Mid(second, sp - i, 1)
        hvf = harfval(fb)
        hvs = harfval(sb)
        If hvf >= hvs Then
             arcmp = True
        Else
             arcmp = False
             Exit Function
        End If
    Next i
    MsgBox "same length"
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
    Dim h As String
    
    h = Left(harf, 1)
    For i = 1 To 44
        If elifbayalin(i) = Left(h, 1) Or _
            elifbabasta(i) = Left(h, 1) Or _
            elifbaortada(i) = Left(h, 1) Or _
            elifbasonda(i) = Left(h, 1) Then
                harfval = i
                Exit Function
        End If
        If h = " " Then
            harfval = 99
            Exit Function
        End If
    Next i
    'Documents("x.doc").Activate
    'Selection.InsertAfter h
    'MsgBox "not found" + h
    harfval = 0
End Function
Private Function normalizearapword(ByVal word As String) As String
    Dim nw As String
    Dim wp As Integer
    Dim nwp As Integer
    Dim wl As Integer
    Dim nwl As Integer
    Dim i As Integer
    'Dim i As Integer
    
    'normalizearapword = word
    'Exit Function
    nw = "                                                  "
    wl = Len(word)
    nwl = Len(nw)
    
    
    nwp = 1
        'Chr(lamelif) Or
    For wp = 1 To wl
        'MsgBox Asc(Mid(word, wp, 1))
        If Asc(Mid(word, wp, 1)) = &HF4 Or _
             Mid(word, wp, 1) = Chr(e_lamelif) Then
                  Mid(nw, nwp, 1) = Chr(lam)
        End If
        If ishareke(Mid(word, wp, 1)) = True Then
            '*nothing* Mid(nw, nwp, 1) = Mid(word, wp, 1)
        Else
            Mid(nw, nwp, 1) = Mid(word, wp, 1)
            nwp = nwp + 1
        End If
    Next wp
    
spagain:
    nwl = Len(nw)
    For i = 1 To nwl
        If Mid(nw, i, 1) = " " Then
            nw = Left(nw, nwl - 1)
            GoTo spagain
        End If
    Next i
    'normalizearapword = Right(nw, nwl - nwp)
    normalizearapword = nw
    'MsgBox ">>>" + normalizearapword + "<<<"
End Function

Private Function ishareke(ByVal c As String) As Boolean
    Dim i As Integer
    For i = 1 To 14
        If Left(c, 1) = elifbahareke(i) Then
            ishareke = True
            Exit Function
        End If
    Next i
    ishareke = False
    Exit Function
End Function


