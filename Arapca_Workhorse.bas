Attribute VB_Name = "Arapca_Workhorse"
Const aelif As Byte = &HAC
Const eelif As Byte = &HB3
Const ielif As Byte = &H26
Const elif As Byte = &HAB
Const be As Byte = &HBB
Const pe As Byte = &HDB
Const te As Byte = &HAE
Const se As Byte = &HC0
Const cim As Byte = &HC3
Const chim As Byte = &H80
Const ha As Byte = &HD5
Const hi As Byte = &H8C
Const dal As Byte = &H9C
Const zel As Byte = &H8B
Const re As Byte = &H97
Const ze As Byte = &H93
Const je As Byte = &H8A
Const ti As Byte = &HB9
Const zi As Byte = &HB4
Const ain As Byte = &H9F
Const gain As Byte = &H8D
Const ar_sin As Byte = &H94
Const shin As Byte = &H91
Const sat As Byte = &H92
Const dat As Byte = &HF7
Const fe As Byte = &H2D
Const ffe As Byte = &H98
Const kaf As Byte = &H82
Const kef As Byte = &H84
Const lam As Byte = &H89
Const mim As Byte = &HC2
Const nun As Byte = &HCA
Const vav As Byte = &HCB
Const hvav As Byte = &H83
Const uvav As Byte = &H99
Const ovav As Byte = &H9B
Const yuvav As Byte = &H34
Const yovav As Byte = &H35
Const pppvav As Byte = &H28
Const he As Byte = &H88
Const the As Byte = &H85
Const hhe As Byte = &HB6
Const hye As Byte = &H54
Const ye As Byte = &HC8
Rem -----------------------------------
Const b_aelif As Byte = aelif
Const b_eelif As Byte = eelif
Const b_ielif As Byte = ielif
Const b_elif As Byte = elif
Const b_be As Byte = &H8F
Const b_pe As Byte = &HC4
Const b_te As Byte = &HD4
Const b_se As Byte = &HDD
Const b_cim As Byte = &H32
Const b_chim As Byte = &H81
Const b_ha As Byte = &H30
Const b_hi As Byte = &H95
Const b_dal As Byte = dal
Const b_zel As Byte = zel
Const b_re As Byte = re
Const b_ze As Byte = ze
Const b_je As Byte = je
Const b_ti As Byte = &HB9
Const b_zi As Byte = &HBA
Const b_ain As Byte = &H90
Const b_gain As Byte = &H9E
Const b_sin As Byte = &H8E
Const b_shin As Byte = &H96
Const b_sat As Byte = &H23
Const b_dat As Byte = &H24
Const b_fe As Byte = &HA7
Const b_ffe As Byte = &HD1
Const b_kaf As Byte = &HAD
Const b_kef As Byte = &HBD
Const b_lam As Byte = &HBC
Const b_mim As Byte = &HA6
Const b_nun As Byte = &HEC
Const b_vav As Byte = vav
Const b_hvav As Byte = hvav
Const b_uvav As Byte = uvav
Const b_ovav As Byte = ovav
Const b_yuvav As Byte = yuvav
Const b_yovav As Byte = yovav
Const b_pppvav As Byte = pppvav
Const b_he As Byte = &HA3
Const b_the As Byte = the
Const b_hhe As Byte = hhe
Const b_ye As Byte = &HB2
Const b_hye As Byte = &H7A
Rem ----------------------------------
Const e_aelif As Byte = elif
Const e_eelif As Byte = eelif
Const e_ielif As Byte = &H53
Const e_elif As Byte = &H55
Const e_be As Byte = &H56
Const e_pe As Byte = &HE9
Const e_te As Byte = &H58
Const e_se As Byte = &H59
Const e_cim As Byte = &H5A
Const e_chim As Byte = &HEA
Const e_ha As Byte = &H60
Const e_hi As Byte = &H61
Const e_dal As Byte = &H62
Const e_zel As Byte = &H63
Const e_re As Byte = &H64
Const e_ze As Byte = &H65
Const e_je As Byte = &HF3
Const e_ti As Byte = &H6A
Const e_zi As Byte = &H6B
Const e_ain As Byte = &H6C
Const e_gain As Byte = &H6D
Const e_sin As Byte = &H66
Const e_shin As Byte = &H67
Const e_sat As Byte = &H68
Const e_dat As Byte = &H69
Const e_fe As Byte = &H6E
Const e_ffe As Byte = &HED
Const e_kaf As Byte = &H6F
Const e_kef As Byte = &H70
Const e_lam As Byte = &H71
Const e_mim As Byte = &H72
Const e_nun As Byte = &H73
Const e_vav As Byte = &H75
Const e_hvav As Byte = hvav
Const e_uvav As Byte = &HA4
Const e_ovav As Byte = &HA9
Const e_yuvav As Byte = &H36
Const e_yovav As Byte = &H37
Const e_pppvav As Byte = &H29
Const e_he As Byte = &HEB
Const e_the As Byte = &H57
Const e_hhe As Byte = hhe
Const e_ye As Byte = &H76
Const e_hye As Byte = &H54
Rem //------------------------------
Const m_aelif As Byte = elif
Const m_eelif As Byte = eelif
Const m_ielif As Byte = e_ielif
Const m_elif As Byte = e_elif
Const m_be As Byte = &HBE
Const m_pe As Byte = &HE1
Const m_te As Byte = &HD7
Const m_se As Byte = &HDE
Const m_cim As Byte = &H9D
Const m_chim As Byte = &HC7
Const m_ha As Byte = &H31
Const m_hi As Byte = &HA5
Const m_dal As Byte = e_dal
Const m_zel As Byte = e_zel
Const m_re As Byte = e_re
Const m_ze As Byte = e_ze
Const m_je As Byte = e_je
Const m_ti As Byte = &H44
Const m_zi As Byte = &H45
Const m_ain As Byte = &H46
Const m_gain As Byte = &H47
Const m_sin As Byte = &H8E
Const m_shin As Byte = &H41
Const m_sat As Byte = &H42
Const m_dat As Byte = &H43
Const m_fe As Byte = &H48
Const m_ffe As Byte = &HE7
Const m_kaf As Byte = &H49
Const m_kef As Byte = &H4A
Const m_lam As Byte = &H2A
Const m_mim As Byte = &H4C
Const m_nun As Byte = &H4D
Const m_vav As Byte = vav
Const m_hvav As Byte = hvav
Const m_uvav As Byte = uvav
Const m_ovav As Byte = ovav
Const m_yuvav As Byte = yuvav
Const m_yovav As Byte = yovav
Const m_pppvav As Byte = pppvav
Const m_he As Byte = &HE4
Const m_the As Byte = the
Const m_hhe As Byte = hhe
Const m_ye As Byte = &H9A
Const m_hye As Byte = &H18
Rem //----------------------------------
Const lamelif As Byte = &HF4
Const b_lamelif As Byte = lamelif
Const e_lamelif As Byte = &HF6
Const m_lamelif As Byte = e_lamelif

Const alamelif As Byte = &HFC
Const b_alamelif As Byte = lamelif
Const e_alamelif As Byte = &H78
Const m_alamelif As Byte = e_lamelif

Const gef As Byte = &HAF
Const b_gef As Byte = &HD6
Const m_gef As Byte = &H7E
Const e_gef As Byte = &H40

Const nazal As Byte = &H21
Const e_nazal As Byte = &H22
Const m_nazal As Byte = &H25
Const b_nazal As Byte = &H27

Const hemze As Byte = &HA1
Rem //------------------------------ hareke
Const sukun As Byte = &HDA
Const ustun As Byte = &HD3
Const kesre As Byte = &HC5
Const sedde As Byte = &HD2
Const tenvin As Byte = &HCE
Const ustsedde As Byte = &H5D
Const asedde As Byte = &HE6
Const likevav As Byte = &H33
Const likevavn As Byte = &HCF

Const likevavsedde As Byte = &H2B
Const aa As Byte = &HD8
Const ii As Byte = &HF5
Const iii As Byte = &HCC
Const ppp As Byte = &H2C

Dim yalin(50) As Byte
Dim basta(50) As Byte
Dim ortada(50) As Byte
Dim sonda(50) As Byte
Dim hareke(20) As Byte
Dim birlesmeyen(50) As Byte



Function arabkeyb(ByVal c As Long) As Long
    If basta(10) <> b_chim Then init_arab
    arabkeyb = c
    
    Select Case c
    Case Asc("a")
        arabkeyb = elif
    Case Asc("b")
        arabkeyb = be
    Case Asc("c")
        arabkeyb = cim
    Case Asc("d")
        arabkeyb = dal
    Case Asc("e")
        arabkeyb = eelif    '//**********
    Case Asc("f")
        arabkeyb = fe
    Case Asc("g")
        arabkeyb = gain
    Case Asc("h")
        arabkeyb = ha
    Case Asc("ý")
        arabkeyb = ye
    Case Asc("i")
        arabkeyb = ye
    Case Asc("j")
        arabkeyb = je
    Case Asc("k")
        arabkeyb = kef
    Case Asc("l")
        arabkeyb = lam
    Case Asc("m")
        arabkeyb = mim
    Case Asc("n")
        arabkeyb = nun
    Case Asc("o")
        arabkeyb = sukun    '//hareke
    Case Asc("p")
        arabkeyb = pe
    Case Asc("q")
        arabkeyb = kaf
    Case Asc("r")
        arabkeyb = re
    Case Asc("s")
        arabkeyb = ar_sin
    Case Asc("t")
        arabkeyb = te
    Case Asc("u")
        arabkeyb = pppvav
    Case Asc("v")
        arabkeyb = vav
    Case Asc("w")
        arabkeyb = sat
    Case Asc("x")
        arabkeyb = dat
    Case Asc("y")
        arabkeyb = ye
    Case Asc("z")
        arabkeyb = ze
    Case 60
        arabkeyb = se
    '//---------------
    Case Asc("A")
        arabkeyb = aelif   '// uzun a
    Case Asc("B")
        arabkeyb = the   '// yvarlak te
    Case Asc("C")
        arabkeyb = chim
    Case Asc("D")
        arabkeyb = zel
    Case Asc("E")
        arabkeyb = hemze
    Case Asc("F")
        arabkeyb = ffe
    Case Asc("G")
        arabkeyb = gef
    Case Asc("H")
        arabkeyb = hi
    Case Asc("I")
        arabkeyb = ielif
    Case Asc("Ý")
        arabkeyb = ielif
    Case Asc("K")
        arabkeyb = he
    Case Asc("L")
        arabkeyb = lamelif
    Case Asc("N")
        arabkeyb = nazal
    Case Asc("O")
        arabkeyb = ovav
    Case Asc("S")
        arabkeyb = shin
    Case Asc("T")
        arabkeyb = ti
    Case Asc("U")
        arabkeyb = uvav
    Case Asc("V")
        arabkeyb = hvav
    Case Asc("Y")
        arabkeyb = hye
    '//---------------
'//    case 1           arabkeyb =  eelif
    Case 7
        arabkeyb = ain
    Case 11
        arabkeyb = hhe
    Case 12
        arabkeyb = alamelif
    Case 15
        arabkeyb = yovav
    Case 20
        arabkeyb = zi
    Case 21
        arabkeyb = yuvav

    '//---------------- hareke
    Case 49
        arabkeyb = ustun
    Case 50
        arabkeyb = kesre
    Case 51
        arabkeyb = sedde
    Case 52
        arabkeyb = ustsedde
    Case 53
        arabkeyb = tenvin
    Case 54
        arabkeyb = likevav
    Case 55
        arabkeyb = aa
    Case 56
        arabkeyb = ii
    Case 57
        arabkeyb = iii
    Case Asc("0")
        arabkeyb = likevavsedde
    End Select
End Function

Private Sub init_arab()

yalin(0) = 255
yalin(1) = aelif
yalin(2) = eelif
yalin(3) = ielif
yalin(4) = elif
yalin(5) = be
yalin(6) = pe
yalin(7) = te
yalin(8) = se
yalin(9) = cim
yalin(10) = chim
yalin(11) = ha
yalin(12) = hi
yalin(13) = dal
yalin(14) = zel
yalin(15) = re
yalin(16) = ze
yalin(17) = je
yalin(18) = ti
yalin(19) = zi
yalin(20) = ain
yalin(21) = gain
yalin(22) = ar_sin
yalin(23) = shin
yalin(24) = sat
yalin(25) = dat
yalin(26) = fe
yalin(27) = ffe
yalin(28) = kaf
yalin(29) = kef
yalin(30) = gef
yalin(31) = nazal
yalin(32) = lam
yalin(33) = mim
yalin(34) = nun
yalin(35) = vav
yalin(36) = hvav
yalin(37) = uvav
yalin(38) = ovav
yalin(39) = yuvav
yalin(40) = yovav
yalin(41) = pppvav
yalin(42) = he
yalin(43) = the
yalin(44) = hhe
yalin(45) = ye
yalin(46) = hye
yalin(47) = lamelif
yalin(48) = alamelif
yalin(49) = 0


basta(0) = 255
basta(1) = b_aelif
basta(2) = b_eelif
basta(3) = b_ielif
basta(4) = b_elif
basta(5) = b_be
basta(6) = b_pe
basta(7) = b_te
basta(8) = b_se
basta(9) = b_cim
basta(10) = b_chim
basta(11) = b_ha
basta(12) = b_hi
basta(13) = b_dal
basta(14) = b_zel
basta(15) = b_re
basta(16) = b_ze
basta(17) = b_je
basta(18) = b_ti
basta(19) = b_zi
basta(20) = b_ain
basta(21) = b_gain
basta(22) = b_sin
basta(23) = b_shin
basta(24) = b_sat
basta(25) = b_dat
basta(26) = b_fe
basta(27) = b_ffe
basta(28) = b_kaf
basta(29) = b_kef
basta(30) = b_gef
basta(31) = b_nazal
basta(32) = b_lam
basta(33) = b_mim
basta(34) = b_nun
basta(35) = b_vav
basta(36) = b_hvav
basta(37) = b_uvav
basta(38) = b_ovav
basta(39) = b_yuvav
basta(40) = b_yovav
basta(41) = b_pppvav
basta(42) = b_he
basta(43) = b_the
basta(44) = b_hhe
basta(45) = b_ye
basta(46) = b_hye
basta(47) = b_lamelif
basta(48) = b_alamelif
basta(49) = 0


ortada(0) = 255
ortada(1) = m_aelif
ortada(2) = m_eelif
ortada(3) = m_ielif
ortada(4) = m_elif
ortada(5) = m_be
ortada(6) = m_pe
ortada(7) = m_te
ortada(8) = m_se
ortada(9) = m_cim
ortada(10) = m_chim
ortada(11) = m_ha
ortada(12) = m_hi
ortada(13) = m_dal
ortada(14) = m_zel
ortada(15) = m_re
ortada(16) = m_ze
ortada(17) = m_je
ortada(18) = m_ti
ortada(19) = m_zi
ortada(20) = m_ain
ortada(21) = m_gain
ortada(22) = m_sin
ortada(23) = m_shin
ortada(24) = m_sat
ortada(25) = m_dat
ortada(26) = m_fe
ortada(27) = m_ffe
ortada(28) = m_kaf
ortada(29) = m_kef
ortada(30) = m_gef
ortada(31) = m_nazal
ortada(32) = m_lam
ortada(33) = m_mim
ortada(34) = m_nun
ortada(35) = m_vav
ortada(36) = m_hvav
ortada(37) = m_uvav
ortada(38) = m_ovav
ortada(39) = m_yuvav
ortada(40) = m_yovav
ortada(41) = m_pppvav
ortada(42) = m_he
ortada(43) = m_the
ortada(44) = m_hhe
ortada(45) = m_ye
ortada(46) = m_hye
ortada(47) = m_lamelif
ortada(48) = m_alamelif
ortada(49) = 0


sonda(0) = 255
sonda(1) = e_aelif
sonda(2) = e_eelif
sonda(3) = e_ielif
sonda(4) = e_elif
sonda(5) = e_be
sonda(6) = e_pe
sonda(7) = e_te
sonda(8) = e_se
sonda(9) = e_cim
sonda(10) = e_chim
sonda(11) = e_ha
sonda(12) = e_hi
sonda(13) = e_dal
sonda(14) = e_zel
sonda(15) = e_re
sonda(16) = e_ze
sonda(17) = e_je
sonda(18) = e_ti
sonda(19) = e_zi
sonda(20) = e_ain
sonda(21) = e_gain
sonda(22) = e_sin
sonda(23) = e_shin
sonda(24) = e_sat
sonda(25) = e_dat
sonda(26) = e_fe
sonda(27) = e_ffe
sonda(28) = e_kaf
sonda(29) = e_kef
sonda(30) = e_gef
sonda(31) = e_nazal
sonda(32) = e_lam
sonda(33) = e_mim
sonda(34) = e_nun
sonda(35) = e_vav
sonda(36) = e_hvav
sonda(37) = e_uvav
sonda(38) = e_ovav
sonda(39) = e_yuvav
sonda(40) = e_yovav
sonda(41) = e_pppvav
sonda(42) = e_he
sonda(43) = e_the
sonda(44) = e_hhe
sonda(45) = e_ye
sonda(46) = e_hye
sonda(47) = e_lamelif
sonda(48) = e_alamelif
sonda(49) = 0


hareke(0) = 255
hareke(1) = sukun
hareke(2) = ustun
hareke(3) = kesre
hareke(4) = sedde
hareke(5) = tenvin
hareke(6) = ustsedde
hareke(7) = hemze
hareke(8) = likevav
hareke(9) = aa
hareke(10) = ii
hareke(11) = iii
hareke(12) = likevavsedde
hareke(13) = likevavn
hareke(14) = asedde
hareke(15) = ppp
hareke(16) = 0


birlesmeyen(0) = aelif
birlesmeyen(1) = eelif
birlesmeyen(2) = ielif
birlesmeyen(3) = elif
birlesmeyen(4) = dal
birlesmeyen(5) = zel
birlesmeyen(6) = re
birlesmeyen(7) = ze
birlesmeyen(8) = je
birlesmeyen(9) = vav
birlesmeyen(10) = hvav
birlesmeyen(11) = uvav
birlesmeyen(12) = ovav
birlesmeyen(13) = yuvav
birlesmeyen(14) = yovav
birlesmeyen(15) = pppvav
birlesmeyen(16) = lamelif
birlesmeyen(17) = alamelif
birlesmeyen(18) = e_aelif
birlesmeyen(19) = e_eelif
birlesmeyen(20) = e_ielif
birlesmeyen(21) = e_elif
birlesmeyen(22) = e_dal
birlesmeyen(23) = e_zel
birlesmeyen(24) = e_re
birlesmeyen(25) = e_ze
birlesmeyen(26) = e_je
birlesmeyen(27) = e_vav
birlesmeyen(28) = e_hvav
birlesmeyen(29) = e_uvav
birlesmeyen(30) = e_ovav
birlesmeyen(31) = e_yuvav
birlesmeyen(32) = e_yovav
birlesmeyen(33) = e_pppvav
birlesmeyen(34) = e_lamelif
birlesmeyen(35) = e_alamelif
birlesmeyen(36) = 0
End Sub

Private Function getprev(ByRef str As String, ByVal pos As Long) As Long
    Dim i As Long
    Dim k As Long
    Dim c As Byte
    
    For i = pos To Len(str)
        c = Asc(Mid(str, i, 1))
        For k = 0 To 20
            If (c <> hareke(k)) Then GoTo end_getprev
        Next k
    Next i
end_getprev:
    getprev = i
End Function

Private Sub setprev(ByRef str As String, ByVal pos As Long, ByRef x As String)
    Dim i As Long
    Dim k As Long
    Dim c As Byte
    
    For i = pos To Len(str)
        c = Asc(Mid(str, i, 1))
        For k = 0 To 20
            If (c <> hareke(k)) Then
            Mid(str, i, 1) = x
            GoTo end_setprev
            End If
        Next k
    Next i
end_setprev:
End Sub

Private Function getharf(ByVal harf As String) As Long
    Dim i As Integer
    Dim c As Byte
    
    c = Asc(Left(harf, 1))
    For i = 1 To 50
        If yalin(i) = c Then
            getharf = i
            GoTo end_getharf
        End If
    Next i
    
    For i = 1 To 50
        If basta(i) = c Then
            getharf = i
            GoTo end_getharf
        End If
    Next i
    
    For i = 1 To 50
        If ortada(i) = c Then
            getharf = i
            GoTo end_getharf
        End If
    Next i
    
    For i = 1 To 50
        If sonda(i) = c Then
            getharf = i
            GoTo end_getharf
        End If
    Next i
    getharf = 0
end_getharf:
    
End Function

'//----------------------
Public Sub doarab(ByRef str As String)
    Dim strend As Long
    Dim CurIdx As Long
    Dim k As Long
    Dim CurHarf As Byte
    Dim PrevHarf As Byte
    
    If basta(10) <> b_chim Then init_arab
    
    For CurIdx = Len(str) To 1 Step -1
        CurHarf = getharf(Mid(str, CurIdx, 1))
        If CurIdx = Len(str) Then
            'setharf
        End If
        PrevHarf = getharf(Mid(str, getprev(str, CurIdx), 1))
        
        If getprev(str, CurIdx) = Len(str) Then GoTo end_main_loop
        
        For k = 0 To 20
            If CurHarf = hareke(k) Then GoTo end_main_loop
        Next k
            
        If PrevHarf = 0 Then
            Mid(str, CurIdx, 1) = Chr(yalin(CurHarf))
            GoTo end_main_loop
        End If

        For k = 0 To 50
            If CurHarf = birlesmeyen(k) Then GoTo end_main_loop
        Next k

    'AAAAAAAAAAAAAARRRRRRRRRRRRRRHHHHHhhhhhhhhhhhh!!!
end_main_loop:
    Next CurIdx

End Sub
        

