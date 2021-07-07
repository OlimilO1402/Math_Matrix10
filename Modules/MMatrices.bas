Attribute VB_Name = "MMatrices"
Option Explicit 'Zeilen: (nov.2019: 2638) (nov.2019: 3674)
'Matrizen-Rechner online:
'https://matrixcalc.org/de/
'oder auch
'https://matrizen-rechner.de/

'Eine Zeile in VB darf maximal 1024 Zeichen lang sein, um Platz zu sparen verwenden wir die kürzest möglichen Variablennamen,
'und das sind reine Buchstaben, jeder Variablennamen muss mit einem Buchstaben beginnen.
'bei Verwendung der Form a11, a12, a13 etc werden bereits 3 Zeichen pro Variable benötigt
'Zeile, Spalte

'Sicherlich werden Sie sich fragen, warum die ganzen Types und nicht einfach nur eine Array?
'Der Grund ist, weil man so den Algo besser erkennen kann, was bei der zwangsweise erforderlichen
'Verschachtelung von Schleifen nicht so schön erkennbar ist.
'jeder kann den Code lesen und lernen wie ein bestimmter Matrizen-Algo arbeitet, und sich dann
'selbst Matrix-funktionen programmieren die auf ein 2D-Array angewendet werden.
'
'
Public Type Matrix2 'Matrix 2x2 bsp "ab" ist das zweite Element der ersten Zeile, oder das erste Element der zweiten Spalte
    aa As Double:    ab As Double
    ba As Double:    bb As Double
End Type
Public Type Matrix3 'Matrix 3x3
    aa As Double:    ab As Double:    ac As Double
    ba As Double:    bb As Double:    bc As Double
    ca As Double:    cb As Double:    cc As Double
End Type
Public Type Matrix4 'Matrix 4x4
    aa As Double:    ab As Double:    ac As Double:    ad As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double
    da As Double:    db As Double:    dc As Double:    dd As Double
End Type
Public Type Matrix5 'Matrix 5x5
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double
End Type
Public Type Matrix6 'Matrix 6x6
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double
End Type
Public Type Matrix7 'Matrix 7x7
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double
End Type
Public Type Matrix8 'Matrix 8x8
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double:    ah As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double:    bh As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double:    ch As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double:    dh As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double:    eh As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double:    fh As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double:    gh As Double
    ha As Double:    hb As Double:    HC As Double:    hd As Double:    he As Double:    hf As Double:    hg As Double:    hh As Double
End Type
Public Type Matrix9 'Matrix 9x9 'here the variablename "if" occurs, this is OK in a type
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double:    ah As Double:    ai As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double:    bh As Double:    bi As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double:    ch As Double:    ci As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double:    dh As Double:    di As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double:    eh As Double:    ei As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double:    fh As Double:    fi As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double:    gh As Double:    gi As Double
    ha As Double:    hb As Double:    HC As Double:    hd As Double:    he As Double:    hf As Double:    hg As Double:    hh As Double:    hi As Double
    ia As Double:    ib As Double:    ic As Double:    id As Double:    ie As Double:    if As Double:    ig As Double:    ih As Double:    ii As Double
End Type
Public Type Matrix10 'Matrix 10x10
    aa As Double:    ab As Double:    ac As Double:    ad As Double:    ae As Double:    af As Double:    ag As Double:    ah As Double:    ai As Double:    aj As Double
    ba As Double:    bb As Double:    bc As Double:    bd As Double:    be As Double:    bf As Double:    bg As Double:    bh As Double:    bi As Double:    bj As Double
    ca As Double:    cb As Double:    cc As Double:    cd As Double:    ce As Double:    cf As Double:    cg As Double:    ch As Double:    ci As Double:    cj As Double
    da As Double:    db As Double:    dc As Double:    dd As Double:    de As Double:    df As Double:    dg As Double:    dh As Double:    di As Double:    dj As Double
    ea As Double:    eb As Double:    ec As Double:    ed As Double:    ee As Double:    ef As Double:    eg As Double:    eh As Double:    ei As Double:    ej As Double
    fa As Double:    fb As Double:    fc As Double:    fd As Double:    fe As Double:    ff As Double:    fg As Double:    fh As Double:    fi As Double:    fj As Double
    ga As Double:    gb As Double:    gc As Double:    gd As Double:    ge As Double:    gf As Double:    gg As Double:    gh As Double:    gi As Double:    gj As Double
    ha As Double:    hb As Double:    HC As Double:    hd As Double:    he As Double:    hf As Double:    hg As Double:    hh As Double:    hi As Double:    hj As Double
    ia As Double:    ib As Double:    ic As Double:    id As Double:    ie As Double:    if As Double:    ig As Double:    ih As Double:    ii As Double:    ij As Double
    ja As Double:    jb As Double:    jc As Double:    jd As Double:    je As Double:    jf As Double:    jg As Double:    jh As Double:    ji As Double:    jj As Double
End Type

'TODO:
'Of course we need Gaussian-Triangulation!!!!!!!!!!!!!!
'I ALSO WANT Singular Value Decomposition (SVD)!!!!!!!!
'WE ALSO NEED EIGENVECTOR AND EIGENVALUE!!!!!!!!!!!!!!!

'Public Type MatrixT
'    a() As Double
'End Type
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal nBytes As Long)
'Alias "cpymem" 'Zeile 106
Private Declare Sub PutOne Lib "msvbvm60.dll" Alias "PutMem8" (ByVal pDst As Long, Optional ByVal Src As Double = 1#)

'Einfache Matrix-Operationen
'Matrizen erzeugen
Public Function Mat2(aa As Double, ab As Double, _
                     ba As Double, bb As Double) As Matrix2
    'Erzeugt eine 2x2 Matrix
    With Mat2: .aa = aa: .ab = ab
               .ba = ba: .bb = bb
    End With
End Function
Public Function Mat3(aa As Double, ab As Double, ac As Double, _
                     ba As Double, bb As Double, bc As Double, _
                     ca As Double, cb As Double, cc As Double) As Matrix3
    'Erzeugt eine 3x3 Matrix
    With Mat3: .aa = aa: .ab = ab: .ac = ac
               .ba = ba: .bb = bb: .bc = bc
               .ca = ca: .cb = cb: .cc = cc
    End With
End Function
Public Function Mat4(aa As Double, ab As Double, ac As Double, ad As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, _
                     da As Double, db As Double, dc As Double, dd As Double) As Matrix4
    'Erzeugt eine 4x4 Matrix
    With Mat4: .aa = aa: .ab = ab: .ac = ac: .ad = ad
               .ba = ba: .bb = bb: .bc = bc: .bd = bd
               .ca = ca: .cb = cb: .cc = cc: .cd = cd
               .da = da: .db = db: .dc = dc: .dd = dd
    End With
End Function
Public Function Mat5(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, _
                     da As Double, db As Double, dc As Double, dd As Double, de As Double, _
                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double) As Matrix5
    'Erzeugt eine 5x5 Matrix
    With Mat5: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae
               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be
               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce
               .da = da: .db = db: .dc = dc: .dd = dd: .de = de
               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee
    End With
End Function
Public Function Mat6(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, af As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, bf As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, cf As Double, _
                     da As Double, db As Double, dc As Double, dd As Double, de As Double, df As Double, _
                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double, ef As Double, _
                     fa As Double, fb As Double, fc As Double, fd As Double, fe As Double, ff As Double) As Matrix6
    'Erzeugt eine 6x6 Matrix
    With Mat6: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae: .af = af
               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be: .bf = bf
               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce: .cf = cf
               .da = da: .db = db: .dc = dc: .dd = dd: .de = de: .df = df
               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee: .ef = ef
               .fa = fa: .fb = fb: .fc = fc: .fd = fd: .fe = fe: .ff = ff
    End With
End Function
Public Function Mat7(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, af As Double, ag As Double, _
                     ba As Double, bb As Double, bc As Double, bd As Double, be As Double, bf As Double, bg As Double, _
                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, cf As Double, cg As Double, _
                     da As Double, db As Double, dc As Double, dd As Double, de As Double, df As Double, dg As Double, _
                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double, ef As Double, eg As Double, _
                     fa As Double, fb As Double, fc As Double, fd As Double, fe As Double, ff As Double, fg As Double, _
                     ga As Double, gb As Double, gc As Double, gd As Double, ge As Double, gf As Double, gg As Double) As Matrix7
    'Erzeugt eine 7x7 Matrix
    With Mat7: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae: .af = af: .ag = ag
               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be: .bf = bf: .bg = bg
               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce: .cf = cf: .cg = cg
               .da = da: .db = db: .dc = dc: .dd = dd: .de = de: .df = df: .dg = dg
               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee: .ef = ef: .eg = eg
               .fa = fa: .fb = fb: .fc = fc: .fd = fd: .fe = fe: .ff = ff: .fg = fg
               .ga = ga: .gb = gb: .gc = gc: .gd = gd: .ge = ge: .gf = gf: .gg = gg
    End With
End Function
'not working, too much function parameters i guess
'Public Function Mat8(aa As Double, ab As Double, ac As Double, ad As Double, ae As Double, af As Double, ag As Double, ah As Double, ba As Double, bb As Double, bc As Double, bd As Double, be As Double, bf As Double, bg As Double, bh As Double, _
'                     ca As Double, cb As Double, cc As Double, cd As Double, ce As Double, cf As Double, cg As Double, ch As Double, da As Double, db As Double, dc As Double, dd As Double, de As Double, df As Double, dg As Double, dh As Double, _
'                     ea As Double, eb As Double, ec As Double, ed As Double, ee As Double, ef As Double, eg As Double, eh As Double, fa As Double, fb As Double, fc As Double, fd As Double, fe As Double, ff As Double, fg As Double, fh As Double, _
'                     ga As Double, gb As Double, gc As Double, gd As Double, ge As Double, gf As Double, gg As Double, gh As Double, ha As Double, hb As Double, hc As Double, hd As Double, he As Double, hf As Double, hg As Double, hh As Double) As Matrix8
'    'Erzeugt eine 8x8 Matrix
'    With Mat8: .aa = aa: .ab = ab: .ac = ac: .ad = ad: .ae = ae: .af = af: .ag = ag: .ah = ah
'               .ba = ba: .bb = bb: .bc = bc: .bd = bd: .be = be: .bf = bf: .bg = bg: .bh = bh
'               .ca = ca: .cb = cb: .cc = cc: .cd = cd: .ce = ce: .cf = cf: .cg = cg: .ch = ch
'               .da = da: .db = db: .dc = dc: .dd = dd: .de = de: .df = df: .dg = dg: .dh = dh
'               .ea = ea: .eb = eb: .ec = ec: .ed = ed: .ee = ee: .ef = ef: .eg = eg: .eh = eh
'               .fa = fa: .fb = fb: .fc = fc: .fd = fd: .fe = fe: .ff = ff: .fg = fg: .fh = fh
'               .ga = ga: .gb = gb: .gc = gc: .gd = gd: .ge = ge: .gf = gf: .gg = gg: .gh = gh
'               .ha = ha: .hb = hb: .hc = hc: .hd = hd: .he = he: .hf = hf: .hg = hg: .hh = hh
'    End With
'End Function
'OK folgende Möglichkeiten:
' * Paramarray, mit Array arbeiten jedes Element ist Variant also cpymem fällt aus
' * irgendwie mit Vektoren arbeiten aber wie Spalten- oder Zeilen-vektoren? OK Zeilenvektoren
Public Function Mat8(Row_a As Vector8, _
                     Row_b As Vector8, _
                     Row_c As Vector8, _
                     Row_d As Vector8, _
                     Row_e As Vector8, _
                     Row_f As Vector8, _
                     Row_g As Vector8, _
                     Row_h As Vector8) As Matrix8
    Dim p As Long: p = VarPtr(Mat8) 'size per row is 8 bytes per double and 8 variables = 8*8=64
    RtlMoveMemory ByVal p, Row_a, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_b, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_c, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_d, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_e, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_f, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_g, 64: p = p + 64
    RtlMoveMemory ByVal p, Row_h, 64
End Function
Public Function Mat9(Row_a As Vector9, _
                     Row_b As Vector9, _
                     Row_c As Vector9, _
                     Row_d As Vector9, _
                     Row_e As Vector9, _
                     Row_f As Vector9, _
                     Row_g As Vector9, _
                     Row_h As Vector9, _
                     Row_i As Vector9) As Matrix9
    Dim p As Long: p = VarPtr(Mat9) 'size per row is 8 bytes per double and 9 variables = 8*9=72
    RtlMoveMemory ByVal p, Row_a, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_b, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_c, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_d, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_e, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_f, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_g, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_h, 72: p = p + 72
    RtlMoveMemory ByVal p, Row_i, 72
End Function
Public Function Mat10(Row_a As Vector10, _
                      Row_b As Vector10, _
                      Row_c As Vector10, _
                      Row_d As Vector10, _
                      Row_e As Vector10, _
                      Row_f As Vector10, _
                      Row_g As Vector10, _
                      Row_h As Vector10, _
                      Row_i As Vector10, _
                      Row_j As Vector10) As Matrix10
    Dim p As Long: p = VarPtr(Mat10) 'size per row is 8 bytes per double and 10 variables = 8*10=80
    RtlMoveMemory ByVal p, Row_a, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_b, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_c, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_d, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_e, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_f, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_g, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_h, 80: p = p + 80
    RtlMoveMemory ByVal p, Row_i, 80
End Function

Public Function Mat2_ij() As Matrix2
    Mat2_ij = Mat2(11, 12, _
                   21, 22)
End Function
Public Function Mat3_ij() As Matrix3
    Mat3_ij = Mat3(11, 12, 13, _
                   21, 22, 23, _
                   31, 32, 33)
End Function
Public Function Mat4_ij() As Matrix4
    Mat4_ij = Mat4(11, 12, 13, 14, _
                   21, 22, 23, 24, _
                   31, 32, 33, 34, _
                   41, 42, 43, 44)
End Function
Public Function Mat5_ij() As Matrix5
    Mat5_ij = Mat5(11, 12, 13, 14, 15, _
                   21, 22, 23, 24, 25, _
                   31, 32, 33, 34, 35, _
                   41, 42, 43, 44, 45, _
                   51, 52, 53, 54, 55)
End Function
Public Function Mat6_ij() As Matrix6
    Mat6_ij = Mat6(11, 12, 13, 14, 15, 16, _
                   21, 22, 23, 24, 25, 26, _
                   31, 32, 33, 34, 35, 36, _
                   41, 42, 43, 44, 45, 46, _
                   51, 52, 53, 54, 55, 56, _
                   61, 62, 63, 64, 65, 66)
End Function
Public Function Mat7_ij() As Matrix7
    Mat7_ij = Mat7(11, 12, 13, 14, 15, 16, 17, _
                   21, 22, 23, 24, 25, 26, 27, _
                   31, 32, 33, 34, 35, 36, 37, _
                   41, 42, 43, 44, 45, 46, 47, _
                   51, 52, 53, 54, 55, 56, 57, _
                   61, 62, 63, 64, 65, 66, 67, _
                   71, 72, 73, 74, 75, 76, 77)
End Function
Public Function Mat8_ij() As Matrix8
    Mat8_ij = Mat8(Vec8(11, 12, 13, 14, 15, 16, 17, 18), _
                   Vec8(21, 22, 23, 24, 25, 26, 27, 28), _
                   Vec8(31, 32, 33, 34, 35, 36, 37, 38), _
                   Vec8(41, 42, 43, 44, 45, 46, 47, 48), _
                   Vec8(51, 52, 53, 54, 55, 56, 57, 58), _
                   Vec8(61, 62, 63, 64, 65, 66, 67, 68), _
                   Vec8(71, 72, 73, 74, 75, 76, 77, 78), _
                   Vec8(81, 82, 83, 84, 85, 86, 87, 88))
End Function
Public Function Mat9_ij() As Matrix9
    Mat9_ij = Mat9(Vec9(11, 12, 13, 14, 15, 16, 17, 18, 19), _
                   Vec9(21, 22, 23, 24, 25, 26, 27, 28, 29), _
                   Vec9(31, 32, 33, 34, 35, 36, 37, 38, 39), _
                   Vec9(41, 42, 43, 44, 45, 46, 47, 48, 49), _
                   Vec9(51, 52, 53, 54, 55, 56, 57, 58, 59), _
                   Vec9(61, 62, 63, 64, 65, 66, 67, 68, 69), _
                   Vec9(71, 72, 73, 74, 75, 76, 77, 78, 79), _
                   Vec9(81, 82, 83, 84, 85, 86, 87, 88, 89), _
                   Vec9(91, 92, 93, 94, 95, 96, 97, 98, 99))
End Function
Public Function Mat10_ij() As Matrix10
    Mat10_ij = Mat10(Vec10(1, 2, 3, 4, 5, 6, 7, 8, 9, 10), _
                     Vec10(11, 12, 13, 14, 15, 16, 17, 18, 19, 20), _
                     Vec10(21, 22, 23, 24, 25, 26, 27, 28, 29, 30), _
                     Vec10(31, 32, 33, 34, 35, 36, 37, 38, 39, 40), _
                     Vec10(41, 42, 43, 44, 45, 46, 47, 48, 49, 50), _
                     Vec10(51, 52, 53, 54, 55, 56, 57, 58, 59, 60), _
                     Vec10(61, 62, 63, 64, 65, 66, 67, 68, 69, 70), _
                     Vec10(71, 72, 73, 74, 75, 76, 77, 78, 79, 80), _
                     Vec10(81, 82, 83, 84, 85, 86, 87, 88, 89, 90), _
                     Vec10(91, 92, 93, 94, 95, 96, 97, 98, 99, 100))
End Function

'Einheits-Matrizen erzeugen
Public Function Mat2_E() As Matrix2
    'Erzeugt eine 2x2 Einheits-Matrix, alle Variablen der Diagonale sind 1 alle Anderen sind 0
    With Mat2_E: .aa = 1: .bb = 1:    End With
End Function
Public Function Mat3_E() As Matrix3
    'Erzeugt eine 3x3 Einheits-Matrix
    With Mat3_E: .aa = 1: .bb = 1: .cc = 1:    End With
End Function
Public Function Mat4_E() As Matrix4
    'Erzeugt eine 4x4 Einheits-Matrix
    With Mat4_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1:    End With
End Function
Public Function Mat5_E() As Matrix5
    'Erzeugt eine 5x5 Einheits-Matrix
    With Mat5_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1:    End With
End Function
Public Function Mat6_E() As Matrix6
    'Erzeugt eine 6x6 Einheits-Matrix
    With Mat6_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1:    End With
End Function
Public Function Mat7_E() As Matrix7
    'Erzeugt eine 7x7 Einheits-Matrix
    With Mat7_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1:    End With
End Function
Public Function Mat8_E() As Matrix8
    'Erzeugt eine 8x8 Einheits-Matrix
    With Mat8_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1: .hh = 1:    End With
End Function
Public Function Mat9_E() As Matrix9
    'Erzeugt eine 9x9 Einheits-Matrix
    With Mat9_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1: .hh = 1: .ii = 1:    End With
End Function
Public Function Mat10_E() As Matrix10
    'Erzeugt eine 10x10 Einheits-Matrix
    With Mat10_E: .aa = 1: .bb = 1: .cc = 1: .dd = 1: .ee = 1: .ff = 1: .gg = 1: .hh = 1: .ii = 1: .jj = 1:    End With
End Function

Public Function Mat2_Ones() As Matrix2
    'Erzeugt eine 2x2-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat2_Ones):    n = 2
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat3_Ones() As Matrix3
    'Erzeugt eine 3x3-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat3_Ones):    n = 3
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat4_Ones() As Matrix4
    'Erzeugt eine 4x4-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat4_Ones):    n = 4
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat5_Ones() As Matrix5
    'Erzeugt eine 5x5-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat5_Ones):    n = 5
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat6_Ones() As Matrix6
    'Erzeugt eine 6x6-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat6_Ones):    n = 6
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat7_Ones() As Matrix7
    'Erzeugt eine 7x7-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat7_Ones):    n = 7
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat8_Ones() As Matrix8
    'Erzeugt eine 8x8-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat8_Ones):    n = 8
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat9_Ones() As Matrix9
    'Erzeugt eine 9x9-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat9_Ones):    n = 9
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function
Public Function Mat10_Ones() As Matrix10
    'Erzeugt eine 10x10-Matriz mit allen Variablen=1
    Dim i As Long, n As Long, p As Long
    p = VarPtr(Mat10_Ones):    n = 10
    For i = 0 To n ^ 2 - 1:   PutOne p: p = p + 8:    Next
End Function

Public Function Mat2_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix2
    'erzeugt eine 2x2-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat2_Rnd = Mat2(rv, rv, _
                    rv, rv)
End Function
Public Function Mat3_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix3
    'erzeugt eine 3x3-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat3_Rnd = Mat3(rv, rv, rv, _
                    rv, rv, rv, _
                    rv, rv, rv)
End Function
Public Function Mat4_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix4
    'erzeugt eine 4x4-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat4_Rnd = Mat4(rv, rv, rv, rv, _
                    rv, rv, rv, rv, _
                    rv, rv, rv, rv, _
                    rv, rv, rv, rv)
End Function
Public Function Mat5_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix5
    'erzeugt eine 5x5-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat5_Rnd = Mat5(rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv)
End Function
Public Function Mat6_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix6
    'erzeugt eine 6x6-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat6_Rnd = Mat6(rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv)
End Function
Public Function Mat7_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix7
    'erzeugt eine 6x6-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat7_Rnd = Mat7(rv, rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, rv, _
                    rv, rv, rv, rv, rv, rv, rv)
End Function
Public Function Mat8_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix8
    'erzeugt eine 6x6-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat8_Rnd = Mat8(Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec8(rv, rv, rv, rv, rv, rv, rv, rv))
End Function
Public Function Mat9_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix9
    'erzeugt eine 6x6-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat9_Rnd = Mat9(Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                    Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv))
End Function
Public Function Mat10_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Matrix10
    'erzeugt eine 6x6-Matrix mit Zufallszahlen im Bereich von dbl_from bis dbl_to
    Rnd_SetFromTo dbl_From, dbl_To
    Mat10_Rnd = Mat10(Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv), _
                      Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv))
End Function


Public Function Mat2_IsSym(m As Matrix2) As Boolean
    With m:     Mat2_IsSym = .ab = .ba:    End With
End Function
Public Function Mat3_IsSym(m As Matrix3) As Boolean
    With m:     Mat3_IsSym = (.ab = .ba) And (.ac = .ca) And (.bc = .cb):    End With
End Function
Public Function Mat4_IsSym(m As Matrix4) As Boolean
    With m
        Mat4_IsSym = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.bc = .cb) And (.bd = .db) And (.cd = .dc)
    End With
End Function
Public Function Mat5_IsSym(m As Matrix5) As Boolean
    Dim b As Boolean
    With m
        b = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.ae = .ea): If Not b Then Exit Function
        b = (.bc = .cb) And (.bd = .db) And (.be = .eb): If Not b Then Exit Function
        b = (.cd = .dc) And (.ce = .ec): If Not b Then Exit Function
        b = (.de = .ed): If Not b Then Exit Function
    End With
    Mat5_IsSym = True
End Function
Public Function Mat6_IsSym(m As Matrix6) As Boolean
    Dim b As Boolean
    With m
        b = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.ae = .ea) And (.af = .fa): If Not b Then Exit Function
        b = (.bc = .cb) And (.bd = .db) And (.be = .eb) And (.bf = .fb): If Not b Then Exit Function
        b = (.cd = .dc) And (.ce = .ec) And (.cf = .fc): If Not b Then Exit Function
        b = (.de = .ed) And (.df = .fd): If Not b Then Exit Function
        b = (.ef = .fe): If Not b Then Exit Function
    End With
    Mat6_IsSym = True
End Function
Public Function Mat7_IsSym(m As Matrix7) As Boolean
    Dim b As Boolean
    With m
        b = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.ae = .ea) And (.af = .fa) And (.ag = .ga): If Not b Then Exit Function
        b = (.bc = .cb) And (.bd = .db) And (.be = .eb) And (.bf = .fb) And (.bg = .gb): If Not b Then Exit Function
        b = (.cd = .dc) And (.ce = .ec) And (.cf = .fc) And (.cg = .gc): If Not b Then Exit Function
        b = (.de = .ed) And (.df = .fd) And (.dg = .gd): If Not b Then Exit Function
        b = (.ef = .fe) And (.eg = .ge): If Not b Then Exit Function
        b = (.fg = .gf): If Not b Then Exit Function
    End With
    Mat7_IsSym = True
End Function
Public Function Mat8_IsSym(m As Matrix8) As Boolean
    Dim b As Boolean
    With m
        b = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.ae = .ea) And (.af = .fa) And (.ag = .ga) And (.ah = .ha): If Not b Then Exit Function
        b = (.bc = .cb) And (.bd = .db) And (.be = .eb) And (.bf = .fb) And (.bg = .gb) And (.bh = .hb): If Not b Then Exit Function
        b = (.cd = .dc) And (.ce = .ec) And (.cf = .fc) And (.cg = .gc) And (.ch = .HC): If Not b Then Exit Function
        b = (.de = .ed) And (.df = .fd) And (.dg = .gd) And (.dh = .hd): If Not b Then Exit Function
        b = (.ef = .fe) And (.eg = .ge) And (.eh = .he): If Not b Then Exit Function
        b = (.fg = .gf) And (.fh = .hf): If Not b Then Exit Function
        b = (.gh = .hg): If Not b Then Exit Function
    End With
    Mat8_IsSym = True
End Function
Public Function Mat9_IsSym(m As Matrix9) As Boolean
    Dim b As Boolean
    With m
        b = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.ae = .ea) And (.af = .fa) And (.ag = .ga) And (.ah = .ha) And (.ai = .ia): If Not b Then Exit Function
        b = (.bc = .cb) And (.bd = .db) And (.be = .eb) And (.bf = .fb) And (.bg = .gb) And (.bh = .hb) And (.bi = .ib): If Not b Then Exit Function
        b = (.cd = .dc) And (.ce = .ec) And (.cf = .fc) And (.cg = .gc) And (.ch = .HC) And (.ci = .ic): If Not b Then Exit Function
        b = (.de = .ed) And (.df = .fd) And (.dg = .gd) And (.dh = .hd) And (.di = .id): If Not b Then Exit Function
        b = (.ef = .fe) And (.eg = .ge) And (.eh = .he) And (.ei = .ie): If Not b Then Exit Function
        b = (.fg = .gf) And (.fh = .hf) And (.fi = .if): If Not b Then Exit Function
        b = (.gh = .hg) And (.gi = .ig): If Not b Then Exit Function
        b = (.hi = .ih): If Not b Then Exit Function
    End With
    Mat9_IsSym = True
End Function
Public Function Mat10_IsSym(m As Matrix10) As Boolean
    Dim b As Boolean
    With m
        b = (.ab = .ba) And (.ac = .ca) And (.ad = .da) And (.ae = .ea) And (.af = .fa) And (.ag = .ga) And (.ah = .ha) And (.ai = .ia) And (.aj = .ja): If Not b Then Exit Function
        b = (.bc = .cb) And (.bd = .db) And (.be = .eb) And (.bf = .fb) And (.bg = .gb) And (.bh = .hb) And (.bi = .ib) And (.bj = .jb): If Not b Then Exit Function
        b = (.cd = .dc) And (.ce = .ec) And (.cf = .fc) And (.cg = .gc) And (.ch = .HC) And (.ci = .ic) And (.cj = .jc): If Not b Then Exit Function
        b = (.de = .ed) And (.df = .fd) And (.dg = .gd) And (.dh = .hd) And (.di = .id) And (.dj = .jd): If Not b Then Exit Function
        b = (.ef = .fe) And (.eg = .ge) And (.eh = .he) And (.ei = .ie) And (.ej = .je): If Not b Then Exit Function
        b = (.fg = .gf) And (.fh = .hf) And (.fi = .if) And (.fj = .jf): If Not b Then Exit Function
        b = (.gh = .hg) And (.gi = .ig) And (.gj = .jg): If Not b Then Exit Function
        b = (.hi = .ih) And (.hj = .jh): If Not b Then Exit Function
        b = (.ij = .ji): If Not b Then Exit Function
    End With
    Mat10_IsSym = True
End Function


Public Function Mat2_trace(m As Matrix2) As Double
    'berechnet von einer 2x2-Matrix die Spur (=Summe der Diagonalelemente)
    Mat2_trace = m.aa + m.bb
End Function
Public Function Mat3_trace(m As Matrix3) As Double
    'berechnet von einer 3x3-Matrix die Spur (=Summe der Diagonalelemente)
    Mat3_trace = m.aa + m.bb + m.cc
End Function
Public Function Mat4_trace(m As Matrix4) As Double
    'berechnet von einer 4x4-Matrix die Spur (=Summe der Diagonalelemente)
    Mat4_trace = m.aa + m.bb + m.cc + m.dd
End Function
Public Function Mat5_trace(m As Matrix5) As Double
    'berechnet von einer 5x5-Matrix die Spur (=Summe der Diagonalelemente)
    Mat5_trace = m.aa + m.bb + m.cc + m.dd + m.ee
End Function
Public Function Mat6_trace(m As Matrix6) As Double
    'berechnet von einer 6x6-Matrix die Spur (=Summe der Diagonalelemente)
    Mat6_trace = m.aa + m.bb + m.cc + m.dd + m.ee + m.ff
End Function
Public Function Mat7_trace(m As Matrix7) As Double
    'berechnet von einer 7x7-Matrix die Spur (=Summe der Diagonalelemente)
    Mat7_trace = m.aa + m.bb + m.cc + m.dd + m.ee + m.ff + m.gg
End Function
Public Function Mat8_trace(m As Matrix8) As Double
    'berechnet von einer 8x8-Matrix die Spur (=Summe der Diagonalelemente)
    Mat8_trace = m.aa + m.bb + m.cc + m.dd + m.ee + m.ff + m.gg + m.hh
End Function
Public Function Mat9_trace(m As Matrix9) As Double
    'berechnet von einer 9x9-Matrix die Spur (=Summe der Diagonalelemente)
    Mat9_trace = m.aa + m.bb + m.cc + m.dd + m.ee + m.ff + m.gg + m.hh + m.ii
End Function
Public Function Mat10_trace(m As Matrix10) As Double
    'berechnet von einer 10x10-Matrix die Spur (=Summe der Diagonalelemente)
    Mat10_trace = m.aa + m.bb + m.cc + m.dd + m.ee + m.ff + m.gg + m.hh + m.ii + m.jj
End Function

Public Function Mat2_Eigenvalues(m As Matrix2) As Vector2
    Dim trcA As Double: trcA = Mat2_trace(m)
    Dim detA As Double: detA = Mat2_det(m)
    With Mat2_Eigenvalues
        .a = (trcA + VBA.Math.Sqr(trcA ^ 2 - 4 * detA)) / 2
        .b = (trcA - VBA.Math.Sqr(trcA ^ 2 - 4 * detA)) / 2
    End With
End Function

Public Function Mat3_Eigenvalues(m As Matrix3) As Vector3
    '
    Dim eig As Matrix3
    Dim p1 As Double: p1 = m.ab ^ 2 + m.ac ^ 2 + m.bc ^ 2
    If p1 = 0 Then
        'm is diagonal all values above the diagonale are 0
        With Mat3_Eigenvalues
            .a = m.aa
            .b = m.bb
            .c = m.cc
        End With
    Else
        Dim q  As Double:  q = Mat3_trace(m) / 3
        Dim p2 As Double: p2 = (m.aa - q) ^ 2 + (m.bb - q) ^ 2 + 2 * p1
        Dim p  As Double:  p = VBA.Sqr(p2 / 6)
        Dim b  As Matrix3: b = Mat3_smul(Mat3_sub(m, Mat3_smul(Mat3_E, q)), 1 / p)
        Dim r  As Double:  r = Mat3_det(b)
        Dim phi As Double
        If (r <= -1) Then
            phi = 3.14159265358979 / 3
        ElseIf (r >= 1) Then
            phi = 0
        Else
            phi = ACos(r) / 3
        End If
        ' % the eigenvalues satisfy eig3 <= eig2 <= eig1
        With Mat3_Eigenvalues
            .a = q + 2 * p * Cos(phi)
            .c = q + 2 * p * Cos(phi + (2 * 3.14159265358979 / 3))
            .b = 3 * q - .a - .c
        End With
    End If
End Function

Public Function ASin(ByVal y As Double) As Double
    Select Case y
    Case 1
        ASin = 0.5 * 3.14159265358979
    Case -1
        ASin = -0.5 * 3.14159265358979
    Case Else
        ASin = VBA.Math.Atn(y / Sqr(1 - y * y))
    End Select
End Function
Public Function ACos(ByVal X As Double) As Double
    ACos = 0.5 * 3.14159265358979 - ASin(X)
End Function

'Matrizen-Addition
Public Function Mat2_add(m1 As Matrix2, M2 As Matrix2) As Matrix2
    'Addiert 2 2x2 Matrizen
    With Mat2_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb
    End With
End Function
Public Function Mat3_add(m1 As Matrix3, M2 As Matrix3) As Matrix3
    'Addiert 2 3x3 Matrizen
    With Mat3_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc
    End With
End Function
Public Function Mat4_add(m1 As Matrix4, M2 As Matrix4) As Matrix4
    'Addiert 2 4x4 Matrizen
    With Mat4_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd
                     .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd
    End With
End Function
Public Function Mat5_add(m1 As Matrix5, M2 As Matrix5) As Matrix5
    'Addiert 2 5x5 Matrizen
    With Mat5_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad:    .ae = m1.ae + M2.ae
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd:    .be = m1.be + M2.be
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd:    .ce = m1.ce + M2.ce
                     .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd:    .de = m1.de + M2.de
                     .ea = m1.ea + M2.ea:    .eb = m1.eb + M2.eb:    .ec = m1.ec + M2.ec:    .ed = m1.ed + M2.ed:    .ee = m1.ee + M2.ee
    End With
End Function
Public Function Mat6_add(m1 As Matrix6, M2 As Matrix6) As Matrix6
    'Addiert 2 6x6 Matrizen
    With Mat6_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad:    .ae = m1.ae + M2.ae:    .af = m1.af + M2.af
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd:    .be = m1.be + M2.be:    .bf = m1.bf + M2.bf
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd:    .ce = m1.ce + M2.ce:    .cf = m1.cf + M2.cf
                     .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd:    .de = m1.de + M2.de:    .df = m1.df + M2.df
                     .ea = m1.ea + M2.ea:    .eb = m1.eb + M2.eb:    .ec = m1.ec + M2.ec:    .ed = m1.ed + M2.ed:    .ee = m1.ee + M2.ee:    .ef = m1.ef + M2.ef
                     .fa = m1.fa + M2.fa:    .fb = m1.fb + M2.fb:    .fc = m1.fc + M2.fc:    .fd = m1.fd + M2.fd:    .fe = m1.fe + M2.fe:    .ff = m1.ff + M2.ff
    End With
End Function
Public Function Mat7_add(m1 As Matrix7, M2 As Matrix7) As Matrix7
    'Addiert 2 7x7 Matrizen
    With Mat7_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad:    .ae = m1.ae + M2.ae:    .af = m1.af + M2.af:    .ag = m1.ag + M2.ag
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd:    .be = m1.be + M2.be:    .bf = m1.bf + M2.bf:    .bg = m1.bg + M2.bg
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd:    .ce = m1.ce + M2.ce:    .cf = m1.cf + M2.cf:    .cg = m1.cg + M2.cg
                     .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd:    .de = m1.de + M2.de:    .df = m1.df + M2.df:    .dg = m1.dg + M2.dg
                     .ea = m1.ea + M2.ea:    .eb = m1.eb + M2.eb:    .ec = m1.ec + M2.ec:    .ed = m1.ed + M2.ed:    .ee = m1.ee + M2.ee:    .ef = m1.ef + M2.ef:    .eg = m1.eg + M2.eg
                     .fa = m1.fa + M2.fa:    .fb = m1.fb + M2.fb:    .fc = m1.fc + M2.fc:    .fd = m1.fd + M2.fd:    .fe = m1.fe + M2.fe:    .ff = m1.ff + M2.ff:    .fg = m1.fg + M2.fg
                     .ga = m1.ga + M2.ga:    .gb = m1.gb + M2.gb:    .gc = m1.gc + M2.gc:    .gd = m1.gd + M2.gd:    .ge = m1.ge + M2.ge:    .gf = m1.gf + M2.gf:    .gg = m1.gg + M2.gg
    End With
End Function
Public Function Mat8_add(m1 As Matrix8, M2 As Matrix8) As Matrix8
    'Addiert 2 8x8 Matrizen
    With Mat8_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad:    .ae = m1.ae + M2.ae:    .af = m1.af + M2.af:    .ag = m1.ag + M2.ag:    .ah = m1.ah + M2.ah
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd:    .be = m1.be + M2.be:    .bf = m1.bf + M2.bf:    .bg = m1.bg + M2.bg:    .bh = m1.bh + M2.bh
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd:    .ce = m1.ce + M2.ce:    .cf = m1.cf + M2.cf:    .cg = m1.cg + M2.cg:    .ch = m1.ch + M2.ch
                     .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd:    .de = m1.de + M2.de:    .df = m1.df + M2.df:    .dg = m1.dg + M2.dg:    .dh = m1.dh + M2.dh
                     .ea = m1.ea + M2.ea:    .eb = m1.eb + M2.eb:    .ec = m1.ec + M2.ec:    .ed = m1.ed + M2.ed:    .ee = m1.ee + M2.ee:    .ef = m1.ef + M2.ef:    .eg = m1.eg + M2.eg:    .eh = m1.eh + M2.eh
                     .fa = m1.fa + M2.fa:    .fb = m1.fb + M2.fb:    .fc = m1.fc + M2.fc:    .fd = m1.fd + M2.fd:    .fe = m1.fe + M2.fe:    .ff = m1.ff + M2.ff:    .fg = m1.fg + M2.fg:    .fh = m1.fh + M2.fh
                     .ga = m1.ga + M2.ga:    .gb = m1.gb + M2.gb:    .gc = m1.gc + M2.gc:    .gd = m1.gd + M2.gd:    .ge = m1.ge + M2.ge:    .gf = m1.gf + M2.gf:    .gg = m1.gg + M2.gg:    .gh = m1.gh + M2.gh
                     .ha = m1.ha + M2.ha:    .hb = m1.hb + M2.hb:    .HC = m1.HC + M2.HC:    .hd = m1.hd + M2.hd:    .he = m1.he + M2.he:    .hf = m1.hf + M2.hf:    .hg = m1.hg + M2.hg:    .hh = m1.hh + M2.hh
    End With
End Function
Public Function Mat9_add(m1 As Matrix9, M2 As Matrix9) As Matrix9
    'Addiert 2 9x9 Matrizen
    With Mat9_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad:    .ae = m1.ae + M2.ae:    .af = m1.af + M2.af:    .ag = m1.ag + M2.ag:    .ah = m1.ah + M2.ah:    .ai = m1.ai + M2.ai
                     .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd:    .be = m1.be + M2.be:    .bf = m1.bf + M2.bf:    .bg = m1.bg + M2.bg:    .bh = m1.bh + M2.bh:    .bi = m1.bi + M2.bi
                     .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd:    .ce = m1.ce + M2.ce:    .cf = m1.cf + M2.cf:    .cg = m1.cg + M2.cg:    .ch = m1.ch + M2.ch:    .ci = m1.ci + M2.ci
                     .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd:    .de = m1.de + M2.de:    .df = m1.df + M2.df:    .dg = m1.dg + M2.dg:    .dh = m1.dh + M2.dh:    .di = m1.di + M2.di
                     .ea = m1.ea + M2.ea:    .eb = m1.eb + M2.eb:    .ec = m1.ec + M2.ec:    .ed = m1.ed + M2.ed:    .ee = m1.ee + M2.ee:    .ef = m1.ef + M2.ef:    .eg = m1.eg + M2.eg:    .eh = m1.eh + M2.eh:    .ei = m1.ei + M2.ei
                     .fa = m1.fa + M2.fa:    .fb = m1.fb + M2.fb:    .fc = m1.fc + M2.fc:    .fd = m1.fd + M2.fd:    .fe = m1.fe + M2.fe:    .ff = m1.ff + M2.ff:    .fg = m1.fg + M2.fg:    .fh = m1.fh + M2.fh:    .fi = m1.fi + M2.fi
                     .ga = m1.ga + M2.ga:    .gb = m1.gb + M2.gb:    .gc = m1.gc + M2.gc:    .gd = m1.gd + M2.gd:    .ge = m1.ge + M2.ge:    .gf = m1.gf + M2.gf:    .gg = m1.gg + M2.gg:    .gh = m1.gh + M2.gh:    .gi = m1.gi + M2.gi
                     .ha = m1.ha + M2.ha:    .hb = m1.hb + M2.hb:    .HC = m1.HC + M2.HC:    .hd = m1.hd + M2.hd:    .he = m1.he + M2.he:    .hf = m1.hf + M2.hf:    .hg = m1.hg + M2.hg:    .hh = m1.hh + M2.hh:    .hi = m1.hi + M2.hi
                     .ia = m1.ia + M2.ia:    .ib = m1.ib + M2.ib:    .ic = m1.ic + M2.ic:    .id = m1.id + M2.id:    .ie = m1.ie + M2.ie:    .if = m1.if + M2.if:    .ig = m1.ig + M2.ig:    .ih = m1.ih + M2.ih:    .ii = m1.ii + M2.ii
    End With
End Function
Public Function Mat10_add(m1 As Matrix10, M2 As Matrix10) As Matrix10
    'Addiert 2 10x10 Matrizen
    With Mat10_add:   .aa = m1.aa + M2.aa:    .ab = m1.ab + M2.ab:    .ac = m1.ac + M2.ac:    .ad = m1.ad + M2.ad:    .ae = m1.ae + M2.ae:    .af = m1.af + M2.af:    .ag = m1.ag + M2.ag:    .ah = m1.ah + M2.ah:    .ai = m1.ai + M2.ai:    .aj = m1.aj + M2.aj
                      .ba = m1.ba + M2.ba:    .bb = m1.bb + M2.bb:    .bc = m1.bc + M2.bc:    .bd = m1.bd + M2.bd:    .be = m1.be + M2.be:    .bf = m1.bf + M2.bf:    .bg = m1.bg + M2.bg:    .bh = m1.bh + M2.bh:    .bi = m1.bi + M2.bi:    .bj = m1.bj + M2.bj
                      .ca = m1.ca + M2.ca:    .cb = m1.cb + M2.cb:    .cc = m1.cc + M2.cc:    .cd = m1.cd + M2.cd:    .ce = m1.ce + M2.ce:    .cf = m1.cf + M2.cf:    .cg = m1.cg + M2.cg:    .ch = m1.ch + M2.ch:    .ci = m1.ci + M2.ci:    .cj = m1.cj + M2.cj
                      .da = m1.da + M2.da:    .db = m1.db + M2.db:    .dc = m1.dc + M2.dc:    .dd = m1.dd + M2.dd:    .de = m1.de + M2.de:    .df = m1.df + M2.df:    .dg = m1.dg + M2.dg:    .dh = m1.dh + M2.dh:    .di = m1.di + M2.di:    .dj = m1.dj + M2.dj
                      .ea = m1.ea + M2.ea:    .eb = m1.eb + M2.eb:    .ec = m1.ec + M2.ec:    .ed = m1.ed + M2.ed:    .ee = m1.ee + M2.ee:    .ef = m1.ef + M2.ef:    .eg = m1.eg + M2.eg:    .eh = m1.eh + M2.eh:    .ei = m1.ei + M2.ei:    .ej = m1.ej + M2.ej
                      .fa = m1.fa + M2.fa:    .fb = m1.fb + M2.fb:    .fc = m1.fc + M2.fc:    .fd = m1.fd + M2.fd:    .fe = m1.fe + M2.fe:    .ff = m1.ff + M2.ff:    .fg = m1.fg + M2.fg:    .fh = m1.fh + M2.fh:    .fi = m1.fi + M2.fi:    .fj = m1.fj + M2.fj
                      .ga = m1.ga + M2.ga:    .gb = m1.gb + M2.gb:    .gc = m1.gc + M2.gc:    .gd = m1.gd + M2.gd:    .ge = m1.ge + M2.ge:    .gf = m1.gf + M2.gf:    .gg = m1.gg + M2.gg:    .gh = m1.gh + M2.gh:    .gi = m1.gi + M2.gi:    .gj = m1.gj + M2.gj
                      .ha = m1.ha + M2.ha:    .hb = m1.hb + M2.hb:    .HC = m1.HC + M2.HC:    .hd = m1.hd + M2.hd:    .he = m1.he + M2.he:    .hf = m1.hf + M2.hf:    .hg = m1.hg + M2.hg:    .hh = m1.hh + M2.hh:    .hi = m1.hi + M2.hi:    .hj = m1.hj + M2.hj
                      .ia = m1.ia + M2.ia:    .ib = m1.ib + M2.ib:    .ic = m1.ic + M2.ic:    .id = m1.id + M2.id:    .ie = m1.ie + M2.ie:    .if = m1.if + M2.if:    .ig = m1.ig + M2.ig:    .ih = m1.ih + M2.ih:    .ii = m1.ii + M2.ii:    .ij = m1.ij + M2.ij
                      .ja = m1.ja + M2.ja:    .jb = m1.jb + M2.jb:    .jc = m1.jc + M2.jc:    .jd = m1.jd + M2.jd:    .je = m1.je + M2.je:    .jf = m1.jf + M2.jf:    .jg = m1.jg + M2.jg:    .jh = m1.jh + M2.jh:    .ji = m1.ji + M2.ji:    .jj = m1.jj + M2.jj
    End With
End Function

'Matrizen-Subraktion
Public Function Mat2_sub(m1 As Matrix2, M2 As Matrix2) As Matrix2
    'Subtrahiert eine 2x2 Matrix m2 von Matrix m1
    With Mat2_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb
    End With
End Function
Public Function Mat3_sub(m1 As Matrix3, M2 As Matrix3) As Matrix3
    'Subtrahiert eine 3x3 Matrix m2 von Matrix m1
    With Mat3_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc
    End With
End Function
Public Function Mat4_sub(m1 As Matrix4, M2 As Matrix4) As Matrix4
    'Subtrahiert eine 4x4 Matrix m2 von Matrix m1
    With Mat4_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd
                     .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd
    End With
End Function
Public Function Mat5_sub(m1 As Matrix5, M2 As Matrix5) As Matrix5
    'Subtrahiert eine 5x5 Matrix m2 von Matrix m1
    With Mat5_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad:    .ae = m1.ae - M2.ae
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd:    .be = m1.be - M2.be
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd:    .ce = m1.ce - M2.ce
                     .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd:    .de = m1.de - M2.de
                     .ea = m1.ea - M2.ea:    .eb = m1.eb - M2.eb:    .ec = m1.ec - M2.ec:    .ed = m1.ed - M2.ed:    .ee = m1.ee - M2.ee
    End With
End Function
Public Function Mat6_sub(m1 As Matrix6, M2 As Matrix6) As Matrix6
    'Subtrahiert eine 6x6 Matrix m2 von Matrix m1
    With Mat6_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad:    .ae = m1.ae - M2.ae:    .af = m1.af - M2.af
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd:    .be = m1.be - M2.be:    .bf = m1.bf - M2.bf
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd:    .ce = m1.ce - M2.ce:    .cf = m1.cf - M2.cf
                     .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd:    .de = m1.de - M2.de:    .df = m1.df - M2.df
                     .ea = m1.ea - M2.ea:    .eb = m1.eb - M2.eb:    .ec = m1.ec - M2.ec:    .ed = m1.ed - M2.ed:    .ee = m1.ee - M2.ee:    .ef = m1.ef - M2.ef
                     .fa = m1.fa - M2.fa:    .fb = m1.fb - M2.fb:    .fc = m1.fc - M2.fc:    .fd = m1.fd - M2.fd:    .fe = m1.fe - M2.fe:    .ff = m1.ff - M2.ff
    End With
End Function
Public Function Mat7_sub(m1 As Matrix7, M2 As Matrix7) As Matrix7
    'Subtrahiert eine 7x7 Matrix m2 von Matrix m1
    With Mat7_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad:    .ae = m1.ae - M2.ae:    .af = m1.af - M2.af:    .ag = m1.ag - M2.ag
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd:    .be = m1.be - M2.be:    .bf = m1.bf - M2.bf:    .bg = m1.bg - M2.bg
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd:    .ce = m1.ce - M2.ce:    .cf = m1.cf - M2.cf:    .cg = m1.cg - M2.cg
                     .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd:    .de = m1.de - M2.de:    .df = m1.df - M2.df:    .dg = m1.dg - M2.dg
                     .ea = m1.ea - M2.ea:    .eb = m1.eb - M2.eb:    .ec = m1.ec - M2.ec:    .ed = m1.ed - M2.ed:    .ee = m1.ee - M2.ee:    .ef = m1.ef - M2.ef:    .eg = m1.eg - M2.eg
                     .fa = m1.fa - M2.fa:    .fb = m1.fb - M2.fb:    .fc = m1.fc - M2.fc:    .fd = m1.fd - M2.fd:    .fe = m1.fe - M2.fe:    .ff = m1.ff - M2.ff:    .fg = m1.fg - M2.fg
                     .ga = m1.ga - M2.ga:    .gb = m1.gb - M2.gb:    .gc = m1.gc - M2.gc:    .gd = m1.gd - M2.gd:    .ge = m1.ge - M2.ge:    .gf = m1.gf - M2.gf:    .gg = m1.gg - M2.gg
    End With
End Function
Public Function Mat8_sub(m1 As Matrix8, M2 As Matrix8) As Matrix8
    'Subtrahiert eine 8x8 Matrix m2 von Matrix m1
    With Mat8_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad:    .ae = m1.ae - M2.ae:    .af = m1.af - M2.af:    .ag = m1.ag - M2.ag:    .ah = m1.ah - M2.ah
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd:    .be = m1.be - M2.be:    .bf = m1.bf - M2.bf:    .bg = m1.bg - M2.bg:    .bh = m1.bh - M2.bh
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd:    .ce = m1.ce - M2.ce:    .cf = m1.cf - M2.cf:    .cg = m1.cg - M2.cg:    .ch = m1.ch - M2.ch
                     .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd:    .de = m1.de - M2.de:    .df = m1.df - M2.df:    .dg = m1.dg - M2.dg:    .dh = m1.dh - M2.dh
                     .ea = m1.ea - M2.ea:    .eb = m1.eb - M2.eb:    .ec = m1.ec - M2.ec:    .ed = m1.ed - M2.ed:    .ee = m1.ee - M2.ee:    .ef = m1.ef - M2.ef:    .eg = m1.eg - M2.eg:    .eh = m1.eh - M2.eh
                     .fa = m1.fa - M2.fa:    .fb = m1.fb - M2.fb:    .fc = m1.fc - M2.fc:    .fd = m1.fd - M2.fd:    .fe = m1.fe - M2.fe:    .ff = m1.ff - M2.ff:    .fg = m1.fg - M2.fg:    .fh = m1.fh - M2.fh
                     .ga = m1.ga - M2.ga:    .gb = m1.gb - M2.gb:    .gc = m1.gc - M2.gc:    .gd = m1.gd - M2.gd:    .ge = m1.ge - M2.ge:    .gf = m1.gf - M2.gf:    .gg = m1.gg - M2.gg:    .gh = m1.gh - M2.gh
                     .ha = m1.ha - M2.ha:    .hb = m1.hb - M2.hb:    .HC = m1.HC - M2.HC:    .hd = m1.hd - M2.hd:    .he = m1.he - M2.he:    .hf = m1.hf - M2.hf:    .hg = m1.hg - M2.hg:    .hh = m1.hh - M2.hh
    End With
End Function
Public Function Mat9_sub(m1 As Matrix9, M2 As Matrix9) As Matrix9
    'Subtrahiert eine 9x9 Matrix m2 von Matrix m1
    With Mat9_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad:    .ae = m1.ae - M2.ae:    .af = m1.af - M2.af:    .ag = m1.ag - M2.ag:    .ah = m1.ah - M2.ah:    .ai = m1.ai - M2.ai
                     .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd:    .be = m1.be - M2.be:    .bf = m1.bf - M2.bf:    .bg = m1.bg - M2.bg:    .bh = m1.bh - M2.bh:    .bi = m1.bi - M2.bi
                     .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd:    .ce = m1.ce - M2.ce:    .cf = m1.cf - M2.cf:    .cg = m1.cg - M2.cg:    .ch = m1.ch - M2.ch:    .ci = m1.ci - M2.ci
                     .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd:    .de = m1.de - M2.de:    .df = m1.df - M2.df:    .dg = m1.dg - M2.dg:    .dh = m1.dh - M2.dh:    .di = m1.di - M2.di
                     .ea = m1.ea - M2.ea:    .eb = m1.eb - M2.eb:    .ec = m1.ec - M2.ec:    .ed = m1.ed - M2.ed:    .ee = m1.ee - M2.ee:    .ef = m1.ef - M2.ef:    .eg = m1.eg - M2.eg:    .eh = m1.eh - M2.eh:    .ei = m1.ei - M2.ei
                     .fa = m1.fa - M2.fa:    .fb = m1.fb - M2.fb:    .fc = m1.fc - M2.fc:    .fd = m1.fd - M2.fd:    .fe = m1.fe - M2.fe:    .ff = m1.ff - M2.ff:    .fg = m1.fg - M2.fg:    .fh = m1.fh - M2.fh:    .fi = m1.fi - M2.fi
                     .ga = m1.ga - M2.ga:    .gb = m1.gb - M2.gb:    .gc = m1.gc - M2.gc:    .gd = m1.gd - M2.gd:    .ge = m1.ge - M2.ge:    .gf = m1.gf - M2.gf:    .gg = m1.gg - M2.gg:    .gh = m1.gh - M2.gh:    .gi = m1.gi - M2.gi
                     .ha = m1.ha - M2.ha:    .hb = m1.hb - M2.hb:    .HC = m1.HC - M2.HC:    .hd = m1.hd - M2.hd:    .he = m1.he - M2.he:    .hf = m1.hf - M2.hf:    .hg = m1.hg - M2.hg:    .hh = m1.hh - M2.hh:    .hi = m1.hi - M2.hi
                     .ia = m1.ia - M2.ia:    .ib = m1.ib - M2.ib:    .ic = m1.ic - M2.ic:    .id = m1.id - M2.id:    .ie = m1.ie - M2.ie:    .if = m1.if - M2.if:    .ig = m1.ig - M2.ig:    .ih = m1.ih - M2.ih:    .ii = m1.ii - M2.ii
    End With
End Function
Public Function Mat10_sub(m1 As Matrix10, M2 As Matrix10) As Matrix10
    'Subtrahiert eine 10x10 Matrix m2 von Matrix m1
    With Mat10_sub:   .aa = m1.aa - M2.aa:    .ab = m1.ab - M2.ab:    .ac = m1.ac - M2.ac:    .ad = m1.ad - M2.ad:    .ae = m1.ae - M2.ae:    .af = m1.af - M2.af:    .ag = m1.ag - M2.ag:    .ah = m1.ah - M2.ah:    .ai = m1.ai - M2.ai:    .aj = m1.aj - M2.aj
                      .ba = m1.ba - M2.ba:    .bb = m1.bb - M2.bb:    .bc = m1.bc - M2.bc:    .bd = m1.bd - M2.bd:    .be = m1.be - M2.be:    .bf = m1.bf - M2.bf:    .bg = m1.bg - M2.bg:    .bh = m1.bh - M2.bh:    .bi = m1.bi - M2.bi:    .bj = m1.bj - M2.bj
                      .ca = m1.ca - M2.ca:    .cb = m1.cb - M2.cb:    .cc = m1.cc - M2.cc:    .cd = m1.cd - M2.cd:    .ce = m1.ce - M2.ce:    .cf = m1.cf - M2.cf:    .cg = m1.cg - M2.cg:    .ch = m1.ch - M2.ch:    .ci = m1.ci - M2.ci:    .cj = m1.cj - M2.cj
                      .da = m1.da - M2.da:    .db = m1.db - M2.db:    .dc = m1.dc - M2.dc:    .dd = m1.dd - M2.dd:    .de = m1.de - M2.de:    .df = m1.df - M2.df:    .dg = m1.dg - M2.dg:    .dh = m1.dh - M2.dh:    .di = m1.di - M2.di:    .dj = m1.dj - M2.dj
                      .ea = m1.ea - M2.ea:    .eb = m1.eb - M2.eb:    .ec = m1.ec - M2.ec:    .ed = m1.ed - M2.ed:    .ee = m1.ee - M2.ee:    .ef = m1.ef - M2.ef:    .eg = m1.eg - M2.eg:    .eh = m1.eh - M2.eh:    .ei = m1.ei - M2.ei:    .ej = m1.ej - M2.ej
                      .fa = m1.fa - M2.fa:    .fb = m1.fb - M2.fb:    .fc = m1.fc - M2.fc:    .fd = m1.fd - M2.fd:    .fe = m1.fe - M2.fe:    .ff = m1.ff - M2.ff:    .fg = m1.fg - M2.fg:    .fh = m1.fh - M2.fh:    .fi = m1.fi - M2.fi:    .fj = m1.fj - M2.fj
                      .ga = m1.ga - M2.ga:    .gb = m1.gb - M2.gb:    .gc = m1.gc - M2.gc:    .gd = m1.gd - M2.gd:    .ge = m1.ge - M2.ge:    .gf = m1.gf - M2.gf:    .gg = m1.gg - M2.gg:    .gh = m1.gh - M2.gh:    .gi = m1.gi - M2.gi:    .gj = m1.gj - M2.gj
                      .ha = m1.ha - M2.ha:    .hb = m1.hb - M2.hb:    .HC = m1.HC - M2.HC:    .hd = m1.hd - M2.hd:    .he = m1.he - M2.he:    .hf = m1.hf - M2.hf:    .hg = m1.hg - M2.hg:    .hh = m1.hh - M2.hh:    .hi = m1.hi - M2.hi:    .hj = m1.hj - M2.hj
                      .ia = m1.ia - M2.ia:    .ib = m1.ib - M2.ib:    .ic = m1.ic - M2.ic:    .id = m1.id - M2.id:    .ie = m1.ie - M2.ie:    .if = m1.if - M2.if:    .ig = m1.ig - M2.ig:    .ih = m1.ih - M2.ih:    .ii = m1.ii - M2.ii:    .ij = m1.ij - M2.ij
                      .ja = m1.ja - M2.ja:    .jb = m1.jb - M2.jb:    .jc = m1.jc - M2.jc:    .jd = m1.jd - M2.jd:    .je = m1.je - M2.je:    .jf = m1.jf - M2.jf:    .jg = m1.jg - M2.jg:    .jh = m1.jh - M2.jh:    .ji = m1.ji - M2.ji:    .jj = m1.jj - M2.jj
    End With
End Function


'Matrizen-Skalar-Multiplikation
Public Function Mat2_smul(m As Matrix2, s As Double) As Matrix2
    'Multipliziert eine 2x2 Matrix mit einem Skalar
    With Mat2_smul:  .aa = m.aa * s: .ab = m.ab * s
                     .ba = m.ba * s: .bb = m.bb * s
    End With
End Function
Public Function Mat3_smul(m As Matrix3, s As Double) As Matrix3
    'Multipliziert eine 3x3 Matrix mit einem Skalar
    With Mat3_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s
    End With
End Function
Public Function Mat4_smul(m As Matrix4, s As Double) As Matrix4
    'Multipliziert eine 4x4 Matrix mit einem Skalar
    With Mat4_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s
                     .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s
    End With
End Function
Public Function Mat5_smul(m As Matrix5, s As Double) As Matrix5
    'Multipliziert eine 5x5 Matrix mit einem Skalar
    With Mat5_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s
                     .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s
                     .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s
    End With
End Function
Public Function Mat6_smul(m As Matrix6, s As Double) As Matrix6
    'Multipliziert eine 6x6 Matrix mit einem Skalar
    With Mat6_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s
                     .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s
                     .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s
                     .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s
    End With
End Function
Public Function Mat7_smul(m As Matrix7, s As Double) As Matrix7
    'Multipliziert eine 7x7 Matrix mit einem Skalar
    With Mat7_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s
                     .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s
                     .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s
                     .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s
                     .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s
    End With
End Function
Public Function Mat8_smul(m As Matrix8, s As Double) As Matrix8
    'Multipliziert eine 8x8 Matrix mit einem Skalar
    With Mat8_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s: .ah = m.ah * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s: .bh = m.bh * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s: .ch = m.ch * s
                     .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s: .dh = m.dh * s
                     .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s: .eh = m.eh * s
                     .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s: .fh = m.fh * s
                     .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s: .gh = m.gh * s
                     .ha = m.ha * s: .hb = m.hb * s: .HC = m.HC * s: .hd = m.hd * s: .he = m.he * s: .hf = m.hf * s: .hg = m.hg * s: .hh = m.hh * s
    End With
End Function
Public Function Mat9_smul(m As Matrix9, s As Double) As Matrix9
    'Multipliziert eine 9x9 Matrix mit einem Skalar
    With Mat9_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s: .ah = m.ah * s: .ai = m.ai * s
                     .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s: .bh = m.bh * s: .bi = m.bi * s
                     .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s: .ch = m.ch * s: .ci = m.ci * s
                     .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s: .dh = m.dh * s: .di = m.di * s
                     .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s: .eh = m.eh * s: .ei = m.ei * s
                     .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s: .fh = m.fh * s: .fi = m.fi * s
                     .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s: .gh = m.gh * s: .gi = m.gi * s
                     .ha = m.ha * s: .hb = m.hb * s: .HC = m.HC * s: .hd = m.hd * s: .he = m.he * s: .hf = m.hf * s: .hg = m.hg * s: .hh = m.hh * s: .hi = m.hi * s
                     .ia = m.ia * s: .ib = m.ib * s: .ic = m.ic * s: .id = m.id * s: .ie = m.ie * s: .if = m.if * s: .ig = m.ig * s: .ih = m.ih * s: .ii = m.ii * s
    End With
End Function
Public Function Mat10_smul(m As Matrix10, s As Double) As Matrix10
    'Multipliziert eine 10x10 Matrix mit einem Skalar
    With Mat10_smul:  .aa = m.aa * s: .ab = m.ab * s: .ac = m.ac * s: .ad = m.ad * s: .ae = m.ae * s: .af = m.af * s: .ag = m.ag * s: .ah = m.ah * s: .ai = m.ai * s: .aj = m.aj * s
                      .ba = m.ba * s: .bb = m.bb * s: .bc = m.bc * s: .bd = m.bd * s: .be = m.be * s: .bf = m.bf * s: .bg = m.bg * s: .bh = m.bh * s: .bi = m.bi * s: .bj = m.bj * s
                      .ca = m.ca * s: .cb = m.cb * s: .cc = m.cc * s: .cd = m.cd * s: .ce = m.ce * s: .cf = m.cf * s: .cg = m.cg * s: .ch = m.ch * s: .ci = m.ci * s: .cj = m.cj * s
                      .da = m.da * s: .db = m.db * s: .dc = m.dc * s: .dd = m.dd * s: .de = m.de * s: .df = m.df * s: .dg = m.dg * s: .dh = m.dh * s: .di = m.di * s: .dj = m.dj * s
                      .ea = m.ea * s: .eb = m.eb * s: .ec = m.ec * s: .ed = m.ed * s: .ee = m.ee * s: .ef = m.ef * s: .eg = m.eg * s: .eh = m.eh * s: .ei = m.ei * s: .ej = m.ej * s
                      .fa = m.fa * s: .fb = m.fb * s: .fc = m.fc * s: .fd = m.fd * s: .fe = m.fe * s: .ff = m.ff * s: .fg = m.fg * s: .fh = m.fh * s: .fi = m.fi * s: .fj = m.fj * s
                      .ga = m.ga * s: .gb = m.gb * s: .gc = m.gc * s: .gd = m.gd * s: .ge = m.ge * s: .gf = m.gf * s: .gg = m.gg * s: .gh = m.gh * s: .gi = m.gi * s: .gj = m.gj * s
                      .ha = m.ha * s: .hb = m.hb * s: .HC = m.HC * s: .hd = m.hd * s: .he = m.he * s: .hf = m.hf * s: .hg = m.hg * s: .hh = m.hh * s: .hi = m.hi * s: .hj = m.hj * s
                      .ia = m.ia * s: .ib = m.ib * s: .ic = m.ic * s: .id = m.id * s: .ie = m.ie * s: .if = m.if * s: .ig = m.ig * s: .ih = m.ih * s: .ii = m.ii * s: .ij = m.ij * s
                      .ja = m.ja * s: .jb = m.jb * s: .jc = m.jc * s: .jd = m.jd * s: .je = m.je * s: .jf = m.jf * s: .jg = m.jg * s: .jh = m.jh * s: .ji = m.ji * s: .jj = m.jj * s
    End With
End Function

'Matrizen-Multiplikation
Public Function Mat2_mul(m1 As Matrix2, M2 As Matrix2) As Matrix2
    'Multipliziert eine 2x2 Matrix m1 mit einer Matrix m2
    With Mat2_mul:   .aa = m1.aa * M2.aa + m1.ab * M2.ba:    .ab = m1.aa * M2.ab + m1.ab * M2.bb
                     .ba = m1.ba * M2.aa + m1.bb * M2.ba:    .bb = m1.ba * M2.ab + m1.bb * M2.bb
    End With
End Function

Public Function Mat3_mul(m1 As Matrix3, M2 As Matrix3) As Matrix3
    'Multipliziert eine 3x3 Matrix m1 mit einer Matrix m2
    With Mat3_mul:   .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc
                     .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc
                     .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc
    End With
End Function
Public Function Mat4_mul(m1 As Matrix4, M2 As Matrix4) As Matrix4
    'Multipliziert eine 4x4 Matrix m1 mit einer Matrix m2
    With Mat4_mul:   .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd
                     .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd
                     .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd
                     .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd
    End With
End Function
Public Function Mat5_mul(m1 As Matrix5, M2 As Matrix5) As Matrix5
    'Multipliziert eine 5x5 Matrix m1 mit einer Matrix m2
    With Mat5_mul:   .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da + m1.ae * M2.ea:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db + m1.ad * M2.eb:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc + m1.ae * M2.ec:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd + m1.ae * M2.ed:    .ae = m1.aa * M2.ae + m1.ab * M2.be + m1.ac * M2.ce + m1.ad * M2.de + m1.ae * M2.ee
                     .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da + m1.be * M2.ea:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db + m1.bd * M2.eb:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc + m1.be * M2.ec:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd + m1.be * M2.ed:    .be = m1.ba * M2.ae + m1.bb * M2.be + m1.bc * M2.ce + m1.bd * M2.de + m1.be * M2.ee
                     .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da + m1.ce * M2.ea:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db + m1.cd * M2.eb:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc + m1.ce * M2.ec:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd + m1.ce * M2.ed:    .ce = m1.ca * M2.ae + m1.cb * M2.be + m1.cc * M2.ce + m1.cd * M2.de + m1.ce * M2.ee
                     .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da + m1.de * M2.ea:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db + m1.dd * M2.eb:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc + m1.de * M2.ec:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd + m1.de * M2.ed:    .de = m1.da * M2.ae + m1.db * M2.be + m1.dc * M2.ce + m1.dd * M2.de + m1.de * M2.ee
                     .ea = m1.ea * M2.aa + m1.eb * M2.ba + m1.ec * M2.ca + m1.ed * M2.da + m1.ee * M2.ea:    .eb = m1.ea * M2.ab + m1.eb * M2.bb + m1.ec * M2.cb + m1.ed * M2.db + m1.ed * M2.eb:    .ec = m1.ea * M2.ac + m1.eb * M2.bc + m1.ec * M2.cc + m1.ed * M2.dc + m1.ee * M2.ec:    .ed = m1.ea * M2.ad + m1.eb * M2.bd + m1.ec * M2.cd + m1.ed * M2.dd + m1.ee * M2.ed:    .ee = m1.ea * M2.ae + m1.eb * M2.be + m1.ec * M2.ce + m1.ed * M2.de + m1.ee * M2.ee
    End With
End Function
Public Function Mat6_mul(m1 As Matrix6, M2 As Matrix6) As Matrix6
    'Multipliziert eine 6x6 Matrix m1 mit einer Matrix m2
    With Mat6_mul:   .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da + m1.ae * M2.ea + m1.af * M2.fa:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db + m1.ad * M2.eb + m1.af * M2.fb:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc + m1.ae * M2.ec + m1.af * M2.fc:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd + m1.ae * M2.ed + m1.af * M2.fd:    .ae = m1.aa * M2.ae + m1.ab * M2.be + m1.ac * M2.ce + m1.ad * M2.de + m1.ae * M2.ee + m1.af * M2.fe:    .af = m1.aa * M2.af + m1.ab * M2.bf + m1.ac * M2.cf + m1.ad * M2.df + m1.ae * M2.ef + m1.af * M2.ff
                     .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da + m1.be * M2.ea + m1.bf * M2.fa:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db + m1.bd * M2.eb + m1.bf * M2.fb:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc + m1.be * M2.ec + m1.bf * M2.fc:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd + m1.be * M2.ed + m1.bf * M2.fd:    .be = m1.ba * M2.ae + m1.bb * M2.be + m1.bc * M2.ce + m1.bd * M2.de + m1.be * M2.ee + m1.bf * M2.fe:    .bf = m1.ba * M2.af + m1.bb * M2.bf + m1.bc * M2.cf + m1.bd * M2.df + m1.be * M2.ef + m1.bf * M2.ff
                     .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da + m1.ce * M2.ea + m1.cf * M2.fa:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db + m1.cd * M2.eb + m1.cf * M2.fb:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc + m1.ce * M2.ec + m1.cf * M2.fc:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd + m1.ce * M2.ed + m1.cf * M2.fd:    .ce = m1.ca * M2.ae + m1.cb * M2.be + m1.cc * M2.ce + m1.cd * M2.de + m1.ce * M2.ee + m1.cf * M2.fe:    .cf = m1.ca * M2.af + m1.cb * M2.bf + m1.cc * M2.cf + m1.cd * M2.df + m1.ce * M2.ef + m1.cf * M2.ff
                     .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da + m1.de * M2.ea + m1.df * M2.fa:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db + m1.dd * M2.eb + m1.df * M2.fb:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc + m1.de * M2.ec + m1.df * M2.fc:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd + m1.de * M2.ed + m1.df * M2.fd:    .de = m1.da * M2.ae + m1.db * M2.be + m1.dc * M2.ce + m1.dd * M2.de + m1.de * M2.ee + m1.df * M2.fe:    .df = m1.da * M2.af + m1.db * M2.bf + m1.dc * M2.cf + m1.dd * M2.df + m1.de * M2.ef + m1.df * M2.ff
                     .ea = m1.ea * M2.aa + m1.eb * M2.ba + m1.ec * M2.ca + m1.ed * M2.da + m1.ee * M2.ea + m1.ef * M2.fa:    .eb = m1.ea * M2.ab + m1.eb * M2.bb + m1.ec * M2.cb + m1.ed * M2.db + m1.ed * M2.eb + m1.ef * M2.fb:    .ec = m1.ea * M2.ac + m1.eb * M2.bc + m1.ec * M2.cc + m1.ed * M2.dc + m1.ee * M2.ec + m1.ef * M2.fc:    .ed = m1.ea * M2.ad + m1.eb * M2.bd + m1.ec * M2.cd + m1.ed * M2.dd + m1.ee * M2.ed + m1.ef * M2.fd:    .ee = m1.ea * M2.ae + m1.eb * M2.be + m1.ec * M2.ce + m1.ed * M2.de + m1.ee * M2.ee + m1.ef * M2.fe:    .ef = m1.ea * M2.af + m1.eb * M2.bf + m1.ec * M2.cf + m1.ed * M2.df + m1.ee * M2.ef + m1.ef * M2.ff
                     .fa = m1.fa * M2.aa + m1.fb * M2.ba + m1.fc * M2.ca + m1.fd * M2.da + m1.fe * M2.ea + m1.ff * M2.fa:    .fb = m1.fa * M2.ab + m1.fb * M2.bb + m1.fc * M2.cb + m1.fd * M2.db + m1.fd * M2.eb + m1.ff * M2.fb:    .fc = m1.fa * M2.ac + m1.fb * M2.bc + m1.fc * M2.cc + m1.fd * M2.dc + m1.fe * M2.ec + m1.ff * M2.fc:    .fd = m1.fa * M2.ad + m1.fb * M2.bd + m1.fc * M2.cd + m1.fd * M2.dd + m1.fe * M2.ed + m1.ff * M2.fd:    .fe = m1.fa * M2.ae + m1.fb * M2.be + m1.fc * M2.ce + m1.fd * M2.de + m1.fe * M2.ee + m1.ff * M2.fe:    .ff = m1.fa * M2.af + m1.fb * M2.bf + m1.fc * M2.cf + m1.fd * M2.df + m1.fe * M2.ef + m1.ff * M2.ff
    End With
End Function
Public Function Mat7_mul(m1 As Matrix7, M2 As Matrix7) As Matrix7
    'Multipliziert eine 7x7 Matrix m1 mit einer Matrix m2
    With Mat7_mul:   .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da + m1.ae * M2.ea + m1.af * M2.fa + m1.ag * M2.ga:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db + m1.ad * M2.eb + m1.af * M2.fb + m1.ag * M2.gb:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc + m1.ae * M2.ec + m1.af * M2.fc + m1.ag * M2.gc:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd + m1.ae * M2.ed + m1.af * M2.fd + m1.ag * M2.gd:    .ae = m1.aa * M2.ae + m1.ab * M2.be + m1.ac * M2.ce + m1.ad * M2.de + m1.ae * M2.ee + m1.af * M2.fe + m1.ag * M2.ge:    .af = m1.aa * M2.af + m1.ab * M2.bf + m1.ac * M2.cf + m1.ad * M2.df + m1.ae * M2.ef + m1.af * M2.ff + m1.ag * M2.gf:    .ag = m1.aa * M2.ag + m1.ab * M2.bg + m1.ac * M2.cg + m1.ad * M2.dg + m1.ae * M2.eg + m1.af * M2.fg + m1.ag * M2.gg
                     .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da + m1.be * M2.ea + m1.bf * M2.fa + m1.bg * M2.ga:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db + m1.bd * M2.eb + m1.bf * M2.fb + m1.bg * M2.gb:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc + m1.be * M2.ec + m1.bf * M2.fc + m1.bg * M2.gc:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd + m1.be * M2.ed + m1.bf * M2.fd + m1.bg * M2.gd:    .be = m1.ba * M2.ae + m1.bb * M2.be + m1.bc * M2.ce + m1.bd * M2.de + m1.be * M2.ee + m1.bf * M2.fe + m1.bg * M2.ge:    .bf = m1.ba * M2.af + m1.bb * M2.bf + m1.bc * M2.cf + m1.bd * M2.df + m1.be * M2.ef + m1.bf * M2.ff + m1.bg * M2.gf:    .bg = m1.ba * M2.ag + m1.bb * M2.bg + m1.bc * M2.cg + m1.bd * M2.dg + m1.be * M2.eg + m1.bf * M2.fg + m1.bg * M2.gg
                     .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da + m1.ce * M2.ea + m1.cf * M2.fa + m1.cg * M2.ga:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db + m1.cd * M2.eb + m1.cf * M2.fb + m1.cg * M2.gb:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc + m1.ce * M2.ec + m1.cf * M2.fc + m1.cg * M2.gc:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd + m1.ce * M2.ed + m1.cf * M2.fd + m1.cg * M2.gd:    .ce = m1.ca * M2.ae + m1.cb * M2.be + m1.cc * M2.ce + m1.cd * M2.de + m1.ce * M2.ee + m1.cf * M2.fe + m1.cg * M2.ge:    .cf = m1.ca * M2.af + m1.cb * M2.bf + m1.cc * M2.cf + m1.cd * M2.df + m1.ce * M2.ef + m1.cf * M2.ff + m1.cg * M2.gf:    .cg = m1.ca * M2.ag + m1.cb * M2.bg + m1.cc * M2.cg + m1.cd * M2.dg + m1.ce * M2.eg + m1.cf * M2.fg + m1.cg * M2.gg
                     .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da + m1.de * M2.ea + m1.df * M2.fa + m1.dg * M2.ga:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db + m1.dd * M2.eb + m1.df * M2.fb + m1.dg * M2.gb:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc + m1.de * M2.ec + m1.df * M2.fc + m1.dg * M2.gc:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd + m1.de * M2.ed + m1.df * M2.fd + m1.dg * M2.gd:    .de = m1.da * M2.ae + m1.db * M2.be + m1.dc * M2.ce + m1.dd * M2.de + m1.de * M2.ee + m1.df * M2.fe + m1.dg * M2.ge:    .df = m1.da * M2.af + m1.db * M2.bf + m1.dc * M2.cf + m1.dd * M2.df + m1.de * M2.ef + m1.df * M2.ff + m1.dg * M2.gf:    .dg = m1.da * M2.ag + m1.db * M2.bg + m1.dc * M2.cg + m1.dd * M2.dg + m1.de * M2.eg + m1.df * M2.fg + m1.dg * M2.gg
                     .ea = m1.ea * M2.aa + m1.eb * M2.ba + m1.ec * M2.ca + m1.ed * M2.da + m1.ee * M2.ea + m1.ef * M2.fa + m1.eg * M2.ga:    .eb = m1.ea * M2.ab + m1.eb * M2.bb + m1.ec * M2.cb + m1.ed * M2.db + m1.ed * M2.eb + m1.ef * M2.fb + m1.eg * M2.gb:    .ec = m1.ea * M2.ac + m1.eb * M2.bc + m1.ec * M2.cc + m1.ed * M2.dc + m1.ee * M2.ec + m1.ef * M2.fc + m1.eg * M2.gc:    .ed = m1.ea * M2.ad + m1.eb * M2.bd + m1.ec * M2.cd + m1.ed * M2.dd + m1.ee * M2.ed + m1.ef * M2.fd + m1.eg * M2.gd:    .ee = m1.ea * M2.ae + m1.eb * M2.be + m1.ec * M2.ce + m1.ed * M2.de + m1.ee * M2.ee + m1.ef * M2.fe + m1.eg * M2.ge:    .ef = m1.ea * M2.af + m1.eb * M2.bf + m1.ec * M2.cf + m1.ed * M2.df + m1.ee * M2.ef + m1.ef * M2.ff + m1.eg * M2.gf:    .eg = m1.ea * M2.ag + m1.eb * M2.bg + m1.ec * M2.cg + m1.ed * M2.dg + m1.ee * M2.eg + m1.ef * M2.fg + m1.eg * M2.gg
                     .fa = m1.fa * M2.aa + m1.fb * M2.ba + m1.fc * M2.ca + m1.fd * M2.da + m1.fe * M2.ea + m1.ff * M2.fa + m1.fg * M2.ga:    .fb = m1.fa * M2.ab + m1.fb * M2.bb + m1.fc * M2.cb + m1.fd * M2.db + m1.fd * M2.eb + m1.ff * M2.fb + m1.fg * M2.gb:    .fc = m1.fa * M2.ac + m1.fb * M2.bc + m1.fc * M2.cc + m1.fd * M2.dc + m1.fe * M2.ec + m1.ff * M2.fc + m1.fg * M2.gc:    .fd = m1.fa * M2.ad + m1.fb * M2.bd + m1.fc * M2.cd + m1.fd * M2.dd + m1.fe * M2.ed + m1.ff * M2.fd + m1.fg * M2.gd:    .fe = m1.fa * M2.ae + m1.fb * M2.be + m1.fc * M2.ce + m1.fd * M2.de + m1.fe * M2.ee + m1.ff * M2.fe + m1.fg * M2.ge:    .ff = m1.fa * M2.af + m1.fb * M2.bf + m1.fc * M2.cf + m1.fd * M2.df + m1.fe * M2.ef + m1.ff * M2.ff + m1.fg * M2.gf:    .fg = m1.fa * M2.ag + m1.fb * M2.bg + m1.fc * M2.cg + m1.fd * M2.dg + m1.fe * M2.eg + m1.ff * M2.fg + m1.fg * M2.gg
                     .ga = m1.ga * M2.aa + m1.gb * M2.ba + m1.gc * M2.ca + m1.gd * M2.da + m1.ge * M2.ea + m1.gf * M2.fa + m1.gg * M2.ga:    .gb = m1.ga * M2.ab + m1.gb * M2.bb + m1.gc * M2.cb + m1.gd * M2.db + m1.gd * M2.eb + m1.gf * M2.fb + m1.gg * M2.gb:    .gc = m1.ga * M2.ac + m1.gb * M2.bc + m1.gc * M2.cc + m1.gd * M2.dc + m1.ge * M2.ec + m1.gf * M2.fc + m1.gg * M2.gc:    .gd = m1.ga * M2.ad + m1.gb * M2.bd + m1.gc * M2.cd + m1.gd * M2.dd + m1.ge * M2.ed + m1.gf * M2.fd + m1.gg * M2.gd:    .ge = m1.ga * M2.ae + m1.gb * M2.be + m1.gc * M2.ce + m1.gd * M2.de + m1.ge * M2.ee + m1.gf * M2.fe + m1.gg * M2.ge:    .gf = m1.ga * M2.af + m1.gb * M2.bf + m1.gc * M2.cf + m1.gd * M2.df + m1.ge * M2.ef + m1.gf * M2.ff + m1.gg * M2.gf:    .gg = m1.ga * M2.ag + m1.gb * M2.bg + m1.gc * M2.cg + m1.gd * M2.dg + m1.ge * M2.eg + m1.gf * M2.fg + m1.gg * M2.gg
    End With
End Function
Public Function Mat8_mul(m1 As Matrix8, M2 As Matrix8) As Matrix8
    'Multipliziert eine 8x8 Matrix m1 mit einer Matrix m2
    With Mat8_mul
        .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da + m1.ae * M2.ea + m1.af * M2.fa + m1.ag * M2.ga + m1.ah * M2.ha:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db + m1.ad * M2.eb + m1.af * M2.fb + m1.ag * M2.gb + m1.ah * M2.hb:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc + m1.ae * M2.ec + m1.af * M2.fc + m1.ag * M2.gc + m1.ah * M2.HC:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd + m1.ae * M2.ed + m1.af * M2.fd + m1.ag * M2.gd + m1.ah * M2.hd
        .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da + m1.be * M2.ea + m1.bf * M2.fa + m1.bg * M2.ga + m1.bh * M2.ha:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db + m1.bd * M2.eb + m1.bf * M2.fb + m1.bg * M2.gb + m1.bh * M2.hb:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc + m1.be * M2.ec + m1.bf * M2.fc + m1.bg * M2.gc + m1.bh * M2.HC:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd + m1.be * M2.ed + m1.bf * M2.fd + m1.bg * M2.gd + m1.bh * M2.hd
        .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da + m1.ce * M2.ea + m1.cf * M2.fa + m1.cg * M2.ga + m1.ch * M2.ha:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db + m1.cd * M2.eb + m1.cf * M2.fb + m1.cg * M2.gb + m1.ch * M2.hb:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc + m1.ce * M2.ec + m1.cf * M2.fc + m1.cg * M2.gc + m1.ch * M2.HC:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd + m1.ce * M2.ed + m1.cf * M2.fd + m1.cg * M2.gd + m1.ch * M2.hd
        .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da + m1.de * M2.ea + m1.df * M2.fa + m1.dg * M2.ga + m1.dh * M2.ha:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db + m1.dd * M2.eb + m1.df * M2.fb + m1.dg * M2.gb + m1.dh * M2.hb:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc + m1.de * M2.ec + m1.df * M2.fc + m1.dg * M2.gc + m1.dh * M2.HC:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd + m1.de * M2.ed + m1.df * M2.fd + m1.dg * M2.gd + m1.dh * M2.hd
        .ea = m1.ea * M2.aa + m1.eb * M2.ba + m1.ec * M2.ca + m1.ed * M2.da + m1.ee * M2.ea + m1.ef * M2.fa + m1.eg * M2.ga + m1.eh * M2.ha:    .eb = m1.ea * M2.ab + m1.eb * M2.bb + m1.ec * M2.cb + m1.ed * M2.db + m1.ed * M2.eb + m1.ef * M2.fb + m1.eg * M2.gb + m1.eh * M2.hb:    .ec = m1.ea * M2.ac + m1.eb * M2.bc + m1.ec * M2.cc + m1.ed * M2.dc + m1.ee * M2.ec + m1.ef * M2.fc + m1.eg * M2.gc + m1.eh * M2.HC:    .ed = m1.ea * M2.ad + m1.eb * M2.bd + m1.ec * M2.cd + m1.ed * M2.dd + m1.ee * M2.ed + m1.ef * M2.fd + m1.eg * M2.gd + m1.eh * M2.hd
        .fa = m1.fa * M2.aa + m1.fb * M2.ba + m1.fc * M2.ca + m1.fd * M2.da + m1.fe * M2.ea + m1.ff * M2.fa + m1.fg * M2.ga + m1.fh * M2.ha:    .fb = m1.fa * M2.ab + m1.fb * M2.bb + m1.fc * M2.cb + m1.fd * M2.db + m1.fd * M2.eb + m1.ff * M2.fb + m1.fg * M2.gb + m1.fh * M2.hb:    .fc = m1.fa * M2.ac + m1.fb * M2.bc + m1.fc * M2.cc + m1.fd * M2.dc + m1.fe * M2.ec + m1.ff * M2.fc + m1.fg * M2.gc + m1.fh * M2.HC:    .fd = m1.fa * M2.ad + m1.fb * M2.bd + m1.fc * M2.cd + m1.fd * M2.dd + m1.fe * M2.ed + m1.ff * M2.fd + m1.fg * M2.gd + m1.fh * M2.hd
        .ga = m1.ga * M2.aa + m1.gb * M2.ba + m1.gc * M2.ca + m1.gd * M2.da + m1.ge * M2.ea + m1.gf * M2.fa + m1.gg * M2.ga + m1.gh * M2.ha:    .gb = m1.ga * M2.ab + m1.gb * M2.bb + m1.gc * M2.cb + m1.gd * M2.db + m1.gd * M2.eb + m1.gf * M2.fb + m1.gg * M2.gb + m1.gh * M2.hb:    .gc = m1.ga * M2.ac + m1.gb * M2.bc + m1.gc * M2.cc + m1.gd * M2.dc + m1.ge * M2.ec + m1.gf * M2.fc + m1.gg * M2.gc + m1.gh * M2.HC:    .gd = m1.ga * M2.ad + m1.gb * M2.bd + m1.gc * M2.cd + m1.gd * M2.dd + m1.ge * M2.ed + m1.gf * M2.fd + m1.gg * M2.gd + m1.gh * M2.hd
        .ha = m1.ha * M2.aa + m1.hb * M2.ba + m1.HC * M2.ca + m1.hd * M2.da + m1.he * M2.ea + m1.hf * M2.fa + m1.hg * M2.ga + m1.hh * M2.ha:    .hb = m1.ha * M2.ab + m1.hb * M2.bb + m1.HC * M2.cb + m1.hd * M2.db + m1.hd * M2.eb + m1.hf * M2.fb + m1.hg * M2.gb + m1.hh * M2.hb:    .HC = m1.ha * M2.ac + m1.hb * M2.bc + m1.HC * M2.cc + m1.hd * M2.dc + m1.he * M2.ec + m1.hf * M2.fc + m1.hg * M2.gc + m1.hh * M2.HC:    .hd = m1.ha * M2.ad + m1.hb * M2.bd + m1.HC * M2.cd + m1.hd * M2.dd + m1.he * M2.ed + m1.hf * M2.fd + m1.hg * M2.gd + m1.hh * M2.hd
        
        .ae = m1.aa * M2.ae + m1.ab * M2.be + m1.ac * M2.ce + m1.ad * M2.de + m1.ae * M2.ee + m1.af * M2.fe + m1.ag * M2.ge + m1.ah * M2.he:    .af = m1.aa * M2.af + m1.ab * M2.bf + m1.ac * M2.cf + m1.ad * M2.df + m1.ae * M2.ef + m1.af * M2.ff + m1.ag * M2.gf + m1.ah * M2.hf:    .ag = m1.aa * M2.ag + m1.ab * M2.bg + m1.ac * M2.cg + m1.ad * M2.dg + m1.ae * M2.eg + m1.af * M2.fg + m1.ag * M2.gg + m1.ah * M2.hg:    .ah = m1.aa * M2.ah + m1.ab * M2.bh + m1.ac * M2.ch + m1.ad * M2.dh + m1.ae * M2.eh + m1.af * M2.fh + m1.ag * M2.gh + m1.ah * M2.hh
        .be = m1.ba * M2.ae + m1.bb * M2.be + m1.bc * M2.ce + m1.bd * M2.de + m1.be * M2.ee + m1.bf * M2.fe + m1.bg * M2.ge + m1.bh * M2.he:    .bf = m1.ba * M2.af + m1.bb * M2.bf + m1.bc * M2.cf + m1.bd * M2.df + m1.be * M2.ef + m1.bf * M2.ff + m1.bg * M2.gf + m1.bh * M2.hf:    .bg = m1.ba * M2.ag + m1.bb * M2.bg + m1.bc * M2.cg + m1.bd * M2.dg + m1.be * M2.eg + m1.bf * M2.fg + m1.bg * M2.gg + m1.bh * M2.hg:    .bh = m1.ba * M2.ah + m1.bb * M2.bh + m1.bc * M2.ch + m1.bd * M2.dh + m1.be * M2.eh + m1.bf * M2.fh + m1.bg * M2.gh + m1.bh * M2.hh
        .ce = m1.ca * M2.ae + m1.cb * M2.be + m1.cc * M2.ce + m1.cd * M2.de + m1.ce * M2.ee + m1.cf * M2.fe + m1.cg * M2.ge + m1.ch * M2.he:    .cf = m1.ca * M2.af + m1.cb * M2.bf + m1.cc * M2.cf + m1.cd * M2.df + m1.ce * M2.ef + m1.cf * M2.ff + m1.cg * M2.gf + m1.ch * M2.hf:    .cg = m1.ca * M2.ag + m1.cb * M2.bg + m1.cc * M2.cg + m1.cd * M2.dg + m1.ce * M2.eg + m1.cf * M2.fg + m1.cg * M2.gg + m1.ch * M2.hg:    .ch = m1.ca * M2.ah + m1.cb * M2.bh + m1.cc * M2.ch + m1.cd * M2.dh + m1.ce * M2.eh + m1.cf * M2.fh + m1.cg * M2.gh + m1.ch * M2.hh
        .de = m1.da * M2.ae + m1.db * M2.be + m1.dc * M2.ce + m1.dd * M2.de + m1.de * M2.ee + m1.df * M2.fe + m1.dg * M2.ge + m1.dh * M2.he:    .df = m1.da * M2.af + m1.db * M2.bf + m1.dc * M2.cf + m1.dd * M2.df + m1.de * M2.ef + m1.df * M2.ff + m1.dg * M2.gf + m1.dh * M2.hf:    .dg = m1.da * M2.ag + m1.db * M2.bg + m1.dc * M2.cg + m1.dd * M2.dg + m1.de * M2.eg + m1.df * M2.fg + m1.dg * M2.gg + m1.dh * M2.hg:    .dh = m1.da * M2.ah + m1.db * M2.bh + m1.dc * M2.ch + m1.dd * M2.dh + m1.de * M2.eh + m1.df * M2.fh + m1.dg * M2.gh + m1.dh * M2.hh
        .ee = m1.ea * M2.ae + m1.eb * M2.be + m1.ec * M2.ce + m1.ed * M2.de + m1.ee * M2.ee + m1.ef * M2.fe + m1.eg * M2.ge + m1.eh * M2.he:    .ef = m1.ea * M2.af + m1.eb * M2.bf + m1.ec * M2.cf + m1.ed * M2.df + m1.ee * M2.ef + m1.ef * M2.ff + m1.eg * M2.gf + m1.eh * M2.hf:    .eg = m1.ea * M2.ag + m1.eb * M2.bg + m1.ec * M2.cg + m1.ed * M2.dg + m1.ee * M2.eg + m1.ef * M2.fg + m1.eg * M2.gg + m1.eh * M2.hg:    .eh = m1.ea * M2.ah + m1.eb * M2.bh + m1.ec * M2.ch + m1.ed * M2.dh + m1.ee * M2.eh + m1.ef * M2.fh + m1.eg * M2.gh + m1.eh * M2.hh
        .fe = m1.fa * M2.ae + m1.fb * M2.be + m1.fc * M2.ce + m1.fd * M2.de + m1.fe * M2.ee + m1.ff * M2.fe + m1.fg * M2.ge + m1.fh * M2.he:    .ff = m1.fa * M2.af + m1.fb * M2.bf + m1.fc * M2.cf + m1.fd * M2.df + m1.fe * M2.ef + m1.ff * M2.ff + m1.fg * M2.gf + m1.fh * M2.hf:    .fg = m1.fa * M2.ag + m1.fb * M2.bg + m1.fc * M2.cg + m1.fd * M2.dg + m1.fe * M2.eg + m1.ff * M2.fg + m1.fg * M2.gg + m1.fh * M2.hg:    .fh = m1.fa * M2.ah + m1.fb * M2.bh + m1.fc * M2.ch + m1.fd * M2.dh + m1.fe * M2.eh + m1.ff * M2.fh + m1.fg * M2.gh + m1.fh * M2.hh
        .ge = m1.ga * M2.ae + m1.gb * M2.be + m1.gc * M2.ce + m1.gd * M2.de + m1.ge * M2.ee + m1.gf * M2.fe + m1.gg * M2.ge + m1.gh * M2.he:    .gf = m1.ga * M2.af + m1.gb * M2.bf + m1.gc * M2.cf + m1.gd * M2.df + m1.ge * M2.ef + m1.gf * M2.ff + m1.gg * M2.gf + m1.gh * M2.hf:    .gg = m1.ga * M2.ag + m1.gb * M2.bg + m1.gc * M2.cg + m1.gd * M2.dg + m1.ge * M2.eg + m1.gf * M2.fg + m1.gg * M2.gg + m1.gh * M2.hg:    .gh = m1.ga * M2.ah + m1.gb * M2.bh + m1.gc * M2.ch + m1.gd * M2.dh + m1.ge * M2.eh + m1.gf * M2.fh + m1.gg * M2.gh + m1.gh * M2.hh
        .he = m1.ha * M2.ae + m1.hb * M2.be + m1.HC * M2.ce + m1.hd * M2.de + m1.he * M2.ee + m1.hf * M2.fe + m1.hg * M2.ge + m1.hh * M2.he:    .hf = m1.ha * M2.af + m1.hb * M2.bf + m1.HC * M2.cf + m1.hd * M2.df + m1.he * M2.ef + m1.hf * M2.ff + m1.hg * M2.gf + m1.hh * M2.hf:    .hg = m1.ga * M2.ag + m1.gb * M2.bg + m1.gc * M2.cg + m1.gd * M2.dg + m1.ge * M2.eg + m1.gf * M2.fg + m1.gg * M2.gg + m1.gh * M2.hg:    .hh = m1.ha * M2.ah + m1.hb * M2.bh + m1.HC * M2.ch + m1.hd * M2.dh + m1.he * M2.eh + m1.hf * M2.fh + m1.hg * M2.gh + m1.hh * M2.hh
    End With
End Function
Public Function Mat9_mul(m1 As Matrix9, M2 As Matrix9) As Matrix9
    'Multipliziert eine 9x9 Matrix m1 mit einer Matrix m2
    With Mat9_mul
        .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da + m1.ae * M2.ea + m1.af * M2.fa + m1.ag * M2.ga + m1.ah * M2.ha + m1.ai * M2.ia:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db + m1.ad * M2.eb + m1.af * M2.fb + m1.ag * M2.gb + m1.ah * M2.hb + m1.ai * M2.ib:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc + m1.ae * M2.ec + m1.af * M2.fc + m1.ag * M2.gc + m1.ah * M2.HC + m1.ai * M2.ic:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd + m1.ae * M2.ed + m1.af * M2.fd + m1.ag * M2.gd + m1.ah * M2.hd + m1.ai * M2.id:    .ae = m1.aa * M2.ae + m1.ab * M2.be + m1.ac * M2.ce + m1.ad * M2.de + m1.ae * M2.ee + m1.af * M2.fe + m1.ag * M2.ge + m1.ah * M2.he + m1.ai * M2.ie
        .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da + m1.be * M2.ea + m1.bf * M2.fa + m1.bg * M2.ga + m1.bh * M2.ha + m1.bi * M2.ia:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db + m1.bd * M2.eb + m1.bf * M2.fb + m1.bg * M2.gb + m1.bh * M2.hb + m1.bi * M2.ib:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc + m1.be * M2.ec + m1.bf * M2.fc + m1.bg * M2.gc + m1.bh * M2.HC + m1.bi * M2.ic:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd + m1.be * M2.ed + m1.bf * M2.fd + m1.bg * M2.gd + m1.bh * M2.hd + m1.bi * M2.id:    .be = m1.ba * M2.ae + m1.bb * M2.be + m1.bc * M2.ce + m1.bd * M2.de + m1.be * M2.ee + m1.bf * M2.fe + m1.bg * M2.ge + m1.bh * M2.he + m1.bi * M2.ie
        .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da + m1.ce * M2.ea + m1.cf * M2.fa + m1.cg * M2.ga + m1.ch * M2.ha + m1.ci * M2.ia:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db + m1.cd * M2.eb + m1.cf * M2.fb + m1.cg * M2.gb + m1.ch * M2.hb + m1.ci * M2.ib:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc + m1.ce * M2.ec + m1.cf * M2.fc + m1.cg * M2.gc + m1.ch * M2.HC + m1.ci * M2.ic:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd + m1.ce * M2.ed + m1.cf * M2.fd + m1.cg * M2.gd + m1.ch * M2.hd + m1.ci * M2.id:    .ce = m1.ca * M2.ae + m1.cb * M2.be + m1.cc * M2.ce + m1.cd * M2.de + m1.ce * M2.ee + m1.cf * M2.fe + m1.cg * M2.ge + m1.ch * M2.he + m1.ci * M2.ie
        .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da + m1.de * M2.ea + m1.df * M2.fa + m1.dg * M2.ga + m1.dh * M2.ha + m1.di * M2.ia:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db + m1.dd * M2.eb + m1.df * M2.fb + m1.dg * M2.gb + m1.dh * M2.hb + m1.di * M2.ib:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc + m1.de * M2.ec + m1.df * M2.fc + m1.dg * M2.gc + m1.dh * M2.HC + m1.di * M2.ic:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd + m1.de * M2.ed + m1.df * M2.fd + m1.dg * M2.gd + m1.dh * M2.hd + m1.di * M2.id:    .de = m1.da * M2.ae + m1.db * M2.be + m1.dc * M2.ce + m1.dd * M2.de + m1.de * M2.ee + m1.df * M2.fe + m1.dg * M2.ge + m1.dh * M2.he + m1.di * M2.ie
        .ea = m1.ea * M2.aa + m1.eb * M2.ba + m1.ec * M2.ca + m1.ed * M2.da + m1.ee * M2.ea + m1.ef * M2.fa + m1.eg * M2.ga + m1.eh * M2.ha + m1.ei * M2.ia:    .eb = m1.ea * M2.ab + m1.eb * M2.bb + m1.ec * M2.cb + m1.ed * M2.db + m1.ed * M2.eb + m1.ef * M2.fb + m1.eg * M2.gb + m1.eh * M2.hb + m1.ei * M2.ib:    .ec = m1.ea * M2.ac + m1.eb * M2.bc + m1.ec * M2.cc + m1.ed * M2.dc + m1.ee * M2.ec + m1.ef * M2.fc + m1.eg * M2.gc + m1.eh * M2.HC + m1.ei * M2.ic:    .ed = m1.ea * M2.ad + m1.eb * M2.bd + m1.ec * M2.cd + m1.ed * M2.dd + m1.ee * M2.ed + m1.ef * M2.fd + m1.eg * M2.gd + m1.eh * M2.hd + m1.ei * M2.id:    .ee = m1.ea * M2.ae + m1.eb * M2.be + m1.ec * M2.ce + m1.ed * M2.de + m1.ee * M2.ee + m1.ef * M2.fe + m1.eg * M2.ge + m1.eh * M2.he + m1.ei * M2.ie
        .fa = m1.fa * M2.aa + m1.fb * M2.ba + m1.fc * M2.ca + m1.fd * M2.da + m1.fe * M2.ea + m1.ff * M2.fa + m1.fg * M2.ga + m1.fh * M2.ha + m1.fi * M2.ia:    .fb = m1.fa * M2.ab + m1.fb * M2.bb + m1.fc * M2.cb + m1.fd * M2.db + m1.fd * M2.eb + m1.ff * M2.fb + m1.fg * M2.gb + m1.fh * M2.hb + m1.fi * M2.ib:    .fc = m1.fa * M2.ac + m1.fb * M2.bc + m1.fc * M2.cc + m1.fd * M2.dc + m1.fe * M2.ec + m1.ff * M2.fc + m1.fg * M2.gc + m1.fh * M2.HC + m1.fi * M2.ic:    .fd = m1.fa * M2.ad + m1.fb * M2.bd + m1.fc * M2.cd + m1.fd * M2.dd + m1.fe * M2.ed + m1.ff * M2.fd + m1.fg * M2.gd + m1.fh * M2.hd + m1.fi * M2.id:    .fe = m1.fa * M2.ae + m1.fb * M2.be + m1.fc * M2.ce + m1.fd * M2.de + m1.fe * M2.ee + m1.ff * M2.fe + m1.fg * M2.ge + m1.fh * M2.he + m1.fi * M2.ie
        .ga = m1.ga * M2.aa + m1.gb * M2.ba + m1.gc * M2.ca + m1.gd * M2.da + m1.ge * M2.ea + m1.gf * M2.fa + m1.gg * M2.ga + m1.gh * M2.ha + m1.gi * M2.ia:    .gb = m1.ga * M2.ab + m1.gb * M2.bb + m1.gc * M2.cb + m1.gd * M2.db + m1.gd * M2.eb + m1.gf * M2.fb + m1.gg * M2.gb + m1.gh * M2.hb + m1.gi * M2.ib:    .gc = m1.ga * M2.ac + m1.gb * M2.bc + m1.gc * M2.cc + m1.gd * M2.dc + m1.ge * M2.ec + m1.gf * M2.fc + m1.gg * M2.gc + m1.gh * M2.HC + m1.gi * M2.ic:    .gd = m1.ga * M2.ad + m1.gb * M2.bd + m1.gc * M2.cd + m1.gd * M2.dd + m1.ge * M2.ed + m1.gf * M2.fd + m1.gg * M2.gd + m1.gh * M2.hd + m1.gi * M2.id:    .ge = m1.ga * M2.ae + m1.gb * M2.be + m1.gc * M2.ce + m1.gd * M2.de + m1.ge * M2.ee + m1.gf * M2.fe + m1.gg * M2.ge + m1.gh * M2.he + m1.gi * M2.ie
        .ha = m1.ha * M2.aa + m1.hb * M2.ba + m1.HC * M2.ca + m1.hd * M2.da + m1.he * M2.ea + m1.hf * M2.fa + m1.hg * M2.ga + m1.hh * M2.ha + m1.hi * M2.ia:    .hb = m1.ha * M2.ab + m1.hb * M2.bb + m1.HC * M2.cb + m1.hd * M2.db + m1.hd * M2.eb + m1.hf * M2.fb + m1.hg * M2.gb + m1.hh * M2.hb + m1.hi * M2.ib:    .HC = m1.ha * M2.ac + m1.hb * M2.bc + m1.HC * M2.cc + m1.hd * M2.dc + m1.he * M2.ec + m1.hf * M2.fc + m1.hg * M2.gc + m1.hh * M2.HC + m1.hi * M2.ic:    .hd = m1.ha * M2.ad + m1.hb * M2.bd + m1.HC * M2.cd + m1.hd * M2.dd + m1.he * M2.ed + m1.hf * M2.fd + m1.hg * M2.gd + m1.hh * M2.hd + m1.hi * M2.id:    .he = m1.ha * M2.ae + m1.hb * M2.be + m1.HC * M2.ce + m1.hd * M2.de + m1.he * M2.ee + m1.hf * M2.fe + m1.hg * M2.ge + m1.hh * M2.he + m1.hi * M2.ie
        .ia = m1.ia * M2.aa + m1.ib * M2.ba + m1.ic * M2.ca + m1.id * M2.da + m1.ie * M2.ea + m1.if * M2.fa + m1.ig * M2.ga + m1.ih * M2.ha + m1.ii * M2.ia:    .ib = m1.ia * M2.ab + m1.ib * M2.bb + m1.ic * M2.cb + m1.id * M2.db + m1.id * M2.eb + m1.if * M2.fb + m1.ig * M2.gb + m1.ih * M2.hb + m1.ii * M2.ib:    .ic = m1.ia * M2.ac + m1.ib * M2.bc + m1.ic * M2.cc + m1.id * M2.dc + m1.ie * M2.ec + m1.if * M2.fc + m1.ig * M2.gc + m1.ih * M2.HC + m1.ii * M2.ic:    .id = m1.ia * M2.ad + m1.ib * M2.bd + m1.ic * M2.cd + m1.id * M2.dd + m1.ie * M2.ed + m1.if * M2.fd + m1.ig * M2.gd + m1.ih * M2.hd + m1.ii * M2.id:    .ie = m1.ia * M2.ae + m1.ib * M2.be + m1.ic * M2.ce + m1.id * M2.de + m1.ie * M2.ee + m1.if * M2.fe + m1.ig * M2.ge + m1.ih * M2.he + m1.ii * M2.ie
        
        .af = m1.aa * M2.af + m1.ab * M2.bf + m1.ac * M2.cf + m1.ad * M2.df + m1.ae * M2.ef + m1.af * M2.ff + m1.ag * M2.gf + m1.ah * M2.hf + m1.ai * M2.if:    .ag = m1.aa * M2.ag + m1.ab * M2.bg + m1.ac * M2.cg + m1.ad * M2.dg + m1.ae * M2.eg + m1.af * M2.fg + m1.ag * M2.gg + m1.ah * M2.hg + m1.ai * M2.ig:    .ah = m1.aa * M2.ah + m1.ab * M2.bh + m1.ac * M2.ch + m1.ad * M2.dh + m1.ae * M2.eh + m1.af * M2.fh + m1.ag * M2.gh + m1.ah * M2.hh + m1.ai * M2.ih:    .ai = m1.aa * M2.ai + m1.ab * M2.bi + m1.ac * M2.ci + m1.ad * M2.di + m1.ae * M2.ei + m1.af * M2.fi + m1.ag * M2.gi + m1.ah * M2.hi + m1.ai * M2.ii
        .bf = m1.ba * M2.af + m1.bb * M2.bf + m1.bc * M2.cf + m1.bd * M2.df + m1.be * M2.ef + m1.bf * M2.ff + m1.bg * M2.gf + m1.bh * M2.hf + m1.bi * M2.if:    .bg = m1.ba * M2.ag + m1.bb * M2.bg + m1.bc * M2.cg + m1.bd * M2.dg + m1.be * M2.eg + m1.bf * M2.fg + m1.bg * M2.gg + m1.bh * M2.hg + m1.bi * M2.ig:    .bh = m1.ba * M2.ah + m1.bb * M2.bh + m1.bc * M2.ch + m1.bd * M2.dh + m1.be * M2.eh + m1.bf * M2.fh + m1.bg * M2.gh + m1.bh * M2.hh + m1.bi * M2.ih:    .bi = m1.ba * M2.ai + m1.bb * M2.bi + m1.bc * M2.ci + m1.bd * M2.di + m1.be * M2.ei + m1.bf * M2.fi + m1.bg * M2.gi + m1.bh * M2.hi + m1.bi * M2.ii
        .cf = m1.ca * M2.af + m1.cb * M2.bf + m1.cc * M2.cf + m1.cd * M2.df + m1.ce * M2.ef + m1.cf * M2.ff + m1.cg * M2.gf + m1.ch * M2.hf + m1.ci * M2.if:    .cg = m1.ca * M2.ag + m1.cb * M2.bg + m1.cc * M2.cg + m1.cd * M2.dg + m1.ce * M2.eg + m1.cf * M2.fg + m1.cg * M2.gg + m1.ch * M2.hg + m1.ci * M2.ig:    .ch = m1.ca * M2.ah + m1.cb * M2.bh + m1.cc * M2.ch + m1.cd * M2.dh + m1.ce * M2.eh + m1.cf * M2.fh + m1.cg * M2.gh + m1.ch * M2.hh + m1.ci * M2.ih:    .ci = m1.ca * M2.ai + m1.cb * M2.bi + m1.cc * M2.ci + m1.cd * M2.di + m1.ce * M2.ei + m1.cf * M2.fi + m1.cg * M2.gi + m1.ch * M2.hi + m1.ci * M2.ii
        .df = m1.da * M2.af + m1.db * M2.bf + m1.dc * M2.cf + m1.dd * M2.df + m1.de * M2.ef + m1.df * M2.ff + m1.dg * M2.gf + m1.dh * M2.hf + m1.di * M2.if:    .dg = m1.da * M2.ag + m1.db * M2.bg + m1.dc * M2.cg + m1.dd * M2.dg + m1.de * M2.eg + m1.df * M2.fg + m1.dg * M2.gg + m1.dh * M2.hg + m1.di * M2.ig:    .dh = m1.da * M2.ah + m1.db * M2.bh + m1.dc * M2.ch + m1.dd * M2.dh + m1.de * M2.eh + m1.df * M2.fh + m1.dg * M2.gh + m1.dh * M2.hh + m1.di * M2.ih:    .di = m1.da * M2.ai + m1.db * M2.bi + m1.dc * M2.ci + m1.dd * M2.di + m1.de * M2.ei + m1.df * M2.fi + m1.dg * M2.gi + m1.dh * M2.hi + m1.di * M2.ii
        .ef = m1.ea * M2.af + m1.eb * M2.bf + m1.ec * M2.cf + m1.ed * M2.df + m1.ee * M2.ef + m1.ef * M2.ff + m1.eg * M2.gf + m1.eh * M2.hf + m1.ei * M2.if:    .eg = m1.ea * M2.ag + m1.eb * M2.bg + m1.ec * M2.cg + m1.ed * M2.dg + m1.ee * M2.eg + m1.ef * M2.fg + m1.eg * M2.gg + m1.eh * M2.hg + m1.ei * M2.ig:    .eh = m1.ea * M2.ah + m1.eb * M2.bh + m1.ec * M2.ch + m1.ed * M2.dh + m1.ee * M2.eh + m1.ef * M2.fh + m1.eg * M2.gh + m1.eh * M2.hh + m1.ei * M2.ih:    .ei = m1.ea * M2.ai + m1.eb * M2.bi + m1.ec * M2.ci + m1.ed * M2.di + m1.ee * M2.ei + m1.ef * M2.fi + m1.eg * M2.gi + m1.eh * M2.hi + m1.ei * M2.ii
        .ff = m1.fa * M2.af + m1.fb * M2.bf + m1.fc * M2.cf + m1.fd * M2.df + m1.fe * M2.ef + m1.ff * M2.ff + m1.fg * M2.gf + m1.fh * M2.hf + m1.fi * M2.if:    .fg = m1.fa * M2.ag + m1.fb * M2.bg + m1.fc * M2.cg + m1.fd * M2.dg + m1.fe * M2.eg + m1.ff * M2.fg + m1.fg * M2.gg + m1.fh * M2.hg + m1.fi * M2.ig:    .fh = m1.fa * M2.ah + m1.fb * M2.bh + m1.fc * M2.ch + m1.fd * M2.dh + m1.fe * M2.eh + m1.ff * M2.fh + m1.fg * M2.gh + m1.fh * M2.hh + m1.fi * M2.ih:    .fi = m1.fa * M2.ai + m1.fb * M2.bi + m1.fc * M2.ci + m1.fd * M2.di + m1.fe * M2.ei + m1.ff * M2.fi + m1.fg * M2.gi + m1.fh * M2.hi + m1.fi * M2.ii
        .gf = m1.ga * M2.af + m1.gb * M2.bf + m1.gc * M2.cf + m1.gd * M2.df + m1.ge * M2.ef + m1.gf * M2.ff + m1.gg * M2.gf + m1.gh * M2.hf + m1.gi * M2.if:    .gg = m1.ga * M2.ag + m1.gb * M2.bg + m1.gc * M2.cg + m1.gd * M2.dg + m1.ge * M2.eg + m1.gf * M2.fg + m1.gg * M2.gg + m1.gh * M2.hg + m1.gi * M2.ig:    .gh = m1.ga * M2.ah + m1.gb * M2.bh + m1.gc * M2.ch + m1.gd * M2.dh + m1.ge * M2.eh + m1.gf * M2.fh + m1.gg * M2.gh + m1.gh * M2.hh + m1.gi * M2.ih:    .gi = m1.ga * M2.ai + m1.gb * M2.bi + m1.gc * M2.ci + m1.gd * M2.di + m1.ge * M2.ei + m1.gf * M2.fi + m1.gg * M2.gi + m1.gh * M2.hi + m1.gi * M2.ii
        .hf = m1.ha * M2.af + m1.hb * M2.bf + m1.HC * M2.cf + m1.hd * M2.df + m1.he * M2.ef + m1.hf * M2.ff + m1.hg * M2.gf + m1.hh * M2.hf + m1.hi * M2.if:    .hg = m1.ha * M2.ag + m1.hb * M2.bg + m1.HC * M2.cg + m1.hd * M2.dg + m1.he * M2.eg + m1.hf * M2.fg + m1.hg * M2.gg + m1.hh * M2.hg + m1.hi * M2.ig:    .hh = m1.ha * M2.ah + m1.hb * M2.bh + m1.HC * M2.ch + m1.hd * M2.dh + m1.he * M2.eh + m1.hf * M2.fh + m1.hg * M2.gh + m1.hh * M2.hh + m1.hi * M2.ih:    .hi = m1.ha * M2.ai + m1.hb * M2.bi + m1.HC * M2.ci + m1.hd * M2.di + m1.he * M2.ei + m1.hf * M2.fi + m1.hg * M2.gi + m1.hh * M2.hi + m1.hi * M2.ii
        .if = m1.ia * M2.af + m1.ib * M2.bf + m1.ic * M2.cf + m1.id * M2.df + m1.ie * M2.ef + m1.if * M2.ff + m1.ig * M2.gf + m1.ih * M2.hf + m1.ii * M2.if:    .ig = m1.ia * M2.ag + m1.ib * M2.bg + m1.ic * M2.cg + m1.id * M2.dg + m1.ie * M2.eg + m1.if * M2.fg + m1.ig * M2.gg + m1.ih * M2.hg + m1.ii * M2.ig:    .ih = m1.ia * M2.ah + m1.ib * M2.bh + m1.ic * M2.ch + m1.id * M2.dh + m1.ie * M2.eh + m1.if * M2.fh + m1.ig * M2.gh + m1.ih * M2.hh + m1.ii * M2.ih:    .ii = m1.ia * M2.ai + m1.ib * M2.bi + m1.ic * M2.ci + m1.id * M2.di + m1.ie * M2.ei + m1.if * M2.fi + m1.ig * M2.gi + m1.ih * M2.hi + m1.ii * M2.ii
    End With
End Function
Public Function Mat10_mul(m1 As Matrix10, M2 As Matrix10) As Matrix10
    'Multipliziert eine 10x10 Matrix m1 mit einer Matrix m2
    With Mat10_mul
        .aa = m1.aa * M2.aa + m1.ab * M2.ba + m1.ac * M2.ca + m1.ad * M2.da + m1.ae * M2.ea + m1.af * M2.fa + m1.ag * M2.ga + m1.ah * M2.ha + m1.ai * M2.ia + m1.aj * M2.ja:    .ab = m1.aa * M2.ab + m1.ab * M2.bb + m1.ac * M2.cb + m1.ad * M2.db + m1.ad * M2.eb + m1.af * M2.fb + m1.ag * M2.gb + m1.ah * M2.hb + m1.ai * M2.ib + m1.aj * M2.jb:    .ac = m1.aa * M2.ac + m1.ab * M2.bc + m1.ac * M2.cc + m1.ad * M2.dc + m1.ae * M2.ec + m1.af * M2.fc + m1.ag * M2.gc + m1.ah * M2.HC + m1.ai * M2.ic + m1.aj * M2.jc:    .ad = m1.aa * M2.ad + m1.ab * M2.bd + m1.ac * M2.cd + m1.ad * M2.dd + m1.ae * M2.ed + m1.af * M2.fd + m1.ag * M2.gd + m1.ah * M2.hd + m1.ai * M2.id + m1.aj * M2.jd:    .ae = m1.aa * M2.ae + m1.ab * M2.be + m1.ac * M2.ce + m1.ad * M2.de + m1.ae * M2.ee + m1.af * M2.fe + m1.ag * M2.ge + m1.ah * M2.he + m1.ai * M2.ie + m1.aj * M2.je
        .ba = m1.ba * M2.aa + m1.bb * M2.ba + m1.bc * M2.ca + m1.bd * M2.da + m1.be * M2.ea + m1.bf * M2.fa + m1.bg * M2.ga + m1.bh * M2.ha + m1.bi * M2.ia + m1.bj * M2.ja:    .bb = m1.ba * M2.ab + m1.bb * M2.bb + m1.bc * M2.cb + m1.bd * M2.db + m1.bd * M2.eb + m1.bf * M2.fb + m1.bg * M2.gb + m1.bh * M2.hb + m1.bi * M2.ib + m1.bj * M2.jb:    .bc = m1.ba * M2.ac + m1.bb * M2.bc + m1.bc * M2.cc + m1.bd * M2.dc + m1.be * M2.ec + m1.bf * M2.fc + m1.bg * M2.gc + m1.bh * M2.HC + m1.bi * M2.ic + m1.bj * M2.jc:    .bd = m1.ba * M2.ad + m1.bb * M2.bd + m1.bc * M2.cd + m1.bd * M2.dd + m1.be * M2.ed + m1.bf * M2.fd + m1.bg * M2.gd + m1.bh * M2.hd + m1.bi * M2.id + m1.bj * M2.jd:    .be = m1.ba * M2.ae + m1.bb * M2.be + m1.bc * M2.ce + m1.bd * M2.de + m1.be * M2.ee + m1.bf * M2.fe + m1.bg * M2.ge + m1.bh * M2.he + m1.bi * M2.ie + m1.bj * M2.je
        .ca = m1.ca * M2.aa + m1.cb * M2.ba + m1.cc * M2.ca + m1.cd * M2.da + m1.ce * M2.ea + m1.cf * M2.fa + m1.cg * M2.ga + m1.ch * M2.ha + m1.ci * M2.ia + m1.cj * M2.ja:    .cb = m1.ca * M2.ab + m1.cb * M2.bb + m1.cc * M2.cb + m1.cd * M2.db + m1.cd * M2.eb + m1.cf * M2.fb + m1.cg * M2.gb + m1.ch * M2.hb + m1.ci * M2.ib + m1.cj * M2.jb:    .cc = m1.ca * M2.ac + m1.cb * M2.bc + m1.cc * M2.cc + m1.cd * M2.dc + m1.ce * M2.ec + m1.cf * M2.fc + m1.cg * M2.gc + m1.ch * M2.HC + m1.ci * M2.ic + m1.cj * M2.jc:    .cd = m1.ca * M2.ad + m1.cb * M2.bd + m1.cc * M2.cd + m1.cd * M2.dd + m1.ce * M2.ed + m1.cf * M2.fd + m1.cg * M2.gd + m1.ch * M2.hd + m1.ci * M2.id + m1.cj * M2.jd:    .ce = m1.ca * M2.ae + m1.cb * M2.be + m1.cc * M2.ce + m1.cd * M2.de + m1.ce * M2.ee + m1.cf * M2.fe + m1.cg * M2.ge + m1.ch * M2.he + m1.ci * M2.ie + m1.cj * M2.je
        .da = m1.da * M2.aa + m1.db * M2.ba + m1.dc * M2.ca + m1.dd * M2.da + m1.de * M2.ea + m1.df * M2.fa + m1.dg * M2.ga + m1.dh * M2.ha + m1.di * M2.ia + m1.dj * M2.ja:    .db = m1.da * M2.ab + m1.db * M2.bb + m1.dc * M2.cb + m1.dd * M2.db + m1.dd * M2.eb + m1.df * M2.fb + m1.dg * M2.gb + m1.dh * M2.hb + m1.di * M2.ib + m1.dj * M2.jb:    .dc = m1.da * M2.ac + m1.db * M2.bc + m1.dc * M2.cc + m1.dd * M2.dc + m1.de * M2.ec + m1.df * M2.fc + m1.dg * M2.gc + m1.dh * M2.HC + m1.di * M2.ic + m1.dj * M2.jc:    .dd = m1.da * M2.ad + m1.db * M2.bd + m1.dc * M2.cd + m1.dd * M2.dd + m1.de * M2.ed + m1.df * M2.fd + m1.dg * M2.gd + m1.dh * M2.hd + m1.di * M2.id + m1.dj * M2.jd:    .de = m1.da * M2.ae + m1.db * M2.be + m1.dc * M2.ce + m1.dd * M2.de + m1.de * M2.ee + m1.df * M2.fe + m1.dg * M2.ge + m1.dh * M2.he + m1.di * M2.ie + m1.dj * M2.je
        .ea = m1.ea * M2.aa + m1.eb * M2.ba + m1.ec * M2.ca + m1.ed * M2.da + m1.ee * M2.ea + m1.ef * M2.fa + m1.eg * M2.ga + m1.eh * M2.ha + m1.ei * M2.ia + m1.ej * M2.ja:    .eb = m1.ea * M2.ab + m1.eb * M2.bb + m1.ec * M2.cb + m1.ed * M2.db + m1.ed * M2.eb + m1.ef * M2.fb + m1.eg * M2.gb + m1.eh * M2.hb + m1.ei * M2.ib + m1.ej * M2.jb:    .ec = m1.ea * M2.ac + m1.eb * M2.bc + m1.ec * M2.cc + m1.ed * M2.dc + m1.ee * M2.ec + m1.ef * M2.fc + m1.eg * M2.gc + m1.eh * M2.HC + m1.ei * M2.ic + m1.ej * M2.jc:    .ed = m1.ea * M2.ad + m1.eb * M2.bd + m1.ec * M2.cd + m1.ed * M2.dd + m1.ee * M2.ed + m1.ef * M2.fd + m1.eg * M2.gd + m1.eh * M2.hd + m1.ei * M2.id + m1.ej * M2.jd:    .ee = m1.ea * M2.ae + m1.eb * M2.be + m1.ec * M2.ce + m1.ed * M2.de + m1.ee * M2.ee + m1.ef * M2.fe + m1.eg * M2.ge + m1.eh * M2.he + m1.ei * M2.ie + m1.ej * M2.je
        .fa = m1.fa * M2.aa + m1.fb * M2.ba + m1.fc * M2.ca + m1.fd * M2.da + m1.fe * M2.ea + m1.ff * M2.fa + m1.fg * M2.ga + m1.fh * M2.ha + m1.fi * M2.ia + m1.fj * M2.ja:    .fb = m1.fa * M2.ab + m1.fb * M2.bb + m1.fc * M2.cb + m1.fd * M2.db + m1.fd * M2.eb + m1.ff * M2.fb + m1.fg * M2.gb + m1.fh * M2.hb + m1.fi * M2.ib + m1.fj * M2.jb:    .fc = m1.fa * M2.ac + m1.fb * M2.bc + m1.fc * M2.cc + m1.fd * M2.dc + m1.fe * M2.ec + m1.ff * M2.fc + m1.fg * M2.gc + m1.fh * M2.HC + m1.fi * M2.ic + m1.fj * M2.jc:    .fd = m1.fa * M2.ad + m1.fb * M2.bd + m1.fc * M2.cd + m1.fd * M2.dd + m1.fe * M2.ed + m1.ff * M2.fd + m1.fg * M2.gd + m1.fh * M2.hd + m1.fi * M2.id + m1.fj * M2.jd:    .fe = m1.fa * M2.ae + m1.fb * M2.be + m1.fc * M2.ce + m1.fd * M2.de + m1.fe * M2.ee + m1.ff * M2.fe + m1.fg * M2.ge + m1.fh * M2.he + m1.fi * M2.ie + m1.fj * M2.je
        .ga = m1.ga * M2.aa + m1.gb * M2.ba + m1.gc * M2.ca + m1.gd * M2.da + m1.ge * M2.ea + m1.gf * M2.fa + m1.gg * M2.ga + m1.gh * M2.ha + m1.gi * M2.ia + m1.gj * M2.ja:    .gb = m1.ga * M2.ab + m1.gb * M2.bb + m1.gc * M2.cb + m1.gd * M2.db + m1.gd * M2.eb + m1.gf * M2.fb + m1.gg * M2.gb + m1.gh * M2.hb + m1.gi * M2.ib + m1.gj * M2.jb:    .gc = m1.ga * M2.ac + m1.gb * M2.bc + m1.gc * M2.cc + m1.gd * M2.dc + m1.ge * M2.ec + m1.gf * M2.fc + m1.gg * M2.gc + m1.gh * M2.HC + m1.gi * M2.ic + m1.gj * M2.jc:    .gd = m1.ga * M2.ad + m1.gb * M2.bd + m1.gc * M2.cd + m1.gd * M2.dd + m1.ge * M2.ed + m1.gf * M2.fd + m1.gg * M2.gd + m1.gh * M2.hd + m1.gi * M2.id + m1.gj * M2.jd:    .ge = m1.ga * M2.ae + m1.gb * M2.be + m1.gc * M2.ce + m1.gd * M2.de + m1.ge * M2.ee + m1.gf * M2.fe + m1.gg * M2.ge + m1.gh * M2.he + m1.gi * M2.ie + m1.gj * M2.je
        .ha = m1.ha * M2.aa + m1.hb * M2.ba + m1.HC * M2.ca + m1.hd * M2.da + m1.he * M2.ea + m1.hf * M2.fa + m1.hg * M2.ga + m1.hh * M2.ha + m1.hi * M2.ia + m1.hj * M2.ja:    .hb = m1.ha * M2.ab + m1.hb * M2.bb + m1.HC * M2.cb + m1.hd * M2.db + m1.hd * M2.eb + m1.hf * M2.fb + m1.hg * M2.gb + m1.hh * M2.hb + m1.hi * M2.ib + m1.hj * M2.jb:    .HC = m1.ha * M2.ac + m1.hb * M2.bc + m1.HC * M2.cc + m1.hd * M2.dc + m1.he * M2.ec + m1.hf * M2.fc + m1.hg * M2.gc + m1.hh * M2.HC + m1.hi * M2.ic + m1.hj * M2.jc:    .hd = m1.ha * M2.ad + m1.hb * M2.bd + m1.HC * M2.cd + m1.hd * M2.dd + m1.he * M2.ed + m1.hf * M2.fd + m1.hg * M2.gd + m1.hh * M2.hd + m1.hi * M2.id + m1.hj * M2.jd:    .he = m1.ha * M2.ae + m1.hb * M2.be + m1.HC * M2.ce + m1.hd * M2.de + m1.he * M2.ee + m1.hf * M2.fe + m1.hg * M2.ge + m1.hh * M2.he + m1.hi * M2.ie + m1.hj * M2.je
        .ia = m1.ia * M2.aa + m1.ib * M2.ba + m1.ic * M2.ca + m1.id * M2.da + m1.ie * M2.ea + m1.if * M2.fa + m1.ig * M2.ga + m1.ih * M2.ha + m1.ii * M2.ia + m1.ij * M2.ja:    .ib = m1.ia * M2.ab + m1.ib * M2.bb + m1.ic * M2.cb + m1.id * M2.db + m1.id * M2.eb + m1.if * M2.fb + m1.ig * M2.gb + m1.ih * M2.hb + m1.ii * M2.ib + m1.ij * M2.jb:    .ic = m1.ia * M2.ac + m1.ib * M2.bc + m1.ic * M2.cc + m1.id * M2.dc + m1.ie * M2.ec + m1.if * M2.fc + m1.ig * M2.gc + m1.ih * M2.HC + m1.ii * M2.ic + m1.ij * M2.jc:    .id = m1.ia * M2.ad + m1.ib * M2.bd + m1.ic * M2.cd + m1.id * M2.dd + m1.ie * M2.ed + m1.if * M2.fd + m1.ig * M2.gd + m1.ih * M2.hd + m1.ii * M2.id + m1.ij * M2.jd:    .ie = m1.ia * M2.ae + m1.ib * M2.be + m1.ic * M2.ce + m1.id * M2.de + m1.ie * M2.ee + m1.if * M2.fe + m1.ig * M2.ge + m1.ih * M2.he + m1.ii * M2.ie + m1.ij * M2.je
        .ja = m1.ja * M2.aa + m1.jb * M2.ba + m1.jc * M2.ca + m1.jd * M2.da + m1.je * M2.ea + m1.jf * M2.fa + m1.jg * M2.ga + m1.jh * M2.ha + m1.ji * M2.ia + m1.jj * M2.ja:    .jb = m1.ja * M2.ab + m1.jb * M2.bb + m1.jc * M2.cb + m1.jd * M2.db + m1.jd * M2.eb + m1.jf * M2.fb + m1.jg * M2.gb + m1.jh * M2.hb + m1.ji * M2.ib + m1.jj * M2.jb:    .jc = m1.ja * M2.ac + m1.jb * M2.bc + m1.jc * M2.cc + m1.jd * M2.dc + m1.je * M2.ec + m1.jf * M2.fc + m1.jg * M2.gc + m1.jh * M2.HC + m1.ji * M2.ic + m1.jj * M2.jc:    .jd = m1.ja * M2.ad + m1.jb * M2.bd + m1.jc * M2.cd + m1.jd * M2.dd + m1.je * M2.ed + m1.jf * M2.fd + m1.jg * M2.gd + m1.jh * M2.hd + m1.ji * M2.id + m1.jj * M2.jd:    .je = m1.ja * M2.ae + m1.jb * M2.be + m1.jc * M2.ce + m1.jd * M2.de + m1.je * M2.ee + m1.jf * M2.fe + m1.jg * M2.ge + m1.jh * M2.he + m1.ji * M2.ie + m1.jj * M2.je
        
        
        .af = m1.aa * M2.af + m1.ab * M2.bf + m1.ac * M2.cf + m1.ad * M2.df + m1.ae * M2.ef + m1.af * M2.ff + m1.ag * M2.gf + m1.ah * M2.hf + m1.ai * M2.if + m1.aj * M2.jf:    .ag = m1.aa * M2.ag + m1.ab * M2.bg + m1.ac * M2.cg + m1.ad * M2.dg + m1.ae * M2.eg + m1.af * M2.fg + m1.ag * M2.gg + m1.ah * M2.hg + m1.ai * M2.ig + m1.aj * M2.jg:    .ah = m1.aa * M2.ah + m1.ab * M2.bh + m1.ac * M2.ch + m1.ad * M2.dh + m1.ae * M2.eh + m1.af * M2.fh + m1.ag * M2.gh + m1.ah * M2.hh + m1.ai * M2.ih + m1.aj * M2.jh:    .ai = m1.aa * M2.ai + m1.ab * M2.bi + m1.ac * M2.ci + m1.ad * M2.di + m1.ae * M2.ei + m1.af * M2.fi + m1.ag * M2.gi + m1.ah * M2.hi + m1.ai * M2.ii + m1.aj * M2.ji:    .aj = m1.aa * M2.aj + m1.ab * M2.bj + m1.ac * M2.cj + m1.ad * M2.dj + m1.ae * M2.ej + m1.af * M2.fj + m1.ag * M2.gj + m1.ah * M2.hj + m1.ai * M2.ij + m1.aj * M2.jj
        .bf = m1.ba * M2.af + m1.bb * M2.bf + m1.bc * M2.cf + m1.bd * M2.df + m1.be * M2.ef + m1.bf * M2.ff + m1.bg * M2.gf + m1.bh * M2.hf + m1.bi * M2.if + m1.bj * M2.jf:    .bg = m1.ba * M2.ag + m1.bb * M2.bg + m1.bc * M2.cg + m1.bd * M2.dg + m1.be * M2.eg + m1.bf * M2.fg + m1.bg * M2.gg + m1.bh * M2.hg + m1.bi * M2.ig + m1.bj * M2.jg:    .bh = m1.ba * M2.ah + m1.bb * M2.bh + m1.bc * M2.ch + m1.bd * M2.dh + m1.be * M2.eh + m1.bf * M2.fh + m1.bg * M2.gh + m1.bh * M2.hh + m1.bi * M2.ih + m1.bj * M2.jh:    .bi = m1.ba * M2.ai + m1.bb * M2.bi + m1.bc * M2.ci + m1.bd * M2.di + m1.be * M2.ei + m1.bf * M2.fi + m1.bg * M2.gi + m1.bh * M2.hi + m1.bi * M2.ii + m1.bj * M2.ji:    .bj = m1.ba * M2.aj + m1.bb * M2.bj + m1.bc * M2.cj + m1.bd * M2.dj + m1.be * M2.ej + m1.bf * M2.fj + m1.bg * M2.gj + m1.bh * M2.hj + m1.bi * M2.ij + m1.bj * M2.jj
        .cf = m1.ca * M2.af + m1.cb * M2.bf + m1.cc * M2.cf + m1.cd * M2.df + m1.ce * M2.ef + m1.cf * M2.ff + m1.cg * M2.gf + m1.ch * M2.hf + m1.ci * M2.if + m1.cj * M2.jf:    .cg = m1.ca * M2.ag + m1.cb * M2.bg + m1.cc * M2.cg + m1.cd * M2.dg + m1.ce * M2.eg + m1.cf * M2.fg + m1.cg * M2.gg + m1.ch * M2.hg + m1.ci * M2.ig + m1.cj * M2.jg:    .ch = m1.ca * M2.ah + m1.cb * M2.bh + m1.cc * M2.ch + m1.cd * M2.dh + m1.ce * M2.eh + m1.cf * M2.fh + m1.cg * M2.gh + m1.ch * M2.hh + m1.ci * M2.ih + m1.cj * M2.jh:    .ci = m1.ca * M2.ai + m1.cb * M2.bi + m1.cc * M2.ci + m1.cd * M2.di + m1.ce * M2.ei + m1.cf * M2.fi + m1.cg * M2.gi + m1.ch * M2.hi + m1.ci * M2.ii + m1.cj * M2.ji:    .cj = m1.ca * M2.aj + m1.cb * M2.bj + m1.cc * M2.cj + m1.cd * M2.dj + m1.ce * M2.ej + m1.cf * M2.fj + m1.cg * M2.gj + m1.ch * M2.hj + m1.ci * M2.ij + m1.cj * M2.jj
        .df = m1.da * M2.af + m1.db * M2.bf + m1.dc * M2.cf + m1.dd * M2.df + m1.de * M2.ef + m1.df * M2.ff + m1.dg * M2.gf + m1.dh * M2.hf + m1.di * M2.if + m1.dj * M2.jf:    .dg = m1.da * M2.ag + m1.db * M2.bg + m1.dc * M2.cg + m1.dd * M2.dg + m1.de * M2.eg + m1.df * M2.fg + m1.dg * M2.gg + m1.dh * M2.hg + m1.di * M2.ig + m1.dj * M2.jg:    .dh = m1.da * M2.ah + m1.db * M2.bh + m1.dc * M2.ch + m1.dd * M2.dh + m1.de * M2.eh + m1.df * M2.fh + m1.dg * M2.gh + m1.dh * M2.hh + m1.di * M2.ih + m1.dj * M2.jh:    .di = m1.da * M2.ai + m1.db * M2.bi + m1.dc * M2.ci + m1.dd * M2.di + m1.de * M2.ei + m1.df * M2.fi + m1.dg * M2.gi + m1.dh * M2.hi + m1.di * M2.ii + m1.dj * M2.ji:    .dj = m1.da * M2.aj + m1.db * M2.bj + m1.dc * M2.cj + m1.dd * M2.dj + m1.de * M2.ej + m1.df * M2.fj + m1.dg * M2.gj + m1.dh * M2.hj + m1.di * M2.ij + m1.dj * M2.jj
        .ef = m1.ea * M2.af + m1.eb * M2.bf + m1.ec * M2.cf + m1.ed * M2.df + m1.ee * M2.ef + m1.ef * M2.ff + m1.eg * M2.gf + m1.eh * M2.hf + m1.ei * M2.if + m1.ej * M2.jf:    .eg = m1.ea * M2.ag + m1.eb * M2.bg + m1.ec * M2.cg + m1.ed * M2.dg + m1.ee * M2.eg + m1.ef * M2.fg + m1.eg * M2.gg + m1.eh * M2.hg + m1.ei * M2.ig + m1.ej * M2.jg:    .eh = m1.ea * M2.ah + m1.eb * M2.bh + m1.ec * M2.ch + m1.ed * M2.dh + m1.ee * M2.eh + m1.ef * M2.fh + m1.eg * M2.gh + m1.eh * M2.hh + m1.ei * M2.ih + m1.ej * M2.jh:    .ei = m1.ea * M2.ai + m1.eb * M2.bi + m1.ec * M2.ci + m1.ed * M2.di + m1.ee * M2.ei + m1.ef * M2.fi + m1.eg * M2.gi + m1.eh * M2.hi + m1.ei * M2.ii + m1.ej * M2.ji:    .ej = m1.ea * M2.aj + m1.eb * M2.bj + m1.ec * M2.cj + m1.ed * M2.dj + m1.ee * M2.ej + m1.ef * M2.fj + m1.eg * M2.gj + m1.eh * M2.hj + m1.ei * M2.ij + m1.ej * M2.jj
        .ff = m1.fa * M2.af + m1.fb * M2.bf + m1.fc * M2.cf + m1.fd * M2.df + m1.fe * M2.ef + m1.ff * M2.ff + m1.fg * M2.gf + m1.fh * M2.hf + m1.fi * M2.if + m1.fj * M2.jf:    .fg = m1.fa * M2.ag + m1.fb * M2.bg + m1.fc * M2.cg + m1.fd * M2.dg + m1.fe * M2.eg + m1.ff * M2.fg + m1.fg * M2.gg + m1.fh * M2.hg + m1.fi * M2.ig + m1.fj * M2.jg:    .fh = m1.fa * M2.ah + m1.fb * M2.bh + m1.fc * M2.ch + m1.fd * M2.dh + m1.fe * M2.eh + m1.ff * M2.fh + m1.fg * M2.gh + m1.fh * M2.hh + m1.fi * M2.ih + m1.fj * M2.jh:    .fi = m1.fa * M2.ai + m1.fb * M2.bi + m1.fc * M2.ci + m1.fd * M2.di + m1.fe * M2.ei + m1.ff * M2.fi + m1.fg * M2.gi + m1.fh * M2.hi + m1.fi * M2.ii + m1.fj * M2.ji:    .fj = m1.fa * M2.aj + m1.fb * M2.bj + m1.fc * M2.cj + m1.fd * M2.dj + m1.fe * M2.ej + m1.ff * M2.fj + m1.fg * M2.gj + m1.fh * M2.hj + m1.fi * M2.ij + m1.fj * M2.jj
        .gf = m1.ga * M2.af + m1.gb * M2.bf + m1.gc * M2.cf + m1.gd * M2.df + m1.ge * M2.ef + m1.gf * M2.ff + m1.gg * M2.gf + m1.gh * M2.hf + m1.gi * M2.if + m1.gj * M2.jf:    .gg = m1.ga * M2.ag + m1.gb * M2.bg + m1.gc * M2.cg + m1.gd * M2.dg + m1.ge * M2.eg + m1.gf * M2.fg + m1.gg * M2.gg + m1.gh * M2.hg + m1.gi * M2.ig + m1.gj * M2.jg:    .gh = m1.ga * M2.ah + m1.gb * M2.bh + m1.gc * M2.ch + m1.gd * M2.dh + m1.ge * M2.eh + m1.gf * M2.fh + m1.gg * M2.gh + m1.gh * M2.hh + m1.gi * M2.ih + m1.gj * M2.jh:    .gi = m1.ga * M2.ai + m1.gb * M2.bi + m1.gc * M2.ci + m1.gd * M2.di + m1.ge * M2.ei + m1.gf * M2.fi + m1.gg * M2.gi + m1.gh * M2.hi + m1.gi * M2.ii + m1.gj * M2.ji:    .gj = m1.ga * M2.aj + m1.gb * M2.bj + m1.gc * M2.cj + m1.gd * M2.dj + m1.ge * M2.ej + m1.gf * M2.fj + m1.gg * M2.gj + m1.gh * M2.hj + m1.gi * M2.ij + m1.gj * M2.jj
        .hf = m1.ha * M2.af + m1.hb * M2.bf + m1.HC * M2.cf + m1.hd * M2.df + m1.he * M2.ef + m1.hf * M2.ff + m1.hg * M2.gf + m1.hh * M2.hf + m1.hi * M2.if + m1.hj * M2.jf:    .hg = m1.ha * M2.ag + m1.hb * M2.bg + m1.HC * M2.cg + m1.hd * M2.dg + m1.he * M2.eg + m1.hf * M2.fg + m1.hg * M2.gg + m1.hh * M2.hg + m1.hi * M2.ig + m1.hj * M2.jg:    .hh = m1.ha * M2.ah + m1.hb * M2.bh + m1.HC * M2.ch + m1.hd * M2.dh + m1.he * M2.eh + m1.hf * M2.fh + m1.hg * M2.gh + m1.hh * M2.hh + m1.hi * M2.ih + m1.hj * M2.jh:    .hi = m1.ha * M2.ai + m1.hb * M2.bi + m1.HC * M2.ci + m1.hd * M2.di + m1.he * M2.ei + m1.hf * M2.fi + m1.hg * M2.gi + m1.hh * M2.hi + m1.hi * M2.ii + m1.hj * M2.ji:    .hj = m1.ha * M2.aj + m1.hb * M2.bj + m1.HC * M2.cj + m1.hd * M2.dj + m1.he * M2.ej + m1.hf * M2.fj + m1.hg * M2.gj + m1.hh * M2.hj + m1.hi * M2.ij + m1.hj * M2.jj
        .if = m1.ia * M2.af + m1.ib * M2.bf + m1.ic * M2.cf + m1.id * M2.df + m1.ie * M2.ef + m1.if * M2.ff + m1.ig * M2.gf + m1.ih * M2.hf + m1.ii * M2.if + m1.ij * M2.jf:    .ig = m1.ia * M2.ag + m1.ib * M2.bg + m1.ic * M2.cg + m1.id * M2.dg + m1.ie * M2.eg + m1.if * M2.fg + m1.ig * M2.gg + m1.ih * M2.hg + m1.ii * M2.ig + m1.ij * M2.jg:    .ih = m1.ia * M2.ah + m1.ib * M2.bh + m1.ic * M2.ch + m1.id * M2.dh + m1.ie * M2.eh + m1.if * M2.fh + m1.ig * M2.gh + m1.ih * M2.hh + m1.ii * M2.ih + m1.ij * M2.jh:    .ii = m1.ia * M2.ai + m1.ib * M2.bi + m1.ic * M2.ci + m1.id * M2.di + m1.ie * M2.ei + m1.if * M2.fi + m1.ig * M2.gi + m1.ih * M2.hi + m1.ii * M2.ii + m1.ij * M2.ji:    .ij = m1.ia * M2.aj + m1.ib * M2.bj + m1.ic * M2.cj + m1.id * M2.dj + m1.ie * M2.ej + m1.if * M2.fj + m1.ig * M2.gj + m1.ih * M2.hj + m1.ii * M2.ij + m1.ij * M2.jj
        .jf = m1.ja * M2.af + m1.jb * M2.bf + m1.jc * M2.cf + m1.jd * M2.df + m1.je * M2.ef + m1.jf * M2.ff + m1.jg * M2.gf + m1.jh * M2.hf + m1.ji * M2.if + m1.jj * M2.jf:    .jg = m1.ja * M2.ag + m1.jb * M2.bg + m1.jc * M2.cg + m1.jd * M2.dg + m1.je * M2.eg + m1.jf * M2.fg + m1.jg * M2.gg + m1.jh * M2.hg + m1.ji * M2.ig + m1.jj * M2.jg:    .jh = m1.ja * M2.ah + m1.jb * M2.bh + m1.jc * M2.ch + m1.jd * M2.dh + m1.je * M2.eh + m1.jf * M2.fh + m1.jg * M2.gh + m1.jh * M2.hh + m1.ji * M2.ih + m1.jj * M2.jh:    .ji = m1.ja * M2.ai + m1.jb * M2.bi + m1.jc * M2.ci + m1.jd * M2.di + m1.je * M2.ei + m1.jf * M2.fi + m1.jg * M2.gi + m1.jh * M2.hi + m1.ji * M2.ii + m1.jj * M2.ji:    .jj = m1.ja * M2.aj + m1.jb * M2.bj + m1.jc * M2.cj + m1.jd * M2.dj + m1.je * M2.ej + m1.jf * M2.fj + m1.jg * M2.gj + m1.jh * M2.hj + m1.ji * M2.ij + m1.jj * M2.jj
    End With
End Function

'Multiplikation Matrix mit Vektor
Public Function Mat2_vmul(m As Matrix2, v As Vector2) As Vector2
    'Multipliziert eine 2x2 Matrix m mit einem 2er-Vektor
    With Mat2_vmul:   .a = m.aa * v.a + m.ab * v.b
                      .b = m.ba * v.a + m.bb * v.b
    End With
End Function
Public Function Mat3_vmul(m As Matrix3, v As Vector3) As Vector3
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat3_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c
    End With
End Function
Public Function Mat4_vmul(m As Matrix4, v As Vector4) As Vector4
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat4_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d
                      .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d
    End With
End Function
Public Function Mat5_vmul(m As Matrix5, v As Vector5) As Vector5
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat5_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e
                      .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e
                      .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e
    End With
End Function
Public Function Mat6_vmul(m As Matrix6, v As Vector6) As Vector6
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat6_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f
                      .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f
                      .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f
                      .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f
    End With
End Function
Public Function Mat7_vmul(m As Matrix7, v As Vector7) As Vector7
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat7_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g
                      .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g
                      .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g
                      .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g
                      .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g
    End With
End Function
Public Function Mat8_vmul(m As Matrix8, v As Vector8) As Vector8
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat8_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g + m.ah * v.H
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g + m.bh * v.H
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g + m.ch * v.H
                      .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g + m.dh * v.H
                      .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g + m.eh * v.H
                      .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g + m.fh * v.H
                      .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g + m.gh * v.H
                      .H = m.ha * v.a + m.hb * v.b + m.HC * v.c + m.hd * v.d + m.he * v.e + m.hf * v.f + m.hg * v.g + m.hh * v.H
    End With
End Function
Public Function Mat9_vmul(m As Matrix9, v As Vector9) As Vector9
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat9_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g + m.ah * v.H + m.ai * v.i
                      .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g + m.bh * v.H + m.bi * v.i
                      .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g + m.ch * v.H + m.ci * v.i
                      .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g + m.dh * v.H + m.di * v.i
                      .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g + m.eh * v.H + m.ei * v.i
                      .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g + m.fh * v.H + m.fi * v.i
                      .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g + m.gh * v.H + m.gi * v.i
                      .H = m.ha * v.a + m.hb * v.b + m.HC * v.c + m.hd * v.d + m.he * v.e + m.hf * v.f + m.hg * v.g + m.hh * v.H + m.hi * v.i
                      .i = m.ia * v.a + m.ib * v.b + m.ic * v.c + m.id * v.d + m.ie * v.e + m.if * v.f + m.ig * v.g + m.ih * v.H + m.ii * v.i
    End With
End Function
Public Function Mat10_vmul(m As Matrix10, v As Vector10) As Vector10
    'Multipliziert eine 3x3 Matrix m mit einem 3er-Vektor
    With Mat10_vmul:   .a = m.aa * v.a + m.ab * v.b + m.ac * v.c + m.ad * v.d + m.ae * v.e + m.af * v.f + m.ag * v.g + m.ah * v.H + m.ai * v.i + m.aj * v.j
                       .b = m.ba * v.a + m.bb * v.b + m.bc * v.c + m.bd * v.d + m.be * v.e + m.bf * v.f + m.bg * v.g + m.bh * v.H + m.bi * v.i + m.bj * v.j
                       .c = m.ca * v.a + m.cb * v.b + m.cc * v.c + m.cd * v.d + m.ce * v.e + m.cf * v.f + m.cg * v.g + m.ch * v.H + m.ci * v.i + m.cj * v.j
                       .d = m.da * v.a + m.db * v.b + m.dc * v.c + m.dd * v.d + m.de * v.e + m.df * v.f + m.dg * v.g + m.dh * v.H + m.di * v.i + m.dj * v.j
                       .e = m.ea * v.a + m.eb * v.b + m.ec * v.c + m.ed * v.d + m.ee * v.e + m.ef * v.f + m.eg * v.g + m.eh * v.H + m.ei * v.i + m.ej * v.j
                       .f = m.fa * v.a + m.fb * v.b + m.fc * v.c + m.fd * v.d + m.fe * v.e + m.ff * v.f + m.fg * v.g + m.fh * v.H + m.fi * v.i + m.fj * v.j
                       .g = m.ga * v.a + m.gb * v.b + m.gc * v.c + m.gd * v.d + m.ge * v.e + m.gf * v.f + m.gg * v.g + m.gh * v.H + m.gi * v.i + m.gj * v.j
                       .H = m.ha * v.a + m.hb * v.b + m.HC * v.c + m.hd * v.d + m.he * v.e + m.hf * v.f + m.hg * v.g + m.hh * v.H + m.hi * v.i + m.hj * v.j
                       .i = m.ia * v.a + m.ib * v.b + m.ic * v.c + m.id * v.d + m.ie * v.e + m.if * v.f + m.ig * v.g + m.ih * v.H + m.ii * v.i + m.ij * v.j
                       .j = m.ia * v.a + m.jb * v.b + m.jc * v.c + m.jd * v.d + m.je * v.e + m.jf * v.f + m.jg * v.g + m.jh * v.H + m.ji * v.i + m.jj * v.j
    End With
End Function

Public Function Mat2_Equilib(m As Matrix2, m_out As Matrix2) As Matrix2
    'führt eine Äquilibrierung für eine 2x2-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab
        .bb = m.ba + m.bb
    End With
    With Mat2_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum
    End With
End Function
Public Function Mat3_Equilib(m As Matrix3, m_out As Matrix3) As Matrix3
    'führt eine Äquilibrierung für eine 3x3-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac
        .bb = m.ba + m.bb + m.bc
        .cc = m.ca + m.cb + m.cc
    End With
    With Mat3_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum
    End With
End Function
Public Function Mat4_Equilib(m As Matrix4, m_out As Matrix4) As Matrix4
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad
        .bb = m.ba + m.bb + m.bc + m.bd
        .cc = m.ca + m.cb + m.cc + m.cd
        .dd = m.da + m.db + m.dc + m.dd
    End With
    With Mat4_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum
    End With
End Function
Public Function Mat5_Equilib(m As Matrix5, m_out As Matrix5) As Matrix5
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad + m.ae
        .bb = m.ba + m.bb + m.bc + m.bd + m.be
        .cc = m.ca + m.cb + m.cc + m.cd + m.ce
        .dd = m.da + m.db + m.dc + m.dd + m.de
        .ee = m.ea + m.eb + m.ec + m.ed + m.ee
    End With
    With Mat5_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum: .ae = m.ae / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum: .be = m.be / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum: .ce = m.ce / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum: .de = m.de / sum
        sum = m_out.ee: .ea = m.ea / sum: .eb = m.eb / sum: .ec = m.ec / sum: .ed = m.ed / sum: .ee = m.ee / sum
    End With
End Function
Public Function Mat6_Equilib(m As Matrix6, m_out As Matrix6) As Matrix6
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad + m.ae + m.af
        .bb = m.ba + m.bb + m.bc + m.bd + m.be + m.bf
        .cc = m.ca + m.cb + m.cc + m.cd + m.ce + m.cf
        .dd = m.da + m.db + m.dc + m.dd + m.de + m.df
        .ee = m.ea + m.eb + m.ec + m.ed + m.ee + m.ef
        .ff = m.fa + m.fb + m.fc + m.fd + m.fe + m.ff
    End With
    With Mat6_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum: .ae = m.ae / sum: .af = m.af / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum: .be = m.be / sum: .bf = m.bf / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum: .ce = m.ce / sum: .cf = m.cf / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum: .de = m.de / sum: .df = m.df / sum
        sum = m_out.ee: .ea = m.ea / sum: .eb = m.eb / sum: .ec = m.ec / sum: .ed = m.ed / sum: .ee = m.ee / sum: .ef = m.ef / sum
        sum = m_out.ff: .fa = m.fa / sum: .fb = m.fb / sum: .fc = m.fc / sum: .fd = m.fd / sum: .fe = m.fe / sum: .ff = m.ff / sum
    End With
End Function
Public Function Mat7_Equilib(m As Matrix7, m_out As Matrix7) As Matrix7
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad + m.ae + m.af + m.ag
        .bb = m.ba + m.bb + m.bc + m.bd + m.be + m.bf + m.bg
        .cc = m.ca + m.cb + m.cc + m.cd + m.ce + m.cf + m.cg
        .dd = m.da + m.db + m.dc + m.dd + m.de + m.df + m.dg
        .ee = m.ea + m.eb + m.ec + m.ed + m.ee + m.ef + m.eg
        .ff = m.fa + m.fb + m.fc + m.fd + m.fe + m.ff + m.fg
        .gg = m.ga + m.gb + m.gc + m.gd + m.ge + m.gf + m.gg
    End With
    With Mat7_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum: .ae = m.ae / sum: .af = m.af / sum: .ag = m.ag / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum: .be = m.be / sum: .bf = m.bf / sum: .bg = m.bg / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum: .ce = m.ce / sum: .cf = m.cf / sum: .cg = m.cg / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum: .de = m.de / sum: .df = m.df / sum: .dg = m.dg / sum
        sum = m_out.ee: .ea = m.ea / sum: .eb = m.eb / sum: .ec = m.ec / sum: .ed = m.ed / sum: .ee = m.ee / sum: .ef = m.ef / sum: .eg = m.eg / sum
        sum = m_out.ff: .fa = m.fa / sum: .fb = m.fb / sum: .fc = m.fc / sum: .fd = m.fd / sum: .fe = m.fe / sum: .ff = m.ff / sum: .fg = m.fg / sum
        sum = m_out.gg: .ga = m.ga / sum: .gb = m.gb / sum: .gc = m.gc / sum: .gd = m.gd / sum: .ge = m.ge / sum: .gf = m.gf / sum: .gg = m.gg / sum
    End With
End Function
Public Function Mat8_Equilib(m As Matrix8, m_out As Matrix8) As Matrix8
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad + m.ae + m.af + m.ag + m.ah
        .bb = m.ba + m.bb + m.bc + m.bd + m.be + m.bf + m.bg + m.bh
        .cc = m.ca + m.cb + m.cc + m.cd + m.ce + m.cf + m.cg + m.ch
        .dd = m.da + m.db + m.dc + m.dd + m.de + m.df + m.dg + m.dh
        .ee = m.ea + m.eb + m.ec + m.ed + m.ee + m.ef + m.eg + m.eh
        .ff = m.fa + m.fb + m.fc + m.fd + m.fe + m.ff + m.fg + m.fh
        .gg = m.ga + m.gb + m.gc + m.gd + m.ge + m.gf + m.gg + m.gh
        .hh = m.ha + m.hb + m.HC + m.hd + m.he + m.hf + m.hg + m.hh
    End With
    With Mat8_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum: .ae = m.ae / sum: .af = m.af / sum: .ag = m.ag / sum: .ah = m.ah / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum: .be = m.be / sum: .bf = m.bf / sum: .bg = m.bg / sum: .bh = m.bh / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum: .ce = m.ce / sum: .cf = m.cf / sum: .cg = m.cg / sum: .ch = m.ch / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum: .de = m.de / sum: .df = m.df / sum: .dg = m.dg / sum: .dh = m.dh / sum
        sum = m_out.ee: .ea = m.ea / sum: .eb = m.eb / sum: .ec = m.ec / sum: .ed = m.ed / sum: .ee = m.ee / sum: .ef = m.ef / sum: .eg = m.eg / sum: .eh = m.eh / sum
        sum = m_out.ff: .fa = m.fa / sum: .fb = m.fb / sum: .fc = m.fc / sum: .fd = m.fd / sum: .fe = m.fe / sum: .ff = m.ff / sum: .fg = m.fg / sum: .fh = m.fh / sum
        sum = m_out.gg: .ga = m.ga / sum: .gb = m.gb / sum: .gc = m.gc / sum: .gd = m.gd / sum: .ge = m.ge / sum: .gf = m.gf / sum: .gg = m.gg / sum: .gh = m.gh / sum
        sum = m_out.hh: .ha = m.ha / sum: .hb = m.hb / sum: .HC = m.HC / sum: .hd = m.hd / sum: .he = m.he / sum: .hf = m.hf / sum: .hg = m.hg / sum: .hh = m.hh / sum
    End With
End Function
Public Function Mat9_Equilib(m As Matrix9, m_out As Matrix9) As Matrix9
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad + m.ae + m.af + m.ag + m.ah + m.ai
        .bb = m.ba + m.bb + m.bc + m.bd + m.be + m.bf + m.bg + m.bh + m.bi
        .cc = m.ca + m.cb + m.cc + m.cd + m.ce + m.cf + m.cg + m.ch + m.ci
        .dd = m.da + m.db + m.dc + m.dd + m.de + m.df + m.dg + m.dh + m.di
        .ee = m.ea + m.eb + m.ec + m.ed + m.ee + m.ef + m.eg + m.eh + m.ei
        .ff = m.fa + m.fb + m.fc + m.fd + m.fe + m.ff + m.fg + m.fh + m.fi
        .gg = m.ga + m.gb + m.gc + m.gd + m.ge + m.gf + m.gg + m.gh + m.gi
        .hh = m.ha + m.hb + m.HC + m.hd + m.he + m.hf + m.hg + m.hh + m.hi
        .ii = m.ia + m.ib + m.ic + m.id + m.ie + m.if + m.ig + m.ih + m.ii
    End With
    With Mat9_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum: .ae = m.ae / sum: .af = m.af / sum: .ag = m.ag / sum: .ah = m.ah / sum: .ai = m.ai / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum: .be = m.be / sum: .bf = m.bf / sum: .bg = m.bg / sum: .bh = m.bh / sum: .bi = m.bi / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum: .ce = m.ce / sum: .cf = m.cf / sum: .cg = m.cg / sum: .ch = m.ch / sum: .ci = m.ci / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum: .de = m.de / sum: .df = m.df / sum: .dg = m.dg / sum: .dh = m.dh / sum: .di = m.di / sum
        sum = m_out.ee: .ea = m.ea / sum: .eb = m.eb / sum: .ec = m.ec / sum: .ed = m.ed / sum: .ee = m.ee / sum: .ef = m.ef / sum: .eg = m.eg / sum: .eh = m.eh / sum: .ei = m.ei / sum
        sum = m_out.ff: .fa = m.fa / sum: .fb = m.fb / sum: .fc = m.fc / sum: .fd = m.fd / sum: .fe = m.fe / sum: .ff = m.ff / sum: .fg = m.fg / sum: .fh = m.fh / sum: .fi = m.fi / sum
        sum = m_out.gg: .ga = m.ga / sum: .gb = m.gb / sum: .gc = m.gc / sum: .gd = m.gd / sum: .ge = m.ge / sum: .gf = m.gf / sum: .gg = m.gg / sum: .gh = m.gh / sum: .gi = m.gi / sum
        sum = m_out.hh: .ha = m.ha / sum: .hb = m.hb / sum: .HC = m.HC / sum: .hd = m.hd / sum: .he = m.he / sum: .hf = m.hf / sum: .hg = m.hg / sum: .hh = m.hh / sum: .hi = m.hi / sum
        sum = m_out.ii: .ia = m.ia / sum: .ib = m.ib / sum: .ic = m.ic / sum: .id = m.id / sum: .ie = m.ie / sum: .if = m.if / sum: .ig = m.ig / sum: .ih = m.ih / sum: .ii = m.ii / sum
    End With
End Function
Public Function Mat10_Equilib(m As Matrix10, m_out As Matrix10) As Matrix10
    'führt eine Äquilibrierung für eine 4x4-Matrix durch
    Dim sum As Double
    With m_out
        .aa = m.aa + m.ab + m.ac + m.ad + m.ae + m.af + m.ag + m.ah + m.ai + m.aj
        .bb = m.ba + m.bb + m.bc + m.bd + m.be + m.bf + m.bg + m.bh + m.bi + m.bj
        .cc = m.ca + m.cb + m.cc + m.cd + m.ce + m.cf + m.cg + m.ch + m.ci + m.cj
        .dd = m.da + m.db + m.dc + m.dd + m.de + m.df + m.dg + m.dh + m.di + m.dj
        .ee = m.ea + m.eb + m.ec + m.ed + m.ee + m.ef + m.eg + m.eh + m.ei + m.ej
        .ff = m.fa + m.fb + m.fc + m.fd + m.fe + m.ff + m.fg + m.fh + m.fi + m.fj
        .gg = m.ga + m.gb + m.gc + m.gd + m.ge + m.gf + m.gg + m.gh + m.gi + m.gj
        .hh = m.ha + m.hb + m.HC + m.hd + m.he + m.hf + m.hg + m.hh + m.hi + m.hj
        .ii = m.ia + m.ib + m.ic + m.id + m.ie + m.if + m.ig + m.ih + m.ii + m.ij
        .jj = m.ja + m.jb + m.jc + m.jd + m.je + m.jf + m.jg + m.jh + m.ji + m.jj
    End With
    With Mat10_Equilib
        sum = m_out.aa: .aa = m.aa / sum: .ab = m.ab / sum: .ac = m.ac / sum: .ad = m.ad / sum: .ae = m.ae / sum: .af = m.af / sum: .ag = m.ag / sum: .ah = m.ah / sum: .ai = m.ai / sum: .aj = m.aj / sum
        sum = m_out.bb: .ba = m.ba / sum: .bb = m.bb / sum: .bc = m.bc / sum: .bd = m.bd / sum: .be = m.be / sum: .bf = m.bf / sum: .bg = m.bg / sum: .bh = m.bh / sum: .bi = m.bi / sum: .bj = m.bj / sum
        sum = m_out.cc: .ca = m.ca / sum: .cb = m.cb / sum: .cc = m.cc / sum: .cd = m.cd / sum: .ce = m.ce / sum: .cf = m.cf / sum: .cg = m.cg / sum: .ch = m.ch / sum: .ci = m.ci / sum: .cj = m.cj / sum
        sum = m_out.dd: .da = m.da / sum: .db = m.db / sum: .dc = m.dc / sum: .dd = m.dd / sum: .de = m.de / sum: .df = m.df / sum: .dg = m.dg / sum: .dh = m.dh / sum: .di = m.di / sum: .dj = m.dj / sum
        sum = m_out.ee: .ea = m.ea / sum: .eb = m.eb / sum: .ec = m.ec / sum: .ed = m.ed / sum: .ee = m.ee / sum: .ef = m.ef / sum: .eg = m.eg / sum: .eh = m.eh / sum: .ei = m.ei / sum: .ej = m.ej / sum
        sum = m_out.ff: .fa = m.fa / sum: .fb = m.fb / sum: .fc = m.fc / sum: .fd = m.fd / sum: .fe = m.fe / sum: .ff = m.ff / sum: .fg = m.fg / sum: .fh = m.fh / sum: .fi = m.fi / sum: .fj = m.fj / sum
        sum = m_out.gg: .ga = m.ga / sum: .gb = m.gb / sum: .gc = m.gc / sum: .gd = m.gd / sum: .ge = m.ge / sum: .gf = m.gf / sum: .gg = m.gg / sum: .gh = m.gh / sum: .gi = m.gi / sum: .gj = m.gj / sum
        sum = m_out.hh: .ha = m.ha / sum: .hb = m.hb / sum: .HC = m.HC / sum: .hd = m.hd / sum: .he = m.he / sum: .hf = m.hf / sum: .hg = m.hg / sum: .hh = m.hh / sum: .hi = m.hi / sum: .hj = m.hj / sum
        sum = m_out.ii: .ia = m.ia / sum: .ib = m.ib / sum: .ic = m.ic / sum: .id = m.id / sum: .ie = m.ie / sum: .if = m.if / sum: .ig = m.ig / sum: .ih = m.ih / sum: .ii = m.ii / sum: .ij = m.ij / sum
        sum = m_out.jj: .ja = m.ja / sum: .jb = m.jb / sum: .jc = m.jc / sum: .jd = m.jd / sum: .je = m.je / sum: .jf = m.jf / sum: .jg = m.jg / sum: .jh = m.jh / sum: .ji = m.ji / sum: .jj = m.jj / sum
    End With
End Function


'Matrizenvergleich
Public Function Mat2_IsEqual(m1 As Matrix2, M2 As Matrix2) As Boolean
    'Liefert True wenn beide 2x2-Matrizen gleich sind
    Mat2_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb)
End Function
Public Function Mat3_IsEqual(m1 As Matrix3, M2 As Matrix3) As Boolean
    'Liefert True wenn beide 3x3-Matrizen gleich sind
    Mat3_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc)
End Function
Public Function Mat4_IsEqual(m1 As Matrix4, M2 As Matrix4) As Boolean
    'Liefert True wenn beide 4x4-Matrizen gleich sind
    Mat4_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd) And _
                   (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd)
End Function
Public Function Mat5_IsEqual(m1 As Matrix5, M2 As Matrix5) As Boolean
    'Liefert True wenn beide 5x5-Matrizen gleich sind
    Mat5_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad And m1.ae = M2.ae) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd And m1.be = M2.be) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd And m1.ce = M2.ce) And _
                   (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd And m1.de = M2.de) And _
                   (m1.ea = M2.ea And m1.eb = M2.eb And m1.ec = M2.ec And m1.ed = M2.ed And m1.ee = M2.ee)
End Function
Public Function Mat6_IsEqual(m1 As Matrix6, M2 As Matrix6) As Boolean
    'Liefert True wenn beide 6x6-Matrizen gleich sind
    Mat6_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad And m1.ae = M2.ae And m1.af = M2.af) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd And m1.be = M2.be And m1.bf = M2.bf) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd And m1.ce = M2.ce And m1.cf = M2.cf) And _
                   (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd And m1.de = M2.de And m1.df = M2.df) And _
                   (m1.ea = M2.ea And m1.eb = M2.eb And m1.ec = M2.ec And m1.ed = M2.ed And m1.ee = M2.ee And m1.ef = M2.ef) And _
                   (m1.fa = M2.fa And m1.fb = M2.fb And m1.fc = M2.fc And m1.fd = M2.fd And m1.fe = M2.fe And m1.ff = M2.ff)
End Function
Public Function Mat7_IsEqual(m1 As Matrix7, M2 As Matrix7) As Boolean
    'Liefert True wenn beide 7x7-Matrizen gleich sind
    Mat7_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad And m1.ae = M2.ae And m1.af = M2.af And m1.ag = M2.ag) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd And m1.be = M2.be And m1.bf = M2.bf And m1.bg = M2.bg) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd And m1.ce = M2.ce And m1.cf = M2.cf And m1.cg = M2.cg) And _
                   (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd And m1.de = M2.de And m1.df = M2.df And m1.dg = M2.dg) And _
                   (m1.ea = M2.ea And m1.eb = M2.eb And m1.ec = M2.ec And m1.ed = M2.ed And m1.ee = M2.ee And m1.ef = M2.ef And m1.eg = M2.eg) And _
                   (m1.fa = M2.fa And m1.fb = M2.fb And m1.fc = M2.fc And m1.fd = M2.fd And m1.fe = M2.fe And m1.ff = M2.ff And m1.fg = M2.fg) And _
                   (m1.fa = M2.ga And m1.gb = M2.gb And m1.gc = M2.gc And m1.gd = M2.gd And m1.ge = M2.ge And m1.gf = M2.gf And m1.gg = M2.gg)
End Function
Public Function Mat8_IsEqual(m1 As Matrix8, M2 As Matrix8) As Boolean
    'Liefert True wenn beide 8x8-Matrizen gleich sind
    Mat8_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad And m1.ae = M2.ae And m1.af = M2.af And m1.ag = M2.ag And m1.ah = M2.ah) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd And m1.be = M2.be And m1.bf = M2.bf And m1.bg = M2.bg And m1.bh = M2.bh) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd And m1.ce = M2.ce And m1.cf = M2.cf And m1.cg = M2.cg And m1.ch = M2.ch) And _
                   (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd And m1.de = M2.de And m1.df = M2.df And m1.dg = M2.dg And m1.dh = M2.dh) And _
                   (m1.ea = M2.ea And m1.eb = M2.eb And m1.ec = M2.ec And m1.ed = M2.ed And m1.ee = M2.ee And m1.ef = M2.ef And m1.eg = M2.eg And m1.eh = M2.eh) And _
                   (m1.fa = M2.fa And m1.fb = M2.fb And m1.fc = M2.fc And m1.fd = M2.fd And m1.fe = M2.fe And m1.ff = M2.ff And m1.fg = M2.fg And m1.fh = M2.fh) And _
                   (m1.fa = M2.ga And m1.gb = M2.gb And m1.gc = M2.gc And m1.gd = M2.gd And m1.ge = M2.ge And m1.gf = M2.gf And m1.gg = M2.gg And m1.gh = M2.gh) And _
                   (m1.ha = M2.ha And m1.hb = M2.hb And m1.HC = M2.HC And m1.hd = M2.hd And m1.he = M2.he And m1.hf = M2.hf And m1.hg = M2.hg And m1.hh = M2.hh)
End Function
Public Function Mat9_IsEqual(m1 As Matrix9, M2 As Matrix9) As Boolean
    'Liefert True wenn beide 9x9-Matrizen gleich sind
    Mat9_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad And m1.ae = M2.ae And m1.af = M2.af And m1.ag = M2.ag And m1.ah = M2.ah And m1.ai = M2.ai) And _
                   (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd And m1.be = M2.be And m1.bf = M2.bf And m1.bg = M2.bg And m1.bh = M2.bh And m1.bi = M2.bi) And _
                   (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd And m1.ce = M2.ce And m1.cf = M2.cf And m1.cg = M2.cg And m1.ch = M2.ch And m1.ci = M2.ci) And _
                   (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd And m1.de = M2.de And m1.df = M2.df And m1.dg = M2.dg And m1.dh = M2.dh And m1.di = M2.di) And _
                   (m1.ea = M2.ea And m1.eb = M2.eb And m1.ec = M2.ec And m1.ed = M2.ed And m1.ee = M2.ee And m1.ef = M2.ef And m1.eg = M2.eg And m1.eh = M2.eh And m1.ei = M2.ei) And _
                   (m1.fa = M2.fa And m1.fb = M2.fb And m1.fc = M2.fc And m1.fd = M2.fd And m1.fe = M2.fe And m1.ff = M2.ff And m1.fg = M2.fg And m1.fh = M2.fh And m1.fi = M2.fi) And _
                   (m1.ga = M2.ga And m1.gb = M2.gb And m1.gc = M2.gc And m1.gd = M2.gd And m1.ge = M2.ge And m1.gf = M2.gf And m1.gg = M2.gg And m1.gh = M2.gh And m1.gi = M2.gi) And _
                   (m1.ha = M2.ha And m1.hb = M2.hb And m1.HC = M2.HC And m1.hd = M2.hd And m1.he = M2.he And m1.hf = M2.hf And m1.hg = M2.hg And m1.hh = M2.hh And m1.hi = M2.hi) And _
                   (m1.ia = M2.ia And m1.ib = M2.ib And m1.ic = M2.ic And m1.id = M2.id And m1.ie = M2.ie And m1.if = M2.if And m1.ig = M2.ig And m1.ih = M2.ih And m1.ii = M2.ii)
End Function
Public Function Mat10_IsEqual(m1 As Matrix10, M2 As Matrix10) As Boolean
    'Liefert True wenn beide 10x10-Matrizen gleich sind
    Mat10_IsEqual = (m1.aa = M2.aa And m1.ab = M2.ab And m1.ac = M2.ac And m1.ad = M2.ad And m1.ae = M2.ae And m1.af = M2.af And m1.ag = M2.ag And m1.ah = M2.ah And m1.ai = M2.ai And m1.aj = M2.aj) And _
                    (m1.ba = M2.ba And m1.bb = M2.bb And m1.bc = M2.bc And m1.bd = M2.bd And m1.be = M2.be And m1.bf = M2.bf And m1.bg = M2.bg And m1.bh = M2.bh And m1.bi = M2.bi And m1.bj = M2.bj) And _
                    (m1.ca = M2.ca And m1.cb = M2.cb And m1.cc = M2.cc And m1.cd = M2.cd And m1.ce = M2.ce And m1.cf = M2.cf And m1.cg = M2.cg And m1.ch = M2.ch And m1.ci = M2.ci And m1.cj = M2.cj) And _
                    (m1.da = M2.da And m1.db = M2.db And m1.dc = M2.dc And m1.dd = M2.dd And m1.de = M2.de And m1.df = M2.df And m1.dg = M2.dg And m1.dh = M2.dh And m1.di = M2.di And m1.dj = M2.dj) And _
                    (m1.ea = M2.ea And m1.eb = M2.eb And m1.ec = M2.ec And m1.ed = M2.ed And m1.ee = M2.ee And m1.ef = M2.ef And m1.eg = M2.eg And m1.eh = M2.eh And m1.ei = M2.ei And m1.ej = M2.ej) And _
                    (m1.fa = M2.fa And m1.fb = M2.fb And m1.fc = M2.fc And m1.fd = M2.fd And m1.fe = M2.fe And m1.ff = M2.ff And m1.fg = M2.fg And m1.fh = M2.fh And m1.fi = M2.fi And m1.fj = M2.fj) And _
                    (m1.ga = M2.ga And m1.gb = M2.gb And m1.gc = M2.gc And m1.gd = M2.gd And m1.ge = M2.ge And m1.gf = M2.gf And m1.gg = M2.gg And m1.gh = M2.gh And m1.gi = M2.gi And m1.gj = M2.gj) And _
                    (m1.ha = M2.ha And m1.hb = M2.hb And m1.HC = M2.HC And m1.hd = M2.hd And m1.he = M2.he And m1.hf = M2.hf And m1.hg = M2.hg And m1.hh = M2.hh And m1.hi = M2.hi And m1.hj = M2.hj) And _
                    (m1.ia = M2.ia And m1.ib = M2.ib And m1.ic = M2.ic And m1.id = M2.id And m1.ie = M2.ie And m1.if = M2.if And m1.ig = M2.ig And m1.ih = M2.ih And m1.ii = M2.ii And m1.ij = M2.ij) And _
                    (m1.ja = M2.ja And m1.jb = M2.jb And m1.jc = M2.jc And m1.jd = M2.jd And m1.je = M2.je And m1.jf = M2.jf And m1.jg = M2.jg And m1.jh = M2.jh And m1.ji = M2.ji And m1.jj = M2.jj)
End Function

'Transponierte Matrix
Public Function Mat2_tra(m As Matrix2) As Matrix2
    'Erzeugt die Transponierte aus einer 2x2-Matrix
    With Mat2_tra:   .aa = m.aa:    .ab = m.ba
                     .ba = m.ab:    .bb = m.bb
    End With
End Function
Public Function Mat3_tra(m As Matrix3) As Matrix3
    'Erzeugt die Transponierte aus einer 3x3-Matrix
    With Mat3_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc
    End With
End Function
Public Function Mat4_tra(m As Matrix4) As Matrix4
    'Erzeugt die Transponierte aus einer 4x4-Matrix
    With Mat4_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc
                     .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd
    End With
End Function
Public Function Mat5_tra(m As Matrix5) As Matrix5
    'Erzeugt die Transponierte aus einer 5x5-Matrix
    With Mat5_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec
                     .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed
                     .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee
    End With
End Function
Public Function Mat6_tra(m As Matrix6) As Matrix6
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Mat6_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc
                     .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd
                     .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe
                     .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff
    End With
End Function
Public Function Mat7_tra(m As Matrix7) As Matrix7
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Mat7_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc
                     .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd
                     .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge
                     .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf
                     .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg
    End With
End Function
Public Function Mat8_tra(m As Matrix8) As Matrix8
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Mat8_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga:    .ah = m.ha
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb:    .bh = m.hb
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc:    .ch = m.HC
                     .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd:    .dh = m.hd
                     .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge:    .eh = m.he
                     .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf:    .fh = m.hf
                     .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg:    .gh = m.hg
                     .ha = m.ah:    .hb = m.bh:    .HC = m.ch:    .hd = m.dh:    .he = m.eh:    .hf = m.fh:    .hg = m.gh:    .hh = m.hh
    End With
End Function
Public Function Mat9_tra(m As Matrix9) As Matrix9
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Mat9_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga:    .ah = m.ha:    .ai = m.ia
                     .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb:    .bh = m.hb:    .bi = m.ib
                     .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc:    .ch = m.HC:    .ci = m.ic
                     .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd:    .dh = m.hd:    .di = m.id
                     .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge:    .eh = m.he:    .ei = m.ie
                     .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf:    .fh = m.hf:    .fi = m.if
                     .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg:    .gh = m.hg:    .gi = m.ig
                     .ha = m.ah:    .hb = m.bh:    .HC = m.ch:    .hd = m.dh:    .he = m.eh:    .hf = m.fh:    .hg = m.gh:    .hh = m.hh:    .hi = m.ih
                     .ia = m.ai:    .ib = m.bi:    .ic = m.ci:    .id = m.di:    .ie = m.ei:    .if = m.fi:    .ig = m.gi:    .ih = m.hi:    .ii = m.ii
    End With
End Function
Public Function Mat10_tra(m As Matrix10) As Matrix10
    'Erzeugt die Transponierte aus einer 6x6-Matrix
    With Mat10_tra:   .aa = m.aa:    .ab = m.ba:    .ac = m.ca:    .ad = m.da:    .ae = m.ea:    .af = m.fa:    .ag = m.ga:    .ah = m.ha:    .ai = m.ia:    .aj = m.ja
                      .ba = m.ab:    .bb = m.bb:    .bc = m.cb:    .bd = m.db:    .be = m.eb:    .bf = m.fb:    .bg = m.gb:    .bh = m.hb:    .bi = m.ib:    .bj = m.jb
                      .ca = m.ac:    .cb = m.bc:    .cc = m.cc:    .cd = m.dc:    .ce = m.ec:    .cf = m.fc:    .cg = m.gc:    .ch = m.HC:    .ci = m.ic:    .cj = m.jc
                      .da = m.ad:    .db = m.bd:    .dc = m.cd:    .dd = m.dd:    .de = m.ed:    .df = m.fd:    .dg = m.gd:    .dh = m.hd:    .di = m.id:    .dj = m.jd
                      .ea = m.ae:    .eb = m.be:    .ec = m.ce:    .ed = m.de:    .ee = m.ee:    .ef = m.fe:    .eg = m.ge:    .eh = m.he:    .ei = m.ie:    .ej = m.je
                      .fa = m.af:    .fb = m.bf:    .fc = m.cf:    .fd = m.df:    .fe = m.ef:    .ff = m.ff:    .fg = m.gf:    .fh = m.hf:    .fi = m.if:    .fj = m.jf
                      .ga = m.ag:    .gb = m.bg:    .gc = m.cg:    .gd = m.dg:    .ge = m.eg:    .gf = m.fg:    .gg = m.gg:    .gh = m.hg:    .gi = m.ig:    .gj = m.jg
                      .ha = m.ah:    .hb = m.bh:    .HC = m.ch:    .hd = m.dh:    .he = m.eh:    .hf = m.fh:    .hg = m.gh:    .hh = m.hh:    .hi = m.ih:    .hj = m.jh
                      .ia = m.ai:    .ib = m.bi:    .ic = m.ci:    .id = m.di:    .ie = m.ei:    .if = m.fi:    .ig = m.gi:    .ih = m.hi:    .ii = m.ii:    .ij = m.ji
                      .ja = m.aj:    .jb = m.bj:    .jc = m.cj:    .jd = m.dj:    .je = m.ej:    .jf = m.fj:    .jg = m.gj:    .jh = m.hj:    .ji = m.ij:    .jj = m.jj
    End With
End Function

'Lesen und Schreiben
'allgemein
'Public Sub Matrix_Parse(t As String, ByVal mRows As Long, ByVal nCols As Long, ByRef a_out() As Double)
Public Function Matrix_Parse(T As String, ByVal mRows As Long, ByVal nCols As Long) As Double()
    ReDim a_out(0 To nCols - 1, 0 To mRows - 1) As Double
    Dim saLines() As String: saLines = Split(DeleteMultiWS(T), vbCrLf)
    Dim sa() As String
    Dim i As Long, j As Long
    For i = 0 To mRows - 1
        If UBound(saLines) < i Then Exit For
        sa = Split(DeleteMultiWS(saLines(i)), " ")
        For j = 0 To nCols - 1
            If UBound(sa) < j Then Exit For
            'a_out(i, j) = DblParse(sa(j))
            a_out(j, i) = DblParse(sa(j)) 'orig ?
        Next
    Next
    Matrix_Parse = a_out
End Function
Public Function Mat2_Parse(T As String) As Matrix2
    Dim mRows As Long: mRows = 2
    Dim nCols As Long: nCols = 2
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat2_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat3_Parse(T As String) As Matrix3
    Dim mRows As Long: mRows = 3
    Dim nCols As Long: nCols = 3
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat3_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat4_Parse(T As String) As Matrix4
    Dim mRows As Long: mRows = 4
    Dim nCols As Long: nCols = 4
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat4_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat5_Parse(T As String) As Matrix5
    Dim mRows As Long: mRows = 5
    Dim nCols As Long: nCols = 5
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat5_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat6_Parse(T As String) As Matrix6
    Dim mRows As Long: mRows = 6
    Dim nCols As Long: nCols = 6
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat6_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat7_Parse(T As String) As Matrix7
    Dim mRows As Long: mRows = 7
    Dim nCols As Long: nCols = 7
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat7_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat8_Parse(T As String) As Matrix8
    Dim mRows As Long: mRows = 8
    Dim nCols As Long: nCols = 8
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat8_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat9_Parse(T As String) As Matrix9
    Dim mRows As Long: mRows = 9
    Dim nCols As Long: nCols = 9
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat9_Parse, a(0, 0), mRows * nCols * 8
End Function
Public Function Mat10_Parse(T As String) As Matrix10
    Dim mRows As Long: mRows = 10
    Dim nCols As Long: nCols = 10
    Dim a() As Double: a = Matrix_Parse(T, mRows, nCols)
    RtlMoveMemory Mat10_Parse, a(0, 0), mRows * nCols * 8
End Function

'Umwandeln in 2d-Array
Public Function Mat2_ToArr(m As Matrix2) As Double()
    Dim d(0 To 1, 0 To 1) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab
        d(1, 0) = .ba: d(1, 1) = .bb
    End With
    Mat2_ToArr = d
End Function
Public Function Mat3_ToArr(m As Matrix3) As Double()
    Dim d(0 To 2, 0 To 2) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc
    End With
    Mat3_ToArr = d
End Function
Public Function Mat4_ToArr(m As Matrix4) As Double()
    Dim d(0 To 3, 0 To 3) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd
    End With
    Mat4_ToArr = d
End Function
Public Function Mat5_ToArr(m As Matrix5) As Double()
    Dim d(0 To 4, 0 To 4) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee
    End With
    Mat5_ToArr = d
End Function
Public Function Mat6_ToArr(m As Matrix6) As Double()
    Dim d(0 To 5, 0 To 5) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff
    End With
    Mat6_ToArr = d
End Function
Public Function Mat7_ToArr(m As Matrix7) As Double()
    Dim d(0 To 6, 0 To 6) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg
    End With
    Mat7_ToArr = d
End Function
Public Function Mat8_ToArr(m As Matrix8) As Double()
    Dim d(0 To 7, 0 To 7) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag: d(0, 7) = .ah
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg: d(1, 7) = .bh
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg: d(2, 7) = .ch
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg: d(3, 7) = .dh
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg: d(4, 7) = .eh
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg: d(5, 7) = .fh
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg: d(6, 7) = .gh
        d(7, 0) = .ha: d(7, 1) = .hb: d(7, 2) = .HC: d(7, 3) = .hd: d(7, 4) = .he: d(7, 5) = .hf: d(7, 6) = .hg: d(7, 7) = .hh
    End With
    Mat8_ToArr = d
End Function
Public Function Mat9_ToArr(m As Matrix9) As Double()
    Dim d(0 To 8, 0 To 8) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag: d(0, 7) = .ah: d(0, 8) = .ai
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg: d(1, 7) = .bh: d(1, 8) = .bi
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg: d(2, 7) = .ch: d(2, 8) = .ci
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg: d(3, 7) = .dh: d(3, 8) = .di
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg: d(4, 7) = .eh: d(4, 8) = .ei
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg: d(5, 7) = .fh: d(5, 8) = .fi
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg: d(6, 7) = .gh: d(6, 8) = .gi
        d(7, 0) = .ha: d(7, 1) = .hb: d(7, 2) = .HC: d(7, 3) = .hd: d(7, 4) = .he: d(7, 5) = .hf: d(7, 6) = .hg: d(7, 7) = .hh: d(7, 8) = .hi
        d(8, 0) = .ia: d(8, 1) = .ib: d(8, 2) = .ic: d(8, 3) = .id: d(8, 4) = .ie: d(8, 5) = .if: d(8, 6) = .ig: d(8, 7) = .ih: d(8, 8) = .ii
    End With
    Mat9_ToArr = d
End Function
Public Function Mat10_ToArr(m As Matrix10) As Double()
    Dim d(0 To 9, 0 To 9) As Double
    With m
        d(0, 0) = .aa: d(0, 1) = .ab: d(0, 2) = .ac: d(0, 3) = .ad: d(0, 4) = .ae: d(0, 5) = .af: d(0, 6) = .ag: d(0, 7) = .ah: d(0, 8) = .ai: d(0, 9) = .aj
        d(1, 0) = .ba: d(1, 1) = .bb: d(1, 2) = .bc: d(1, 3) = .bd: d(1, 4) = .be: d(1, 5) = .bf: d(1, 6) = .bg: d(1, 7) = .bh: d(1, 8) = .bi: d(1, 9) = .bj
        d(2, 0) = .ca: d(2, 1) = .cb: d(2, 2) = .cc: d(2, 3) = .cd: d(2, 4) = .ce: d(2, 5) = .cf: d(2, 6) = .cg: d(2, 7) = .ch: d(2, 8) = .ci: d(2, 9) = .cj
        d(3, 0) = .da: d(3, 1) = .db: d(3, 2) = .dc: d(3, 3) = .dd: d(3, 4) = .de: d(3, 5) = .df: d(3, 6) = .dg: d(3, 7) = .dh: d(3, 8) = .di: d(3, 9) = .dj
        d(4, 0) = .ea: d(4, 1) = .eb: d(4, 2) = .ec: d(4, 3) = .ed: d(4, 4) = .ee: d(4, 5) = .ef: d(4, 6) = .eg: d(4, 7) = .eh: d(4, 8) = .ei: d(4, 9) = .ej
        d(5, 0) = .fa: d(5, 1) = .fb: d(5, 2) = .fc: d(5, 3) = .fd: d(5, 4) = .fe: d(5, 5) = .ff: d(5, 6) = .fg: d(5, 7) = .fh: d(5, 8) = .fi: d(5, 9) = .fj
        d(6, 0) = .ga: d(6, 1) = .gb: d(6, 2) = .gc: d(6, 3) = .gd: d(6, 4) = .ge: d(6, 5) = .gf: d(6, 6) = .gg: d(6, 7) = .gh: d(6, 8) = .gi: d(6, 9) = .gj
        d(7, 0) = .ha: d(7, 1) = .hb: d(7, 2) = .HC: d(7, 3) = .hd: d(7, 4) = .he: d(7, 5) = .hf: d(7, 6) = .hg: d(7, 7) = .hh: d(7, 8) = .hi: d(7, 9) = .hj
        d(8, 0) = .ia: d(8, 1) = .ib: d(8, 2) = .ic: d(8, 3) = .id: d(8, 4) = .ie: d(8, 5) = .if: d(8, 6) = .ig: d(8, 7) = .ih: d(8, 8) = .ii: d(8, 9) = .ij
        d(9, 0) = .ja: d(9, 1) = .jb: d(9, 2) = .jc: d(9, 3) = .jd: d(9, 4) = .je: d(9, 5) = .jf: d(9, 6) = .jg: d(9, 7) = .jh: d(9, 8) = .ji: d(9, 9) = .jj
    End With
    Mat10_ToArr = d
End Function


'Ausgabefunktionen
Public Function Mat2_ToStr(m As Matrix2, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat2_ToStr = MatrixA_ToStr(Mat2_ToArr(m), Min(mRows, 2), Min(nCols, 2), dFormat)
    'Mat2_ToStr = Matrix_ToStr(VarPtr(m), 2, 2)
End Function
Public Function Mat3_ToStr(m As Matrix3, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat3_ToStr = MatrixA_ToStr(Mat3_ToArr(m), Min(mRows, 3), Min(nCols, 3), dFormat)
    'Mat3_ToStr = Matrix_ToStr(VarPtr(m), 3, 3)
End Function
Public Function Mat4_ToStr(m As Matrix4, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat4_ToStr = MatrixA_ToStr(Mat4_ToArr(m), Min(mRows, 4), Min(nCols, 4), dFormat)
    'Mat4_ToStr = Matrix_ToStr(VarPtr(m), 4, 4)
End Function
Public Function Mat5_ToStr(m As Matrix5, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat5_ToStr = MatrixA_ToStr(Mat5_ToArr(m), Min(mRows, 5), Min(nCols, 5), dFormat)
    'Mat5_ToStr = Matrix_ToStr(VarPtr(m), 5, 5)
End Function
Public Function Mat6_ToStr(m As Matrix6, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat6_ToStr = MatrixA_ToStr(Mat6_ToArr(m), Min(mRows, 6), Min(nCols, 6), dFormat)
    'Mat6_ToStr = Matrix_ToStr(VarPtr(m), 6, 6)
End Function
Public Function Mat7_ToStr(m As Matrix7, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat7_ToStr = MatrixA_ToStr(Mat7_ToArr(m), Min(mRows, 7), Min(nCols, 7), dFormat)
    'Mat7_ToStr = Matrix_ToStr(VarPtr(m), 7, 7)
End Function
Public Function Mat8_ToStr(m As Matrix8, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat8_ToStr = MatrixA_ToStr(Mat8_ToArr(m), Min(mRows, 8), Min(nCols, 8), dFormat)
    'Mat8_ToStr = Matrix_ToStr(VarPtr(m), 8, 8)
End Function
Public Function Mat9_ToStr(m As Matrix9, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat9_ToStr = MatrixA_ToStr(Mat9_ToArr(m), Min(mRows, 9), Min(nCols, 9), dFormat)
    'Mat9_ToStr = Matrix_ToStr(VarPtr(m), 9, 9)
End Function
Public Function Mat10_ToStr(m As Matrix10, ByVal mRows As Byte, ByVal nCols As Byte, Optional ByVal dFormat As Integer = -1) As String
    Mat10_ToStr = MatrixA_ToStr(Mat10_ToArr(m), Min(mRows, 10), Min(nCols, 10), dFormat)
    'Mat10_ToStr = Matrix_ToStr(VarPtr(m), 10, 10)
End Function

Public Function MatrixA_ToStr(a() As Double, ByVal mRows As Long, ByVal nCols As Long, Optional dFormat As Integer = -1) As String

    Dim s As String ': s = ""
    Dim sl As String, vs As String, vsa() As String
    If mRows = 0 Then Exit Function
    ReDim msa(0 To mRows - 1) As String
    Dim i As Long, j As Long
    ReDim ca(0 To mRows - 1) As Double
    For j = 0 To nCols - 1
        For i = 0 To mRows - 1
            'ca(i) = a(j, i)
            ca(i) = a(i, j)
        Next
        vsa = Split(VectorFormat(ca, 0, , dFormat), vbCrLf)
        For i = 0 To mRows - 1
            msa(i) = msa(i) & " " & vsa(i)
        Next
    Next
    MatrixA_ToStr = Join(msa, vbCrLf) 's

End Function

'Function MatA_ToStr(ByVal pMat As Long, ByVal mRows As Long, ByVal nCols As Long)
'    ReDim ma(0 To mRows - 1, 0 To nCols - 1) As Double
'    RtlMoveMemory ma(0, 0), ByVal pMat, mRows * nCols * 8
'    Dim i As Long, j As Long
'    Dim s As String
'    mRows = mRows - 1
'    nCols = nCols - 1
'    For i = 0 To mRows
'        For j = 0 To nCols
'            s = s & ma(j, i)
'            If j < nCols Then s = s & " "
'        Next
'        If i < mRows Then s = s & vbCrLf
'    Next
'    MatA_ToStr = s
'End Function

Public Function Matrix_ToStr(ByVal p_Matrix As Long, ByVal mRows As Long, ByVal nCols As Long) As String
    ReDim mA(0 To mRows - 1, 0 To nCols - 1) As Double
    RtlMoveMemory mA(0, 0), ByVal p_Matrix, mRows * nCols * 8
    Dim i As Long, j As Long
    Dim s As String
    mRows = mRows - 1
    nCols = nCols - 1
    For i = 0 To mRows
        For j = 0 To nCols
            s = s & mA(j, i)
            If j < nCols Then s = s & " "
        Next
        If i < mRows Then s = s & vbCrLf
    Next
    Matrix_ToStr = s



'    'die allgemeine mathematische Anordnung  ist a(iZeile, jSpalte)
'    'vgl die Speicheranordnung von VB-Arrays ist a(jSpalte, iZeile)
'    If mRows = 0 Or nCols = 0 Then Exit Function
'    ReDim a(0 To nCols - 1, 0 To mRows - 1) As Double
'    RtlMoveMemory a(0, 0), ByVal p_Matrix, mRows * nCols * 8
'    Matrix_ToStr = MatrixA_ToStr(a, mRows, nCols)
End Function
'Zeile 588



'Fortgeschrittene Matrix-Operationen
'Berechnung der Determinante
Public Function Mat2_det(m As Matrix2) As Double
    'Berechnet die Determinante einer 2x2-Matrix
    With m
        Mat2_det = .aa * .bb - .ab * .ba
    End With
End Function
Public Function Mat3_det(m As Matrix3) As Double
    'Berechnet die Determinante einer 3x3-Matrix
    Dim d1 As Double, d2 As Double, d3 As Double
    Dim d4 As Double, d5 As Double, d6 As Double
    
    With m
        d1 = .aa * .bb * .cc
        d2 = .ab * .bc * .ca
        d3 = .ac * .ba * .cb
        d4 = -.ac * .bb * .ca
        d5 = -.ab * .ba * .cc
        d6 = -.aa * .bc * .cb
    End With
    Mat3_det = d1 + d2 + d3 + d4 + d5 + d6

'    With m
'        Mat3_det = .aa * .bb * .cc + .ab * .bc * .ca + .ac * .ba * .cb _
'                 - .ac * .bb * .ca - .ab * .ba * .cc - .aa * .bc * .cb
'    End With
End Function
'oder alternativ:
'Public Function Mat3_det2(m As Matrix3) As Double
'    'Berechnet die Determinante einer 3x3-Matrix
'    With m
'        Mat3_det2 = .aa * Mat2_det(Mat2(.bb, .bc, .cb, .cc)) _
'                     - .ab * Mat2_det(Mat2(.ba, .bc, .ca, .cc)) _
'                     + .ac * Mat2_det(Mat2(.ba, .bb, .ca, .cb))
'    End With
'End Function
Public Function Mat4_det(m As Matrix4) As Double
    'Berechnet die Determinante einer 4x4-Matrix
    'Entwicklung nach der letzten Zeile 'Achtung die Vorzeichen sind dann immer anders
    Dim md As Matrix3
    With md
        .aa = m.ab: .ab = m.ac: .ac = m.ad
        .ba = m.bb: .bb = m.bc: .bc = m.bd
        .ca = m.cb: .cb = m.cc: .cc = m.cd
    End With
    Dim det_a As Double: det_a = Mat3_det(md)
    With md
        .aa = m.aa: .ab = m.ac: .ac = m.ad
        .ba = m.ba: .bb = m.bc: .bc = m.bd
        .ca = m.ca: .cb = m.cc: .cc = m.cd
    End With
    Dim det_b As Double: det_b = Mat3_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ad
        .ba = m.ba: .bb = m.bb: .bc = m.bd
        .ca = m.ca: .cb = m.cb: .cc = m.cd
    End With
    Dim det_c As Double: det_c = Mat3_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac
        .ba = m.ba: .bb = m.bb: .bc = m.bc
        .ca = m.ca: .cb = m.cb: .cc = m.cc
    End With
    Dim det_d As Double: det_d = Mat3_det(md)
         '   +   -   +   -   +
    With m
        Mat4_det = -.da * det_a _
                  + .db * det_b _
                  - .dc * det_c _
                  + .dd * det_d
    End With
End Function
Public Function Mat5_det(m As Matrix5) As Double
    'Berechnet die Determinante einer 5x5-Matrix
    'Entwicklung nach der letzten Zeile 'Achtung die Vorzeichen sind dann immer anders
    Dim md As Matrix4
    With md
        .aa = m.ab: .ab = m.ac: .ac = m.ad: .ad = m.ae
        .ba = m.bb: .bb = m.bc: .bc = m.bd: .bd = m.be
        .ca = m.cb: .cb = m.cc: .cc = m.cd: .cd = m.ce
        .da = m.db: .db = m.dc: .dc = m.dd: .dd = m.de
    End With
    Dim det_a As Double: det_a = Mat4_det(md)
    With md
        .aa = m.aa: .ab = m.ac: .ac = m.ad: .ad = m.ae
        .ba = m.ba: .bb = m.bc: .bc = m.bd: .bd = m.be
        .ca = m.ca: .cb = m.cc: .cc = m.cd: .cd = m.ce
        .da = m.da: .db = m.dc: .dc = m.dd: .dd = m.de
    End With
    Dim det_b As Double: det_b = Mat4_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ad: .ad = m.ae
        .ba = m.ba: .bb = m.bb: .bc = m.bd: .bd = m.be
        .ca = m.ca: .cb = m.cb: .cc = m.cd: .cd = m.ce
        .da = m.da: .db = m.db: .dc = m.dd: .dd = m.de
    End With
    Dim det_c As Double: det_c = Mat4_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ae
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.be
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.ce
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.de
    End With
    Dim det_d As Double: det_d = Mat4_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd
    End With
    Dim det_e As Double: det_e = Mat4_det(md)
         '   +   -   +   -   +
    With m
        Mat5_det = .ea * det_a _
                 - .eb * det_b _
                 + .ec * det_c _
                 - .ed * det_d _
                 + .ee * det_e
    End With
End Function
Public Function Mat6_det(m As Matrix6) As Double
    'Berechnet die Determinante einer 6x6-Matrix
    'Entwicklung nach der letzten Zeile 'Achtung die Vorzeichen sind dann immer anders
    Dim md As Matrix5
    With md
        .aa = m.ab: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af
        .ba = m.bb: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf
        .ca = m.cb: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf
        .da = m.db: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df
        .ea = m.eb: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef
    End With
    Dim det_a As Double: det_a = Mat5_det(md)
    With md
        .aa = m.aa: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af
        .ba = m.ba: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf
        .ca = m.ca: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf
        .da = m.da: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df
        .ea = m.ea: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef
    End With
    Dim det_b As Double: det_b = Mat5_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ad: .ad = m.ae: .ae = m.af
        .ba = m.ba: .bb = m.bb: .bc = m.bd: .bd = m.be: .be = m.bf
        .ca = m.ca: .cb = m.cb: .cc = m.cd: .cd = m.ce: .ce = m.cf
        .da = m.da: .db = m.db: .dc = m.dd: .dd = m.de: .de = m.df
        .ea = m.ea: .eb = m.eb: .ec = m.ed: .ed = m.ee: .ee = m.ef
    End With
    Dim det_c As Double: det_c = Mat5_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ae: .ae = m.af
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.be: .be = m.bf
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.ce: .ce = m.cf
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.de: .de = m.df
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ee: .ee = m.ef
    End With
    Dim det_d As Double: det_d = Mat5_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.af
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.bf
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.cf
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.df
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ef
    End With
    Dim det_e As Double: det_e = Mat5_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee
    End With
    Dim det_f As Double: det_f = Mat5_det(md)
         '   -   +   -   +   -   +
    With m
        Mat6_det = -.fa * det_a _
                  + .fb * det_b _
                  - .fc * det_c _
                  + .fd * det_d _
                  - .fe * det_e _
                  + .ff * det_f
    End With
End Function
Public Function Mat7_det(m As Matrix7) As Double
    'Berechnet die Determinante einer 7x7-Matrix
    'Entwicklung nach der letzten Zeile 'Achtung die Vorzeichen sind dann immer anders
    Dim md As Matrix6
    With md
        .aa = m.ab: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag
        .ba = m.bb: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg
        .ca = m.cb: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg
        .da = m.db: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg
        .ea = m.eb: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg
        .fa = m.fb: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg
    End With
    Dim det_a As Double: det_a = Mat6_det(md)
    With md
        .aa = m.aa: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag
        .ba = m.ba: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg
        .ca = m.ca: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg
        .da = m.da: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg
        .ea = m.ea: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg
        .fa = m.fa: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg
    End With
    Dim det_b As Double: det_b = Mat6_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag
        .ba = m.ba: .bb = m.bb: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg
        .ca = m.ca: .cb = m.cb: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg
        .da = m.da: .db = m.db: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg
        .ea = m.ea: .eb = m.eb: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg
        .fa = m.fa: .fb = m.fb: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg
    End With
    Dim det_c As Double: det_c = Mat6_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ae: .ae = m.af: .af = m.ag
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.be: .be = m.bf: .bf = m.bg
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.ce: .ce = m.cf: .cf = m.cg
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.de: .de = m.df: .df = m.dg
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ee: .ee = m.ef: .ef = m.eg
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fe: .fe = m.ff: .ff = m.fg
    End With
    Dim det_d As Double: det_d = Mat6_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.af: .af = m.ag
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.bf: .bf = m.bg
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.cf: .cf = m.cg
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.df: .df = m.dg
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ef: .ef = m.eg
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.ff: .ff = m.fg
    End With
    Dim det_e As Double: det_e = Mat6_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.ag
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bg
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cg
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.dg
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.eg
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.fg
    End With
    Dim det_f As Double: det_f = Mat6_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff
    End With
    Dim det_g As Double: det_g = Mat6_det(md)
         '   +   -   +   -   +   -   +   -   +
    With m
        Mat7_det = .ga * det_a _
                 - .gb * det_b _
                 + .gc * det_c _
                 - .gd * det_d _
                 + .ge * det_e _
                 - .gf * det_f _
                 + .gg * det_g
    End With
End Function
Public Function Mat8_det(m As Matrix8) As Double
    'Berechnet die Determinante einer 8x8-Matrix
    'Entwicklung nach der letzten Zeile 'Achtung die Vorzeichen sind dann immer anders
    Dim md As Matrix7
    With md
        .aa = m.ab: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah
        .ba = m.bb: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh
        .ca = m.cb: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch
        .da = m.db: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh
        .ea = m.eb: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh
        .fa = m.fb: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh
        .ga = m.gb: .gb = m.gc: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh
    End With
    Dim det_a As Double: det_a = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah
        .ba = m.ba: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh
        .ca = m.ca: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch
        .da = m.da: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh
        .ea = m.ea: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh
        .fa = m.fa: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh
        .ga = m.ga: .gb = m.gc: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh
    End With
    Dim det_b As Double: det_b = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah
        .ba = m.ba: .bb = m.bb: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh
        .ca = m.ca: .cb = m.cb: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch
        .da = m.da: .db = m.db: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh
        .ea = m.ea: .eb = m.eb: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh
        .fa = m.fa: .fb = m.fb: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh
        .ga = m.ga: .gb = m.gb: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh
    End With
    Dim det_c As Double: det_c = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh
    End With
    Dim det_d As Double: det_d = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.af: .af = m.ag: .ag = m.ah
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.bf: .bf = m.bg: .bg = m.bh
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.cf: .cf = m.cg: .cg = m.ch
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.df: .df = m.dg: .dg = m.dh
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ef: .ef = m.eg: .eg = m.eh
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.ff: .ff = m.fg: .fg = m.fh
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.gf: .gf = m.gg: .gg = m.gh
    End With
    Dim det_e As Double: det_e = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.ag: .ag = m.ah
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bg: .bg = m.bh
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cg: .cg = m.ch
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.dg: .dg = m.dh
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.eg: .eg = m.eh
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.fg: .fg = m.fh
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gg: .gg = m.gh
    End With
    Dim det_f As Double: det_f = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ah
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bh
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.ch
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dh
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eh
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fh
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gh
    End With
    Dim det_g As Double: det_g = Mat7_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ag
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bg
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.cg
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dg
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eg
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fg
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gg
    End With
    Dim det_h As Double: det_h = Mat7_det(md)
         '+   -   +   -   +   -   +   -   +
    With m
        Mat8_det = -.ha * det_a _
                  + .hb * det_b _
                  - .HC * det_c _
                  + .hd * det_d _
                  - .he * det_e _
                  + .hf * det_f _
                  - .hg * det_g _
                  + .hh * det_h
    End With
End Function


Public Function Mat9_det(m As Matrix9) As Double
    Dim md As Matrix8
    With md
        .aa = m.ab: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
        .ba = m.bb: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
        .ca = m.cb: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
        .da = m.db: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
        .ea = m.eb: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
        .fa = m.fb: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
        .ga = m.gb: .gb = m.gc: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
        .ha = m.hb: .hb = m.HC: .HC = m.hd: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
    End With
    Dim det_a As Double: det_a = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ac: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
        .ba = m.ba: .bb = m.bc: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
        .ca = m.ca: .cb = m.cc: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
        .da = m.da: .db = m.dc: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
        .ea = m.ea: .eb = m.ec: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
        .fa = m.fa: .fb = m.fc: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
        .ga = m.ga: .gb = m.gc: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
        .ha = m.ha: .hb = m.HC: .HC = m.hd: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
    End With
    Dim det_b As Double: det_b = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ad: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
        .ba = m.ba: .bb = m.bb: .bc = m.bd: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
        .ca = m.ca: .cb = m.cb: .cc = m.cd: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
        .da = m.da: .db = m.db: .dc = m.dd: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
        .ea = m.ea: .eb = m.eb: .ec = m.ed: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
        .fa = m.fa: .fb = m.fb: .fc = m.fd: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
        .ga = m.ga: .gb = m.gb: .gc = m.gd: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
        .ha = m.ha: .hb = m.hb: .HC = m.hd: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
    End With
    Dim det_c As Double: det_c = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ae: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.be: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.ce: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.de: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ee: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fe: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.ge: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.he: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
    End With
    Dim det_d As Double: det_d = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.af: .af = m.ag: .ag = m.ah: .ah = m.ai
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.bf: .bf = m.bg: .bg = m.bh: .bh = m.bi
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.cf: .cf = m.cg: .cg = m.ch: .ch = m.ci
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.df: .df = m.dg: .dg = m.dh: .dh = m.di
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ef: .ef = m.eg: .eg = m.eh: .eh = m.ei
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.ff: .ff = m.fg: .fg = m.fh: .fh = m.fi
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.gf: .gf = m.gg: .gg = m.gh: .gh = m.gi
        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.hf: .hf = m.hg: .hg = m.hh: .hh = m.hi
    End With
    Dim det_e As Double: det_e = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.ag: .ag = m.ah: .ah = m.ai
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bg: .bg = m.bh: .bh = m.bi
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cg: .cg = m.ch: .ch = m.ci
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.dg: .dg = m.dh: .dh = m.di
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.eg: .eg = m.eh: .eh = m.ei
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.fg: .fg = m.fh: .fh = m.fi
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gg: .gg = m.gh: .gh = m.gi
        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hg: .hg = m.hh: .hh = m.hi
    End With
    Dim det_f As Double: det_f = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ah: .ah = m.ai
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bh: .bh = m.bi
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.ch: .ch = m.ci
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dh: .dh = m.di
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eh: .eh = m.ei
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fh: .fh = m.fi
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gh: .gh = m.gi
        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hf: .hg = m.hh: .hh = m.hi
    End With
    Dim det_g As Double: det_g = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ag: .ah = m.ai
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bg: .bh = m.bi
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.cg: .ch = m.ci
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dg: .dh = m.di
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eg: .eh = m.ei
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fg: .fh = m.fi
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gg: .gh = m.gi
        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hf: .hg = m.hg: .hh = m.hi
    End With
    Dim det_h As Double: det_h = Mat8_det(md)
    With md
        .aa = m.aa: .ab = m.ab: .ac = m.ac: .ad = m.ad: .ae = m.ae: .af = m.af: .ag = m.ag: .ah = m.ah
        .ba = m.ba: .bb = m.bb: .bc = m.bc: .bd = m.bd: .be = m.be: .bf = m.bf: .bg = m.bg: .bh = m.bh
        .ca = m.ca: .cb = m.cb: .cc = m.cc: .cd = m.cd: .ce = m.ce: .cf = m.cf: .cg = m.cg: .ch = m.ch
        .da = m.da: .db = m.db: .dc = m.dc: .dd = m.dd: .de = m.de: .df = m.df: .dg = m.dg: .dh = m.dh
        .ea = m.ea: .eb = m.eb: .ec = m.ec: .ed = m.ed: .ee = m.ee: .ef = m.ef: .eg = m.eg: .eh = m.eh
        .fa = m.fa: .fb = m.fb: .fc = m.fc: .fd = m.fd: .fe = m.fe: .ff = m.ff: .fg = m.fg: .fh = m.fh
        .ga = m.ga: .gb = m.gb: .gc = m.gc: .gd = m.gd: .ge = m.ge: .gf = m.gf: .gg = m.gg: .gh = m.gh
        .ha = m.ha: .hb = m.hb: .HC = m.HC: .hd = m.hd: .he = m.he: .hf = m.hf: .hg = m.hg: .hh = m.hh
    End With
    Dim det_i As Double: det_i = Mat8_det(md)
         '+   -   +   -   +   -   +   -   +
    With m
        Mat9_det = .ia * det_a _
                 - .ib * det_b _
                 + .ic * det_c _
                 - .id * det_d _
                 + .ie * det_e _
                 - .if * det_f _
                 + .ig * det_g _
                 - .ih * det_h _
                 + .ii * det_i
    End With
End Function

Public Function Mat10_det(m As Matrix10) As Double
    'Berechnet die Determinante einer 10x10-Matrix
    'Entwicklung nach der ersten Zeile
    'eigentlich ist Entwicklung nach der letzen Zeile einfacher, weil aa=aa: .ab=m.ab
    Dim M9 As Matrix9
    With M9
        .aa = m.bb: .ab = m.bc: .ac = m.bd: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.cb: .bb = m.cc: .bc = m.cd: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.db: .cb = m.dc: .cc = m.dd: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.eb: .db = m.ec: .dc = m.ed: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fb: .eb = m.fc: .ec = m.fd: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.gb: .fb = m.gc: .fc = m.gd: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.hb: .gb = m.HC: .gc = m.hd: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ib: .hb = m.ic: .HC = m.id: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.jb: .ib = m.jc: .ic = m.jd: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_a As Double: det_a = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bc: .ac = m.bd: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cc: .bc = m.cd: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.dc: .cc = m.dd: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.ec: .dc = m.ed: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fc: .ec = m.fd: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gc: .fc = m.gd: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.HC: .gc = m.hd: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ic: .HC = m.id: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jc: .ic = m.jd: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_b As Double: det_b = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bd: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cd: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dd: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ed: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fd: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gd: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.hd: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.id: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jd: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_c As Double: det_c = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.be: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.ce: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.de: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ee: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fe: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.ge: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.he: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.ie: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.je: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_d As Double: det_d = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.bf: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.cf: .bf = m.cg: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.df: .cf = m.dg: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ef: .df = m.eg: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.ff: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.gf: .ff = m.gg: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.hd: .ge = m.hf: .gf = m.hg: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.id: .he = m.if: .hf = m.ig: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.jf: .if = m.jg: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_e As Double: det_e = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bg: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cg: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.dg: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.eg: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.fg: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gg: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.hd: .ge = m.he: .gf = m.hg: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.id: .he = m.ie: .hf = m.ig: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jg: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_f As Double: det_f = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bh: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.ch: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dh: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eh: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fh: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gh: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hh: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ih: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jh: .ih = m.ji: .ii = m.jj
    End With
    Dim det_g As Double: det_g = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bg: .ah = m.bi: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.cg: .bh = m.ci: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dg: .ch = m.di: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eg: .dh = m.ei: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fg: .eh = m.fi: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gg: .fh = m.gi: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hg: .gh = m.hi: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ig: .hh = m.ii: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jg: .ih = m.ji: .ii = m.jj
    End With
    Dim det_h As Double: det_h = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bg: .ah = m.bh: .ai = m.bj
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.cg: .bh = m.ch: .bi = m.cj
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dg: .ch = m.dh: .ci = m.dj
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eg: .dh = m.eh: .di = m.ej
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fg: .eh = m.fh: .ei = m.fj
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gg: .fh = m.gh: .fi = m.gj
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hg: .gh = m.hh: .gi = m.hj
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ig: .hh = m.ih: .hi = m.ij
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jg: .ih = m.jh: .ii = m.jj
    End With
    Dim det_i As Double: det_i = Mat9_det(M9)
    With M9
        .aa = m.ba: .ab = m.bb: .ac = m.bc: .ad = m.bd: .ae = m.be: .af = m.bf: .ag = m.bg: .ah = m.bh: .ai = m.bi
        .ba = m.ca: .bb = m.cb: .bc = m.cc: .bd = m.cd: .be = m.ce: .bf = m.cf: .bg = m.cg: .bh = m.ch: .bi = m.ci
        .ca = m.da: .cb = m.db: .cc = m.dc: .cd = m.dd: .ce = m.de: .cf = m.df: .cg = m.dg: .ch = m.dh: .ci = m.di
        .da = m.ea: .db = m.eb: .dc = m.ec: .dd = m.ed: .de = m.ee: .df = m.ef: .dg = m.eg: .dh = m.eh: .di = m.ei
        .ea = m.fa: .eb = m.fb: .ec = m.fc: .ed = m.fd: .ee = m.fe: .ef = m.ff: .eg = m.fg: .eh = m.fh: .ei = m.fi
        .fa = m.ga: .fb = m.gb: .fc = m.gc: .fd = m.gd: .fe = m.ge: .ff = m.gf: .fg = m.gg: .fh = m.gh: .fi = m.gi
        .ga = m.ha: .gb = m.hb: .gc = m.HC: .gd = m.hd: .ge = m.he: .gf = m.hf: .gg = m.hg: .gh = m.hh: .gi = m.hi
        .ha = m.ia: .hb = m.ib: .HC = m.ic: .hd = m.id: .he = m.ie: .hf = m.if: .hg = m.ig: .hh = m.ih: .hi = m.ii
        .ia = m.ja: .ib = m.jb: .ic = m.jc: .id = m.jd: .ie = m.je: .if = m.jf: .ig = m.jg: .ih = m.jh: .ii = m.ji
    End With
    Dim det_j As Double: det_j = Mat9_det(M9)
        
    With m
        Mat10_det = .aa * det_a _
                     - .ab * det_b _
                     + .ac * det_c _
                     - .ad * det_d _
                     + .ae * det_e _
                     - .af * det_f _
                     + .ag * det_g _
                     - .ah * det_h _
                     + .ai * det_i _
                     - .aj * det_j
    End With
End Function

'Lesen/Schreiben von Zeilen oder Spalten aus/in eine Matrix
Public Property Get Mat2_Row(m As Matrix2, ByVal Index As Long) As Vector2
    With m
        Select Case Index
        Case 0: Mat2_Row = Vec2(.aa, .ab)
        Case 1: Mat2_Row = Vec2(.ba, .bb)
        End Select
    End With
End Property
Public Property Let Mat2_Row(m As Matrix2, ByVal Index As Long, v As Vector2)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b
        Case 1: .ba = v.a: .bb = v.b
        End Select
    End With
End Property

Public Property Get Mat2_Col(m As Matrix2, ByVal Index As Long) As Vector2
    With m
        Select Case Index
        Case 0: Mat2_Col = Vec2(.aa, .ba)
        Case 1: Mat2_Col = Vec2(.ab, .bb)
        End Select
    End With
End Property
Public Property Let Mat2_Col(m As Matrix2, ByVal Index As Long, v As Vector2)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b
        Case 1: .ab = v.a: .bb = v.b
        End Select
    End With
End Property

Public Property Get Mat3_Row(m As Matrix3, ByVal Index As Long) As Vector3
    With m
        Select Case Index
        Case 0: Mat3_Row = Vec3(.aa, .ab, .ac)
        Case 1: Mat3_Row = Vec3(.ba, .bb, .bc)
        Case 2: Mat3_Row = Vec3(.ca, .cb, .cc)
        End Select
    End With
End Property
Public Property Let Mat3_Row(m As Matrix3, ByVal Index As Long, v As Vector3)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c
        End Select
    End With
End Property
Public Property Get Mat3_Col(m As Matrix3, ByVal Index As Long) As Vector3
    With m
        Select Case Index
        Case 0: Mat3_Col = Vec3(.aa, .ba, .ca)
        Case 1: Mat3_Col = Vec3(.ab, .bb, .cb)
        Case 2: Mat3_Col = Vec3(.ac, .bc, .cc)
        End Select
    End With
End Property
Public Property Let Mat3_Col(m As Matrix3, ByVal Index As Long, v As Vector3)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c
        End Select
    End With
End Property

Public Property Get Mat4_Row(m As Matrix4, ByVal Index As Long) As Vector4
    With m
        Select Case Index
        Case 0: Mat4_Row = Vec4(.aa, .ab, .ac, .ad)
        Case 1: Mat4_Row = Vec4(.ba, .bb, .bc, .bd)
        Case 2: Mat4_Row = Vec4(.ca, .cb, .cc, .cd)
        Case 3: Mat4_Row = Vec4(.da, .db, .dc, .dd)
        End Select
    End With
End Property
Public Property Let Mat4_Row(m As Matrix4, ByVal Index As Long, v As Vector4)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d
        End Select
    End With
End Property
Public Property Get Mat4_Col(m As Matrix4, ByVal Index As Long) As Vector4
    With m
        Select Case Index
        Case 0: Mat4_Col = Vec4(.aa, .ba, .ca, .da)
        Case 1: Mat4_Col = Vec4(.ab, .bb, .cb, .db)
        Case 2: Mat4_Col = Vec4(.ac, .bc, .cc, .dc)
        Case 3: Mat4_Col = Vec4(.ad, .bd, .cd, .dd)
        End Select
    End With
End Property
Public Property Let Mat4_Col(m As Matrix4, ByVal Index As Long, v As Vector4)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d
        End Select
    End With
End Property

Public Property Get Mat5_Row(m As Matrix5, ByVal Index As Long) As Vector5
    With m
        Select Case Index
        Case 0: Mat5_Row = Vec5(.aa, .ab, .ac, .ad, .ae)
        Case 1: Mat5_Row = Vec5(.ba, .bb, .bc, .bd, .be)
        Case 2: Mat5_Row = Vec5(.ca, .cb, .cc, .cd, .ce)
        Case 3: Mat5_Row = Vec5(.da, .db, .dc, .dd, .de)
        Case 4: Mat5_Row = Vec5(.ea, .eb, .ec, .ed, .ee)
        End Select
    End With
End Property
Public Property Let Mat5_Row(m As Matrix5, ByVal Index As Long, v As Vector5)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e
        End Select
    End With
End Property
Public Property Get Mat5_Col(m As Matrix5, ByVal Index As Long) As Vector5
    With m
        Select Case Index
        Case 0: Mat5_Col = Vec5(.aa, .ba, .ca, .da, .ea)
        Case 1: Mat5_Col = Vec5(.ab, .bb, .cb, .db, .eb)
        Case 2: Mat5_Col = Vec5(.ac, .bc, .cc, .dc, .ec)
        Case 3: Mat5_Col = Vec5(.ad, .bd, .cd, .dd, .ed)
        Case 4: Mat5_Col = Vec5(.ae, .be, .ce, .de, .ee)
        End Select
    End With
End Property
Public Property Let Mat5_Col(m As Matrix5, ByVal Index As Long, v As Vector5)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e
        End Select
    End With
End Property

Public Property Get Mat6_Row(m As Matrix6, ByVal Index As Long) As Vector6
    With m
        Select Case Index
        Case 0: Mat6_Row = Vec6(.aa, .ab, .ac, .ad, .ae, .af)
        Case 1: Mat6_Row = Vec6(.ba, .bb, .bc, .bd, .be, .bf)
        Case 2: Mat6_Row = Vec6(.ca, .cb, .cc, .cd, .ce, .cf)
        Case 3: Mat6_Row = Vec6(.da, .db, .dc, .dd, .de, .df)
        Case 4: Mat6_Row = Vec6(.ea, .eb, .ec, .ed, .ee, .ef)
        Case 5: Mat6_Row = Vec6(.fa, .fb, .fc, .fd, .fe, .ff)
        End Select
    End With
End Property
Public Property Let Mat6_Row(m As Matrix6, ByVal Index As Long, v As Vector6)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f
        End Select
    End With
End Property
Public Property Get Mat6_Col(m As Matrix6, ByVal Index As Long) As Vector6
    With m
        Select Case Index
        Case 0: Mat6_Col = Vec6(.aa, .ba, .ca, .da, .ea, .fa)
        Case 1: Mat6_Col = Vec6(.ab, .bb, .cb, .db, .eb, .fb)
        Case 2: Mat6_Col = Vec6(.ac, .bc, .cc, .dc, .ec, .fc)
        Case 3: Mat6_Col = Vec6(.ad, .bd, .cd, .dd, .ed, .fd)
        Case 4: Mat6_Col = Vec6(.ae, .be, .ce, .de, .ee, .fe)
        Case 5: Mat6_Col = Vec6(.af, .bf, .cf, .df, .ef, .ff)
        End Select
    End With
End Property
Public Property Let Mat6_Col(m As Matrix6, ByVal Index As Long, v As Vector6)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f
        End Select
    End With
End Property

Public Property Get Mat7_Row(m As Matrix7, ByVal Index As Long) As Vector7
    With m
        Select Case Index
        Case 0: Mat7_Row = Vec7(.aa, .ab, .ac, .ad, .ae, .af, .ag)
        Case 1: Mat7_Row = Vec7(.ba, .bb, .bc, .bd, .be, .bf, .bg)
        Case 2: Mat7_Row = Vec7(.ca, .cb, .cc, .cd, .ce, .cf, .cg)
        Case 3: Mat7_Row = Vec7(.da, .db, .dc, .dd, .de, .df, .dg)
        Case 4: Mat7_Row = Vec7(.ea, .eb, .ec, .ed, .ee, .ef, .eg)
        Case 5: Mat7_Row = Vec7(.fa, .fb, .fc, .fd, .fe, .ff, .fg)
        Case 6: Mat7_Row = Vec7(.ga, .gb, .gc, .gd, .ge, .gf, .gg)
        End Select
    End With
End Property
Public Property Let Mat7_Row(m As Matrix7, ByVal Index As Long, v As Vector7)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g
        End Select
    End With
End Property
Public Property Get Mat7_Col(m As Matrix7, ByVal Index As Long) As Vector7
    With m
        Select Case Index
        Case 0: Mat7_Col = Vec7(.aa, .ba, .ca, .da, .ea, .fa, .ga)
        Case 1: Mat7_Col = Vec7(.ab, .bb, .cb, .db, .eb, .fb, .gb)
        Case 2: Mat7_Col = Vec7(.ac, .bc, .cc, .dc, .ec, .fc, .gc)
        Case 3: Mat7_Col = Vec7(.ad, .bd, .cd, .dd, .ed, .fd, .gd)
        Case 4: Mat7_Col = Vec7(.ae, .be, .ce, .de, .ee, .fe, .ge)
        Case 5: Mat7_Col = Vec7(.af, .bf, .cf, .df, .ef, .ff, .gf)
        Case 6: Mat7_Col = Vec7(.ag, .bg, .cg, .dg, .eg, .fg, .gg)
        End Select
    End With
End Property
Public Property Let Mat7_Col(m As Matrix7, ByVal Index As Long, v As Vector7)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g
        End Select
    End With
End Property

Public Property Get Mat8_Row(m As Matrix8, ByVal Index As Long) As Vector8
    With m
        Select Case Index
        Case 0: Mat8_Row = Vec8(.aa, .ab, .ac, .ad, .ae, .af, .ag, .ah)
        Case 1: Mat8_Row = Vec8(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh)
        Case 2: Mat8_Row = Vec8(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch)
        Case 3: Mat8_Row = Vec8(.da, .db, .dc, .dd, .de, .df, .dg, .dh)
        Case 4: Mat8_Row = Vec8(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh)
        Case 5: Mat8_Row = Vec8(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh)
        Case 6: Mat8_Row = Vec8(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh)
        Case 7: Mat8_Row = Vec8(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hh)
        End Select
    End With
End Property
Public Property Let Mat8_Row(m As Matrix8, ByVal Index As Long, v As Vector8)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g: .ah = v.H
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g: .bh = v.H
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g: .ch = v.H
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g: .dh = v.H
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g: .eh = v.H
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g: .fh = v.H
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g: .gh = v.H
        Case 7: .ha = v.a: .hb = v.b: .HC = v.c: .hd = v.d: .he = v.e: .hf = v.f: .hg = v.g: .hh = v.H
        End Select
    End With
End Property
Public Property Get Mat8_Col(m As Matrix8, ByVal Index As Long) As Vector8
    With m
        Select Case Index
        Case 0: Mat8_Col = Vec8(.aa, .ba, .ca, .da, .ea, .fa, .ga, .ha)
        Case 1: Mat8_Col = Vec8(.ab, .bb, .cb, .db, .eb, .fb, .gb, .hb)
        Case 2: Mat8_Col = Vec8(.ac, .bc, .cc, .dc, .ec, .fc, .gc, .HC)
        Case 3: Mat8_Col = Vec8(.ad, .bd, .cd, .dd, .ed, .fd, .gd, .hd)
        Case 4: Mat8_Col = Vec8(.ae, .be, .ce, .de, .ee, .fe, .ge, .he)
        Case 5: Mat8_Col = Vec8(.af, .bf, .cf, .df, .ef, .ff, .gf, .hf)
        Case 6: Mat8_Col = Vec8(.ag, .bg, .cg, .dg, .eg, .fg, .gg, .hg)
        Case 7: Mat8_Col = Vec8(.ah, .bh, .ch, .dh, .eh, .fh, .gh, .hh)
        End Select
    End With
End Property
Public Property Let Mat8_Col(m As Matrix8, ByVal Index As Long, v As Vector8)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g: .ha = v.H
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g: .hb = v.H
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g: .HC = v.H
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g: .hd = v.H
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g: .he = v.H
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g: .hf = v.H
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g: .hg = v.H
        Case 7: .ah = v.a: .bh = v.b: .ch = v.c: .dh = v.d: .eh = v.e: .fh = v.f: .gh = v.g: .hh = v.H
        End Select
    End With
End Property

Public Property Get Mat9_Row(m As Matrix9, ByVal Index As Long) As Vector9
    With m
        Select Case Index
        Case 0: Mat9_Row = Vec9(.aa, .ab, .ac, .ad, .ae, .af, .ag, .ah, .ai)
        Case 1: Mat9_Row = Vec9(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi)
        Case 2: Mat9_Row = Vec9(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci)
        Case 3: Mat9_Row = Vec9(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di)
        Case 4: Mat9_Row = Vec9(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei)
        Case 5: Mat9_Row = Vec9(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi)
        Case 6: Mat9_Row = Vec9(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi)
        Case 7: Mat9_Row = Vec9(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hh, .hi)
        Case 8: Mat9_Row = Vec9(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii)
        End Select
    End With
End Property
Public Property Let Mat9_Row(m As Matrix9, ByVal Index As Long, v As Vector9)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g: .ah = v.H: .ai = v.i
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g: .bh = v.H: .bi = v.i
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g: .ch = v.H: .ci = v.i
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g: .dh = v.H: .di = v.i
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g: .eh = v.H: .ei = v.i
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g: .fh = v.H: .fi = v.i
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g: .gh = v.H: .gi = v.i
        Case 7: .ha = v.a: .hb = v.b: .HC = v.c: .hd = v.d: .he = v.e: .hf = v.f: .hg = v.g: .hh = v.H: .hi = v.i
        Case 8: .ia = v.a: .ib = v.b: .ic = v.c: .id = v.d: .ie = v.e: .if = v.f: .ig = v.g: .ih = v.H: .ii = v.i
        End Select
    End With
End Property
Public Property Get Mat9_Col(m As Matrix9, ByVal Index As Long) As Vector9
    With m
        Select Case Index
        Case 0: Mat9_Col = Vec9(.aa, .ba, .ca, .da, .ea, .fa, .ga, .ha, .ia)
        Case 1: Mat9_Col = Vec9(.ab, .bb, .cb, .db, .eb, .fb, .gb, .hb, .ib)
        Case 2: Mat9_Col = Vec9(.ac, .bc, .cc, .dc, .ec, .fc, .gc, .HC, .ic)
        Case 3: Mat9_Col = Vec9(.ad, .bd, .cd, .dd, .ed, .fd, .gd, .hd, .id)
        Case 4: Mat9_Col = Vec9(.ae, .be, .ce, .de, .ee, .fe, .ge, .he, .ie)
        Case 5: Mat9_Col = Vec9(.af, .bf, .cf, .df, .ef, .ff, .gf, .hf, .if)
        Case 6: Mat9_Col = Vec9(.ag, .bg, .cg, .dg, .eg, .fg, .gg, .hg, .ig)
        Case 7: Mat9_Col = Vec9(.ah, .bh, .ch, .dh, .eh, .fh, .gh, .hh, .ih)
        Case 8: Mat9_Col = Vec9(.ai, .bi, .ci, .di, .ei, .fi, .gi, .hi, .ii)
        End Select
    End With
End Property
Public Property Let Mat9_Col(m As Matrix9, ByVal Index As Long, v As Vector9)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g: .ha = v.H: .ia = v.i
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g: .hb = v.H: .ib = v.i
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g: .HC = v.H: .ic = v.i
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g: .hd = v.H: .id = v.i
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g: .he = v.H: .ie = v.i
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g: .hf = v.H: .if = v.i
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g: .hg = v.H: .ig = v.i
        Case 7: .ah = v.a: .bh = v.b: .ch = v.c: .dh = v.d: .eh = v.e: .fh = v.f: .gh = v.g: .hh = v.H: .ih = v.i
        Case 8: .ai = v.a: .bi = v.b: .ci = v.c: .di = v.d: .ei = v.e: .fi = v.f: .gi = v.g: .hi = v.H: .ii = v.i
        End Select
    End With
End Property

Public Property Get Mat10_Row(m As Matrix10, ByVal Index As Long) As Vector10
    With m
        Select Case Index
        Case 0: Mat10_Row = Vec10(.aa, .ab, .ac, .ad, .ae, .af, .ag, .ah, .ai, .aj)
        Case 1: Mat10_Row = Vec10(.ba, .bb, .bc, .bd, .be, .bf, .bg, .bh, .bi, .bj)
        Case 2: Mat10_Row = Vec10(.ca, .cb, .cc, .cd, .ce, .cf, .cg, .ch, .ci, .cj)
        Case 3: Mat10_Row = Vec10(.da, .db, .dc, .dd, .de, .df, .dg, .dh, .di, .dj)
        Case 4: Mat10_Row = Vec10(.ea, .eb, .ec, .ed, .ee, .ef, .eg, .eh, .ei, .ej)
        Case 5: Mat10_Row = Vec10(.fa, .fb, .fc, .fd, .fe, .ff, .fg, .fh, .fi, .fj)
        Case 6: Mat10_Row = Vec10(.ga, .gb, .gc, .gd, .ge, .gf, .gg, .gh, .gi, .gj)
        Case 7: Mat10_Row = Vec10(.ha, .hb, .HC, .hd, .he, .hf, .hg, .hh, .hi, .hj)
        Case 8: Mat10_Row = Vec10(.ia, .ib, .ic, .id, .ie, .if, .ig, .ih, .ii, .ij)
        Case 9: Mat10_Row = Vec10(.ja, .jb, .jc, .jd, .je, .jf, .jg, .jh, .ji, .jj)
        End Select
    End With
End Property
Public Property Let Mat10_Row(m As Matrix10, ByVal Index As Long, v As Vector10)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ab = v.b: .ac = v.c: .ad = v.d: .ae = v.e: .af = v.f: .ag = v.g: .ah = v.H: .ai = v.i: .aj = v.j
        Case 1: .ba = v.a: .bb = v.b: .bc = v.c: .bd = v.d: .be = v.e: .bf = v.f: .bg = v.g: .bh = v.H: .bi = v.i: .bj = v.j
        Case 2: .ca = v.a: .cb = v.b: .cc = v.c: .cd = v.d: .ce = v.e: .cf = v.f: .cg = v.g: .ch = v.H: .ci = v.i: .cj = v.j
        Case 3: .da = v.a: .db = v.b: .dc = v.c: .dd = v.d: .de = v.e: .df = v.f: .dg = v.g: .dh = v.H: .di = v.i: .dj = v.j
        Case 4: .ea = v.a: .eb = v.b: .ec = v.c: .ed = v.d: .ee = v.e: .ef = v.f: .eg = v.g: .eh = v.H: .ei = v.i: .ej = v.j
        Case 5: .fa = v.a: .fb = v.b: .fc = v.c: .fd = v.d: .fe = v.e: .ff = v.f: .fg = v.g: .fh = v.H: .fi = v.i: .fj = v.j
        Case 6: .ga = v.a: .gb = v.b: .gc = v.c: .gd = v.d: .ge = v.e: .gf = v.f: .gg = v.g: .gh = v.H: .gi = v.i: .gj = v.j
        Case 7: .ha = v.a: .hb = v.b: .HC = v.c: .hd = v.d: .he = v.e: .hf = v.f: .hg = v.g: .hh = v.H: .hi = v.i: .hj = v.j
        Case 8: .ia = v.a: .ib = v.b: .ic = v.c: .id = v.d: .ie = v.e: .if = v.f: .ig = v.g: .ih = v.H: .ii = v.i: .ij = v.j
        Case 9: .ja = v.a: .jb = v.b: .jc = v.c: .jd = v.d: .je = v.e: .jf = v.f: .jg = v.g: .jh = v.H: .ji = v.i: .jj = v.j
        End Select
    End With
End Property
Public Property Get Mat10_Col(m As Matrix10, ByVal Index As Long) As Vector10
    With m
        Select Case Index
        Case 0: Mat10_Col = Vec10(.aa, .ba, .ca, .da, .ea, .fa, .ga, .ha, .ia, .ja)
        Case 1: Mat10_Col = Vec10(.ab, .bb, .cb, .db, .eb, .fb, .gb, .hb, .ib, .jb)
        Case 2: Mat10_Col = Vec10(.ac, .bc, .cc, .dc, .ec, .fc, .gc, .HC, .ic, .jc)
        Case 3: Mat10_Col = Vec10(.ad, .bd, .cd, .dd, .ed, .fd, .gd, .hd, .id, .jd)
        Case 4: Mat10_Col = Vec10(.ae, .be, .ce, .de, .ee, .fe, .ge, .he, .ie, .je)
        Case 5: Mat10_Col = Vec10(.af, .bf, .cf, .df, .ef, .ff, .gf, .hf, .if, .jf)
        Case 6: Mat10_Col = Vec10(.ag, .bg, .cg, .dg, .eg, .fg, .gg, .hg, .ig, .jg)
        Case 7: Mat10_Col = Vec10(.ah, .bh, .ch, .dh, .eh, .fh, .gh, .hh, .ih, .jh)
        Case 8: Mat10_Col = Vec10(.ai, .bi, .ci, .di, .ei, .fi, .gi, .hi, .ii, .ji)
        Case 9: Mat10_Col = Vec10(.aj, .bj, .cj, .dj, .ej, .fj, .gj, .hj, .ij, .jj)
        End Select
    End With
End Property
Public Property Let Mat10_Col(m As Matrix10, ByVal Index As Long, v As Vector10)
    With m
        Select Case Index
        Case 0: .aa = v.a: .ba = v.b: .ca = v.c: .da = v.d: .ea = v.e: .fa = v.f: .ga = v.g: .ha = v.H: .ia = v.i: .ja = v.j
        Case 1: .ab = v.a: .bb = v.b: .cb = v.c: .db = v.d: .eb = v.e: .fb = v.f: .gb = v.g: .hb = v.H: .ib = v.i: .jb = v.j
        Case 2: .ac = v.a: .bc = v.b: .cc = v.c: .dc = v.d: .ec = v.e: .fc = v.f: .gc = v.g: .HC = v.H: .ic = v.i: .jc = v.j
        Case 3: .ad = v.a: .bd = v.b: .cd = v.c: .dd = v.d: .ed = v.e: .fd = v.f: .gd = v.g: .hd = v.H: .id = v.i: .jd = v.j
        Case 4: .ae = v.a: .be = v.b: .ce = v.c: .de = v.d: .ee = v.e: .fe = v.f: .ge = v.g: .he = v.H: .ie = v.i: .je = v.j
        Case 5: .af = v.a: .bf = v.b: .cf = v.c: .df = v.d: .ef = v.e: .ff = v.f: .gf = v.g: .hf = v.H: .if = v.i: .jf = v.j
        Case 6: .ag = v.a: .bg = v.b: .cg = v.c: .dg = v.d: .eg = v.e: .fg = v.f: .gg = v.g: .hg = v.H: .ig = v.i: .jg = v.j
        Case 7: .ah = v.a: .bh = v.b: .ch = v.c: .dh = v.d: .eh = v.e: .fh = v.f: .gh = v.g: .hh = v.H: .ih = v.i: .jh = v.j
        Case 8: .ai = v.a: .bi = v.b: .ci = v.c: .di = v.d: .ei = v.e: .fi = v.f: .gi = v.g: .hi = v.H: .ii = v.i: .ji = v.j
        Case 9: .aj = v.a: .bj = v.b: .cj = v.c: .dj = v.d: .ej = v.e: .fj = v.f: .gj = v.g: .hj = v.H: .ij = v.i: .jj = v.j
        End Select
    End With
End Property

Public Sub SwapB(V1_inout As Byte, V2_inout As Byte)
    Dim tmp As Byte: tmp = V1_inout: V1_inout = V2_inout: V2_inout = tmp
End Sub

Public Sub Swap(V1_inout As Double, V2_inout As Double)
    Dim tmp As Double: tmp = V1_inout: V1_inout = V2_inout: V2_inout = tmp
End Sub
Public Function Mat2_RowSwap(m As Matrix2, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix2
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    If r1 = r2 Then Mat2_RowSwap = m: Exit Function
    With Mat2_RowSwap
        .aa = m.ba: .ab = m.bb
        .ba = m.aa: .bb = m.ab
    End With
End Function
Public Function Mat3_RowSwap(m As Matrix3, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix3
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat3_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat3_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc
            End Select
        End Select
    End With
End Function
Public Function Mat4_RowSwap(m As Matrix4, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix4
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat4_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat4_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd
            Case 4: Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd
            Case 4: Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd
            End Select
        Case 3
            Select Case r2
            Case 4: Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd
            End Select
        End Select
    End With
End Function
Public Function Mat5_RowSwap(m As Matrix5, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix5
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat5_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat5_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd: Swap .ae, .be
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd: Swap .ae, .ce
            Case 4: Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd: Swap .ae, .de
            Case 5: Swap .aa, .ea: Swap .ab, .eb: Swap .ac, .ec: Swap .ad, .ed: Swap .ae, .ee
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd: Swap .be, .ce
            Case 4: Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd: Swap .be, .de
            Case 5: Swap .ba, .ea: Swap .bb, .eb: Swap .bc, .ec: Swap .bd, .ed: Swap .be, .ee
            End Select
        Case 3
            Select Case r2
            Case 4: Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd: Swap .ce, .de
            Case 5: Swap .ca, .ea: Swap .cb, .eb: Swap .cc, .ec: Swap .cd, .ed: Swap .ce, .ee
            End Select
        Case 4
            Select Case r2
            Case 5: Swap .da, .ea: Swap .db, .eb: Swap .dc, .ec: Swap .dd, .ed: Swap .de, .ee
            End Select
        End Select
    End With
End Function
Public Function Mat6_RowSwap(m As Matrix6, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix6
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat6_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat6_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd: Swap .ae, .be: Swap .af, .bf
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd: Swap .ae, .ce: Swap .af, .cf
            Case 4: Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd: Swap .ae, .de: Swap .af, .df
            Case 5: Swap .aa, .ea: Swap .ab, .eb: Swap .ac, .ec: Swap .ad, .ed: Swap .ae, .ee: Swap .af, .ef
            Case 6: Swap .aa, .fa: Swap .ab, .fb: Swap .ac, .fc: Swap .ad, .fd: Swap .ae, .fe: Swap .af, .ff
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd: Swap .be, .ce: Swap .bf, .cf
            Case 4: Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd: Swap .be, .de: Swap .bf, .df
            Case 5: Swap .ba, .ea: Swap .bb, .eb: Swap .bc, .ec: Swap .bd, .ed: Swap .be, .ee: Swap .bf, .ef
            Case 6: Swap .ba, .fa: Swap .bb, .fb: Swap .bc, .fc: Swap .bd, .fd: Swap .be, .fe: Swap .bf, .ff
            End Select
        Case 3
            Select Case r2
            Case 4: Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd: Swap .ce, .de: Swap .cf, .df
            Case 5: Swap .ca, .ea: Swap .cb, .eb: Swap .cc, .ec: Swap .cd, .ed: Swap .ce, .ee: Swap .cf, .ef
            Case 6: Swap .ca, .fa: Swap .cb, .fb: Swap .cc, .fc: Swap .cd, .fd: Swap .ce, .fe: Swap .cf, .ff
            End Select
        Case 4
            Select Case r2
            Case 5: Swap .da, .ea: Swap .db, .eb: Swap .dc, .ec: Swap .dd, .ed: Swap .de, .ee: Swap .df, .ef
            Case 6: Swap .da, .fa: Swap .db, .fb: Swap .dc, .fc: Swap .dd, .fd: Swap .de, .fe: Swap .df, .ff
            End Select
        Case 5
            Select Case r2
            Case 6: Swap .ea, .fa: Swap .eb, .fb: Swap .ec, .fc: Swap .ed, .fd: Swap .ee, .fe: Swap .ef, .ff
            End Select
        End Select
    End With
End Function
Public Function Mat7_RowSwap(m As Matrix7, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix7
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat7_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat7_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd: Swap .ae, .be: Swap .af, .bf: Swap .ag, .bg
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd: Swap .ae, .ce: Swap .af, .cf: Swap .ag, .cg
            Case 4: Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd: Swap .ae, .de: Swap .af, .df: Swap .ag, .dg
            Case 5: Swap .aa, .ea: Swap .ab, .eb: Swap .ac, .ec: Swap .ad, .ed: Swap .ae, .ee: Swap .af, .ef: Swap .ag, .eg
            Case 6: Swap .aa, .fa: Swap .ab, .fb: Swap .ac, .fc: Swap .ad, .fd: Swap .ae, .fe: Swap .af, .ff: Swap .ag, .fg
            Case 7: Swap .aa, .ga: Swap .ab, .gb: Swap .ac, .gc: Swap .ad, .gd: Swap .ae, .ge: Swap .af, .gf: Swap .ag, .gg
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd: Swap .be, .ce: Swap .bf, .cf: Swap .bg, .cg
            Case 4: Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd: Swap .be, .de: Swap .bf, .df: Swap .bg, .dg
            Case 5: Swap .ba, .ea: Swap .bb, .eb: Swap .bc, .ec: Swap .bd, .ed: Swap .be, .ee: Swap .bf, .ef: Swap .bg, .eg
            Case 6: Swap .ba, .fa: Swap .bb, .fb: Swap .bc, .fc: Swap .bd, .fd: Swap .be, .fe: Swap .bf, .ff: Swap .bg, .fg
            Case 7: Swap .ba, .ga: Swap .bb, .gb: Swap .bc, .gc: Swap .bd, .gd: Swap .be, .ge: Swap .bf, .gf: Swap .bg, .gg
            End Select
        Case 3
            Select Case r2
            Case 4: Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd: Swap .ce, .de: Swap .cf, .df: Swap .cg, .dg
            Case 5: Swap .ca, .ea: Swap .cb, .eb: Swap .cc, .ec: Swap .cd, .ed: Swap .ce, .ee: Swap .cf, .ef: Swap .cg, .eg
            Case 6: Swap .ca, .fa: Swap .cb, .fb: Swap .cc, .fc: Swap .cd, .fd: Swap .ce, .fe: Swap .cf, .ff: Swap .cg, .fg
            Case 7: Swap .ca, .ga: Swap .cb, .gb: Swap .cc, .gc: Swap .cd, .gd: Swap .ce, .ge: Swap .cf, .gf: Swap .cg, .gg
            End Select
        Case 4
            Select Case r2
            Case 5: Swap .da, .ea: Swap .db, .eb: Swap .dc, .ec: Swap .dd, .ed: Swap .de, .ee: Swap .df, .ef: Swap .dg, .eg
            Case 6: Swap .da, .fa: Swap .db, .fb: Swap .dc, .fc: Swap .dd, .fd: Swap .de, .fe: Swap .df, .ff: Swap .dg, .fg
            Case 7: Swap .da, .ga: Swap .db, .gb: Swap .dc, .gc: Swap .dd, .gd: Swap .de, .ge: Swap .df, .gf: Swap .dg, .gg
            End Select
        Case 5
            Select Case r2
            Case 6: Swap .ea, .fa: Swap .eb, .fb: Swap .ec, .fc: Swap .ed, .fd: Swap .ee, .fe: Swap .ef, .ff: Swap .eg, .fg
            Case 7: Swap .ea, .ga: Swap .eb, .gb: Swap .ec, .gc: Swap .ed, .gd: Swap .ee, .ge: Swap .ef, .gf: Swap .eg, .gg
            End Select
        Case 6
            Select Case r2
            Case 7: Swap .fa, .ga: Swap .fb, .gb: Swap .fc, .gc: Swap .fd, .gd: Swap .fe, .ge: Swap .ff, .gf: Swap .fg, .gg
            End Select
        End Select
    End With
End Function
Public Function Mat8_RowSwap(m As Matrix8, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix8
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat8_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat8_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd: Swap .ae, .be: Swap .af, .bf: Swap .ag, .bg: Swap .ah, .bh
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd: Swap .ae, .ce: Swap .af, .cf: Swap .ag, .cg: Swap .ah, .ch
            Case 4: Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd: Swap .ae, .de: Swap .af, .df: Swap .ag, .dg: Swap .ah, .dh
            Case 5: Swap .aa, .ea: Swap .ab, .eb: Swap .ac, .ec: Swap .ad, .ed: Swap .ae, .ee: Swap .af, .ef: Swap .ag, .eg: Swap .ah, .eh
            Case 6: Swap .aa, .fa: Swap .ab, .fb: Swap .ac, .fc: Swap .ad, .fd: Swap .ae, .fe: Swap .af, .ff: Swap .ag, .fg: Swap .ah, .fh
            Case 7: Swap .aa, .ga: Swap .ab, .gb: Swap .ac, .gc: Swap .ad, .gd: Swap .ae, .ge: Swap .af, .gf: Swap .ag, .gg: Swap .ah, .gh
            Case 8: Swap .aa, .ha: Swap .ab, .hb: Swap .ac, .HC: Swap .ad, .hd: Swap .ae, .he: Swap .af, .hf: Swap .ag, .hg: Swap .ah, .hh
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd: Swap .be, .ce: Swap .bf, .cf: Swap .bg, .cg: Swap .bh, .ch
            Case 4: Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd: Swap .be, .de: Swap .bf, .df: Swap .bg, .dg: Swap .bh, .dh
            Case 5: Swap .ba, .ea: Swap .bb, .eb: Swap .bc, .ec: Swap .bd, .ed: Swap .be, .ee: Swap .bf, .ef: Swap .bg, .eg: Swap .bh, .eh
            Case 6: Swap .ba, .fa: Swap .bb, .fb: Swap .bc, .fc: Swap .bd, .fd: Swap .be, .fe: Swap .bf, .ff: Swap .bg, .fg: Swap .bh, .fh
            Case 7: Swap .ba, .ga: Swap .bb, .gb: Swap .bc, .gc: Swap .bd, .gd: Swap .be, .ge: Swap .bf, .gf: Swap .bg, .gg: Swap .bh, .gh
            Case 8: Swap .ba, .ha: Swap .bb, .hb: Swap .bc, .HC: Swap .bd, .hd: Swap .be, .he: Swap .bf, .hf: Swap .bg, .hg: Swap .bh, .hh
            End Select
        Case 3
            Select Case r2
            Case 4: Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd: Swap .ce, .de: Swap .cf, .df: Swap .cg, .dg: Swap .ch, .dh
            Case 5: Swap .ca, .ea: Swap .cb, .eb: Swap .cc, .ec: Swap .cd, .ed: Swap .ce, .ee: Swap .cf, .ef: Swap .cg, .eg: Swap .ch, .eh
            Case 6: Swap .ca, .fa: Swap .cb, .fb: Swap .cc, .fc: Swap .cd, .fd: Swap .ce, .fe: Swap .cf, .ff: Swap .cg, .fg: Swap .ch, .fh
            Case 7: Swap .ca, .ga: Swap .cb, .gb: Swap .cc, .gc: Swap .cd, .gd: Swap .ce, .ge: Swap .cf, .gf: Swap .cg, .gg: Swap .ch, .gh
            Case 8: Swap .ca, .ha: Swap .cb, .hb: Swap .cc, .HC: Swap .cd, .hd: Swap .ce, .he: Swap .cf, .hf: Swap .cg, .hg: Swap .ch, .hh
            End Select
        Case 4
            Select Case r2
            Case 5: Swap .da, .ea: Swap .db, .eb: Swap .dc, .ec: Swap .dd, .ed: Swap .de, .ee: Swap .df, .ef: Swap .dg, .eg: Swap .dh, .eh
            Case 6: Swap .da, .fa: Swap .db, .fb: Swap .dc, .fc: Swap .dd, .fd: Swap .de, .fe: Swap .df, .ff: Swap .dg, .fg: Swap .dh, .fh
            Case 7: Swap .da, .ga: Swap .db, .gb: Swap .dc, .gc: Swap .dd, .gd: Swap .de, .ge: Swap .df, .gf: Swap .dg, .gg: Swap .dh, .gh
            Case 8: Swap .da, .ha: Swap .db, .hb: Swap .dc, .HC: Swap .hd, .hd: Swap .de, .he: Swap .df, .hf: Swap .dg, .hg: Swap .dh, .hh
            End Select
        Case 5
            Select Case r2
            Case 6: Swap .ea, .fa: Swap .eb, .fb: Swap .ec, .fc: Swap .ed, .fd: Swap .ee, .fe: Swap .ef, .ff: Swap .eg, .fg: Swap .eh, .fh
            Case 7: Swap .ea, .ga: Swap .eb, .gb: Swap .ec, .gc: Swap .ed, .gd: Swap .ee, .ge: Swap .ef, .gf: Swap .eg, .gg: Swap .eh, .gh
            Case 8: Swap .ea, .ha: Swap .eb, .hb: Swap .ec, .HC: Swap .ed, .hd: Swap .ee, .he: Swap .ef, .hf: Swap .eg, .hg: Swap .eh, .hh
            End Select
        Case 6
            Select Case r2
            Case 7: Swap .fa, .ga: Swap .fb, .gb: Swap .fc, .gc: Swap .fd, .gd: Swap .fe, .ge: Swap .ff, .gf: Swap .fg, .gg: Swap .fh, .gh
            Case 8: Swap .fa, .ha: Swap .fb, .hb: Swap .fc, .HC: Swap .fd, .hd: Swap .fe, .he: Swap .ff, .hf: Swap .fg, .hg: Swap .fh, .hh
            End Select
        Case 7
            Select Case r2
            Case 8: Swap .ga, .ha: Swap .gb, .hb: Swap .gc, .HC: Swap .gd, .hd: Swap .ge, .he: Swap .gf, .hf: Swap .gg, .hg: Swap .gh, .hh
            End Select
        End Select
    End With
End Function
Public Function Mat9_RowSwap(m As Matrix9, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix9
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat9_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat9_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2: Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd: Swap .ae, .be: Swap .af, .bf: Swap .ag, .bg: Swap .ah, .bh: Swap .ai, .bi
            Case 3: Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd: Swap .ae, .ce: Swap .af, .cf: Swap .ag, .cg: Swap .ah, .ch: Swap .ai, .ci
            Case 4: Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd: Swap .ae, .de: Swap .af, .df: Swap .ag, .dg: Swap .ah, .dh: Swap .ai, .di
            Case 5: Swap .aa, .ea: Swap .ab, .eb: Swap .ac, .ec: Swap .ad, .ed: Swap .ae, .ee: Swap .af, .ef: Swap .ag, .eg: Swap .ah, .eh: Swap .ai, .ei
            Case 6: Swap .aa, .fa: Swap .ab, .fb: Swap .ac, .fc: Swap .ad, .fd: Swap .ae, .fe: Swap .af, .ff: Swap .ag, .fg: Swap .ah, .fh: Swap .ai, .fi
            Case 7: Swap .aa, .ga: Swap .ab, .gb: Swap .ac, .gc: Swap .ad, .gd: Swap .ae, .ge: Swap .af, .gf: Swap .ag, .gg: Swap .ah, .gh: Swap .ai, .gi
            Case 8: Swap .aa, .ha: Swap .ab, .hb: Swap .ac, .HC: Swap .ad, .hd: Swap .ae, .he: Swap .af, .hf: Swap .ag, .hg: Swap .ah, .hh: Swap .ai, .hi
            Case 9: Swap .aa, .ia: Swap .ab, .ib: Swap .ac, .ic: Swap .ad, .id: Swap .ae, .ie: Swap .af, .if: Swap .ag, .ig: Swap .ah, .ih: Swap .ai, .ii
            End Select
        Case 2
            Select Case r2
            Case 3: Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd: Swap .be, .ce: Swap .bf, .cf: Swap .bg, .cg: Swap .bh, .ch: Swap .bi, .ci
            Case 4: Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd: Swap .be, .de: Swap .bf, .df: Swap .bg, .dg: Swap .bh, .dh: Swap .bi, .di
            Case 5: Swap .ba, .ea: Swap .bb, .eb: Swap .bc, .ec: Swap .bd, .ed: Swap .be, .ee: Swap .bf, .ef: Swap .bg, .eg: Swap .bh, .eh: Swap .bi, .ei
            Case 6: Swap .ba, .fa: Swap .bb, .fb: Swap .bc, .fc: Swap .bd, .fd: Swap .be, .fe: Swap .bf, .ff: Swap .bg, .fg: Swap .bh, .fh: Swap .bi, .fi
            Case 7: Swap .ba, .ga: Swap .bb, .gb: Swap .bc, .gc: Swap .bd, .gd: Swap .be, .ge: Swap .bf, .gf: Swap .bg, .gg: Swap .bh, .gh: Swap .bi, .gi
            Case 8: Swap .ba, .ha: Swap .bb, .hb: Swap .bc, .HC: Swap .bd, .hd: Swap .be, .he: Swap .bf, .hf: Swap .bg, .hg: Swap .bh, .hh: Swap .bi, .hi
            Case 9: Swap .ba, .ia: Swap .bb, .ib: Swap .bc, .ic: Swap .bd, .id: Swap .be, .ie: Swap .bf, .if: Swap .bg, .ig: Swap .bh, .ih: Swap .bi, .ii
            End Select
        Case 3
            Select Case r2
            Case 4: Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd: Swap .ce, .de: Swap .cf, .df: Swap .cg, .dg: Swap .ch, .dh: Swap .ci, .di
            Case 5: Swap .ca, .ea: Swap .cb, .eb: Swap .cc, .ec: Swap .cd, .ed: Swap .ce, .ee: Swap .cf, .ef: Swap .cg, .eg: Swap .ch, .eh: Swap .ci, .ei
            Case 6: Swap .ca, .fa: Swap .cb, .fb: Swap .cc, .fc: Swap .cd, .fd: Swap .ce, .fe: Swap .cf, .ff: Swap .cg, .fg: Swap .ch, .fh: Swap .ci, .fi
            Case 7: Swap .ca, .ga: Swap .cb, .gb: Swap .cc, .gc: Swap .cd, .gd: Swap .ce, .ge: Swap .cf, .gf: Swap .cg, .gg: Swap .ch, .gh: Swap .ci, .gi
            Case 8: Swap .ca, .ha: Swap .cb, .hb: Swap .cc, .HC: Swap .cd, .hd: Swap .ce, .he: Swap .cf, .hf: Swap .cg, .hg: Swap .ch, .hh: Swap .ci, .hi
            Case 9: Swap .ca, .ia: Swap .cb, .ib: Swap .cc, .ic: Swap .cd, .id: Swap .ce, .ie: Swap .cf, .if: Swap .cg, .ig: Swap .ch, .ih: Swap .ci, .ii
            End Select
        Case 4
            Select Case r2
            Case 5: Swap .da, .ea: Swap .db, .eb: Swap .dc, .ec: Swap .dd, .ed: Swap .de, .ee: Swap .df, .ef: Swap .dg, .eg: Swap .dh, .eh: Swap .di, .ei
            Case 6: Swap .da, .fa: Swap .db, .fb: Swap .dc, .fc: Swap .dd, .fd: Swap .de, .fe: Swap .df, .ff: Swap .dg, .fg: Swap .dh, .fh: Swap .di, .fi
            Case 7: Swap .da, .ga: Swap .db, .gb: Swap .dc, .gc: Swap .dd, .gd: Swap .de, .ge: Swap .df, .gf: Swap .dg, .gg: Swap .dh, .gh: Swap .di, .gi
            Case 8: Swap .da, .ha: Swap .db, .hb: Swap .dc, .HC: Swap .hd, .hd: Swap .de, .he: Swap .df, .hf: Swap .dg, .hg: Swap .dh, .hh: Swap .di, .hi
            Case 9: Swap .da, .ia: Swap .db, .ib: Swap .dc, .ic: Swap .hd, .id: Swap .de, .ie: Swap .df, .if: Swap .dg, .ig: Swap .dh, .ih: Swap .di, .ii
            End Select
        Case 5
            Select Case r2
            Case 6: Swap .ea, .fa: Swap .eb, .fb: Swap .ec, .fc: Swap .ed, .fd: Swap .ee, .fe: Swap .ef, .ff: Swap .eg, .fg: Swap .eh, .fh: Swap .ei, .fi
            Case 7: Swap .ea, .ga: Swap .eb, .gb: Swap .ec, .gc: Swap .ed, .gd: Swap .ee, .ge: Swap .ef, .gf: Swap .eg, .gg: Swap .eh, .gh: Swap .ei, .gi
            Case 8: Swap .ea, .ha: Swap .eb, .hb: Swap .ec, .HC: Swap .ed, .hd: Swap .ee, .he: Swap .ef, .hf: Swap .eg, .hg: Swap .eh, .hh: Swap .ei, .hi
            Case 9: Swap .ea, .ia: Swap .eb, .ib: Swap .ec, .ic: Swap .ed, .id: Swap .ee, .ie: Swap .ef, .if: Swap .eg, .ig: Swap .eh, .ih: Swap .ei, .ii
            End Select
        Case 6
            Select Case r2
            Case 7: Swap .fa, .ga: Swap .fb, .gb: Swap .fc, .gc: Swap .fd, .gd: Swap .fe, .ge: Swap .ff, .gf: Swap .fg, .gg: Swap .fh, .gh: Swap .fi, .gi
            Case 8: Swap .fa, .ha: Swap .fb, .hb: Swap .fc, .HC: Swap .fd, .hd: Swap .fe, .he: Swap .ff, .hf: Swap .fg, .hg: Swap .fh, .hh: Swap .fi, .hi
            Case 9: Swap .fa, .ia: Swap .fb, .ib: Swap .fc, .ic: Swap .fd, .id: Swap .fe, .ie: Swap .ff, .if: Swap .fg, .ig: Swap .fh, .ih: Swap .fi, .ii
            End Select
        Case 7
            Select Case r2
            Case 8: Swap .ga, .ha: Swap .gb, .hb: Swap .gc, .HC: Swap .gd, .hd: Swap .ge, .he: Swap .gf, .hf: Swap .gg, .hg: Swap .gh, .hh: Swap .gi, .hi
            Case 9: Swap .ga, .ia: Swap .gb, .ib: Swap .gc, .ic: Swap .gd, .id: Swap .ge, .ie: Swap .gf, .if: Swap .gg, .ig: Swap .gh, .ih: Swap .gi, .ii
            End Select
        Case 8
            Select Case r2
            Case 9: Swap .ha, .ia: Swap .hb, .ib: Swap .HC, .ic: Swap .hd, .id: Swap .he, .ie: Swap .hf, .if: Swap .hg, .ig: Swap .hh, .ih: Swap .hi, .ii
            End Select
        End Select
    End With
End Function
Public Function Mat10_RowSwap(m As Matrix10, ByVal r1 As Byte, ByVal r2 As Byte) As Matrix10
    'Gibt die Matrix m mit 2 vertauschten Zeilen zurück
    Mat10_RowSwap = m
    If r1 = r2 Then Exit Function
    If r2 < r1 Then SwapB r1, r2
    With Mat10_RowSwap
        Select Case r1
        Case 1
            Select Case r2
            Case 2:  Swap .aa, .ba: Swap .ab, .bb: Swap .ac, .bc: Swap .ad, .bd: Swap .ae, .be: Swap .af, .bf: Swap .ag, .bg: Swap .ah, .bh: Swap .ai, .bi: Swap .aj, .bj
            Case 3:  Swap .aa, .ca: Swap .ab, .cb: Swap .ac, .cc: Swap .ad, .cd: Swap .ae, .ce: Swap .af, .cf: Swap .ag, .cg: Swap .ah, .ch: Swap .ai, .ci: Swap .aj, .cj
            Case 4:  Swap .aa, .da: Swap .ab, .db: Swap .ac, .dc: Swap .ad, .dd: Swap .ae, .de: Swap .af, .df: Swap .ag, .dg: Swap .ah, .dh: Swap .ai, .di: Swap .aj, .dj
            Case 5:  Swap .aa, .ea: Swap .ab, .eb: Swap .ac, .ec: Swap .ad, .ed: Swap .ae, .ee: Swap .af, .ef: Swap .ag, .eg: Swap .ah, .eh: Swap .ai, .ei: Swap .aj, .ej
            Case 6:  Swap .aa, .fa: Swap .ab, .fb: Swap .ac, .fc: Swap .ad, .fd: Swap .ae, .fe: Swap .af, .ff: Swap .ag, .fg: Swap .ah, .fh: Swap .ai, .fi: Swap .aj, .fj
            Case 7:  Swap .aa, .ga: Swap .ab, .gb: Swap .ac, .gc: Swap .ad, .gd: Swap .ae, .ge: Swap .af, .gf: Swap .ag, .gg: Swap .ah, .gh: Swap .ai, .gi: Swap .aj, .gj
            Case 8:  Swap .aa, .ha: Swap .ab, .hb: Swap .ac, .HC: Swap .ad, .hd: Swap .ae, .he: Swap .af, .hf: Swap .ag, .hg: Swap .ah, .hh: Swap .ai, .hi: Swap .aj, .hj
            Case 9:  Swap .aa, .ia: Swap .ab, .ib: Swap .ac, .ic: Swap .ad, .id: Swap .ae, .ie: Swap .af, .if: Swap .ag, .ig: Swap .ah, .ih: Swap .ai, .ii: Swap .aj, .ij
            Case 10: Swap .aa, .ja: Swap .ab, .jb: Swap .ac, .jc: Swap .ad, .jd: Swap .ae, .je: Swap .af, .jf: Swap .ag, .jg: Swap .ah, .jh: Swap .ai, .ji: Swap .aj, .jj
            End Select
        Case 2
            Select Case r2
            Case 3:  Swap .ba, .ca: Swap .bb, .cb: Swap .bc, .cc: Swap .bd, .cd: Swap .be, .ce: Swap .bf, .cf: Swap .bg, .cg: Swap .bh, .ch: Swap .bi, .ci: Swap .bj, .cj
            Case 4:  Swap .ba, .da: Swap .bb, .db: Swap .bc, .dc: Swap .bd, .dd: Swap .be, .de: Swap .bf, .df: Swap .bg, .dg: Swap .bh, .dh: Swap .bi, .di: Swap .bj, .dj
            Case 5:  Swap .ba, .ea: Swap .bb, .eb: Swap .bc, .ec: Swap .bd, .ed: Swap .be, .ee: Swap .bf, .ef: Swap .bg, .eg: Swap .bh, .eh: Swap .bi, .ei: Swap .bj, .ej
            Case 6:  Swap .ba, .fa: Swap .bb, .fb: Swap .bc, .fc: Swap .bd, .fd: Swap .be, .fe: Swap .bf, .ff: Swap .bg, .fg: Swap .bh, .fh: Swap .bi, .fi: Swap .bj, .fj
            Case 7:  Swap .ba, .ga: Swap .bb, .gb: Swap .bc, .gc: Swap .bd, .gd: Swap .be, .ge: Swap .bf, .gf: Swap .bg, .gg: Swap .bh, .gh: Swap .bi, .gi: Swap .bj, .gj
            Case 8:  Swap .ba, .ha: Swap .bb, .hb: Swap .bc, .HC: Swap .bd, .hd: Swap .be, .he: Swap .bf, .hf: Swap .bg, .hg: Swap .bh, .hh: Swap .bi, .hi: Swap .bj, .hj
            Case 9:  Swap .ba, .ia: Swap .bb, .ib: Swap .bc, .ic: Swap .bd, .id: Swap .be, .ie: Swap .bf, .if: Swap .bg, .ig: Swap .bh, .ih: Swap .bi, .ii: Swap .bj, .ij
            Case 10: Swap .ba, .ja: Swap .bb, .jb: Swap .bc, .jc: Swap .bd, .jd: Swap .be, .je: Swap .bf, .jf: Swap .bg, .jg: Swap .bh, .jh: Swap .bi, .ji: Swap .bj, .jj
            End Select
        Case 3
            Select Case r2
            Case 4:  Swap .ca, .da: Swap .cb, .db: Swap .cc, .dc: Swap .cd, .dd: Swap .ce, .de: Swap .cf, .df: Swap .cg, .dg: Swap .ch, .dh: Swap .ci, .di: Swap .cj, .dj
            Case 5:  Swap .ca, .ea: Swap .cb, .eb: Swap .cc, .ec: Swap .cd, .ed: Swap .ce, .ee: Swap .cf, .ef: Swap .cg, .eg: Swap .ch, .eh: Swap .ci, .ei: Swap .cj, .ej
            Case 6:  Swap .ca, .fa: Swap .cb, .fb: Swap .cc, .fc: Swap .cd, .fd: Swap .ce, .fe: Swap .cf, .ff: Swap .cg, .fg: Swap .ch, .fh: Swap .ci, .fi: Swap .cj, .fj
            Case 7:  Swap .ca, .ga: Swap .cb, .gb: Swap .cc, .gc: Swap .cd, .gd: Swap .ce, .ge: Swap .cf, .gf: Swap .cg, .gg: Swap .ch, .gh: Swap .ci, .gi: Swap .cj, .gj
            Case 8:  Swap .ca, .ha: Swap .cb, .hb: Swap .cc, .HC: Swap .cd, .hd: Swap .ce, .he: Swap .cf, .hf: Swap .cg, .hg: Swap .ch, .hh: Swap .ci, .hi: Swap .cj, .hj
            Case 9:  Swap .ca, .ia: Swap .cb, .ib: Swap .cc, .ic: Swap .cd, .id: Swap .ce, .ie: Swap .cf, .if: Swap .cg, .ig: Swap .ch, .ih: Swap .ci, .ii: Swap .cj, .ij
            Case 10: Swap .ca, .ja: Swap .cb, .jb: Swap .cc, .jc: Swap .cd, .jd: Swap .ce, .je: Swap .cf, .jf: Swap .cg, .jg: Swap .ch, .jh: Swap .ci, .ji: Swap .cj, .jj
            End Select
        Case 4
            Select Case r2
            Case 5:  Swap .da, .ea: Swap .db, .eb: Swap .dc, .ec: Swap .dd, .ed: Swap .de, .ee: Swap .df, .ef: Swap .dg, .eg: Swap .dh, .eh: Swap .di, .ei: Swap .dj, .ej
            Case 6:  Swap .da, .fa: Swap .db, .fb: Swap .dc, .fc: Swap .dd, .fd: Swap .de, .fe: Swap .df, .ff: Swap .dg, .fg: Swap .dh, .fh: Swap .di, .fi: Swap .dj, .fj
            Case 7:  Swap .da, .ga: Swap .db, .gb: Swap .dc, .gc: Swap .dd, .gd: Swap .de, .ge: Swap .df, .gf: Swap .dg, .gg: Swap .dh, .gh: Swap .di, .gi: Swap .dj, .gj
            Case 8:  Swap .da, .ha: Swap .db, .hb: Swap .dc, .HC: Swap .hd, .hd: Swap .de, .he: Swap .df, .hf: Swap .dg, .hg: Swap .dh, .hh: Swap .di, .hi: Swap .dj, .hj
            Case 9:  Swap .da, .ia: Swap .db, .ib: Swap .dc, .ic: Swap .hd, .id: Swap .de, .ie: Swap .df, .if: Swap .dg, .ig: Swap .dh, .ih: Swap .di, .ii: Swap .dj, .ij
            Case 10: Swap .da, .ja: Swap .db, .jb: Swap .dc, .jc: Swap .hd, .jd: Swap .de, .je: Swap .df, .jf: Swap .dg, .jg: Swap .dh, .jh: Swap .di, .ji: Swap .dj, .jj
            End Select
        Case 5
            Select Case r2
            Case 6:  Swap .ea, .fa: Swap .eb, .fb: Swap .ec, .fc: Swap .ed, .fd: Swap .ee, .fe: Swap .ef, .ff: Swap .eg, .fg: Swap .eh, .fh: Swap .ei, .fi: Swap .ej, .fj
            Case 7:  Swap .ea, .ga: Swap .eb, .gb: Swap .ec, .gc: Swap .ed, .gd: Swap .ee, .ge: Swap .ef, .gf: Swap .eg, .gg: Swap .eh, .gh: Swap .ei, .gi: Swap .ej, .gj
            Case 8:  Swap .ea, .ha: Swap .eb, .hb: Swap .ec, .HC: Swap .ed, .hd: Swap .ee, .he: Swap .ef, .hf: Swap .eg, .hg: Swap .eh, .hh: Swap .ei, .hi: Swap .ej, .hj
            Case 9:  Swap .ea, .ia: Swap .eb, .ib: Swap .ec, .ic: Swap .ed, .id: Swap .ee, .ie: Swap .ef, .if: Swap .eg, .ig: Swap .eh, .ih: Swap .ei, .ii: Swap .ej, .ij
            Case 10: Swap .ea, .ja: Swap .eb, .jb: Swap .ec, .jc: Swap .ed, .jd: Swap .ee, .je: Swap .ef, .jf: Swap .eg, .jg: Swap .eh, .jh: Swap .ei, .ji: Swap .ej, .jj
            End Select
        Case 6
            Select Case r2
            Case 7:  Swap .fa, .ga: Swap .fb, .gb: Swap .fc, .gc: Swap .fd, .gd: Swap .fe, .ge: Swap .ff, .gf: Swap .fg, .gg: Swap .fh, .gh: Swap .fi, .gi: Swap .fj, .gj
            Case 8:  Swap .fa, .ha: Swap .fb, .hb: Swap .fc, .HC: Swap .fd, .hd: Swap .fe, .he: Swap .ff, .hf: Swap .fg, .hg: Swap .fh, .hh: Swap .fi, .hi: Swap .fj, .hj
            Case 9:  Swap .fa, .ia: Swap .fb, .ib: Swap .fc, .ic: Swap .fd, .id: Swap .fe, .ie: Swap .ff, .if: Swap .fg, .ig: Swap .fh, .ih: Swap .fi, .ii: Swap .fj, .ij
            Case 10: Swap .fa, .ja: Swap .fb, .jb: Swap .fc, .jc: Swap .fd, .jd: Swap .fe, .je: Swap .ff, .jf: Swap .fg, .jg: Swap .fh, .jh: Swap .fi, .ji: Swap .fj, .jj
            End Select
        Case 7
            Select Case r2
            Case 8:  Swap .ga, .ha: Swap .gb, .hb: Swap .gc, .HC: Swap .gd, .hd: Swap .ge, .he: Swap .gf, .hf: Swap .gg, .hg: Swap .gh, .hh: Swap .gi, .hi: Swap .gj, .hj
            Case 9:  Swap .ga, .ia: Swap .gb, .ib: Swap .gc, .ic: Swap .gd, .id: Swap .ge, .ie: Swap .gf, .if: Swap .gg, .ig: Swap .gh, .ih: Swap .gi, .ii: Swap .gj, .ij
            Case 10: Swap .ga, .ja: Swap .gb, .jb: Swap .gc, .jc: Swap .gd, .jd: Swap .ge, .je: Swap .gf, .jf: Swap .gg, .jg: Swap .gh, .jh: Swap .gi, .ji: Swap .gj, .jj
            End Select
        Case 8
            Select Case r2
            Case 9:  Swap .ha, .ia: Swap .hb, .ib: Swap .HC, .ic: Swap .hd, .id: Swap .he, .ie: Swap .hf, .if: Swap .hg, .ig: Swap .hh, .ih: Swap .hi, .ii: Swap .hj, .ij
            Case 10: Swap .ha, .ja: Swap .hb, .jb: Swap .HC, .jc: Swap .hd, .jd: Swap .he, .je: Swap .hf, .jf: Swap .hg, .jg: Swap .hh, .jh: Swap .hi, .ji: Swap .hj, .jj
            End Select
        Case 9
            Select Case r2
            Case 10: Swap .ia, .ja: Swap .ib, .jb: Swap .ic, .jc: Swap .id, .jd: Swap .ie, .je: Swap .if, .jf: Swap .ig, .jg: Swap .ih, .jh: Swap .ii, .ji: Swap .hj, .jj
            End Select
        End Select
    End With
End Function

Public Function Mat2_ColSwap(m As Matrix2, ByVal c1 As Byte, ByVal c2 As Byte) As Matrix2
    'Gibt die Matrix m mit 2 vertauschten Spalten zurück
    If c1 = c2 Then Mat2_ColSwap = m: Exit Function
    With Mat2_ColSwap
        .aa = m.ab: .ab = m.aa
        .ab = m.bb: .bb = m.ab
    End With
End Function
Public Function Mat3_ColSwap(m As Matrix3, ByVal c1 As Byte, ByVal c2 As Byte) As Matrix3
    'Gibt die Matrix m mit 2 vertauschten Spalten zurück
    'TODO: hier weiter!
    Mat3_ColSwap = m
    If c1 = c2 Then Exit Function
    If c2 < c1 Then SwapB c1, c2
    With Mat3_ColSwap
        Select Case c1
        Case 1
            Select Case c2
            Case 2: Swap .aa, .ab: Swap .ba, .bb: Swap .ca, .cb
            Case 3: Swap .aa, .ac: Swap .ba, .bc: Swap .ca, .cc
            End Select
        Case 2
            Select Case c2
            Case 3: Swap .ab, .ac: Swap .bb, .bc: Swap .cb, .cc
            End Select
        End Select
    End With
End Function
Public Function Mat4_ColSwap(m As Matrix4, ByVal c1 As Byte, ByVal c2 As Byte) As Matrix4
    'Gibt die Matrix m mit 2 vertauschten Spalten zurück
    'TODO: hier weiter!
    Mat4_ColSwap = m
    If c1 = c2 Then Exit Function
    If c2 < c1 Then SwapB c1, c2
    With Mat4_ColSwap
        Select Case c1
        Case 1
            Select Case c2
            Case 2: Swap .aa, .ab: Swap .ba, .bb: Swap .ca, .cb: Swap .da, .db
            Case 3: Swap .aa, .ac: Swap .ba, .bc: Swap .ca, .cc: Swap .da, .dc
            Case 4: Swap .aa, .ad: Swap .ba, .bd: Swap .ca, .cd: Swap .da, .dd
            End Select
        Case 2
            Select Case c2
            Case 3: Swap .ab, .ac: Swap .bb, .bc: Swap .cb, .cc: Swap .db, .dc
            Case 4: Swap .ab, .ad: Swap .bb, .bd: Swap .cb, .cd: Swap .db, .dd
            End Select
        Case 3
            Select Case c2
            Case 4: Swap .ac, .ad: Swap .bc, .bd: Swap .cc, .cd: Swap .dc, .dd
            End Select
        End Select
    End With
End Function


'Untermatrix
Public Function Mat2_umat(m As Matrix2, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    Select Case r_ex: Case 0: Select Case c_ex: Case 0: Mat2_umat = m.bb
                                                Case 1: Mat2_umat = m.ba: End Select
                      Case 1: Select Case c_ex: Case 0: Mat2_umat = m.ab
                                                Case 1: Mat2_umat = m.aa: End Select: End Select
End Function
Public Function Mat3_umat(m As Matrix3, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix2
    'Liefert aus einer 3x3-Matrix die 2x2-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat2_Row(Mat3_umat, icex) = Vec3_uvec(Mat3_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat2_Row(Mat3_umat, icex) = Vec3_uvec(Mat3_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat2_Row(Mat3_umat, icex) = Vec3_uvec(Mat3_Row(m, 2), c_ex): icex = icex + 1
End Function
Public Function Mat4_umat(m As Matrix4, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix3
    'Liefert aus einer 4x4-Matrix die 3x3-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat3_Row(Mat4_umat, icex) = Vec4_uvec(Mat4_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat3_Row(Mat4_umat, icex) = Vec4_uvec(Mat4_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat3_Row(Mat4_umat, icex) = Vec4_uvec(Mat4_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat3_Row(Mat4_umat, icex) = Vec4_uvec(Mat4_Row(m, 3), c_ex): icex = icex + 1
End Function
Public Function Mat5_umat(m As Matrix5, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix4
    'Liefert aus einer 5x5-Matrix die 4x4-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat4_Row(Mat5_umat, icex) = Vec5_uvec(Mat5_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat4_Row(Mat5_umat, icex) = Vec5_uvec(Mat5_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat4_Row(Mat5_umat, icex) = Vec5_uvec(Mat5_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat4_Row(Mat5_umat, icex) = Vec5_uvec(Mat5_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat4_Row(Mat5_umat, icex) = Vec5_uvec(Mat5_Row(m, 4), c_ex): icex = icex + 1
End Function
Public Function Mat6_umat(m As Matrix6, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix5
    'Liefert aus einer 6x6-Matrix die 5x5-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat5_Row(Mat6_umat, icex) = Vec6_uvec(Mat6_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat5_Row(Mat6_umat, icex) = Vec6_uvec(Mat6_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat5_Row(Mat6_umat, icex) = Vec6_uvec(Mat6_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat5_Row(Mat6_umat, icex) = Vec6_uvec(Mat6_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat5_Row(Mat6_umat, icex) = Vec6_uvec(Mat6_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat5_Row(Mat6_umat, icex) = Vec6_uvec(Mat6_Row(m, 5), c_ex): icex = icex + 1
End Function
Public Function Mat7_umat(m As Matrix7, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix6
    'Liefert aus einer 7x7-Matrix die 6x6-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat6_Row(Mat7_umat, icex) = Vec7_uvec(Mat7_Row(m, 6), c_ex): icex = icex + 1
End Function
Public Function Mat8_umat(m As Matrix8, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix7
    'Liefert aus einer 8x8-Matrix die 7x7-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 6), c_ex): icex = icex + 1
    If r_ex <> 7 Then Mat7_Row(Mat8_umat, icex) = Vec8_uvec(Mat8_Row(m, 7), c_ex): icex = icex + 1
End Function
Public Function Mat9_umat(m As Matrix9, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix8
    'Liefert aus einer 9x9-Matrix die 8x8-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 6), c_ex): icex = icex + 1
    If r_ex <> 7 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 7), c_ex): icex = icex + 1
    If r_ex <> 8 Then Mat8_Row(Mat9_umat, icex) = Vec9_uvec(Mat9_Row(m, 8), c_ex): icex = icex + 1
End Function
Public Function Mat10_umat(m As Matrix10, ByVal r_ex As Long, ByVal c_ex As Long) As Matrix9
    'Liefert aus einer 10x10-Matrix die 9x9-Untermatrix, indem die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Dim icex As Long
    If r_ex <> 0 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 0), c_ex): icex = icex + 1
    If r_ex <> 1 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 1), c_ex): icex = icex + 1
    If r_ex <> 2 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 2), c_ex): icex = icex + 1
    If r_ex <> 3 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 3), c_ex): icex = icex + 1
    If r_ex <> 4 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 4), c_ex): icex = icex + 1
    If r_ex <> 5 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 5), c_ex): icex = icex + 1
    If r_ex <> 6 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 6), c_ex): icex = icex + 1
    If r_ex <> 7 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 7), c_ex): icex = icex + 1
    If r_ex <> 8 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 8), c_ex): icex = icex + 1
    If r_ex <> 9 Then Mat9_Row(Mat10_umat, icex) = Vec10_uvec(Mat10_Row(m, 9), c_ex): icex = icex + 1
End Function

'Berechnet die Minoren = Determinanten der Untermatrizen
Public Function Mat2_min(m As Matrix2, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 2x2-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat2_min = Mat2_umat(m, r_ex, c_ex)
End Function
Public Function Mat3_min(m As Matrix3, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 3x3-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat3_min = Mat2_det(Mat3_umat(m, r_ex, c_ex))
End Function
Public Function Mat4_min(m As Matrix4, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 4x4-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat4_min = Mat3_det(Mat4_umat(m, r_ex, c_ex))
End Function
Public Function Mat5_min(m As Matrix5, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 5x5-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat5_min = Mat4_det(Mat5_umat(m, r_ex, c_ex))
End Function
Public Function Mat6_min(m As Matrix6, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat6_min = Mat5_det(Mat6_umat(m, r_ex, c_ex))
End Function
Public Function Mat7_min(m As Matrix7, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat7_min = Mat6_det(Mat7_umat(m, r_ex, c_ex))
End Function
Public Function Mat8_min(m As Matrix8, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat8_min = Mat7_det(Mat8_umat(m, r_ex, c_ex))
End Function
Public Function Mat9_min(m As Matrix9, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat9_min = Mat8_det(Mat9_umat(m, r_ex, c_ex))
End Function
Public Function Mat10_min(m As Matrix10, ByVal r_ex As Long, ByVal c_ex As Long) As Double
    'Berechnet die Minore einer 6x6-Matrix bei der die Zeile r_ex und die Spalte c_ex ausgelassen werden.
    Mat10_min = Mat9_det(Mat10_umat(m, r_ex, c_ex))
End Function

'Adjunkte
Public Function Mat2_Adj(m As Matrix2) As Matrix2
    With m
        Mat2_Adj = Mat2(.bb, -.ab, _
                           -.ba, .aa)
    End With
End Function
Public Function Mat3_Adj(m As Matrix3) As Matrix3
    Mat3_Adj = Mat3(Mat3_min(m, 0, 0), -Mat3_min(m, 1, 0), Mat3_min(m, 2, 0), _
                   -Mat3_min(m, 0, 1), Mat3_min(m, 1, 1), -Mat3_min(m, 2, 1), _
                    Mat3_min(m, 0, 2), -Mat3_min(m, 1, 2), Mat3_min(m, 2, 2))
End Function
Public Function Mat4_Adj(m As Matrix4) As Matrix4
    Mat4_Adj = Mat4(Mat4_min(m, 0, 0), -Mat4_min(m, 1, 0), Mat4_min(m, 2, 0), -Mat4_min(m, 3, 0), _
                   -Mat4_min(m, 0, 1), Mat4_min(m, 1, 1), -Mat4_min(m, 2, 1), Mat4_min(m, 3, 1), _
                    Mat4_min(m, 0, 2), -Mat4_min(m, 1, 2), Mat4_min(m, 2, 2), -Mat4_min(m, 3, 2), _
                   -Mat4_min(m, 0, 3), Mat4_min(m, 1, 3), -Mat4_min(m, 2, 3), Mat4_min(m, 3, 3))
End Function
Public Function Mat5_Adj(m As Matrix5) As Matrix5
    Mat5_Adj = Mat5(Mat5_min(m, 0, 0), -Mat5_min(m, 1, 0), Mat5_min(m, 2, 0), -Mat5_min(m, 3, 0), Mat5_min(m, 4, 0), _
                   -Mat5_min(m, 0, 1), Mat5_min(m, 1, 1), -Mat5_min(m, 2, 1), Mat5_min(m, 3, 1), -Mat5_min(m, 4, 1), _
                    Mat5_min(m, 0, 2), -Mat5_min(m, 1, 2), Mat5_min(m, 2, 2), -Mat5_min(m, 3, 2), Mat5_min(m, 4, 2), _
                   -Mat5_min(m, 0, 3), Mat5_min(m, 1, 3), -Mat5_min(m, 2, 3), Mat5_min(m, 3, 3), -Mat5_min(m, 4, 3), _
                    Mat5_min(m, 0, 4), -Mat5_min(m, 1, 4), Mat5_min(m, 2, 4), -Mat5_min(m, 3, 4), Mat5_min(m, 4, 4))
End Function
Public Function Mat6_Adj(m As Matrix6) As Matrix6
    Mat6_Adj = Mat6(Mat6_min(m, 0, 0), -Mat6_min(m, 1, 0), Mat6_min(m, 2, 0), -Mat6_min(m, 3, 0), Mat6_min(m, 4, 0), -Mat6_min(m, 5, 0), _
                   -Mat6_min(m, 0, 1), Mat6_min(m, 1, 1), -Mat6_min(m, 2, 1), Mat6_min(m, 3, 1), -Mat6_min(m, 4, 1), Mat6_min(m, 5, 1), _
                    Mat6_min(m, 0, 2), -Mat6_min(m, 1, 2), Mat6_min(m, 2, 2), -Mat6_min(m, 3, 2), Mat6_min(m, 4, 2), -Mat6_min(m, 5, 2), _
                   -Mat6_min(m, 0, 3), Mat6_min(m, 1, 3), -Mat6_min(m, 2, 3), Mat6_min(m, 3, 3), -Mat6_min(m, 4, 3), Mat6_min(m, 5, 3), _
                    Mat6_min(m, 0, 4), -Mat6_min(m, 1, 4), Mat6_min(m, 2, 4), -Mat6_min(m, 3, 4), Mat6_min(m, 4, 4), -Mat6_min(m, 5, 4), _
                   -Mat6_min(m, 0, 5), Mat6_min(m, 1, 5), -Mat6_min(m, 2, 5), Mat6_min(m, 3, 5), -Mat6_min(m, 4, 5), Mat6_min(m, 5, 5))
End Function
Public Function Mat7_Adj(m As Matrix7) As Matrix7
    Mat7_Adj = Mat7(Mat7_min(m, 0, 0), -Mat7_min(m, 1, 0), Mat7_min(m, 2, 0), -Mat7_min(m, 3, 0), Mat7_min(m, 4, 0), -Mat7_min(m, 5, 0), Mat7_min(m, 6, 0), _
                   -Mat7_min(m, 0, 1), Mat7_min(m, 1, 1), -Mat7_min(m, 2, 1), Mat7_min(m, 3, 1), -Mat7_min(m, 4, 1), Mat7_min(m, 5, 1), -Mat7_min(m, 6, 1), _
                    Mat7_min(m, 0, 2), -Mat7_min(m, 1, 2), Mat7_min(m, 2, 2), -Mat7_min(m, 3, 2), Mat7_min(m, 4, 2), -Mat7_min(m, 5, 2), Mat7_min(m, 6, 2), _
                   -Mat7_min(m, 0, 3), Mat7_min(m, 1, 3), -Mat7_min(m, 2, 3), Mat7_min(m, 3, 3), -Mat7_min(m, 4, 3), Mat7_min(m, 5, 3), -Mat7_min(m, 6, 3), _
                    Mat7_min(m, 0, 4), -Mat7_min(m, 1, 4), Mat7_min(m, 2, 4), -Mat7_min(m, 3, 4), Mat7_min(m, 4, 4), -Mat7_min(m, 5, 4), Mat7_min(m, 6, 4), _
                   -Mat7_min(m, 0, 5), Mat7_min(m, 1, 5), -Mat7_min(m, 2, 5), Mat7_min(m, 3, 5), -Mat7_min(m, 4, 5), Mat7_min(m, 5, 5), -Mat7_min(m, 6, 5), _
                    Mat7_min(m, 0, 6), -Mat7_min(m, 1, 6), Mat7_min(m, 2, 6), -Mat7_min(m, 3, 6), Mat7_min(m, 4, 6), -Mat7_min(m, 5, 6), Mat7_min(m, 6, 6))
End Function
Public Function Mat8_Adj(m As Matrix8) As Matrix8
    Mat8_Adj = Mat8(Vec8(Mat8_min(m, 0, 0), -Mat8_min(m, 1, 0), Mat8_min(m, 2, 0), -Mat8_min(m, 3, 0), Mat8_min(m, 4, 0), -Mat8_min(m, 5, 0), Mat8_min(m, 6, 0), -Mat8_min(m, 7, 0)), _
                    Vec8(-Mat8_min(m, 0, 1), Mat8_min(m, 1, 1), -Mat8_min(m, 2, 1), Mat8_min(m, 3, 1), -Mat8_min(m, 4, 1), Mat8_min(m, 5, 1), -Mat8_min(m, 6, 1), Mat8_min(m, 7, 1)), _
                    Vec8(Mat8_min(m, 0, 2), -Mat8_min(m, 1, 2), Mat8_min(m, 2, 2), -Mat8_min(m, 3, 2), Mat8_min(m, 4, 2), -Mat8_min(m, 5, 2), Mat8_min(m, 6, 2), -Mat8_min(m, 7, 2)), _
                    Vec8(-Mat8_min(m, 0, 3), Mat8_min(m, 1, 3), -Mat8_min(m, 2, 3), Mat8_min(m, 3, 3), -Mat8_min(m, 4, 3), Mat8_min(m, 5, 3), -Mat8_min(m, 6, 3), Mat8_min(m, 7, 3)), _
                    Vec8(Mat8_min(m, 0, 4), -Mat8_min(m, 1, 4), Mat8_min(m, 2, 4), -Mat8_min(m, 3, 4), Mat8_min(m, 4, 4), -Mat8_min(m, 5, 4), Mat8_min(m, 6, 4), -Mat8_min(m, 7, 4)), _
                    Vec8(-Mat8_min(m, 0, 5), Mat8_min(m, 1, 5), -Mat8_min(m, 2, 5), Mat8_min(m, 3, 5), -Mat8_min(m, 4, 5), Mat8_min(m, 5, 5), -Mat8_min(m, 6, 5), Mat8_min(m, 7, 5)), _
                    Vec8(Mat8_min(m, 0, 6), -Mat8_min(m, 1, 6), Mat8_min(m, 2, 6), -Mat8_min(m, 3, 6), Mat8_min(m, 4, 6), -Mat8_min(m, 5, 6), Mat8_min(m, 6, 6), -Mat8_min(m, 7, 6)), _
                    Vec8(-Mat8_min(m, 0, 7), Mat8_min(m, 1, 7), -Mat8_min(m, 2, 7), Mat8_min(m, 3, 7), -Mat8_min(m, 4, 7), Mat8_min(m, 5, 7), -Mat8_min(m, 6, 7), Mat8_min(m, 7, 7)))
End Function
Public Function Mat9_Adj(m As Matrix9) As Matrix9
    Mat9_Adj = Mat9(Vec9(Mat9_min(m, 0, 0), -Mat9_min(m, 1, 0), Mat9_min(m, 2, 0), -Mat9_min(m, 3, 0), Mat9_min(m, 4, 0), -Mat9_min(m, 5, 0), Mat9_min(m, 6, 0), -Mat9_min(m, 7, 0), Mat9_min(m, 8, 0)), _
                    Vec9(-Mat9_min(m, 0, 1), Mat9_min(m, 1, 1), -Mat9_min(m, 2, 1), Mat9_min(m, 3, 1), -Mat9_min(m, 4, 1), Mat9_min(m, 5, 1), -Mat9_min(m, 6, 1), Mat9_min(m, 7, 1), -Mat9_min(m, 8, 1)), _
                    Vec9(Mat9_min(m, 0, 2), -Mat9_min(m, 1, 2), Mat9_min(m, 2, 2), -Mat9_min(m, 3, 2), Mat9_min(m, 4, 2), -Mat9_min(m, 5, 2), Mat9_min(m, 6, 2), -Mat9_min(m, 7, 2), Mat9_min(m, 8, 2)), _
                    Vec9(-Mat9_min(m, 0, 3), Mat9_min(m, 1, 3), -Mat9_min(m, 2, 3), Mat9_min(m, 3, 3), -Mat9_min(m, 4, 3), Mat9_min(m, 5, 3), -Mat9_min(m, 6, 3), Mat9_min(m, 7, 3), -Mat9_min(m, 8, 3)), _
                    Vec9(Mat9_min(m, 0, 4), -Mat9_min(m, 1, 4), Mat9_min(m, 2, 4), -Mat9_min(m, 3, 4), Mat9_min(m, 4, 4), -Mat9_min(m, 5, 4), Mat9_min(m, 6, 4), -Mat9_min(m, 7, 4), Mat9_min(m, 8, 4)), _
                    Vec9(-Mat9_min(m, 0, 5), Mat9_min(m, 1, 5), -Mat9_min(m, 2, 5), Mat9_min(m, 3, 5), -Mat9_min(m, 4, 5), Mat9_min(m, 5, 5), -Mat9_min(m, 6, 5), Mat9_min(m, 7, 5), -Mat9_min(m, 8, 5)), _
                    Vec9(Mat9_min(m, 0, 6), -Mat9_min(m, 1, 6), Mat9_min(m, 2, 6), -Mat9_min(m, 3, 6), Mat9_min(m, 4, 6), -Mat9_min(m, 5, 6), Mat9_min(m, 6, 6), -Mat9_min(m, 7, 6), Mat9_min(m, 8, 6)), _
                    Vec9(-Mat9_min(m, 0, 7), Mat9_min(m, 1, 7), -Mat9_min(m, 2, 7), Mat9_min(m, 3, 7), -Mat9_min(m, 4, 7), Mat9_min(m, 5, 7), -Mat9_min(m, 6, 7), Mat9_min(m, 7, 7), -Mat9_min(m, 8, 7)), _
                    Vec9(Mat9_min(m, 0, 8), -Mat9_min(m, 1, 8), Mat9_min(m, 2, 8), -Mat9_min(m, 3, 8), Mat9_min(m, 4, 8), -Mat9_min(m, 5, 8), Mat9_min(m, 6, 8), -Mat9_min(m, 7, 8), Mat9_min(m, 8, 8)))
End Function
Public Function Mat10_Adj(m As Matrix10) As Matrix10
    Mat10_Adj = Mat10(Vec10(Mat10_min(m, 0, 0), -Mat10_min(m, 1, 0), Mat10_min(m, 2, 0), -Mat10_min(m, 3, 0), Mat10_min(m, 4, 0), -Mat10_min(m, 5, 0), Mat10_min(m, 6, 0), -Mat10_min(m, 7, 0), Mat10_min(m, 8, 0), -Mat10_min(m, 9, 0)), _
                      Vec10(-Mat10_min(m, 0, 1), Mat10_min(m, 1, 1), -Mat10_min(m, 2, 1), Mat10_min(m, 3, 1), -Mat10_min(m, 4, 1), Mat10_min(m, 5, 1), -Mat10_min(m, 6, 1), Mat10_min(m, 7, 1), -Mat10_min(m, 8, 1), Mat10_min(m, 9, 1)), _
                      Vec10(Mat10_min(m, 0, 2), -Mat10_min(m, 1, 2), Mat10_min(m, 2, 2), -Mat10_min(m, 3, 2), Mat10_min(m, 4, 2), -Mat10_min(m, 5, 2), Mat10_min(m, 6, 2), -Mat10_min(m, 7, 2), Mat10_min(m, 8, 2), -Mat10_min(m, 9, 2)), _
                      Vec10(-Mat10_min(m, 0, 3), Mat10_min(m, 1, 3), -Mat10_min(m, 2, 3), Mat10_min(m, 3, 3), -Mat10_min(m, 4, 3), Mat10_min(m, 5, 3), -Mat10_min(m, 6, 3), Mat10_min(m, 7, 3), -Mat10_min(m, 8, 3), Mat10_min(m, 9, 3)), _
                      Vec10(Mat10_min(m, 0, 4), -Mat10_min(m, 1, 4), Mat10_min(m, 2, 4), -Mat10_min(m, 3, 4), Mat10_min(m, 4, 4), -Mat10_min(m, 5, 4), Mat10_min(m, 6, 4), -Mat10_min(m, 7, 4), Mat10_min(m, 8, 4), -Mat10_min(m, 9, 4)), _
                      Vec10(-Mat10_min(m, 0, 5), Mat10_min(m, 1, 5), -Mat10_min(m, 2, 5), Mat10_min(m, 3, 5), -Mat10_min(m, 4, 5), Mat10_min(m, 5, 5), -Mat10_min(m, 6, 5), Mat10_min(m, 7, 5), -Mat10_min(m, 8, 5), Mat10_min(m, 9, 5)), _
                      Vec10(Mat10_min(m, 0, 6), -Mat10_min(m, 1, 6), Mat10_min(m, 2, 6), -Mat10_min(m, 3, 6), Mat10_min(m, 4, 6), -Mat10_min(m, 5, 6), Mat10_min(m, 6, 6), -Mat10_min(m, 7, 6), Mat10_min(m, 8, 6), -Mat10_min(m, 9, 6)), _
                      Vec10(-Mat10_min(m, 0, 7), Mat10_min(m, 1, 7), -Mat10_min(m, 2, 7), Mat10_min(m, 3, 7), -Mat10_min(m, 4, 7), Mat10_min(m, 5, 7), -Mat10_min(m, 6, 7), Mat10_min(m, 7, 7), -Mat10_min(m, 8, 7), Mat10_min(m, 9, 7)), _
                      Vec10(Mat10_min(m, 0, 8), -Mat10_min(m, 1, 8), Mat10_min(m, 2, 8), -Mat10_min(m, 3, 8), Mat10_min(m, 4, 8), -Mat10_min(m, 5, 8), Mat10_min(m, 6, 8), -Mat10_min(m, 7, 8), Mat10_min(m, 8, 8), -Mat10_min(m, 9, 8)), _
                      Vec10(-Mat10_min(m, 0, 9), Mat10_min(m, 1, 9), -Mat10_min(m, 2, 9), Mat10_min(m, 3, 9), -Mat10_min(m, 4, 9), Mat10_min(m, 5, 9), -Mat10_min(m, 6, 9), Mat10_min(m, 7, 9), -Mat10_min(m, 8, 9), Mat10_min(m, 9, 9)))
End Function

'Inverse
Public Function Mat2_inv(m As Matrix2) As Matrix2
    Dim det As Double: det = Mat2_det(m): If det = 0 Then Exit Function
    Mat2_inv = Mat2_smul(Mat2_Adj(m), 1 / det)
End Function
Public Function Mat3_inv(m As Matrix3) As Matrix3
    Dim det As Double: det = Mat3_det(m): If det = 0 Then Exit Function
    Mat3_inv = Mat3_smul(Mat3_Adj(m), 1 / det)
End Function
Public Function Mat4_inv(m As Matrix4) As Matrix4
    Dim det As Double: det = Mat4_det(m): If det = 0 Then Exit Function
    Mat4_inv = Mat4_smul(Mat4_Adj(m), 1 / det)
End Function
Public Function Mat5_inv(m As Matrix5) As Matrix5
    Dim det As Double: det = Mat5_det(m): If det = 0 Then Exit Function
    Mat5_inv = Mat5_smul(Mat5_Adj(m), 1 / det)
End Function
Public Function Mat6_inv(m As Matrix6) As Matrix6
    Dim det As Double: det = Mat6_det(m): If det = 0 Then Exit Function
    Mat6_inv = Mat6_smul(Mat6_Adj(m), 1 / det)
End Function
Public Function Mat7_inv(m As Matrix7) As Matrix7
    Dim det As Double: det = Mat7_det(m): If det = 0 Then Exit Function
    Mat7_inv = Mat7_smul(Mat7_Adj(m), 1 / det)
End Function
Public Function Mat8_inv(m As Matrix8) As Matrix8
    Dim det As Double: det = Mat8_det(m): If det = 0 Then Exit Function
    Mat8_inv = Mat8_smul(Mat8_Adj(m), 1 / det)
End Function
Public Function Mat9_inv(m As Matrix9) As Matrix9
    Dim det As Double: det = Mat9_det(m): If det = 0 Then Exit Function
    Mat9_inv = Mat9_smul(Mat9_Adj(m), 1 / det)
End Function
Public Function Mat10_inv(m As Matrix10) As Matrix10
    Dim det As Double: det = Mat10_det(m): If det = 0 Then Exit Function
    Mat10_inv = Mat10_smul(Mat10_Adj(m), 1 / det)
End Function

'Lösen von LGS
Public Function Mat2_solve(m As Matrix2, b As Vector2) As Vector2
    Mat2_solve = Mat2_vmul(Mat2_inv(m), b)
End Function
Public Function Mat3_solve(m As Matrix3, b As Vector3) As Vector3
    Mat3_solve = Mat3_vmul(Mat3_inv(m), b)
End Function
Public Function Mat4_solve(m As Matrix4, b As Vector4) As Vector4
    Mat4_solve = Mat4_vmul(Mat4_inv(m), b)
End Function
Public Function Mat5_solve(m As Matrix5, b As Vector5) As Vector5
    Mat5_solve = Mat5_vmul(Mat5_inv(m), b)
End Function
Public Function Mat6_solve(m As Matrix6, b As Vector6) As Vector6
    Mat6_solve = Mat6_vmul(Mat6_inv(m), b)
End Function
Public Function Mat7_solve(m As Matrix7, b As Vector7) As Vector7
    Mat7_solve = Mat7_vmul(Mat7_inv(m), b)
End Function
Public Function Mat8_solve(m As Matrix8, b As Vector8) As Vector8
    Mat8_solve = Mat8_vmul(Mat8_inv(m), b)
End Function
Public Function Mat9_solve(m As Matrix9, b As Vector9) As Vector9
    Mat9_solve = Mat9_vmul(Mat9_inv(m), b)
End Function
Public Function Mat10_solve(m As Matrix10, b As Vector10) As Vector10
    Mat10_solve = Mat10_vmul(Mat10_inv(m), b)
End Function

'Public Function Mat2_Rnd() As Matrix2
'    Dim d() As Double: d = Matrix_Random(2)
'    RtlMoveMemory Mat2_Rnd, d(0), 2 ^ 2 * 8
'End Function
'Public Function Mat3_Rnd() As Matrix3
'    Dim d() As Double: d = Matrix_Random(3)
'    RtlMoveMemory Mat3_Rnd, d(0), 3 ^ 2 * 8
'End Function
'Public Function Mat4_Rnd() As Matrix4
'    Dim d() As Double: d = Matrix_Random(4)
'    RtlMoveMemory Mat4_Rnd, d(0), 4 ^ 2 * 8
'End Function
'Public Function Mat5_Rnd() As Matrix5
'    Dim d() As Double: d = Matrix_Random(5)
'    RtlMoveMemory Mat5_Rnd, d(0), 5 ^ 2 * 8
'End Function
'Public Function Mat6_Rnd() As Matrix6
'    Dim d() As Double: d = Matrix_Random(6)
'    RtlMoveMemory Mat6_Rnd, d(0), 6 ^ 2 * 8
'End Function
'Public Function Mat7_Rnd() As Matrix7
'    Dim d() As Double: d = Matrix_Random(7)
'    RtlMoveMemory Mat7_Rnd, d(0), 7 ^ 2 * 8
'End Function
'Public Function Mat8_Rnd() As Matrix8
'    Dim d() As Double: d = Matrix_Random(8)
'    RtlMoveMemory Mat8_Rnd, d(0), 8 ^ 2 * 8
'End Function
'Public Function Mat9_Rnd() As Matrix9
'    Dim d() As Double: d = Matrix_Random(9)
'    RtlMoveMemory Mat9_Rnd, d(0), 9 ^ 2 * 8
'End Function
'Public Function Mat10_Rnd() As Matrix10
'    Dim d() As Double: d = Matrix_Random(10)
'    RtlMoveMemory Mat10_Rnd, d(0), 10 ^ 2 * 8
'End Function
'
'Public Function Matrix_Random(ByVal rc As Byte) As Double()
'    Dim u As Long: u = rc * rc - 1
'    ReDim d(0 To u) As Double
'    Randomize
'    Dim i As Long
'    For i = 0 To u
'        d(i) = Rnd() * 200 - 100
'    Next
'    Matrix_Random = d
'End Function
'
Public Function Mat2_Gauss(m As Matrix2) As Matrix2
    'OM: TODO
    
    With Mat2_Gauss
        m.aa = .aa
    End With
End Function
