Attribute VB_Name = "MVector"
Option Explicit
Public Type Vector2
    a As Double
    b As Double
End Type
Public Type Vector3
    a As Double
    b As Double
    c As Double
End Type
Public Type Vector4
    a As Double
    b As Double
    c As Double
    d As Double
End Type
Public Type Vector5
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
End Type
Public Type Vector6
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
End Type
Public Type Vector7
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
    g As Double
End Type
Public Type Vector8
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
    g As Double
    H As Double
End Type
Public Type Vector9
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
    g As Double
    H As Double
    i As Double
End Type
Public Type Vector10
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
    g As Double
    H As Double
    i As Double
    j As Double
End Type


'Vektoren erzeugen
Public Function Vec2(a As Double, b As Double) As Vector2
    'Erzeugt einen Vector mit 2 Elementen
    With Vec2: .a = a: .b = b: End With
End Function
Public Function Vec3(a As Double, b As Double, c As Double) As Vector3
    'Erzeugt einen Vector mit 3 Elementen
    With Vec3: .a = a: .b = b: .c = c: End With
End Function
Public Function Vec4(a As Double, b As Double, _
                     c As Double, d As Double) As Vector4
    'Erzeugt einen Vector mit 4 Elementen
    With Vec4: .a = a: .b = b: .c = c: .d = d: End With
End Function
Public Function Vec5(a As Double, b As Double, c As Double, _
                     d As Double, e As Double) As Vector5
    'Erzeugt einen Vector mit 5 Elementen
    With Vec5: .a = a: .b = b: .c = c: .d = d: .e = e: End With
End Function
Public Function Vec6(a As Double, b As Double, c As Double, _
                     d As Double, e As Double, f As Double) As Vector6
    'Erzeugt einen Vector mit 6 Elementen
    With Vec6: .a = a: .b = b: .c = c: .d = d: .e = e: .f = f: End With
End Function
Public Function Vec7(a As Double, b As Double, c As Double, d As Double, _
                     e As Double, f As Double, g As Double) As Vector7
    'Erzeugt einen Vector mit 7 Elementen
    With Vec7: .a = a: .b = b: .c = c: .d = d: .e = e: .f = f: .g = g: End With
End Function
Public Function Vec8(a As Double, b As Double, c As Double, d As Double, _
                     e As Double, f As Double, g As Double, H As Double) As Vector8
    'Erzeugt einen Vector mit 8 Elementen
    With Vec8: .a = a: .b = b: .c = c: .d = d: .e = e: .f = f: .g = g: .H = H: End With
End Function
Public Function Vec9(a As Double, b As Double, c As Double, d As Double, e As Double, _
                     f As Double, g As Double, H As Double, i As Double) As Vector9
    'Erzeugt einen Vector mit 8 Elementen
    With Vec9: .a = a: .b = b: .c = c: .d = d: .e = e: .f = f: .g = g: .H = H: .i = i: End With
End Function
Public Function Vec10(a As Double, b As Double, c As Double, d As Double, e As Double, _
                      f As Double, g As Double, H As Double, i As Double, j As Double) As Vector10
    'Erzeugt einen Vector mit 8 Elementen
    With Vec10: .a = a: .b = b: .c = c: .d = d: .e = e: .f = f: .g = g: .H = H: .i = i: .j = j: End With
End Function

'Random Values
Public Function Vec2_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector2
    Rnd_SetFromTo dbl_From, dbl_To
    Vec2_Rnd = Vec2(rv, rv)
End Function
Public Function Vec3_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector3
    Rnd_SetFromTo dbl_From, dbl_To
    Vec3_Rnd = Vec3(rv, rv, rv)
End Function
Public Function Vec4_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector4
    Rnd_SetFromTo dbl_From, dbl_To
    Vec4_Rnd = Vec4(rv, rv, rv, rv)
End Function
Public Function Vec5_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector5
    Rnd_SetFromTo dbl_From, dbl_To
    Vec5_Rnd = Vec5(rv, rv, rv, rv, rv)
End Function
Public Function Vec6_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector6
    Rnd_SetFromTo dbl_From, dbl_To
    Vec6_Rnd = Vec6(rv, rv, rv, rv, rv, rv)
End Function
Public Function Vec7_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector7
    Rnd_SetFromTo dbl_From, dbl_To
    Vec7_Rnd = Vec7(rv, rv, rv, rv, rv, rv, rv)
End Function
Public Function Vec8_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector8
    Rnd_SetFromTo dbl_From, dbl_To
    Vec8_Rnd = Vec8(rv, rv, rv, rv, rv, rv, rv, rv)
End Function
Public Function Vec9_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector9
    Rnd_SetFromTo dbl_From, dbl_To
    Vec9_Rnd = Vec9(rv, rv, rv, rv, rv, rv, rv, rv, rv)
End Function
Public Function Vec10_Rnd(Optional ByVal dbl_From As Double = -200, Optional ByVal dbl_To As Double = 200) As Vector10
    Rnd_SetFromTo dbl_From, dbl_To
    Vec10_Rnd = Vec10(rv, rv, rv, rv, rv, rv, rv, rv, rv, rv)
End Function

'Vektoren-Addition
Public Function Vec2_add(V1 As Vector2, V2 As Vector2) As Vector2
    'Addiert 2 2er-Vectoren, liefert einen neuen Vektor
    With Vec2_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    End With
End Function
Public Function Vec3_add(V1 As Vector3, V2 As Vector3) As Vector3
    'Addiert 2 3er-Vectoren, liefert einen neuen Vektor
    With Vec3_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    End With
End Function
Public Function Vec4_add(V1 As Vector4, V2 As Vector4) As Vector4
    'Addiert 2 4er-Vectoren, liefert einen neuen Vektor
    With Vec4_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    End With
End Function
Public Function Vec5_add(V1 As Vector5, V2 As Vector5) As Vector5
    'Addiert 2 5er-Vectoren, liefert einen neuen Vektor
    With Vec5_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    .e = V1.e + V2.e:    End With
End Function
Public Function Vec6_add(V1 As Vector6, V2 As Vector6) As Vector6
    'Addiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec6_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    .e = V1.e + V2.e:    .f = V1.f + V2.f:    End With
End Function
Public Function Vec7_add(V1 As Vector7, V2 As Vector7) As Vector7
    'Addiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec7_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    .e = V1.e + V2.e:    .f = V1.f + V2.f:    .g = V1.g + V2.g:    End With
End Function
Public Function Vec8_add(V1 As Vector8, V2 As Vector8) As Vector8
    'Addiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec8_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    .e = V1.e + V2.e:    .f = V1.f + V2.f:    .g = V1.g + V2.g:    .H = V1.H + V2.H:    End With
End Function
Public Function Vec9_add(V1 As Vector9, V2 As Vector9) As Vector9
    'Addiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec9_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    .e = V1.e + V2.e:    .f = V1.f + V2.f:    .g = V1.g + V2.g:    .H = V1.H + V2.H:    .i = V1.i + V2.i:    End With
End Function
Public Function Vec10_add(V1 As Vector10, V2 As Vector10) As Vector10
    'Addiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec10_add:    .a = V1.a + V2.a:    .b = V1.b + V2.b:    .c = V1.c + V2.c:    .d = V1.d + V2.d:    .e = V1.e + V2.e:    .f = V1.f + V2.f:    .g = V1.g + V2.g:    .H = V1.H + V2.H:    .i = V1.i + V2.i:    .j = V1.j + V2.j:    End With
End Function

'Vektoren-Sutraktion
Public Function Vec2_sub(V1 As Vector2, V2 As Vector2) As Vector2
    'Subtrahiert 2 2er-Vectoren, liefert einen neuen Vektor
    With Vec2_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    End With
End Function
Public Function Vec3_sub(V1 As Vector3, V2 As Vector3) As Vector3
    'Subtrahiert 2 3er-Vectoren, liefert einen neuen Vektor
    With Vec3_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    End With
End Function
Public Function Vec4_sub(V1 As Vector4, V2 As Vector4) As Vector4
    'Subtrahiert 2 4er-Vectoren, liefert einen neuen Vektor
    With Vec4_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    End With
End Function
Public Function Vec5_sub(V1 As Vector5, V2 As Vector5) As Vector5
    'Subtrahiert 2 5er-Vectoren, liefert einen neuen Vektor
    With Vec5_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    .e = V1.e - V2.e:    End With
End Function
Public Function Vec6_sub(V1 As Vector6, V2 As Vector6) As Vector6
    'Subtrahiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec6_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    .e = V1.e - V2.e:    .f = V1.f - V2.f:    End With
End Function
Public Function Vec7_sub(V1 As Vector7, V2 As Vector7) As Vector7
    'Subtrahiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec7_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    .e = V1.e - V2.e:    .f = V1.f - V2.f:    .g = V1.g - V2.g:    End With
End Function
Public Function Vec8_sub(V1 As Vector8, V2 As Vector8) As Vector8
    'Subtrahiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec8_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    .e = V1.e - V2.e:    .f = V1.f - V2.f:    .g = V1.g - V2.g:    .H = V1.H - V2.H:    End With
End Function
Public Function Vec9_sub(V1 As Vector9, V2 As Vector9) As Vector9
    'Subtrahiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec9_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    .e = V1.e - V2.e:    .f = V1.f - V2.f:    .g = V1.g - V2.g:    .H = V1.H - V2.H:    .i = V1.i - V2.i:    End With
End Function
Public Function Vec10_sub(V1 As Vector10, V2 As Vector10) As Vector10
    'Subtrahiert 2 6er-Vectoren, liefert einen neuen Vektor
    With Vec10_sub:   .a = V1.a - V2.a:    .b = V1.b - V2.b:    .c = V1.c - V2.c:    .d = V1.d - V2.d:    .e = V1.e - V2.e:    .f = V1.f - V2.f:    .g = V1.g - V2.g:    .H = V1.H - V2.H:    .i = V1.i - V2.i:    .j = V1.j - V2.j:    End With
End Function

'Vektoren mit Skalar multiplizieren
Public Function Vec2_smul(v As Vector2, ByVal s As Double) As Vector2
    'Multipliziert Alle Elemente eines 2er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec2_smul: .a = v.a * s: .b = v.b * s: End With
End Function
Public Function Vec3_smul(v As Vector3, ByVal s As Double) As Vector3
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec3_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: End With
End Function
Public Function Vec4_smul(v As Vector4, ByVal s As Double) As Vector4
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec4_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: End With
End Function
Public Function Vec5_smul(v As Vector5, ByVal s As Double) As Vector5
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec5_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: .e = v.e * s: End With
End Function
Public Function Vec6_smul(v As Vector6, ByVal s As Double) As Vector6
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec6_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: .e = v.e * s: .f = v.f * s: End With
End Function
Public Function Vec7_smul(v As Vector7, ByVal s As Double) As Vector7
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec7_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: .e = v.e * s: .f = v.f * s: .g = v.g * s: End With
End Function
Public Function Vec8_smul(v As Vector8, ByVal s As Double) As Vector8
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec8_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: .e = v.e * s: .f = v.f * s: .g = v.g * s: .H = v.H * s: End With
End Function
Public Function Vec9_smul(v As Vector9, ByVal s As Double) As Vector9
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec9_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: .e = v.e * s: .f = v.f * s: .g = v.g * s: .H = v.H * s: .i = v.i * s: End With
End Function
Public Function Vec10_smul(v As Vector10, ByVal s As Double) As Vector10
    'Multipliziert Alle Elemente eines 3er-Vectors mit einem Skalar, liefert einen neuen Vektor
    With Vec10_smul: .a = v.a * s: .b = v.b * s: .c = v.c * s: .d = v.d * s: .e = v.e * s: .f = v.f * s: .g = v.g * s: .H = v.H * s: .i = v.i * s: .j = v.j * s: End With
End Function

Public Function Vec2_len(v As Vector2) As Double
    With v: Vec2_len = VBA.Math.Sqr(.a * .a + .b * .b): End With
End Function
Public Function Vec3_len(v As Vector3) As Double
    With v: Vec3_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c): End With
End Function
Public Function Vec4_len(v As Vector4) As Double
    With v: Vec4_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d): End With
End Function
Public Function Vec5_len(v As Vector5) As Double
    With v: Vec5_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d + .e * .e): End With
End Function
Public Function Vec6_len(v As Vector6) As Double
    With v: Vec6_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d + .e * .e + .f * .f): End With
End Function
Public Function Vec7_len(v As Vector7) As Double
    With v: Vec7_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d + .e * .e + .f * .f + .g * .g): End With
End Function
Public Function Vec8_len(v As Vector8) As Double
    With v: Vec8_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d + .e * .e + .f * .f + .g * .g + .H * .H): End With
End Function
Public Function Vec9_len(v As Vector9) As Double
    With v: Vec9_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d + .e * .e + .f * .f + .g * .g + .H * .H + .i * .i): End With
End Function
Public Function Vec10_len(v As Vector10) As Double
    With v: Vec10_len = VBA.Math.Sqr(.a * .a + .b * .b + .c * .c + .d * .d + .e * .e + .f * .f + .g * .g + .H * .H + .i * .i + .j * .j): End With
End Function

' Das Kreuzprodukt ist nicht kommutativ;
' werden a und b vertauscht so ‰ndert sich das Vorzeichen.
' Man sagt auch, es sei antikommutativ
' F¸r das Kreuzprodukt gilt das Distributivgesetz
' Das Kreuzprodukt ist nicht assoziativ

Public Function Vec2_cross(V1 As Vector2, V2 As Vector2) As Double
    Vec2_cross = V1.a * V2.b - V2.a * V1.b
End Function
Public Function Vec3_cross(V1 As Vector3, V2 As Vector3) As Vector3
    With Vec3_cross
        .a = V1.b * V2.c - V1.c * V2.b
        .b = V1.c * V2.a - V1.a * V2.c
        .c = V1.a * V2.b - V1.b * V2.a
    End With
End Function
Public Function Vec4_cross(V1 As Vector4, V2 As Vector4, V3 As Vector4) As Vector4
    Dim m As Matrix4
    Dim i As Long
    Mat4_Row(m, i) = V1: i = i + 1
    Mat4_Row(m, i) = V2: i = i + 1
    Mat4_Row(m, i) = V3: i = i + 1
    Mat4_Row(m, i) = Vec4(1#, 1#, 1#, 1#)
    Vec4_cross = Mat4_Col(Mat4_Adj(m), i)
End Function
Public Function Vec5_cross(V1 As Vector5, V2 As Vector5, V3 As Vector5, V4 As Vector5) As Vector5
    Dim m As Matrix5
    Dim i As Long
    Mat5_Row(m, i) = V1: i = i + 1
    Mat5_Row(m, i) = V2: i = i + 1
    Mat5_Row(m, i) = V3: i = i + 1
    Mat5_Row(m, i) = V4: i = i + 1
    Mat5_Row(m, i) = Vec5(1#, 1#, 1#, 1#, 1#)
    Vec5_cross = Mat5_Col(Mat5_Adj(m), i)
End Function
Public Function Vec6_cross(V1 As Vector6, V2 As Vector6, V3 As Vector6, V4 As Vector6, V5 As Vector6) As Vector6
    Dim m As Matrix6
    Dim i As Long
    Mat6_Row(m, i) = V1: i = i + 1
    Mat6_Row(m, i) = V2: i = i + 1
    Mat6_Row(m, i) = V3: i = i + 1
    Mat6_Row(m, i) = V4: i = i + 1
    Mat6_Row(m, i) = V5: i = i + 1
    Mat6_Row(m, i) = Vec6(1#, 1#, 1#, 1#, 1#, 1#)
    Vec6_cross = Mat6_Col(Mat6_Adj(m), i)
End Function
Public Function Vec7_cross(V1 As Vector7, V2 As Vector7, V3 As Vector7, V4 As Vector7, V5 As Vector7, V6 As Vector7) As Vector7
    Dim m As Matrix7
    Dim i As Long
    Mat7_Row(m, i) = V1: i = i + 1
    Mat7_Row(m, i) = V2: i = i + 1
    Mat7_Row(m, i) = V3: i = i + 1
    Mat7_Row(m, i) = V4: i = i + 1
    Mat7_Row(m, i) = V5: i = i + 1
    Mat7_Row(m, i) = V6: i = i + 1
    Mat7_Row(m, i) = Vec7(1#, 1#, 1#, 1#, 1#, 1#, 1#)
    Vec7_cross = Mat7_Col(Mat7_Adj(m), i)
End Function
Public Function Vec8_cross(V1 As Vector8, V2 As Vector8, V3 As Vector8, V4 As Vector8, V5 As Vector8, V6 As Vector8, V7 As Vector8) As Vector8
    Dim m As Matrix8
    Dim i As Long
    Mat8_Row(m, i) = V1: i = i + 1
    Mat8_Row(m, i) = V2: i = i + 1
    Mat8_Row(m, i) = V3: i = i + 1
    Mat8_Row(m, i) = V4: i = i + 1
    Mat8_Row(m, i) = V5: i = i + 1
    Mat8_Row(m, i) = V6: i = i + 1
    Mat8_Row(m, i) = V7: i = i + 1
    Mat8_Row(m, i) = Vec8(1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#)
    Vec8_cross = Mat8_Col(Mat8_Adj(m), i)
End Function
Public Function Vec9_cross(V1 As Vector9, V2 As Vector9, V3 As Vector9, V4 As Vector9, V5 As Vector9, V6 As Vector9, V7 As Vector9, V8 As Vector9) As Vector9
    Dim m As Matrix9
    Dim i As Long
    Mat9_Row(m, i) = V1: i = i + 1
    Mat9_Row(m, i) = V2: i = i + 1
    Mat9_Row(m, i) = V3: i = i + 1
    Mat9_Row(m, i) = V4: i = i + 1
    Mat9_Row(m, i) = V5: i = i + 1
    Mat9_Row(m, i) = V6: i = i + 1
    Mat9_Row(m, i) = V7: i = i + 1
    Mat9_Row(m, i) = V8: i = i + 1
    Mat9_Row(m, i) = Vec9(1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#)
    Vec9_cross = Mat9_Col(Mat9_Adj(m), i)
End Function
Public Function Vec10_cross(V1 As Vector10, V2 As Vector10, V3 As Vector10, V4 As Vector10, V5 As Vector10, V6 As Vector10, V7 As Vector10, V8 As Vector10, V9 As Vector10) As Vector10
    Dim m As Matrix10
    Dim i As Long
    Mat10_Row(m, i) = V1: i = i + 1
    Mat10_Row(m, i) = V2: i = i + 1
    Mat10_Row(m, i) = V3: i = i + 1
    Mat10_Row(m, i) = V4: i = i + 1
    Mat10_Row(m, i) = V5: i = i + 1
    Mat10_Row(m, i) = V6: i = i + 1
    Mat10_Row(m, i) = V7: i = i + 1
    Mat10_Row(m, i) = V8: i = i + 1
    Mat10_Row(m, i) = V9: i = i + 1
    Mat10_Row(m, i) = Vec10(1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#, 1#)
    Vec10_cross = Mat10_Col(Mat10_Adj(m), i)
End Function


Public Function Vec2_norm(v As Vector2) As Vector2
    Dim vlen As Double: vlen = Vec2_len(v)
    With Vec2_norm: .a = .a / vlen: .b = .b / vlen: End With
End Function
Public Function Vec3_norm(v As Vector3) As Vector3
    Dim vlen  As Double: vlen = Vec3_len(v)
    With Vec3_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: End With
End Function
Public Function Vec4_norm(v As Vector4) As Vector4
    Dim vlen  As Double: vlen = Vec4_len(v)
    With Vec4_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: End With
End Function
Public Function Vec5_norm(v As Vector5) As Vector5
    Dim vlen As Double: vlen = Vec5_len(v)
    With Vec5_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: .e = .e / vlen: End With
End Function
Public Function Vec6_norm(v As Vector6) As Vector6
    Dim vlen As Double: vlen = Vec6_len(v)
    With Vec6_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: .e = .e / vlen: .f = .f / vlen: End With
End Function
Public Function Vec7_norm(v As Vector7) As Vector7
    Dim vlen As Double: vlen = Vec7_len(v)
    With Vec7_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: .e = .e / vlen: .f = .f / vlen: .g = .g / vlen: End With
End Function
Public Function Vec8_norm(v As Vector8) As Vector8
    Dim vlen As Double: vlen = Vec8_len(v)
    With Vec8_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: .e = .e / vlen: .f = .f / vlen: .g = .g / vlen: .H = .H / vlen: End With
End Function
Public Function Vec9_norm(v As Vector9) As Vector9
    Dim vlen As Double: vlen = Vec9_len(v)
    With Vec9_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: .e = .e / vlen: .f = .f / vlen: .g = .g / vlen: .H = .H / vlen: .i = .i / vlen: End With
End Function
Public Function Vec10_norm(v As Vector10) As Vector10
    Dim vlen As Double: vlen = Vec10_len(v)
    With Vec10_norm: .a = .a / vlen: .b = .b / vlen: .c = .c / vlen: .d = .d / vlen: .e = .e / vlen: .f = .f / vlen: .g = .g / vlen: .H = .H / vlen: .i = .i / vlen: .j = .j / vlen: End With
End Function







'Aus einem Vektor Untervektoren Rauskopieren
Public Function Vec3_uvec(v As Vector3, ByVal ex As Long) As Vector2
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec3_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c
        Case 1: .a = v.a: .b = v.c
        Case 2: .a = v.a: .b = v.b
        End Select
    End With
End Function
Public Function Vec4_uvec(v As Vector4, ByVal ex As Long) As Vector3
    'kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec4_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d
        Case 1: .a = v.a: .b = v.c: .c = v.d
        Case 2: .a = v.a: .b = v.b: .c = v.d
        Case 3: .a = v.a: .b = v.b: .c = v.c
        End Select
    End With
End Function
Public Function Vec5_uvec(v As Vector5, ByVal ex As Long) As Vector4
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec5_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d
        End Select
    End With
End Function
Public Function Vec6_uvec(v As Vector6, ByVal ex As Long) As Vector5
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec6_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e: .e = v.f
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e: .e = v.f
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e: .e = v.f
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e: .e = v.f
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.f
        Case 5: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e
        End Select
    End With
End Function
Public Function Vec7_uvec(v As Vector7, ByVal ex As Long) As Vector6
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec7_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e: .e = v.f: .f = v.g
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e: .e = v.f: .f = v.g
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.f: .f = v.g
        Case 5: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.g
        Case 6: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f
        End Select
    End With
End Function
Public Function Vec8_uvec(v As Vector8, ByVal ex As Long) As Vector7
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec8_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e: .e = v.f: .f = v.g: .g = v.H
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.f: .f = v.g: .g = v.H
        Case 5: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.g: .g = v.H
        Case 6: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.H
        Case 7: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.g
        End Select
    End With
End Function
Public Function Vec9_uvec(v As Vector9, ByVal ex As Long) As Vector8
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec9_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.f: .f = v.g: .g = v.H: .H = v.i
        Case 5: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.g: .g = v.H: .H = v.i
        Case 6: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.H: .H = v.i
        Case 7: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.g: .H = v.i
        Case 8: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.g: .H = v.H
        End Select
    End With
End Function
Public Function Vec10_uvec(v As Vector10, ByVal ex As Long) As Vector9
    'Kopiert alle Elemente auﬂer ex in einen kleineren Vektor
    With Vec10_uvec
        Select Case ex
        Case 0: .a = v.b: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i: .i = v.j
        Case 1: .a = v.a: .b = v.c: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i: .i = v.j
        Case 2: .a = v.a: .b = v.b: .c = v.d: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i: .i = v.j
        Case 3: .a = v.a: .b = v.b: .c = v.c: .d = v.e: .e = v.f: .f = v.g: .g = v.H: .H = v.i: .i = v.j
        Case 4: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.f: .f = v.g: .g = v.H: .H = v.i: .i = v.j
        Case 5: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.g: .g = v.H: .H = v.i: .i = v.j
        Case 6: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.H: .H = v.i: .i = v.j
        Case 7: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.g: .H = v.i: .i = v.j
        Case 8: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.g: .H = v.H: .i = v.j
        Case 9: .a = v.a: .b = v.b: .c = v.c: .d = v.d: .e = v.e: .f = v.f: .g = v.g: .H = v.H: .i = v.i
        End Select
    End With
End Function

'einen Vektor in eine array kopieren, m¸ﬂte mit CopyMem leichter zu programmieren sein
'aber weil wir ja nicht faul sind.
Public Function Vec2_ToArr(v As Vector2) As Double()
    Dim da(0 To 1) As Double
    With v: da(0) = .a: da(1) = .b: End With
    Vec2_ToArr = da
End Function
Public Function Vec3_ToArr(v As Vector3) As Double()
    Dim da(0 To 2) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: End With
    Vec3_ToArr = da
End Function
Public Function Vec4_ToArr(v As Vector4) As Double()
    Dim da(0 To 3) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: End With
    Vec4_ToArr = da
End Function
Public Function Vec5_ToArr(v As Vector5) As Double()
    Dim da(0 To 4) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: da(4) = .e: End With
    Vec5_ToArr = da
End Function
Public Function Vec6_ToArr(v As Vector6) As Double()
    Dim da(0 To 5) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: da(4) = .e: da(5) = .f: End With
    Vec6_ToArr = da
End Function
Public Function Vec7_ToArr(v As Vector7) As Double()
    Dim da(0 To 6) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: da(4) = .e: da(5) = .f: da(6) = .g: End With
    Vec7_ToArr = da
End Function
Public Function Vec8_ToArr(v As Vector8) As Double()
    Dim da(0 To 7) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: da(4) = .e: da(5) = .f: da(6) = .g: da(7) = .H: End With
    Vec8_ToArr = da
End Function
Public Function Vec9_ToArr(v As Vector9) As Double()
    Dim da(0 To 8) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: da(4) = .e: da(5) = .f: da(6) = .g: da(7) = .H: da(8) = .i: End With
    Vec9_ToArr = da
End Function
Public Function Vec10_ToArr(v As Vector10) As Double()
    Dim da(0 To 9) As Double
    With v: da(0) = .a: da(1) = .b: da(2) = .c: da(3) = .d: da(4) = .e: da(5) = .f: da(6) = .g: da(7) = .H: da(8) = .i: da(9) = .j: End With
    Vec10_ToArr = da
End Function

Public Function Vec2_ToStr(v As Vector2, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
'    Dim sa(0 To 1) As String
'    With v: sa(0) = CStr(.a): sa(1) = CStr(.b): End With
'    Vec2_ToStr = VectorStrPadLR(sa, bIsLineVec)
    Vec2_ToStr = VectorFormat(Vec2_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec3_ToStr(v As Vector3, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
'    Dim sa(0 To 2) As String
'    With v: sa(0) = CStr(.a): sa(1) = CStr(.b): sa(2) = CStr(.c): End With
'    Vec3_ToStr = VectorStrPadLR(sa, bIsLineVec)
    Vec3_ToStr = VectorFormat(Vec3_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec4_ToStr(v As Vector4, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
'    Dim sa(0 To 3) As String
'    With v: sa(0) = CStr(.a): sa(1) = CStr(.b): sa(2) = CStr(.c): sa(3) = CStr(.d): End With
'    Vec4_ToStr = VectorStrPadLR(sa, bIsLineVec)
    Vec4_ToStr = VectorFormat(Vec4_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec5_ToStr(v As Vector5, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
'    Dim sa(0 To 4) As String
'    With v: sa(0) = CStr(.a): sa(1) = CStr(.b): sa(2) = CStr(.c): sa(3) = CStr(.d): sa(4) = CStr(.e): End With
'    Vec5_ToStr = VectorStrPadLR(sa, bIsLineVec)
    Vec5_ToStr = VectorFormat(Vec5_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec6_ToStr(v As Vector6, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
'    Dim sa(0 To 5) As String
'    With v: sa(0) = CStr(.a): sa(1) = CStr(.b): sa(2) = CStr(.c): sa(3) = CStr(.d): sa(4) = CStr(.e): sa(5) = CStr(.f): End With
'    Vec6_ToStr = VectorStrPadLR(sa, 0, bIsLineVec)
    Vec6_ToStr = VectorFormat(Vec6_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec7_ToStr(v As Vector7, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
    Vec7_ToStr = VectorFormat(Vec7_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec8_ToStr(v As Vector8, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
    Vec8_ToStr = VectorFormat(Vec8_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec9_ToStr(v As Vector9, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = 3) As String
    Vec9_ToStr = VectorFormat(Vec9_ToArr(v), 0, bIsLineVec, dFormat)
End Function
Public Function Vec10_ToStr(v As Vector10, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = -1) As String
    Vec10_ToStr = VectorFormat(Vec10_ToArr(v), 0, bIsLineVec, dFormat)
End Function
'Public Function VectorStrPadLR(sa() As String, totalwidth As Long, Optional bIsLineVec As Boolean = False) As String
'    Dim i As Long, maxlenL As Long, maxlenR As Long
'    For i = LBound(sa) To UBound(sa): maxlenL = Max(maxlen, Len(sa(i))): Next
'    For i = LBound(sa) To UBound(sa):  sa(i) = PadLeft(sa(i), maxlen): Next
'    VectorStrPadLR = Join(sa, IIf(bIsLineVec, " ", vbCrLf))
'End Function

Public Function VectorFormat(da() As Double, totalWidth As Long, Optional bIsLineVec As Boolean = False, Optional ByVal dFormat As Integer = -1) As String
    Dim i As Long, maxlenL As Long, maxlenR As Long
    Dim s As String, sa() As String
    Dim sdi As String, sdf As String
    Dim l As Long: l = LBound(da)
    Dim u As Long: u = UBound(da)
    For i = l To u
        'If dFormat < 0 Then
            s = Trim(Str(da(i)))
        'Else
        '    s = Replace(Format(da(i), "0." & String$(dFormat, "0")), ",", ".")
        'End If
        If InStr(1, s, ".") Then
            sa = Split(s, ".")
            sdi = sa(0): If Len(sdi) = 0 Then sdi = "0"
            sdf = sa(1)
        Else
            sdi = s: sdf = ""
        End If
        maxlenL = Max(maxlenL, Len(sdi))
        If dFormat >= 0 Then
            maxlenR = Max(maxlenR, dFormat)
        Else
            maxlenR = Max(maxlenR, Len(sdf))
        End If
    Next
    ReDim sar(l To u) As String
    For i = l To u
        If dFormat < 0 Then
            s = Trim(Str(da(i)))
        ElseIf dFormat = 255 Then
            'bei 255 soll die wissenschaftliche E-Notation erfolgen
            s = Format(Str(da(i)), "scientific")
            Debug.Print s
        Else
            s = Replace(Format(da(i), "0." & String$(dFormat, "0")), ",", ".")
        End If
        If InStr(1, s, ".") Then
            sa = Split(s, ".")
            sdi = sa(0): If Len(sdi) = 0 Then sdi = "0"
            sdf = sa(1)
            sar(i) = PadLeft(sdi, maxlenL) & "." & PadRight(sdf, maxlenR)
        Else
            sdi = s: sdf = ""
            sar(i) = PadLeft(sdi, maxlenL) & PadRight(sdf, maxlenR + 1)
        End If
    Next
    VectorFormat = Join(sar, IIf(bIsLineVec, " ", vbCrLf))
End Function

Public Function Vector_Parse(T As String, ByVal n As Long) As Double()
    ReDim a_out(0 To n - 1) As Double
    Dim s As String: s = DeleteMultiWS(DeleteCRLF(T))
    Dim sa() As String: sa = Split(s, " ")
    Dim i As Long
    For i = 0 To n - 1
        If UBound(sa) < i Then Exit For
        a_out(i) = DblParse(sa(i)) 'orig ?
    Next
    Vector_Parse = a_out
End Function
Public Function Vec2_Parse(T As String) As Vector2
    Dim n As Long: n = 2
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec2_Parse, a(0), n * 8
End Function
Public Function Vec3_Parse(T As String) As Vector3
    Dim n As Long: n = 3
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec3_Parse, a(0), n * 8
End Function
Public Function Vec4_Parse(T As String) As Vector4
    Dim n As Long: n = 4
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec4_Parse, a(0), n * 8
End Function
Public Function Vec5_Parse(T As String) As Vector5
    Dim n As Long: n = 5
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec5_Parse, a(0), n * 8
End Function
Public Function Vec6_Parse(T As String) As Vector6
    Dim n As Long: n = 6
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec6_Parse, a(0), n * 8
End Function
Public Function Vec7_Parse(T As String) As Vector7
    Dim n As Long: n = 7
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec7_Parse, a(0), n * 8
End Function
Public Function Vec8_Parse(T As String) As Vector8
    Dim n As Long: n = 8
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec8_Parse, a(0), n * 8
End Function
Public Function Vec9_Parse(T As String) As Vector9
    Dim n As Long: n = 9
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec9_Parse, a(0), n * 8
End Function
Public Function Vec10_Parse(T As String) As Vector10
    Dim n As Long: n = 10
    Dim a() As Double: a = Vector_Parse(T, n)
    RtlMoveMemory Vec10_Parse, a(0), n * 8
End Function



