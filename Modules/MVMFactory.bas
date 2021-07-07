Attribute VB_Name = "MVMFactory"
Option Explicit

Function CMatRnd(ByVal r As Integer, ByVal c As Integer, dfr As Double, dto As Double) As CMatOp
    'r und c nicht größer als 255, aber -1 möglich
    Dim mv As CMatOp
    Dim maxrc As Long: maxrc = Max(r, c)
    If r = 1 Or c = 1 Then 'Vektor
        Select Case maxrc
        Case 1:  Set mv = MNew.CMatOpByVal(VarPtr(GetNextRandom))
        Case 2:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec2_Rnd(dfr, dto)), r, c)
        Case 3:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec3_Rnd(dfr, dto)), r, c)
        Case 4:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec4_Rnd(dfr, dto)), r, c)
        Case 5:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec5_Rnd(dfr, dto)), r, c)
        Case 6:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec6_Rnd(dfr, dto)), r, c)
        Case 7:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec7_Rnd(dfr, dto)), r, c)
        Case 8:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec8_Rnd(dfr, dto)), r, c)
        Case 9:  Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec9_Rnd(dfr, dto)), r, c)
        Case 10: Set mv = MNew.CMatOpByVal(VarPtr(MVector.Vec10_Rnd(dfr, dto)), r, c)
        End Select
    Else
        Select Case maxrc
        Case 2:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat2_Rnd(dfr, dto)), r, c)
        Case 3:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat3_Rnd(dfr, dto)), r, c)
        Case 4:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat4_Rnd(dfr, dto)), r, c)
        Case 5:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat5_Rnd(dfr, dto)), r, c)
        Case 6:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat6_Rnd(dfr, dto)), r, c)
        Case 7:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat7_Rnd(dfr, dto)), r, c)
        Case 8:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat8_Rnd(dfr, dto)), r, c)
        Case 9:   Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat9_Rnd(dfr, dto)), r, c)
        Case 10:  Set mv = MNew.CMatOpByVal(VarPtr(MMatrices.Mat10_Rnd(dfr, dto)), r, c)
        End Select
    End If
    Set CMatRnd = mv
End Function
