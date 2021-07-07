Attribute VB_Name = "MFont"
Option Explicit
Private Const LF_FontNameSIZE As Long = 128&
Private Const DEFAULT_CHARSET As Long = 1&
Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFontName(LF_FontNameSIZE) As Byte
End Type
Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExW" ( _
    ByVal hDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, _
    ByVal lParam As Collection, Optional ByVal dw As Long = 0) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByRef lpString As Byte) As Long

Private LastFont As String

Public Function GetAllFonts(ByVal hDC As Long) As Collection
    Set GetAllFonts = New Collection
    Dim lf As LOGFONT: lf.lfCharSet = DEFAULT_CHARSET
    EnumFontFamiliesEx hDC, lf, AddressOf EnumFontFamExProc, GetAllFonts
End Function

Public Function EnumFontFamExProc(ByRef lpElfe As LOGFONT, ByVal lpntme As Long, ByVal FontType As Long, ByVal lPCol As Collection) As Long
    Dim fnam As String: fnam = Left(lpElfe.lfFontName, lstrlenW(lpElfe.lfFontName(0)))
    If Not LastFont = fnam Then
        lPCol.Add fnam
    End If
    LastFont = fnam
    EnumFontFamExProc = 1
End Function
