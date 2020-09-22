VERSION 5.00
Begin VB.UserControl VectorPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   FontTransparent =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
End
Attribute VB_Name = "VectorPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim token As Long

Dim mObj() As VectorObject
Dim mObjCount As Integer
Dim mObjBox As RECT

Dim mDrawMode As eDrawMode
Dim mBorderWidth As Long
Dim mRendered As Boolean
Dim mTransparent As Byte
Dim mShowProgress As Boolean

Private Type POINT_API
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Enum POINT_TYPE
    ptMove = &H6
    ptLine = &H2
    ptBezier = &H4
End Enum

Private Type VectorObject
    Point() As POINT_API
    PointType() As POINT_TYPE
    PointCount As Long
    Color As Long
End Type

Public Enum eDrawMode
    BlackLines = 0
    ColoredBorder = 1
    Filled = 2
End Enum

Public Event OpenProgress(ByVal Val As Long, ByVal Max As Long)

Private Function OpenPLTFile(FileName As String) As Boolean
On Error Resume Next
Dim vI As Long
Dim FF As Long
Dim tmpFile As String
Dim tLine() As String
Dim tLineCount As Long
Dim tStr As String
Dim tC As Integer
Dim tCol As Long
Dim OldC As Long
Dim tH As Long
Dim toCancel As Boolean

FF = FreeFile

tmpFile = Space(FileLen(FileName))
Open FileName For Binary As #FF
Get #FF, , tmpFile
Close #FF

tLine = Split(tmpFile, Chr(13) & Chr(10))

tLineCount = UBound(tLine)

tH = 5000

For vI = 0 To tLineCount
RaiseEvent OpenProgress(vI, tLineCount)
DrawProg vI, tLineCount
tLine(vI) = Replace(tLine(vI), ";", "")
tStr = Left(tLine(vI), 2)
    Select Case LCase(tStr)
        Case "sp"
        tC = Int(Right(tLine(vI), Len(tLine(vI)) - 2))
        tCol = GetPenColor(tC)
        Case "pu"
        mObjCount = mObjCount + 1
        ReDim Preserve mObj(mObjCount)
        mObj(mObjCount - 1).PointCount = 1
        ReDim mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount)
        ReDim mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X = Right(Trim(Split(tLine(vI), " ")(0)), Len(Trim(Split(tLine(vI), " ")(0))) - 2)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y = tH - (Trim(Split(tLine(vI), " ")(1)))
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount) = ptMove
        mObj(mObjCount - 1).Color = tCol
        Case "pd"
        mObj(mObjCount - 1).PointCount = mObj(mObjCount - 1).PointCount + 1
        ReDim Preserve mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount)
        ReDim Preserve mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X = Right(Trim(Split(tLine(vI), " ")(0)), Len(Trim(Split(tLine(vI), " ")(0))) - 2)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y = tH - (Trim(Split(tLine(vI), " ")(1)))
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount) = ptLine
    End Select
Next vI

OpenPLTFile = CBool(mObjCount)

End Function


Public Function OpenVectorFile(FileName As String) As Boolean
Dim Ext As String

Ext = LCase(Right(FileName, 3))
Ext = Replace(Ext, ".", "")

mObjCount = 0
ReDim mObj(mObjCount)
UserControl.BackColor = vbButtonFace
UserControl.Cls
DoEvents

Select Case Ext
    Case "plt"
        OpenVectorFile = OpenPLTFile(FileName)
    Case "eps", "ai"
        OpenVectorFile = OpenPSFile(FileName)
    Case Else
        OpenVectorFile = False
End Select

If mObjCount > 0 Then
GetObjBox
MoveAndSize
DrawVector
End If

End Function

Private Function GetPenColor(PenNo As Integer) As Long
Select Case PenNo - 1
    Case 0
    GetPenColor = vbBlack
    Case 1
    GetPenColor = vbBlue
    Case 2
    GetPenColor = vbRed
    Case 3
    GetPenColor = 32768 ' Green
    Case 4
    GetPenColor = vbMagenta
    Case 5
    GetPenColor = vbYellow
    Case 6
    GetPenColor = vbCyan
    Case 7
    GetPenColor = 34815 ' Orange
    Case 8
    GetPenColor = &HFCFCFC ' vbWhite
    Case 9
    GetPenColor = 16777215 ' Gray
    Case Else
    GetPenColor = vbBlack
End Select
End Function

Private Sub UserControl_Initialize()
Dim GpInput As GdiplusStartupInput
GpInput.GdiplusVersion = 1
If GdiplusStartup(token, GpInput) <> Ok Then
   MsgBox "Error loading GDI+!", vbCritical
End If
End Sub



Private Sub GetObjBox()
Dim vI As Long
Dim vJ As Long
Dim tBox As RECT

tBox.Left = 100000000
tBox.Top = 100000000
tBox.Right = -100000000
tBox.Bottom = -100000000


For vI = 0 To mObjCount - 1
    For vJ = 1 To mObj(vI).PointCount
        If tBox.Left > mObj(vI).Point(vJ).X Then
        tBox.Left = mObj(vI).Point(vJ).X
        End If
        If tBox.Right < mObj(vI).Point(vJ).X Then
        tBox.Right = mObj(vI).Point(vJ).X
        End If
        If tBox.Top > mObj(vI).Point(vJ).Y Then
        tBox.Top = mObj(vI).Point(vJ).Y
        End If
        If tBox.Bottom < mObj(vI).Point(vJ).Y Then
        tBox.Bottom = mObj(vI).Point(vJ).Y
        End If
    Next vJ
Next vI
mObjBox = tBox
End Sub

Private Sub MoveAndSize()
Dim vI As Long
Dim vJ As Long
Dim movX As Long
Dim movY As Long
Dim Ratio As Double

movX = mObjBox.Left - mBorderWidth
movY = mObjBox.Top - mBorderWidth

For vI = 0 To mObjCount - 1
    For vJ = 1 To mObj(vI).PointCount
    mObj(vI).Point(vJ).X = mObj(vI).Point(vJ).X - movX
    mObj(vI).Point(vJ).Y = mObj(vI).Point(vJ).Y - movY
    Next vJ
Next vI

If (mObjBox.Right - mObjBox.Left) >= (mObjBox.Bottom - mObjBox.Top) Then
Ratio = (UserControl.ScaleWidth - (mBorderWidth * 2)) / (mObjBox.Right - mObjBox.Left)
movX = 0
movY = ((UserControl.ScaleHeight - (mBorderWidth * 2)) - ((mObjBox.Bottom - mObjBox.Top) * Ratio)) / 2
Else
Ratio = (UserControl.ScaleHeight - (mBorderWidth * 2)) / (mObjBox.Bottom - mObjBox.Top)
movX = ((UserControl.ScaleWidth - (mBorderWidth * 2)) - ((mObjBox.Right - mObjBox.Left) * Ratio)) / 2
movY = 0
End If

For vI = 0 To mObjCount - 1
    For vJ = 1 To mObj(vI).PointCount
    mObj(vI).Point(vJ).X = ((mObj(vI).Point(vJ).X - mBorderWidth) * Ratio) + mBorderWidth + movX
    mObj(vI).Point(vJ).Y = ((mObj(vI).Point(vJ).Y - mBorderWidth) * Ratio) + mBorderWidth + movY
    Next vJ
Next vI

End Sub

Private Function OpenPSFile(FileName As String) As Boolean
On Error Resume Next
Dim vI As Long
Dim vJ As Long
Dim FF As Long
Dim tmpFile As String
Dim tLine() As String
Dim tLineCount As Long
Dim tStr As String
Dim tCol As Long
Dim tQty As Integer
Dim OldC As Long
Dim tH As Long
Dim toCancel As Boolean

tH = 5000

FF = FreeFile

tmpFile = Space(FileLen(FileName))
Open FileName For Binary As #FF
Get #FF, , tmpFile
Close #FF

tmpFile = Split(tmpFile, "%%EndSetup", , vbTextCompare)(1)

tLine = Split(tmpFile, Chr(13))

tLineCount = UBound(tLine)

For vI = 0 To tLineCount
RaiseEvent OpenProgress(vI, tLineCount)
DrawProg vI, tLineCount
tLine(vI) = Replace(tLine(vI), Chr(10), "")
    If Len(tLine(vI)) > 0 Then
    tStr = Split(tLine(vI), " ")(UBound(Split(tLine(vI), " ")))
    tStr = Replace(tStr, "_", "")
    tLine(vI) = Replace(tLine(vI), Chr(13), "")
    Else
    tStr = ""
    End If
    Select Case tStr
        Case "Xa", "setrgbcolor"
        tCol = RGB(Split(tLine(vI), " ")(0) * 255, Split(tLine(vI), " ")(1) * 255, Split(tLine(vI), " ")(2) * 255)
            
        Case "k", "setfillcolor", "K"
            If LCase(tStr) = "k" Then
            tCol = RGBfromCMYK(Left(tLine(vI), InStrRev(tLine(vI), "k", , vbTextCompare) - 1))
            Else
            tCol = RGBfromCMYK(Left(tLine(vI), InStrRev(tLine(vI), "set") - 1))
            End If

        Case "m", "moveto"
        Dim tSL As String
        If mObjCount > 0 Then
        tSL = Replace(Replace((tLine(vI)), Chr(13), ""), Chr(10), "")
            If CLng(Trim(Split(tSL, " ")(0)) * 10) = mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X And _
            tH - CLng(Trim(Split(tSL, " ")(1)) * 10) = mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y Then
            GoTo Nx
            End If
        End If
        tQty = tQty + 1
        mObjCount = mObjCount + 1
        ReDim Preserve mObj(mObjCount)
        mObj(mObjCount - 1).PointCount = 1
        ReDim mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount)
        ReDim mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X = (Trim(Split(tLine(vI), " ")(0)))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y = tH - (Trim(Split(tLine(vI), " ")(1)))
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount) = ptMove
        mObj(mObjCount - 1).Color = tCol

        Case "l", "L", "lineto"
        mObj(mObjCount - 1).PointCount = mObj(mObjCount - 1).PointCount + 1
        ReDim Preserve mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount)
        ReDim Preserve mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X = (Trim(Split(tLine(vI), " ")(0)))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y = tH - (Trim(Split(tLine(vI), " ")(1)))
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount) = ptLine
        mObj(mObjCount - 1).Color = tCol

        Case "c", "C", "curveto"
        mObj(mObjCount - 1).PointCount = mObj(mObjCount - 1).PointCount + 3
        ReDim Preserve mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount)
        ReDim Preserve mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 2).X = _
        (Split(tLine(vI), " ")(0))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 2).Y = tH - _
        (Split(tLine(vI), " ")(1))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 1).X = _
        (Split(tLine(vI), " ")(2))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 1).Y = tH - _
        (Split(tLine(vI), " ")(3))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X = _
        (Split(tLine(vI), " ")(4))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y = tH - _
        (Split(tLine(vI), " ")(5))
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount - 2) = ptBezier
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount - 1) = ptBezier
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount) = ptBezier

        Case "v", "V", "y", "Y", "arc"
        mObj(mObjCount - 1).PointCount = mObj(mObjCount - 1).PointCount + 3
        ReDim Preserve mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount)
        ReDim Preserve mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount)
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 2).X = _
        (Split(tLine(vI), " ")(0))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 2).Y = tH - _
        (Split(tLine(vI), " ")(1))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 1).X = _
        (Split(tLine(vI), " ")(0))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount - 1).Y = tH - _
        (Split(tLine(vI), " ")(1))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).X = _
        (Split(tLine(vI), " ")(2))
        mObj(mObjCount - 1).Point(mObj(mObjCount - 1).PointCount).Y = tH - _
        (Split(tLine(vI), " ")(3))
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount - 2) = ptBezier
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount - 1) = ptBezier
        mObj(mObjCount - 1).PointType(mObj(mObjCount - 1).PointCount) = ptBezier
    End Select
Nx:
Next vI

OpenPSFile = CBool(mObjCount)
End Function

Private Function RGBfromCMYK(ByVal cString As String) As Long
Dim iRed As Integer
Dim iGreen As Integer
Dim iBlue As Integer
Dim iC As Integer
Dim iM As Integer
Dim iY As Integer
Dim iB As Integer
Dim tC As Long
Dim tM As Long
Dim tY As Long
Dim tB As Long
Dim tRed As Integer
Dim tGreen As Integer
Dim tBlue As Integer

If InStr(cString, "(") > 0 Then
cString = Left(cString, InStr(cString, "(") - 1)
End If

cString = Replace(cString, Chr(10), "")
cString = Replace(cString, Chr(13), "")
cString = Replace(cString, "[", "")
cString = Replace(cString, "]", "")
cString = Replace(cString, "null", "")
cString = Trim(cString)

If UBound(Split(cString, " ")) = 4 Then
tC = (Split(cString, " ")(1)) * 100
tM = (Split(cString, " ")(2)) * 100
tY = (Split(cString, " ")(3)) * 100
tB = (Split(cString, " ")(4)) * 100
ElseIf UBound(Split(cString, " ")) = 3 Then
tC = (Split(cString, " ")(0)) * 100
tM = (Split(cString, " ")(1)) * 100
tY = (Split(cString, " ")(2)) * 100
tB = (Split(cString, " ")(3)) * 100
Else
RGBfromCMYK = 0
Exit Function
End If

iC = tC * 2.55
iM = tM * 2.55
iY = tY * 2.55
iB = tB * 2.55

iRed = (255 - iC - iB)
iGreen = (255 - iM - iB)
iBlue = (255 - iY - iB)

If iRed < 0 Then tRed = iRed / -1 Else tRed = iRed
If iGreen < 0 Then tGreen = iGreen / -1 Else tGreen = iGreen
If iBlue < 0 Then tBlue = iBlue / -1 Else tBlue = iBlue

RGBfromCMYK = RGB(tRed, tGreen, tBlue)
End Function


Private Sub DrawVector()
Dim vI As Long
Dim vJ As Long
Dim vK As Long
Dim vL As Long
Dim pQ As Long
Dim Pt() As POINTF
Dim PTtype() As POINT_TYPE
Dim ptMove As POINTF
Dim vGraphics As Long
Dim vPen As Long
Dim vPath As Long
Dim vBrush As Long
Dim tCol As Long

UserControl.Cls
UserControl.BackColor = vbWhite
If mObjCount = 0 Then Exit Sub

GdipCreateFromHDC UserControl.hDC, vGraphics
Call GdipCreatePath(FillModeAlternate, vPath)

If mRendered = True Then Call GdipSetSmoothingMode(vGraphics, SmoothingModeAntiAlias)

For vI = 0 To mObjCount - 1
pQ = 0
tCol = mObj(vI).Color
    For vJ = 1 To mObj(vI).PointCount
    pQ = pQ + 1
    ReDim Preserve Pt(pQ)
    ReDim Preserve PTtype(pQ)
    Pt(pQ).X = mObj(vI).Point(vJ).X
    Pt(pQ).Y = mObj(vI).Point(vJ).Y
    PTtype(pQ) = mObj(vI).PointType(vJ)
    Next vJ

    For vK = 1 To pQ
        Select Case PTtype(vK)
            Case 6
            GdipStartPathFigure vPath
            ptMove = Pt(vK)
            Case 2
            GdipAddPathLine vPath, ptMove.X, ptMove.Y, Pt(vK).X, Pt(vK).Y
            ptMove = Pt(vK)
            Case 4
            GdipAddPathBezier vPath, ptMove.X, ptMove.Y, Pt(vK).X, Pt(vK).Y, _
            Pt(vK + 1).X, Pt(vK + 1).Y, Pt(vK + 2).X, Pt(vK + 2).Y
            ptMove = Pt(vK + 2)
            vK = vK + 2
        End Select
    Next vK

If tCol < 0 Then
tCol = 0
End If

    Select Case mDrawMode
        Case BlackLines
            Call GdipCreatePen1(GetRGB_VB2GDIP(vbBlack), 1, UnitPixel, vPen)
            Call GdipDrawPath(vGraphics, vPen, vPath)
            Call GdipDeletePen(vPen)
        Case ColoredBorder
            If tCol = vbWhite Then
            tCol = &HF0F0F0
            End If
            Call GdipCreatePen1(GetRGB_VB2GDIP(tCol), 1, UnitPixel, vPen)
            Call GdipDrawPath(vGraphics, vPen, vPath)
            Call GdipDeletePen(vPen)
        Case Filled
                If mTransparent < 255 Then
                Call GdipCreatePen1(GetRGB_VB2GDIP(tCol), 1, UnitPixel, vPen)
                Call GdipDrawPath(vGraphics, vPen, vPath)
                Call GdipDeletePen(vPen)
                End If
            If tCol = vbWhite Then
            tCol = &HF0F0F0
            End If
            Call GdipCreateSolidFill(GetRGB_VB2GDIP(tCol, mTransparent), vBrush)
            Call GdipFillPath(vGraphics, vBrush, vPath)
            Call GdipDeleteBrush(vBrush)
    End Select
GdipResetPath vPath
Next vI

Call GdipDeleteGraphics(vGraphics)
Call GdipDeletePen(vPen)
Call GdipDeletePath(vPath)

UserControl.Refresh
End Sub

Public Property Get hDC() As Long
hDC = UserControl.hDC
End Property
Private Sub UserControl_InitProperties()
mBorderWidth = 10
mDrawMode = Filled
mRendered = True
mTransparent = 255
mShowProgress = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    mBorderWidth = .ReadProperty("BorderWidth", 10)
    mDrawMode = .ReadProperty("DrawMode", 2)
    mRendered = .ReadProperty("Rendered", True)
    mTransparent = .ReadProperty("Transparent", 255)
    mShowProgress = .ReadProperty("ShowProgress", False)
End With
End Sub




Private Sub UserControl_Terminate()
Call GdiplusShutdown(token)
End Sub



Public Property Get BorderWidth() As Long
BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(ByVal vNewValue As Long)
mBorderWidth = vNewValue
PropertyChanged "BorderWidth"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "BorderWidth", mBorderWidth, 10
    .WriteProperty "DrawMode", mDrawMode, 2
    .WriteProperty "Rendered", mRendered, True
    .WriteProperty "Transparent", mTransparent, 255
    .WriteProperty "ShowProgress", mShowProgress, False
End With
End Sub



Public Property Get DrawMode() As eDrawMode
DrawMode = mDrawMode
End Property

Public Property Let DrawMode(ByVal vNewValue As eDrawMode)
mDrawMode = vNewValue
DrawVector
PropertyChanged "DrawMode"
End Property

Public Property Get Rendered() As Boolean
Rendered = mRendered
End Property

Public Property Let Rendered(ByVal vNewValue As Boolean)
mRendered = vNewValue
DrawVector
PropertyChanged "Rendered"
End Property

Public Property Get Transparent() As Byte
Transparent = mTransparent
End Property

Public Property Let Transparent(ByVal vNewValue As Byte)
mTransparent = vNewValue
DrawVector
PropertyChanged "Transparent"
End Property

Public Sub SaveImage(BMPFileName As String)
'You can easily save to jpeg,gif,png... with gdiplus
SavePicture UserControl.image, BMPFileName
End Sub

Public Property Get ShowProgress() As Boolean
ShowProgress = mShowProgress
End Property

Public Property Let ShowProgress(ByVal vNewValue As Boolean)
mShowProgress = vNewValue
PropertyChanged "ShowProgress"
End Property

Private Sub DrawProg(ByVal Val As Long, ByVal Max As Long)
Dim Percent As String
Static LastP As String

If mShowProgress = False Then Exit Sub

Percent = Int((Val / Max) * 100) & "%"

If LastP = Percent Then Exit Sub

LastP = Percent

With UserControl
    .CurrentX = (.ScaleWidth - .TextWidth(Percent)) / 2
    .CurrentY = (.ScaleHeight - .TextHeight(Percent)) / 2
End With

UserControl.Print Percent
DoEvents
End Sub
