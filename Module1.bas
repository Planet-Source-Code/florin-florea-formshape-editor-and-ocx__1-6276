Attribute VB_Name = "Module1"
Global Const ALTERNATE = 1
Global Const WINDING = 2
Global Const RGN_DIFF = 4

Type POINTDATA
    X As Long
    Y As Long
End Type

Public Celula(1 To 3, 1 To 3) As POINTDATA

Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTDATA, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public CuloareFond As Long
Public SCuloareFond As Boolean
Public t As Long
Public XV As Long
Public YV As Long
Public XX As Boolean
Public StabilirePunct As Boolean
Public F As Long
Public Desenare As Boolean
Public Centru As Boolean
Public Raza As Boolean
Public Xc As Double
Public Yc As Double
Public R As Double
Public XO As Double
Public YO As Double
Public OldXP As Double, OldYP As Double
Public MoveItP As Boolean
Public Points() As POINTDATA
Public Points1() As POINTDATA
Public Points2() As POINTDATA
Public Points3() As POINTDATA
Public Points4() As POINTDATA
Public Points5() As POINTDATA

Public NumPoints As Long
Public NumPoints1 As Long
Public NumPoints2 As Long
Public NumPoints3 As Long
Public NumPoints4 As Long
Public NumPoints5 As Long

Public PolyRegion As Long
Public ReturnVal As Long
Public Method As Long
Public Interfata As Double

Public Puncte1() As POINTDATA
Public NrPuncte As Long
Public culoare As Long
Public PickColor As Boolean
Public StangaSus As Boolean
Public DreaptaJos As Boolean
Public Xmin As Double, Ymin As Double, XMax As Double, Ymax As Double
Public Xinitial As Long, Yinitial As Long





Public Function Min(a As Double, b As Double) As Double
If a < b Then
    Min = a
Else
Min = b
End If
End Function

Public Function Max(a As Double, b As Double) As Double
If a > b Then
    Max = a
Else
    Max = b
End If
End Function



Public Sub TrasareContur()
Dim i As Long
Dim j As Long
Dim X As Long
Dim Y As Long
Dim p As Long
Dim Gata As Boolean
Dim XI As Long, YI As Long

i = 0
j = YV

While j < frmMain.Picture1.Height - 1
    If Not Gata Then
        j = j + 1
    Else
        j = frmMain.Picture1.Height - 1
    End If
    
    For i = XV To frmMain.Picture1.Width - 1
        p = frmMain.Picture1.Point(i, j)
        If p = 255 Then
            X = i - 1
            Y = j
            Gata = True
            Exit For
        End If
    Next i
Wend
NumPoints = 0
ReDim Preserve Points(NumPoints)
XI = X
YI = Y
Points(0).X = X
Points(0).Y = Y

Select Case IntoarceVecin(XI, YI)
Case 1:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI - 1
    Points(NumPoints).Y = YI
    XI = XI - 1
    YI = YI
Case 2:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI - 1
    Points(NumPoints).Y = YI - 1
    XI = XI - 1
    YI = YI - 1
Case 3:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI
    Points(NumPoints).Y = YI - 1
    XI = XI
    YI = YI - 1
Case 4:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI + 1
    Points(NumPoints).Y = YI - 1
    XI = XI + 1
    YI = YI - 1
Case 5:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI + 1
    Points(NumPoints).Y = YI
    XI = XI + 1
    YI = YI
Case 6:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI + 1
    Points(NumPoints).Y = YI - 1
    XI = XI + 1
    YI = YI - 1
Case 7:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI
    Points(NumPoints).Y = YI - 1
    XI = XI
    YI = YI - 1
Case 8:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI - 1
    Points(NumPoints).Y = YI - 1
    XI = XI - 1
    YI = YI - 1
End Select

While ((XI <> X) And (YI <> Y))
Select Case IntoarceVecin(XI, YI)
Case 1:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI - 1
    Points(NumPoints).Y = YI
    XI = XI - 1
    YI = YI
Case 2:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI - 1
    Points(NumPoints).Y = YI - 1
    XI = XI - 1
    YI = YI - 1
Case 3:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI
    Points(NumPoints).Y = YI - 1
    XI = XI
    YI = YI - 1
Case 4:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI + 1
    Points(NumPoints).Y = YI - 1
    XI = XI + 1
    YI = YI - 1
Case 5:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI + 1
    Points(NumPoints).Y = YI
    XI = XI + 1
    YI = YI
Case 6:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI + 1
    Points(NumPoints).Y = YI - 1
    XI = XI + 1
    YI = YI - 1
Case 7:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI
    Points(NumPoints).Y = YI - 1
    XI = XI
    YI = YI - 1
Case 8:
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = XI - 1
    Points(NumPoints).Y = YI - 1
    XI = XI - 1
    YI = YI - 1
End Select
Wend

PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)

End Sub



Public Function IntoarceVecin(X As Long, Y As Long) As Long
Dim p As Double
p = frmMain.Picture1.Point(X - 1, Y)
If p = 255 Then
    IntoarceVecin = 1
    Exit Function
End If

p = frmMain.Picture1.Point(X - 1, Y - 1)
If p = 255 Then
    IntoarceVecin = 2
    Exit Function
End If

p = frmMain.Picture1.Point(X, Y + 1)
If p = 255 Then
    IntoarceVecin = 3
    Exit Function
End If

p = frmMain.Picture1.Point(X + 1, Y + 1)
If p = 255 Then
    IntoarceVecin = 4
    Exit Function
End If

p = frmMain.Picture1.Point(X + 1, Y)
If p = 255 Then
    IntoarceVecin = 5
    Exit Function
End If

p = frmMain.Picture1.Point(X + 1, Y - 1)
If p = 255 Then
    IntoarceVecin = 6
    Exit Function
End If

p = frmMain.Picture1.Point(X, Y - 1)
If p = 255 Then
    IntoarceVecin = 7
    Exit Function
End If

p = frmMain.Picture1.Point(X - 1, Y - 1)
If p = 255 Then
    IntoarceVecin = 8
    Exit Function
End If
End Function
