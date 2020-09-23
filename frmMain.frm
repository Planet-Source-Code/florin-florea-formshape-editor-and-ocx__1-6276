VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Orice Forma"
   ClientHeight    =   9975
   ClientLeft      =   255
   ClientTop       =   825
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   12345
   Begin VB.CommandButton Command13 
      Caption         =   ">"
      Height          =   285
      Left            =   11730
      TabIndex        =   21
      Top             =   5190
      Width           =   300
   End
   Begin VB.CommandButton Command12 
      Caption         =   "<"
      Height          =   285
      Left            =   11115
      TabIndex        =   20
      Top             =   5205
      Width           =   300
   End
   Begin VB.CommandButton Command11 
      Caption         =   "\/"
      Height          =   285
      Left            =   11415
      TabIndex        =   19
      Top             =   5430
      Width           =   300
   End
   Begin VB.CommandButton Command10 
      Caption         =   "^"
      Height          =   285
      Left            =   11415
      TabIndex        =   18
      Top             =   4965
      Width           =   300
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   7335
      TabIndex        =   17
      Top             =   120
      Width           =   1185
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Intersect"
      Height          =   375
      Left            =   11010
      TabIndex        =   16
      Top             =   6810
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Animate"
      Height          =   375
      Left            =   11010
      TabIndex        =   15
      Top             =   6330
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load frames"
      Height          =   375
      Left            =   11010
      TabIndex        =   14
      Top             =   5850
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11130
      Top             =   5370
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hide"
      Height          =   375
      Left            =   11010
      TabIndex        =   9
      Top             =   1950
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load"
      Height          =   375
      Left            =   11010
      TabIndex        =   7
      Top             =   2895
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   375
      Left            =   11025
      TabIndex        =   6
      Top             =   2415
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11010
      Top             =   3855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Alternate"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11025
      TabIndex        =   5
      Top             =   285
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Winding"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11025
      TabIndex        =   4
      Top             =   30
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   11010
      TabIndex        =   3
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Left            =   11010
      TabIndex        =   2
      Top             =   990
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   9135
      Left            =   75
      ScaleHeight     =   9075
      ScaleWidth      =   10695
      TabIndex        =   0
      Top             =   510
      Width           =   10755
      Begin VB.PictureBox Punct 
         BackColor       =   &H00FFFFFF&
         Height          =   70
         Index           =   0
         Left            =   570
         MouseIcon       =   "frmMain.frx":0000
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   8
         Top             =   330
         Visible         =   0   'False
         Width           =   70
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   11595
      Top             =   3855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp, *.gif,*.jpg|*.bmp;*.gif;*.jpg"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   6900
      TabIndex        =   24
      Top             =   165
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move picture"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   11100
      TabIndex        =   23
      Top             =   4575
      Width           =   930
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Method"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   10215
      TabIndex        =   22
      Top             =   60
      Width           =   630
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   150
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   150
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Draw shape here or load a picture"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2400
   End
   Begin VB.Menu mnuMatrita 
      Caption         =   "Load picture"
   End
   Begin VB.Menu mnu1 
      Caption         =   ""
   End
   Begin VB.Menu mnuTraseazaContur 
      Caption         =   "Trace"
   End
   Begin VB.Menu mnu2 
      Caption         =   ""
   End
   Begin VB.Menu mnuCuloare 
      Caption         =   "Color"
   End
   Begin VB.Menu mnu3 
      Caption         =   ""
   End
   Begin VB.Menu mnuLegaturi 
      Caption         =   "Set Limits"
      Begin VB.Menu mnuStangaSus 
         Caption         =   "UpperLeft"
      End
      Begin VB.Menu mnuDreaptaJos 
         Caption         =   "LowerRight"
      End
   End
   Begin VB.Menu mnu4 
      Caption         =   ""
   End
   Begin VB.Menu mnuUmbra 
      Caption         =   "Shadow"
   End
   Begin VB.Menu mnufond 
      Caption         =   "Background Color"
   End
   Begin VB.Menu mnuShrink 
      Caption         =   "Shrink"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Shrink(z As Integer)
Dim i As Long
For i = 0 To NumPoints
    X = X + Points(i).X
    Y = Y + Points(i).Y
Next i
X = X / NumPoints
Y = Y / NumPoints

For i = 0 To NumPoints
    If Points(i).X > X Then
        Points(i).X = Points(i).X - z ' Screen.TwipsPerPixelX
    Else
    If Points(i).X < X Then
        Points(i).X = Points(i).X + z 'Screen.TwipsPerPixelX
    End If
    End If
    If Points(i).Y > Y Then
        Points(i).Y = Points(i).Y - z 'Screen.TwipsPerPixelY
    Else
    If Points(i).Y < Y Then
        Points(i).Y = Points(i).Y + z 'Screen.TwipsPerPixelY
    End If
    End If
Next i
End Sub

Private Sub Command1_Click()
    If NumPoints < 2 Then
        MsgBox "At least 3 points are needed."
        Exit Sub
    End If
    Load frmTest
    If Option1.Value Then Method = WINDING Else Method = ALTERNATE
    PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
    ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
    frmTest.Show
    Picture1.Enabled = False
End Sub

Private Sub Command10_Click()
frmTest.Image1.Top = frmTest.Image1.Top - 1
End Sub

Private Sub Command11_Click()
frmTest.Image1.Top = frmTest.Image1.Top + 1
End Sub

Private Sub Command12_Click()
frmTest.Image1.Left = frmTest.Image1.Left - 1
End Sub

Private Sub Command13_Click()
frmTest.Image1.Left = frmTest.Image1.Left + 1
End Sub

Private Sub Command2_Click()
    Dim i As Long
    Dim R As Integer
    On Error Resume Next
R = MsgBox("Are you shure?", vbYesNo)
If R = 6 Then
    ReDim Points(0)
    For i = 1 To NumPoints + 1
        Unload Punct(i)
    Next i
    NumPoints = -1
    Picture1.Cls
    Unload frmTest
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Unload frmTest
Me.Picture1.Enabled = True
End Sub

Private Sub Command4_Click()
Dim Fisier As String
Dim i As Long

Fisier = InputBox("Nume Fisier", "Save")
If Fisier <> "" Then
If Dir(Fisier) <> "" Then
Kill Fisier
F = FreeFile
Open Fisier For Random As F
For i = 0 To NumPoints
Put F, , Points(i)
Next i
Close F
Else
F = FreeFile
Open Fisier For Random As F
For i = 0 To NumPoints
Put F, , Points(i)
Next i
Close F
End If
End If
End Sub

Private Sub Command5_Click()
Dim S As String
Dim Fisier As String
Dim i As Long
Dim R As POINTDATA
Fisier = ""
S = Me.CommonDialog1.Filter
Me.CommonDialog1.Filter = "*.txt|*.txt"
Me.CommonDialog1.ShowOpen
Fisier = Me.CommonDialog1.FileName
If Fisier <> "" Then
Call Command2_Click

F = FreeFile
Open Fisier For Random As F
i = 0
While Not EOF(F)
ReDim Preserve Points(i)
Get F, , R
Points(i) = R
i = i + 1
Wend
NumPoints = i - 1
    If NumPoints < 2 Then
        MsgBox "At least 3 points are needed."
        Exit Sub
    End If
    On Error Resume Next

    Picture1.Cls
    Picture1.PSet (Points(0).X * Screen.TwipsPerPixelX, Points(0).Y * Screen.TwipsPerPixelX)
    For i = 1 To NumPoints
            Picture1.Line -(Points(i).X * Screen.TwipsPerPixelX, Points(i).Y * Screen.TwipsPerPixelY)
    Next i
    Load frmTest
    If Option1.Value Then Method = WINDING Else Method = ALTERNATE
    PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
    ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
    frmTest.BackColor = Picture2.BackColor
    frmTest.Show
    Me.Picture1.Enabled = False
End If
Close F
End Sub

Private Sub Command6_Click()
Dim S As String
Dim Fisier As String
Dim i As Long
Dim R As POINTDATA
Dim Ind As Long
On Error GoTo xxx
Ind = InputBox("Frame = ", "")
Fisier = ""
S = Me.CommonDialog1.Filter
Me.CommonDialog1.Filter = "*.txt|*.txt"
Me.CommonDialog1.ShowOpen
Fisier = Me.CommonDialog1.FileName
If Fisier <> "" Then


F = FreeFile
Open Fisier For Random As F
i = 0
Select Case Ind
Case 1
While Not EOF(F)
ReDim Preserve Points1(i)
Get F, , R
Points1(i) = R
i = i + 1
Wend
NumPoints1 = i - 1
    If NumPoints1 < 2 Then
        MsgBox "At least 3 points1 are needed."
        Exit Sub
    End If


Case 2
While Not EOF(F)
ReDim Preserve Points2(i)
Get F, , R
Points2(i) = R
i = i + 1
Wend
NumPoints2 = i - 1
    If NumPoints2 < 2 Then
        MsgBox "At least 3 points1 are needed."
        Exit Sub
    End If


Case 3
While Not EOF(F)
ReDim Preserve Points3(i)
Get F, , R
Points3(i) = R
i = i + 1
Wend
NumPoints3 = i - 1
    If NumPoints3 < 2 Then
        MsgBox "At least 3 points3 are needed."
        Exit Sub
    End If


Case 4
While Not EOF(F)
ReDim Preserve Points4(i)
Get F, , R
Points4(i) = R
i = i + 1
Wend
NumPoints4 = i - 1
    If NumPoints4 < 2 Then
        MsgBox "At least 3 points4 are needed."
        Exit Sub
    End If


End Select
End If
Close F
Exit Sub
xxx:
MsgBox Err.Description
End Sub

Private Sub Command7_Click()
If Timer1.Enabled = False Then
Me.Timer1.Enabled = True
Else
Me.Timer1.Enabled = False
End If
End Sub

Private Sub Command8_Click()
Const RGN_DIFF = 4
Dim combined_rgn  As Long
Dim outer_rgn As Long
Dim inner_rgn As Long
On Error GoTo xxx
inner_rgn = CreatePolygonRgn(Points1(0), NumPoints1 + 1, Method)
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
    CombineRgn combined_rgn, PolyRegion, inner_rgn, RGN_DIFF
    Load frmTest
     SetWindowRgn frmTest.hWnd, combined_rgn, True
     frmTest.Show
Exit Sub
xxx:
MsgBox "Load Frame 1"
End Sub



Private Sub Form_Load()
Dim R As Long

StabilirePunct = True
NumPoints = -1
Interfata = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTest
    End
End Sub

Private Sub mnuCuloare_Click()
PickColor = True
End Sub

Private Sub mnuDreaptaJos_Click()
DreaptaJos = True
End Sub

Private Sub mnufond_Click()
SCuloareFond = True
End Sub

Private Sub mnuMatrita_Click()
Dim S As Long
On Error GoTo xxx
Me.CommonDialog2.ShowOpen
If Me.CommonDialog2.FileName <> "" Then
Me.Picture1.Picture = LoadPicture(Me.CommonDialog2.FileName)
frmTest.Image1.Picture = LoadPicture(Me.CommonDialog2.FileName)
End If
Exit Sub
xxx:
MsgBox Err.Description
End Sub

Private Sub mnuShrink_Click()
Dim z As Long
z = Val(InputBox("Zoom = ", ""))
Shrink (z)
End Sub

Private Sub mnuStangaSus_Click()
StangaSus = True
End Sub

Private Sub mnuTraseazaContur_Click()
Dim X As Long, Y As Long
Dim R As Long
Xinitial = -20000
Yinitial = -20000

frmMain.Picture1.ForeColor = vbBlack
For X = Xmin To XMax Step Screen.TwipsPerPixelX
    For Y = Ymin To Ymax Step Screen.TwipsPerPixelY
       
       If Me.Picture1.Point(X, Y) <> 16777215 Then
        With Me.Picture1
            If .Point(X, Y) = culoare Then

If frmMain.Picture1.Point(X - 1 * Screen.TwipsPerPixelX, Y - 1 * Screen.TwipsPerPixelY) <> culoare Then
    frmMain.Picture1.PSet (X - 1 * Screen.TwipsPerPixelX, Y - 1 * Screen.TwipsPerPixelY)
End If
If frmMain.Picture1.Point(X - 1, Y) <> culoare Then
    frmMain.Picture1.PSet (X - 1 * Screen.TwipsPerPixelX, Y)
End If
If frmMain.Picture1.Point(X - 1 * Screen.TwipsPerPixelX, Y + 1 * Screen.TwipsPerPixelY) <> culoare Then
    frmMain.Picture1.PSet (X - 1 * Screen.TwipsPerPixelX, Y + 1 * Screen.TwipsPerPixelY)
End If
If frmMain.Picture1.Point(X, Y + 1 * Screen.TwipsPerPixelY) <> culoare Then
    frmMain.Picture1.PSet (X, Y + 1 * Screen.TwipsPerPixelY)
End If
If frmMain.Picture1.Point(X + 1 * Screen.TwipsPerPixelX, Y + 1 * Screen.TwipsPerPixelY) <> culoare Then
    frmMain.Picture1.PSet (X + 1 * Screen.TwipsPerPixelX, Y + 1 * Screen.TwipsPerPixelY)
End If
If frmMain.Picture1.Point(X + 1 * Screen.TwipsPerPixelX, Y) <> culoare Then
    frmMain.Picture1.PSet (X + 1 * Screen.TwipsPerPixelX, Y)
End If
If frmMain.Picture1.Point(X + 1 * Screen.TwipsPerPixelX, Y - 1 * Screen.TwipsPerPixelY) <> culoare Then
    frmMain.Picture1.PSet (X + 1 * Screen.TwipsPerPixelX, Y - 1 * Screen.TwipsPerPixelY)
End If
If frmMain.Picture1.Point(X, Y - 1 * Screen.TwipsPerPixelY) <> culoare Then
    frmMain.Picture1.PSet (X, Y - 1 * Screen.TwipsPerPixelY)
End If
            End If
        End With
       End If
    Next Y

Text2.Text = X
Me.Picture1.Refresh
DoEvents
Next X
Me.Picture1.Refresh
Beep
Beep
Beep

For X = Xmin + 1 To XMax - 1 Step Screen.TwipsPerPixelX
    For Y = Ymin + 1 To Ymax - 1 Step Screen.TwipsPerPixelY
    If Me.Picture1.Point(X, Y) = vbBlack Then
        Xinitial = X
        Yinitial = Y
        Contur Xinitial, Yinitial
        Exit Sub
    End If
    Next Y
Next X
End Sub

Private Sub mnuUmbra_Click()
Dim X As Long
Dim Y As Long
Me.Picture1.ForeColor = vbBlue
For X = Xmin To XMax Step Screen.TwipsPerPixelX
    For Y = Ymin To Ymax Step Screen.TwipsPerPixelY
        If Me.Picture1.Point(X, Y) <> CuloareFond Then
            Me.Picture1.PSet (X, Y)
        End If
    Next Y
Me.Text1.Text = X
DoEvents
Next X
culoare = vbBlue
End Sub

Private Sub Option1_Click()
    Option2.Value = False
End Sub

Private Sub Option2_Click()
    Option1.Value = False
End Sub

Private Sub Picture0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Desenare Then
If Centru Then
Xc = X
Yc = Y
Centru = False
Else
Desenare = False
R = Sqr((Xc - X) * (Xc - X) + (Yc - Y) * (Yc - Y))
Me.Picture1.Circle (Xc, Yc), R
Me.Picture1.ForeColor = 0
Picture1.PSet (XO, YO)
End If
Else
If XX Then
XV = X
YV = Y
StabilirePunct = False
Else
    
    NumPoints = NumPoints + 1
    
    Load Punct(NumPoints + 1)
    Punct(NumPoints + 1).Left = X - Punct(NumPoints + 1).Width / 2
    Punct(NumPoints + 1).Top = Y - Punct(NumPoints + 1).Height / 2
    Punct(NumPoints + 1).Visible = True
    
    ReDim Preserve Points(NumPoints)
    
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
    
    If NumPoints = 0 Then Picture1.PSet (X, Y)
    Picture1.Line -(X, Y)
    XO = X
    YO = Y
    End If
End If
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PickColor Then
culoare = Me.Picture1.Point(X, Y)
PickColor = False
Else
If StangaSus Then
    Xmin = X
    Ymin = Y
    StangaSus = False
    Else
            If DreaptaJos Then
                XMax = X
                Ymax = Y
                DreaptaJos = False
            Else
                    If SCuloareFond Then
                        CuloareFond = Picture1.Point(X, Y)
                    Else
                            NumPoints = NumPoints + 1
                            ReDim Preserve Points(NumPoints)
                            Points(NumPoints).X = X / Screen.TwipsPerPixelX
                            Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
                            If NumPoints = 0 Then Picture1.PSet (X, Y)
                                Picture1.Line -(X, Y)
                            End If
                    End If
            End If
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Integer
Text1.Text = X
Text2.Text = Y
c1& = Me.Picture1.Point(X, Y)
Text3.Text = c1

End Sub





Public Sub Punct_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        OldXP = X
        OldYP = Y
        MoveItP = True
    End If
End Sub

Private Sub Punct_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
   If MoveItP Then
    
        Punct(Index).Left = Punct(Index).Left + X - OldXP
        Punct(Index).Top = Punct(Index).Top + Y - OldYP
        Points(Index - 1).X = (Punct(Index).Left + Punct(Index).Width / 2) / Screen.TwipsPerPixelX
        Points(Index - 1).Y = (Punct(Index).Top + Punct(Index).Height / 2) / Screen.TwipsPerPixelY
        Picture1.Cls
        Picture1.PSet (Points(0).X * Screen.TwipsPerPixelX, Points(0).Y * Screen.TwipsPerPixelX)
        For i = 1 To NumPoints
                Picture1.Line -(Points(i).X * Screen.TwipsPerPixelX, Points(i).Y * Screen.TwipsPerPixelY)
        Next i
    Else
    MousePointer = 99
    Punct(Index).Visible = True
    End If

End Sub

Private Sub Punct_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveItP = False
    Punct(Index).Visible = False
End Sub

Private Sub Timer1_Timer()
On Error GoTo xxx
Select Case Interfata
Case 0
PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.Show
Interfata = 1

Case 1
PolyRegion = CreatePolygonRgn(Points1(0), NumPoints1 + 1, Method)
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.Show
Interfata = 2

Case 2
PolyRegion = CreatePolygonRgn(Points2(0), NumPoints2 + 1, Method)
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.Show
Interfata = 3

Case 3
PolyRegion = CreatePolygonRgn(Points3(0), NumPoints3 + 1, Method)
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.Show
Interfata = 4

Case 4
PolyRegion = CreatePolygonRgn(Points4(0), NumPoints4 + 1, Method)
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.Show
Interfata = 0
End Select
Exit Sub
xxx:
Timer1.Enabled = False
End Sub



Public Sub Contur(Xinitial As Long, Yinitial As Long)

Dim X As Long
Dim Y As Long
Dim p  As Long
Dim Xtemp As Long, Ytemp As Long
Dim i As Long
Dim j As Long

X = Xinitial
Y = Yinitial
Xtemp = X
Ytemp = Y

NumPoints = NumPoints + 1
ReDim Preserve Points(NumPoints)
Points(NumPoints).X = Xinitial / Screen.TwipsPerPixelX
Points(NumPoints).Y = Yinitial / Screen.TwipsPerPixelY
    
    
With Me.Picture1
.ForeColor = vbRed
p = .Point(X - 1 * Screen.TwipsPerPixelX, Y)
If p = vbBlack Then
    Me.Picture1.PSet (X - 1 * Screen.TwipsPerPixelX, Y)
    X = X - 1 * Screen.TwipsPerPixelX
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X, Y + 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X, Y + 1 * Screen.TwipsPerPixelY)
    Y = Y + 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X + 1 * Screen.TwipsPerPixelX, Y)
If p = vbBlack Then
    Me.Picture1.PSet (X + 1 * Screen.TwipsPerPixelX, Y)
    X = X + 1 * Screen.TwipsPerPixelX
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X, Y - 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X, Y - 1 * Screen.TwipsPerPixelY)
    Y = Y - 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
End If
End If
End If
End If

While Not (X = Xinitial And Y = Yinitial)
p = .Point(X - 1 * Screen.TwipsPerPixelX, Y)
If p = vbBlack Then
    Me.Picture1.PSet (X - 1 * Screen.TwipsPerPixelX, Y)
    X = X - 1 * Screen.TwipsPerPixelX
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X, Y + 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X, Y + 1 * Screen.TwipsPerPixelY)
    Y = Y + 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X + 1 * Screen.TwipsPerPixelX, Y)
If p = vbBlack Then
    Me.Picture1.PSet (X + 1 * Screen.TwipsPerPixelX, Y)
    X = X + 1 * Screen.TwipsPerPixelX
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X, Y - 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X, Y - 1 * Screen.TwipsPerPixelY)
    Y = Y - 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X - 1 * Screen.TwipsPerPixelX, Y + 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X - 1 * Screen.TwipsPerPixelX, Y)
    X = X - 1 * Screen.TwipsPerPixelX
    Y = Y + 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X + 1 * Screen.TwipsPerPixelX, Y + 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X, Y + 1 * Screen.TwipsPerPixelY)
    X = X + 1 * Screen.TwipsPerPixelX
    Y = Y + 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X + 1 * Screen.TwipsPerPixelX, Y - 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X + 1 * Screen.TwipsPerPixelX, Y)
    X = X + 1 * Screen.TwipsPerPixelX
    Y = Y - 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY
Else
p = .Point(X - 1 * Screen.TwipsPerPixelX, Y - 1 * Screen.TwipsPerPixelY)
If p = vbBlack Then
    Me.Picture1.PSet (X, Y - 1 * Screen.TwipsPerPixelY)
    X = X - 1 * Screen.TwipsPerPixelX
    Y = Y - 1 * Screen.TwipsPerPixelY
    NumPoints = NumPoints + 1
    ReDim Preserve Points(NumPoints)
    Points(NumPoints).X = X / Screen.TwipsPerPixelX
    Points(NumPoints).Y = Y / Screen.TwipsPerPixelY

End If
End If
End If
End If

End If
End If
End If
End If
DoEvents
If Xtemp = X And Ytemp = Y Then
    X = Xinitial
    Y = Yinitial
End If
Wend
End With

X = 0
Y = 0
For i = 0 To NumPoints
    X = X + Points(i).X
    Y = Y + Points(i).Y
Next i
X = X / NumPoints
Y = Y / NumPoints

For i = 0 To NumPoints
    If Points(i).X > X Then
        Points(i).X = Points(i).X - 1
    Else
    If Points(i).X < X Then
        Points(i).X = Points(i).X + 1
    End If
    End If
    If Points(i).Y > Y Then
        Points(i).Y = Points(i).Y - 1
    Else
    If Points(i).Y < Y Then
        Points(i).Y = Points(i).Y + 1
    End If
    End If
Next i

End Sub

Public Sub Smooth()
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

For i = 1 To Me.Picture.Width / Screen.TwipsPerPixelX Step 3
    For j = 1 To Me.Picture1.Height / Screen.TwipsPerPixelY Step 3
        For k = 1 To 3
            For l = 1 To 3
            
            Next l
        Next k
    Next j
Next i

End Sub
