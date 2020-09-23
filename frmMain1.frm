VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain1 
   Caption         =   "Orice Forma"
   ClientHeight    =   9435
   ClientLeft      =   255
   ClientTop       =   825
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   11145
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8910
      Top             =   -15
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Fade"
      Height          =   270
      Left            =   2145
      TabIndex        =   20
      Top             =   105
      Width           =   870
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   7095
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   120
      Width           =   1080
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Cut"
      Height          =   375
      Left            =   9840
      TabIndex        =   18
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   9840
      TabIndex        =   17
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load"
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9960
      Top             =   5400
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hide"
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load"
      Height          =   375
      Left            =   9840
      TabIndex        =   9
      Top             =   3345
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   375
      Left            =   9855
      TabIndex        =   8
      Top             =   2865
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   10800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Alternate"
      Height          =   255
      Left            =   9840
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Winding"
      Height          =   255
      Left            =   9840
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9840
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp|*.bmp"
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   8775
      Left            =   90
      ScaleHeight     =   8715
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   495
      Width           =   9615
      Begin VB.PictureBox Punct 
         BackColor       =   &H00FFFFFF&
         Height          =   105
         Index           =   0
         Left            =   600
         MouseIcon       =   "frmMain1.frx":0000
         ScaleHeight     =   45
         ScaleWidth      =   45
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "BackColor"
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Deseneaza forma aici"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1590
   End
   Begin VB.Menu mnuDesen 
      Caption         =   "Desene ajutatoare"
      Begin VB.Menu mnuCerc 
         Caption         =   "Cerc"
      End
   End
   Begin VB.Menu mnuMatrita 
      Caption         =   "Incarca Matrita"
   End
   Begin VB.Menu mnuArc 
      Caption         =   "Arc"
   End
   Begin VB.Menu mnuTraseazaContur 
      Caption         =   "Traseaza Contur"
   End
End
Attribute VB_Name = "frmMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    If NumPoints < 2 Then
        MsgBox "At least 3 points are needed."
        Exit Sub
    End If
    Load frmTest
    'Set the fill method
    'WINDING-Fills overlapping regions
    'ALTERNATE-Doesn't fill overlapping regions
    'Try drawing a pentagon to see the difference between the two methods.
    If Option1.Value Then Method = WINDING Else Method = ALTERNATE
    'Create the Region using predefined points
    PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
    'Set the window region of our form
    ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
    frmTest.BackColor = Picture2.BackColor
    frmTest.Show
    Picture1.Enabled = False
End Sub

Private Sub Command2_Click()
    Dim i As Long
    Dim R As Integer
    On Error Resume Next
    'Clear the picture box and point data
R = MsgBox("Sunteti sigur ca doriti sa stergeti?", vbYesNo)
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
    For i = 0 To NumPoints
    Load Punct(i + 1)
    Punct(i + 1).Left = Points(i).x * Screen.TwipsPerPixelX - Punct(i + 1).Width / 2
    Punct(i + 1).Top = Points(i).y * Screen.TwipsPerPixelX - Punct(i + 1).Height / 2
    Punct(i + 1).Visible = True
    Next i
    Picture1.Cls
    Picture1.PSet (Points(0).x * Screen.TwipsPerPixelX, Points(0).y * Screen.TwipsPerPixelX)
    For i = 1 To NumPoints
            Picture1.Line -(Points(i).x * Screen.TwipsPerPixelX, Points(i).y * Screen.TwipsPerPixelY)
    Next i
    Load frmTest
    'Set the fill method
    'WINDING-Fills overlapping regions
    'ALTERNATE-Doesn't fill overlapping regions
    'Try drawing a pentagon to see the difference between the two methods.
    If Option1.Value Then Method = WINDING Else Method = ALTERNATE
    'Create the Region using predefined points
    PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
    'Set the window region of our form
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
inner_rgn = CreatePolygonRgn(Points1(0), NumPoints1 + 1, Method)
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
    CombineRgn combined_rgn, PolyRegion, inner_rgn, RGN_DIFF
    Load frmTest
     SetWindowRgn frmTest.hWnd, combined_rgn, True
     frmTest.Show
End Sub

Private Sub Command9_Click()
Me.Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Dim R As Long
StabilirePunct = True
NumPoints = -1
Interfata = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTest
End Sub

Private Sub mnuArc_Click()
Arc = True
End Sub

Private Sub mnuCerc_Click()
If Desenare = False Then
Me.Picture1.ForeColor = 4000
Desenare = True
Centru = True
Else
Me.Picture1.ForeColor = 0
Desenare = False
End If
End Sub

Private Sub mnuMatrita_Click()
Dim S As Long
Me.CommonDialog2.ShowOpen
If Me.CommonDialog2.FileName <> "" Then
Me.Picture1.Picture = LoadPicture(Me.CommonDialog2.FileName)
End If
End Sub

Private Sub mnuTraseazaContur_Click()
'Dim i As Long
'Dim j As Long
'Dim Gata As Boolean
'Dim x As Long
'Dim y As Long
'Gata = False
'i = 0
'j = 0
'While Not Gata
'If i < Me.Picture.Width Then
'        i = i + 1
        
'Else
'    i = 1
'End If

'If j < Me.Picture1.Height Then
'        j = j + 1
'Else
'        Gata = True
'End If
'Wend
If StabilirePunct Then
XX = True
Else
StabilirePunct = True
TrasareContur
End If
End Sub

Private Sub Option1_Click()
    Option2.Value = False
End Sub

Private Sub Option2_Click()
    Option1.Value = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
If Desenare Then
If Centru Then
Xc = x
Yc = y
Centru = False
Else
Desenare = False
R = Sqr((Xc - x) * (Xc - x) + (Yc - y) * (Yc - y))
Me.Picture1.Circle (Xc, Yc), R
Me.Picture1.ForeColor = 0
Picture1.PSet (XO, YO)
End If
Else
If XX Then
XV = x
YV = y
StabilirePunct = False
Else
    'Increase the number of points
    NumPoints = NumPoints + 1
    
    Load Punct(NumPoints + 1)
    Punct(NumPoints + 1).Left = x - Punct(NumPoints + 1).Width / 2
    Punct(NumPoints + 1).Top = y - Punct(NumPoints + 1).Height / 2
    Punct(NumPoints + 1).Visible = True
    
    ReDim Preserve Points(NumPoints)
    'Add the new point to the Point data
    Points(NumPoints).x = x / Screen.TwipsPerPixelX
    Points(NumPoints).y = y / Screen.TwipsPerPixelY
    'Draw the new point ob the picturebox.
    If NumPoints = 0 Then Picture1.PSet (x, y)
    Picture1.Line -(x, y)
    XO = x
    YO = y
    End If
End If
Else
    If Arc Then
        XInceputArc = x
        YInceputArc = y
        Arc = False
        SfarsitArc = True
    Else
        If SfarsitArc Then
            XSfarsitArc = x
            YSfarsitArc = y
            SfarsitArc = False
            CentruArc = True
        Else
            If CentruArc Then
                XCentruArc = x
                YCentruArc = y
                DesenareArc XInceputArc, YInceputArc, XSfarsitArc, YSfarsitArc, XCentruArc, YCentruArc, 20
                CentruArc = False
            End If
        End If
    End If
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim c As Integer
Text1.Text = x
Text2.Text = y
c1& = Me.Picture1.Point(x, y)
Text3.Text = c1
End Sub

Private Sub Picture2_Click()
    On Error Resume Next
    CommonDialog1.ShowColor
    If Err.Number <> 0 Then Exit Sub
    'Set form backcolor
    frmTest.BackColor = CommonDialog1.Color
    Picture2.BackColor = CommonDialog1.Color
    Picture1.ForeColor = CommonDialog1.Color
    If frmTest.Visible Then frmTest.Show
End Sub

Public Sub Punct_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'If the button is the right mouse button then unload the form
    'else keep the mouse coordinates to move the form.
    If Button = vbLeftButton Then
        OldXP = x
        OldYP = y
        MoveItP = True
    End If
End Sub

Private Sub Punct_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
   If MoveItP Then
        'Set new window position
        Punct(Index).Left = Punct(Index).Left + x - OldXP
        Punct(Index).Top = Punct(Index).Top + y - OldYP
        Points(Index - 1).x = (Punct(Index).Left + Punct(Index).Width / 2) / Screen.TwipsPerPixelX
        Points(Index - 1).y = (Punct(Index).Top + Punct(Index).Height / 2) / Screen.TwipsPerPixelY
        Picture1.Cls
        Picture1.PSet (Points(0).x * Screen.TwipsPerPixelX, Points(0).y * Screen.TwipsPerPixelX)
        For i = 1 To NumPoints
                Picture1.Line -(Points(i).x * Screen.TwipsPerPixelX, Points(i).y * Screen.TwipsPerPixelY)
        Next i
    Else
    MousePointer = 99
    End If
End Sub

Private Sub Punct_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveItP = False
End Sub

Private Sub Timer1_Timer()
On Error GoTo xxx
Select Case Interfata
Case 0
PolyRegion = CreatePolygonRgn(Points(0), NumPoints + 1, Method)
'Set the window region of our form
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.BackColor = Picture2.BackColor
frmTest.Show
Interfata = 1

Case 1
PolyRegion = CreatePolygonRgn(Points1(0), NumPoints1 + 1, Method)
'Set the window region of our form
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.BackColor = Picture2.BackColor
frmTest.Show
Interfata = 2

Case 2
PolyRegion = CreatePolygonRgn(Points2(0), NumPoints2 + 1, Method)
'Set the window region of our form
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.BackColor = Picture2.BackColor
frmTest.Show
Interfata = 3

Case 3
PolyRegion = CreatePolygonRgn(Points3(0), NumPoints3 + 1, Method)
'Set the window region of our form
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.BackColor = Picture2.BackColor
frmTest.Show
Interfata = 4

Case 4
PolyRegion = CreatePolygonRgn(Points4(0), NumPoints4 + 1, Method)
'Set the window region of our form
ReturnVal = SetWindowRgn(frmTest.hWnd, PolyRegion, True)
frmTest.BackColor = Picture2.BackColor
frmTest.Show
Interfata = 0
End Select
Exit Sub
xxx:
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If t < 15 Then
t = t + 1
Else
t = 8
End If
'MsgBox t
Fade (t)
End Sub
