VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9030
   ClientLeft      =   810
   ClientTop       =   405
   ClientWidth     =   11070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00BF9151&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   11070
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   -22
      Top             =   -21
      Width           =   3315
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Double, OldY As Double
Dim MoveIt As Boolean


Private Sub ActiveMovie1_Click()
Beep
End Sub

Private Sub ActiveMovie1_DblClick()
Beep
End Sub

Private Sub ActiveMovie1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    Else
        Unload frmTest
    End If
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    Else
        Unload frmTest
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveIt Then
        
        frmTest.Left = frmTest.Left + X - OldX
        frmTest.Top = frmTest.Top + Y - OldY
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    Else
        Unload frmTest
    End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveIt Then
        frmTest.Left = frmTest.Left + X - OldX
        frmTest.Top = frmTest.Top + Y - OldY
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
    frmTest.Visible = False

    frmTest.Visible = True
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    Else
        Unload frmTest
    End If
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveIt Then
        
        frmTest.Left = frmTest.Left + X - OldX
        frmTest.Top = frmTest.Top + Y - OldY
    End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
End Sub
