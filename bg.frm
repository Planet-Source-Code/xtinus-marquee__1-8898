VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "bg.frx":0000
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   1560
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "commandline=text ($$T=time,$$D=date)"
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.PictureBox P_Alpha 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1800
      Picture         =   "bg.frx":B112
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   7335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code was written using Chris Yates' autoshape form.

Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim CurRgn, TempRgn As Long  ' Region variables



Private Sub Form_Load()
If Len(Command) > 0 Then Text1.Text = Command
AutoFormShape frmMain, RGB(0, 0, 0)  ' Shape the form so that all areas that are bright purple become transparent.

HScroll1.Max = Me.ScaleWidth / 3 + Len(Text1.Text) * 6

Flag% = SWP_NOMOVE Or SWP_NOSIZE
lSetPos = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FRM_Titlebar.MNU_Back.Checked = False Then Exit Sub
ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)  ' This tells the OS to pick up the form to be moved
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FRM_Titlebar.MNU_Back.Checked = True Then Exit Sub
Timer2.Enabled = False
FRM_Titlebar.Width = Me.Width
FRM_Titlebar.Top = Me.Top - 300
FRM_Titlebar.Height = Me.Height + 320
FRM_Titlebar.Left = Me.Left
FRM_Titlebar.Show


Timer2.Enabled = True

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
  PopupMenu FRM_Titlebar.OptMenu
End If

End Sub

Private Sub Text1_Change()
HScroll1.Max = Me.ScaleWidth / 3 + Len(Text1.Text) * 6
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If HScroll1.Value = HScroll1.Max Then HScroll1.Value = HScroll1.Min
nt = ""
For X = 1 To Len(Text1.Text)
  char_index = Asc(UCase(Mid(Text1.Text, X, 1)))

  If char_index = Asc("$") And X < Len(Text1.Text) Then
    Select Case UCase(Mid(Text1.Text, X + 1, 1))
    Case "T"
    nt = nt & Format(Now, "hh:nn:ss")
    Case "D"
    nt = nt & Format(Now, "dd-mm-yy")
    Case "$"
    nt = nt & "$"
    
    End Select
    X = X + 1
  Else
    nt = nt & Chr(char_index)
  End If
Next
HScroll1.Max = Me.ScaleWidth / 3 + Len(nt) * 6
For X = 1 To Len(nt)
  chara = Asc(UCase(Mid(nt, X, 1)))
  Select Case chara
  Case Is < 32
  char_index = 0
  Case 32 To 93
  char_index = Asc(UCase(Mid(nt, X, 1))) - 32

  Case Else
  char_index = 0
  End Select
res_offs_x = (X - 1) * 18
alp_offs_x = char_index * 18 + 0
s_pos = Me.ScaleWidth - HScroll1.Value * 3
Me.PaintPicture P_Alpha, res_offs_x + s_pos, 0, 18, 24, alp_offs_x, 0, 18, 24
Next X
 

HScroll1.Value = HScroll1.Value + 1
End If
End Sub

Private Sub Timer1_Timer()
speed = 3

Text1.Text = RTrim(Text1.Text) & Space(CInt(speed / 4 + 0.5))

If HScroll1.Value > HScroll1.Max - speed Then HScroll1.Value = HScroll1.Min
nt = ""
For X = 1 To Len(Text1.Text)
  char_index = Asc(UCase(Mid(Text1.Text, X, 1)))

  If char_index = Asc("$") And X < Len(Text1.Text) Then
    Select Case UCase(Mid(Text1.Text, X + 1, 1))
    Case "T"
    nt = nt & Format(Now, "hh:nn:ss")
    Case "D"
    nt = nt & Format(Now, "dd-mm-yy")
    Case "$"
    nt = nt & "$"
    
    End Select
    X = X + 1
  Else
    nt = nt & Chr(char_index)
  End If
Next

HScroll1.Max = Me.ScaleWidth / 3 + Len(nt) * 6

For X = 1 To Len(nt)
  chara = Asc(UCase(Mid(nt, X, 1)))
  Select Case chara
  Case Is < 32
  char_index = 0
  Case 32 To 93
  char_index = Asc(UCase(Mid(nt, X, 1))) - 32
  'Format(Time, "HH:MM:SS")
  Case Else
  char_index = 0
  End Select
res_offs_x = (X - 1) * 18
alp_offs_x = char_index * 18 + 0
s_pos = Me.ScaleWidth - HScroll1.Value * 3
Me.PaintPicture P_Alpha, res_offs_x + s_pos, 0, 18, 24, alp_offs_x, 0, 18, 24
Next X

If HScroll1.Max - HScroll1.Value < speed Then
  HScroll1.Value = HScroll1.Max

Else
  HScroll1.Value = HScroll1.Value + speed
End If



End Sub

Private Sub Timer2_Timer()
FRM_Titlebar.Hide
Timer2.Enabled = False

End Sub

