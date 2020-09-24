VERSION 5.00
Begin VB.Form FRM_Titlebar 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Marquee"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu OptMenu 
      Caption         =   "OptMenu"
      Visible         =   0   'False
      Begin VB.Menu MNU_Back 
         Caption         =   "Never show Background"
      End
      Begin VB.Menu mnu_s1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FRM_Titlebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim O_x
Dim O_Y
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Sub Form_LostFocus()
'If the user lost focus, enable timer so the background can dissapear
frmMain.Timer2.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Reset the timer if there is any movement
frmMain.Timer2.Enabled = False
frmMain.Timer2.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'Rightmouseclick
  PopupMenu OptMenu 'Show the menu
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMain
End
End Sub

Private Sub MNU_Back_Click()
MNU_Back.Checked = Not MNU_Back.Checked 'Never show background (other form)
If MNU_Back.Checked = True Then 'Change cursor into size all
  frmMain.MousePointer = 15
  Me.Hide
Else
  frmMain.MousePointer = 0 'Reset cursor
  
End If

End Sub

Private Sub Mnu_Exit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
dirty = 0 'Timer that keeps track of FRM_Titlebar
If O_x <> Me.Left Then 'Move the marquee at the center of the titlebar
  O_x = Me.Left
  frmMain.Left = Me.Left
  dirty = 1
End If
If O_Y <> Me.Top Then
  O_Y = Me.Top
  frmMain.Top = Me.Top + 300
  dirty = 1
End If
If dirty = 1 Then 'If the form was moved
  dirty = 0
  frmMain.Timer2.Enabled = False 'Reset the 'disapear'-timer
  
End If
frmMain.Timer2.Enabled = True
End Sub
