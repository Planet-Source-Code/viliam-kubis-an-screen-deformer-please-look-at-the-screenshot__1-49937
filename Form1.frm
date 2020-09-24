VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   11610
      Left            =   0
      ScaleHeight     =   11610
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   0
      Width           =   3780
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ULW_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_OPAQUE = &H4
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_WINDOWEDGE = &H100&
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Sub Form_Load()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Resize()
Picture1.Width = Form1.Width
End Sub

Private Sub Picture1_Click()
SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetWindowPos Me.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_FRAMECHANGED + SWP_NOZORDER
SetLayeredWindowAttributes Me.hwnd, &HFF&, 0, ULW_COLORKEY
End Sub

Private Sub Timer1_Timer()
Picture1.Top = Picture1.Top - 10
Picture1.Height = Picture1.Height + 10
If Picture1.Top + Picture1.Height < 0 Then Picture1.Top = 0
End Sub
