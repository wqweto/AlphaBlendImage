VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1764
      Top             =   840
   End
   Begin Project1.AlphaBlendImage AlphaBlendImage1 
      Height          =   768
      Left            =   0
      Top             =   0
      Width           =   768
      _ExtentX        =   1355
      _ExtentY        =   1355
      AutoRedraw      =   -1  'True
      Rotation        =   60
      Zoom            =   2
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lX                As Long
Private m_lY                As Long
Private m_sngDelta          As Single

Private Sub Form_Load()
    Set AlphaBlendImage1.Picture = AlphaBlendImage1.GdipLoadPicture(App.Path & "\garden.png")
    Width = AlphaBlendImage1.Width
    Height = AlphaBlendImage1.Height
    AlphaBlendImage1.GdipUpdateLayeredWindow hWnd
    m_sngDelta = 0.01
End Sub

Private Sub AlphaBlendImage1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_lX = X
        m_lY = Y
    End If
End Sub

Private Sub AlphaBlendImage1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) <> 0 Then
        Move Left + X - m_lX, Top + Y - m_lY
    End If
End Sub

Private Sub AlphaBlendImage1_DblClick()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    AlphaBlendImage1.Opacity = AlphaBlendImage1.Opacity + m_sngDelta
    If AlphaBlendImage1.Opacity <= 0 Or AlphaBlendImage1.Opacity >= 1 Then
        m_sngDelta = -m_sngDelta
    End If
End Sub
