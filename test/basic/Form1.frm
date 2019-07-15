VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4872
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6324
   LinkTopic       =   "Form1"
   ScaleHeight     =   4872
   ScaleWidth      =   6324
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2772
      Top             =   756
   End
   Begin Project1.AlphaBlendImage AlphaBlendImage2 
      Height          =   1440
      Left            =   4452
      Top             =   2856
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   2540
      Picture         =   "Form1.frx":0000
   End
   Begin Project1.AlphaBlendImage AlphaBlendImage1 
      Height          =   768
      Left            =   2184
      Top             =   2016
      Width           =   768
      _ExtentX        =   1355
      _ExtentY        =   1355
      Opacity         =   0.5
      Zoom            =   2
      Picture         =   "Form1.frx":0018
   End
   Begin VB.Image Image1 
      Height          =   1440
      Left            =   168
      Top             =   84
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error GoTo EH
    Set AlphaBlendImage1.Picture = AlphaBlendImage1.GdipLoadPictureArray(ReadBinaryFile(App.Path & "\bbb.png"))
    AlphaBlendImage1.Tag = -120
    Set Image1.Picture = AlphaBlendImage1.GdipLoadPicture(App.Path & "\bbb.png")
    Image1.Tag = 80
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub AlphaBlendImage1_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Timer1_Timer()
    AlphaBlendImage1.Rotation = AlphaBlendImage1.Rotation + 13
    AlphaBlendImage1.Left = AlphaBlendImage1.Left + AlphaBlendImage1.Tag
    If AlphaBlendImage1.Left + AlphaBlendImage1.Width > ScaleWidth Then
        AlphaBlendImage1.Tag = -Abs(AlphaBlendImage1.Tag)
    ElseIf AlphaBlendImage1.Left < 0 Then
        AlphaBlendImage1.Tag = Abs(AlphaBlendImage1.Tag)
    End If
    Caption = AlphaBlendImage1.Rotation
    Image1.Left = Image1.Left + Image1.Tag
    If Image1.Left + Image1.Width > ScaleWidth And Image1.Left > 0 Then
        Image1.Tag = -Abs(Image1.Tag)
    ElseIf Image1.Left < 0 And Image1.Left + Image1.Width < ScaleWidth Then
        Image1.Tag = Abs(Image1.Tag)
    End If
End Sub

Private Function ReadBinaryFile(sFile As String) As Byte()
    Dim nFile           As Integer
    Dim baBuffer()      As Byte
    
    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    ReadBinaryFile = baBuffer
End Function
