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

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal NoSecurity As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

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
    Const GENERIC_READ  As Long = &H80000000
    Const FILE_SHARE_READ As Long = &H1
    Const FILE_SHARE_WRITE As Long = &H2
    Const OPEN_EXISTING As Long = &H3
    Const INVALID_HANDLE_VALUE As Long = -1
    Const FILE_BEGIN    As Long = 0
    Const FILE_END      As Long = 2
    Dim hFile           As Long
    Dim baBuffer()      As Byte
    
    hFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    ReDim baBuffer(0 To SetFilePointer(hFile, 0, 0, FILE_END) - 1) As Byte
    Call SetFilePointer(hFile, 0, 0, FILE_BEGIN)
    Call ReadFile(hFile, baBuffer(0), UBound(baBuffer) + 1, 0, 0)
    Call CloseHandle(hFile)
    ReadBinaryFile = baBuffer
End Function

