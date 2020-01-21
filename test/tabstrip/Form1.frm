VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6948
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6948
   StartUpPosition =   3  'Windows Default
   Begin Project1.AlphaBlendTabStrip AlphaBlendTabStrip1 
      Height          =   432
      Left            =   504
      Top             =   2436
      Width           =   5136
      _ExtentX        =   9059
      _ExtentY        =   762
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   "Printers|Configuration|Logs"
   End
   Begin Project1.AlphaBlendLabel AlphaBlendLabel1 
      Height          =   1608
      Left            =   336
      Top             =   252
      Width           =   4464
      _ExtentX        =   7874
      _ExtentY        =   2836
      Caption         =   "This is a test of the wrap mode This is a test of the wrap mode"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PT Sans Narrow"
         Size            =   19.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483627
      ShadowColor     =   -2147483630
      ShadowOpacity   =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AlphaBlendTabStrip1_BeforeClick(TabIndex As Long, Cancel As Boolean)
    If TabIndex = AlphaBlendTabStrip1.CurrentTab Then
        Cancel = True
    ElseIf AlphaBlendTabStrip1.CurrentTab = 1 Then
        Select Case MsgBox("Do you want to save configuration?", vbQuestion Or vbYesNoCancel)
        Case vbYes
            AlphaBlendTabStrip1.TabCaption(1) = Replace(AlphaBlendTabStrip1.TabCaption(1), "*", vbNullString)
        Case vbCancel
            Cancel = True
        End Select
    End If
End Sub

Private Sub AlphaBlendTabStrip1_Click()
    Debug.Print "AlphaBlendTabStrip1.CurrentTab=" & AlphaBlendTabStrip1.CurrentTab
    If AlphaBlendTabStrip1.CurrentTab = 1 And Right$(AlphaBlendTabStrip1.TabCaption(1), 1) <> "*" Then
        AlphaBlendTabStrip1.TabCaption(1) = AlphaBlendTabStrip1.TabCaption(1) & "*"
    End If
End Sub

