VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Show Form"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Message"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 01-Feb-05 21:23
' Author    : Abhishek
' Purpose   : To Use .NET Classes in VB
' System    : Dev: Windows 2000+, User: Windows 98+
'---------------------------------------------------------------------------------------

Option Explicit

Private MyDot As Class1

Private Sub Command1_Click()
    MyDot.ShowMessage
End Sub

Private Sub Command2_Click()
    MyDot.ShowForm
End Sub

Private Sub Form_Load()
    Set MyDot = New Class1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MyDot = Nothing
End Sub
