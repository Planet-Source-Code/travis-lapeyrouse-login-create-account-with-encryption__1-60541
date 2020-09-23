VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Please login!"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2880
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Account"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Create Account"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Pass:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ListBox Users 
      Height          =   450
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this encryption I found on psc somewhere
'vote for me

Private Sub Command1_Click()
If FileExists(App.Path & "\users\" & Text1.Text & ".txt") Then
Users.Clear
Open (App.Path & "\users\" & Text1.Text & ".txt") For Input As #1
   While Not EOF(1)
      Input #1, Doit$
        DoEvents
        Text3.Text = Doit$
        Wend
        Close #1
If Text2.Text = decrypt(Text3.Text) Then
Form3.Show
Else
MsgBox "Incorrect Password!", vbCritical, "Error"
End If
Else
MsgBox "Username doesnt exist!", vbCritical, "Error"
End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub
