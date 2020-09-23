VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Create Account"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Account Info"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Create It!"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Pass:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "User:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "You must enter a username."
End If
If Not FileExists(App.Path & "\users\" & Text1.Text & ".txt") Then
    On Error Resume Next
        Open App.Path & "\users\" & Text1.Text & ".txt" For Output As #1
        Print #1, encrypt(Text2.Text)
      Close #1
      MsgBox "Account created!", vbInformation, "YAY!"
      Else
      MsgBox "Username already taken!", vbCritical, "ERROR"
      End If
End Sub
