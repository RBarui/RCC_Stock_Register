VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "ENTER YOUR USERNAME AND PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2880
      TabIndex        =   2
      Top             =   3840
      Width           =   8415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "DEPT OF COMPUTER SCIENCE"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   12135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "RCC INSTITUTE OF INFORMATION TECHNOLOGY"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   12015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "admin" And Text2.Text = "admin" Then
Form1.Hide
Form2.Show
Else
MsgBox " The Username And/Or Password Entered Is Incorrect "
End If
End Sub
