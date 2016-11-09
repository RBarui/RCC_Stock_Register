VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000014&
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13650
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "ABOUT US"
      Height          =   1335
      Left            =   7080
      TabIndex        =   6
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LAB EQUIPMENTS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "ROOM INVENTORY"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7080
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MONITORS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COMPUTERS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "MASTER REPORT"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "DEPARTMENTAL STOCK REGISTER"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   9855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form3.Show

End Sub

Private Sub Command2_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub Command3_Click()
Form2.Hide
Form8.Show
End Sub

Private Sub Command4_Click()
Form12.Show
Form2.Hide

End Sub

Private Sub Command5_Click()
Form2.Hide
Form16.Show
End Sub

Private Sub Command6_Click()
Close
End Sub

Private Sub Command7_Click()
Form2.Hide
Form25.Show
End Sub
