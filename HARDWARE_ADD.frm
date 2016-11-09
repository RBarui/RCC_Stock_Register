VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H8000000E&
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   LinkTopic       =   "Form13"
   ScaleHeight     =   8175
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   10200
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD ANOTHER"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   1
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      TabIndex        =   0
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Stock_Register\HARD.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HARD"
      Top             =   7680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "ADD A NEW HARDWARE"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   24
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "ROOM NO"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   23
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "ASSET NO"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   22
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   21
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "MFG"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "DETAILS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   19
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "DATE OF INSTALL"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   17
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000014&
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   16
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000014&
      Caption         =   "#Insert ' NA'  in the fields, you don't have any data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   6000
      Width           =   5055
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000014&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000014&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3120
      Width           =   375
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = Text5.Text
Data1.Recordset.Fields(5) = Text6.Text
Data1.Recordset.Fields(6) = Text7.Text
Data1.Recordset.Fields(7) = Text8.Text
Data1.Recordset.Update
    MsgBox "Data Succesfully Added"
End Sub

Private Sub Command2_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "

End Sub

Private Sub Command3_Click()
Form12.Show
Form13.Hide
End Sub

Private Sub Command4_Click()
Form13.Hide
Form2.Show

End Sub

