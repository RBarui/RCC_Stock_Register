VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H80000014&
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   LinkTopic       =   "Form22"
   ScaleHeight     =   9210
   ScaleWidth      =   15300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   10440
      TabIndex        =   30
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   10440
      TabIndex        =   29
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   10440
      TabIndex        =   28
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   10440
      TabIndex        =   27
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   10440
      TabIndex        =   26
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   10440
      TabIndex        =   25
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   1680
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
      Top             =   7440
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
      Left            =   4080
      TabIndex        =   2
      Top             =   7440
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
      Left            =   9360
      TabIndex        =   1
      Top             =   7440
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
      Left            =   11520
      TabIndex        =   0
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Stock_Register\LAB_PC.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LAB_PC"
      Top             =   8760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label19 
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
      Left            =   9480
      TabIndex        =   36
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label18 
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
      Left            =   8520
      TabIndex        =   35
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000014&
      Caption         =   "MOTHERBOARD"
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
      Left            =   8640
      TabIndex        =   34
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000014&
      Caption         =   "HDD"
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
      Left            =   9840
      TabIndex        =   33
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000014&
      Caption         =   "RAM"
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
      Left            =   9840
      TabIndex        =   32
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000014&
      Caption         =   "PROCESSOR"
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
      TabIndex        =   31
      Top             =   2520
      Width           =   2295
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
      Left            =   2760
      TabIndex        =   24
      Top             =   360
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
      Left            =   2040
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
      Left            =   2040
      TabIndex        =   22
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "SERVICE TAG"
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
      Left            =   1680
      TabIndex        =   21
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "MAC ID"
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
      TabIndex        =   20
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "IP ADRESS"
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
      TabIndex        =   19
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "SYS MODEL"
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
      Left            =   1800
      TabIndex        =   18
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "TYPE"
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
      TabIndex        =   17
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label9 
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
      Left            =   9720
      TabIndex        =   16
      Top             =   1800
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
      Left            =   8760
      TabIndex        =   15
      Top             =   6840
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
      Left            =   6120
      TabIndex        =   14
      Top             =   1800
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
      Left            =   6120
      TabIndex        =   13
      Top             =   2520
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
      Left            =   6120
      TabIndex        =   12
      Top             =   3240
      Width           =   375
   End
End
Attribute VB_Name = "Form22"
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
Data1.Recordset.Fields(8) = Text9.Text
Data1.Recordset.Fields(9) = Text10.Text
Data1.Recordset.Fields(10) = Text11.Text
Data1.Recordset.Fields(11) = Text12.Text
Data1.Recordset.Fields(12) = Text13.Text
Data1.Recordset.Fields(13) = Text14.Text
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
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
Text13.Text = " "
Text14.Text = " "
End Sub

Private Sub Command3_Click()
Form22.Hide
Form21.Show

End Sub

Private Sub Command4_Click()
Form22.Hide
Form2.Show
End Sub
