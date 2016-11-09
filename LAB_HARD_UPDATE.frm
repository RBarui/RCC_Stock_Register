VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H80000014&
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   LinkTopic       =   "Form20"
   ScaleHeight     =   8820
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   10200
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "UPDATE"
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
      TabIndex        =   3
      Top             =   6840
      Width           =   2055
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Stock_Register\LAB_HARD.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LAB_HARD"
      Top             =   8040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "UPDATE  A  HARDWARE"
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
      TabIndex        =   21
      Top             =   240
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
      TabIndex        =   20
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "ENTER ASSET NO OF THE HARDWARE YOU WANT TO UPDATE"
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
      Left            =   2520
      TabIndex        =   19
      Top             =   1440
      Width           =   6495
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
      TabIndex        =   18
      Top             =   3480
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
      TabIndex        =   17
      Top             =   4200
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
      TabIndex        =   16
      Top             =   4920
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
      TabIndex        =   15
      Top             =   2880
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
      TabIndex        =   14
      Top             =   3720
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
      TabIndex        =   13
      Top             =   4560
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
      TabIndex        =   12
      Top             =   5880
      Width           =   5055
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.Edit
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = Text5.Text
Data1.Recordset.Fields(5) = Text6.Text
Data1.Recordset.Fields(6) = Text7.Text
Data1.Recordset.Fields(7) = Text8.Text
 Data1.Recordset.Update
    MsgBox "Data Succesfully updated"
    Data1.Refresh
   cltxt
End Sub

Private Sub Command2_Click()
Dim flag As Integer
flag = 0
Data1.Recordset.MoveFirst
While Not Data1.Recordset.EOF
    If Trim(Text2.Text) = Trim(Data1.Recordset!ASSET_NO) Then
    MsgBox ("SUCCESS! RECORD FOUND")
    Store
    flag = 1
    GoTo fin
    End If
    Data1.Recordset.MoveNext
    Wend
If flag = 0 Then
MsgBox "Sorry! Record Not Found!"
cltxt
End If
fin:
End Sub
Private Sub cltxt()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "

End Sub
Private Sub Store()
Text1.Text = Data1.Recordset.Fields(0)
Text3.Text = Data1.Recordset.Fields(2)
Text4.Text = Data1.Recordset.Fields(3)
Text5.Text = Data1.Recordset.Fields(4)
Text6.Text = Data1.Recordset.Fields(5)
Text7.Text = Data1.Recordset.Fields(6)
Text8.Text = Data1.Recordset.Fields(7)

End Sub

Private Sub Command3_Click()
Form20.Hide
Form17.Show
End Sub

Private Sub Command4_Click()
Form20.Hide
Form2.Show
End Sub

