VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H8000000E&
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15585
   LinkTopic       =   "Form7"
   ScaleHeight     =   8775
   ScaleWidth      =   15585
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Stock_Register\PC.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PC"
      Top             =   8160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BACK"
      Height          =   615
      Left            =   11640
      TabIndex        =   28
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MENU"
      Height          =   615
      Left            =   10080
      TabIndex        =   27
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      Height          =   615
      Left            =   8280
      TabIndex        =   26
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   12000
      TabIndex        =   25
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   24
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   23
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   22
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   21
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   20
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   19
      Top             =   7080
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   18
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   17
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   16
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   15
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   14
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   1800
      Width           =   3735
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
      Left            =   8520
      TabIndex        =   29
      Top             =   6960
      Width           =   5055
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   13
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "DATE OF REPAIR"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   11
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "PROBLEM"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "DATE OF INSTALL"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "SOFTWARE SPECS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "HARDWARE SPECS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "OS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   6
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "IP ADRESS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "SERVICE TAG"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "ROOM NO"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "ENTER THE ASSET NO OF THE PC YOU WANT TO UPDATE"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "UPDATE   COMPUTER"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   7335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim flag As Integer
flag = 0
Data1.Recordset.MoveFirst
While Not Data1.Recordset.EOF
    If Trim(Text1.Text) = Trim(Data1.Recordset!ASSET_NO) Then
    MsgBox ("SUCCESS! MACHINE FOUND")
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
Private Sub Store()
Text1.Text = Data1.Recordset.Fields(0)
Text3.Text = Data1.Recordset.Fields(2)
Text4.Text = Data1.Recordset.Fields(3)
Text5.Text = Data1.Recordset.Fields(4)
Text6.Text = Data1.Recordset.Fields(5)
Text7.Text = Data1.Recordset.Fields(6)
Text8.Text = Data1.Recordset.Fields(7)
Text9.Text = Data1.Recordset.Fields(8)
Text10.Text = Data1.Recordset.Fields(9)
Text11.Text = Data1.Recordset.Fields(10)
Text12.Text = Data1.Recordset.Fields(11)
End Sub

Private Sub Command2_Click()
Data1.Recordset.Edit
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
    Data1.Recordset.Update
    MsgBox "Data Succesfully updated"
    Data1.Refresh
   cltxt
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
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
End Sub

Private Sub Command3_Click()
Form7.Hide
Form2.Show
End Sub

Private Sub Command4_Click()
Form7.Hide
Form4.Show
End Sub
