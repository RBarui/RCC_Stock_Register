VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000014&
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   LinkTopic       =   "Form3"
   ScaleHeight     =   8655
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
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
      Left            =   8760
      TabIndex        =   46
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Data MASTER 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Stock_Register\MASTER.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MASTER"
      Top             =   7080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text25 
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
      Left            =   12840
      TabIndex        =   41
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text24 
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
      Left            =   12840
      TabIndex        =   40
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text23 
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
      Left            =   12840
      TabIndex        =   39
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text22 
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
      Left            =   4800
      TabIndex        =   38
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text21 
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
      Left            =   12840
      TabIndex        =   37
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text20 
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
      Left            =   12840
      TabIndex        =   36
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text19 
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
      Left            =   4800
      TabIndex        =   35
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text18 
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
      Left            =   4800
      TabIndex        =   34
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text17 
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
      Left            =   4800
      TabIndex        =   33
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text16 
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
      Left            =   4800
      TabIndex        =   32
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text15 
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
      Left            =   4800
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text14 
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
      Left            =   4800
      TabIndex        =   30
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RESET"
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
      Left            =   12480
      TabIndex        =   29
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MAIN MENU"
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
      Left            =   11400
      TabIndex        =   28
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text13 
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
      Left            =   11400
      TabIndex        =   24
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11400
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text11 
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
      Left            =   11400
      TabIndex        =   22
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text10 
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
      Left            =   3360
      TabIndex        =   16
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text9 
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
      Left            =   11400
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text8 
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
      Left            =   11400
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text7 
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
      Left            =   3360
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text6 
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
      Left            =   3360
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text5 
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
      Left            =   3360
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text4 
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
      Left            =   3360
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
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
      Left            =   3360
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
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
      Left            =   10560
      TabIndex        =   5
      Top             =   840
      Width           =   1455
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
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000014&
      Caption         =   "FAULTY"
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
      Left            =   12960
      TabIndex        =   45
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000014&
      Caption         =   "TOTAL"
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
      Left            =   11640
      TabIndex        =   44
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000014&
      Caption         =   "FAULTY"
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
      Left            =   4920
      TabIndex        =   43
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000014&
      Caption         =   "TOTAL"
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
      Left            =   3600
      TabIndex        =   42
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000014&
      Caption         =   "PROJECTOR"
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
      Left            =   9960
      TabIndex        =   27
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000014&
      Caption         =   "SCANNER"
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
      Left            =   10200
      TabIndex        =   26
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000014&
      Caption         =   "PRINTER"
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
      Left            =   10320
      TabIndex        =   25
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000014&
      Caption         =   "ALMIRAH"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000014&
      Caption         =   "COMPUTER"
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
      Left            =   10080
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000014&
      Caption         =   "INTERNET LINK PORT"
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
      Left            =   9240
      TabIndex        =   19
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "BOARD"
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
      Left            =   2400
      TabIndex        =   18
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "STUDENT DESK"
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
      Left            =   1680
      TabIndex        =   17
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "TABLE"
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
      Left            =   2400
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "CHAIR"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "FAN"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "TUBE"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "ENTER THE ROOM NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "ROOM STOCK"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim flag As Integer
flag = 0
MASTER.Recordset.MoveFirst
While Not MASTER.Recordset.EOF
    If Trim(Text1.Text) = Trim(MASTER.Recordset!room) Then
    MsgBox ("Yes! Record Available")
    Store
    flag = 1
    GoTo fin
    End If
   MASTER.Recordset.MoveNext
    Wend
If flag = 0 Then
MsgBox "Sorry! Record Not Found!"
cltxt
End If
fin:
End Sub
Private Sub Store()
Text2.Text = MASTER.Recordset.Fields(1)
Text14.Text = MASTER.Recordset.Fields(2)
Text3.Text = MASTER.Recordset.Fields(3)
Text15.Text = MASTER.Recordset.Fields(4)
Text4.Text = MASTER.Recordset.Fields(5)
Text16.Text = MASTER.Recordset.Fields(6)
Text5.Text = MASTER.Recordset.Fields(7)
Text17.Text = MASTER.Recordset.Fields(8)
Text6.Text = MASTER.Recordset.Fields(9)
Text18.Text = MASTER.Recordset.Fields(10)
Text7.Text = MASTER.Recordset.Fields(11)
Text19.Text = MASTER.Recordset.Fields(12)
Text10.Text = MASTER.Recordset.Fields(13)
Text22.Text = MASTER.Recordset.Fields(14)
Text9.Text = MASTER.Recordset.Fields(15)
Text21.Text = MASTER.Recordset.Fields(16)
Text8.Text = MASTER.Recordset.Fields(17)
Text20.Text = MASTER.Recordset.Fields(18)
Text11.Text = MASTER.Recordset.Fields(19)
Text23.Text = MASTER.Recordset.Fields(20)
Text12.Text = MASTER.Recordset.Fields(21)
Text24.Text = MASTER.Recordset.Fields(22)
Text13.Text = MASTER.Recordset.Fields(23)
Text25.Text = MASTER.Recordset.Fields(24)

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
Text13.Text = " "
Text14.Text = " "
Text15.Text = " "
Text16.Text = " "
Text17.Text = " "
Text18.Text = " "
Text19.Text = " "
Text20.Text = " "
Text21.Text = " "
Text22.Text = " "
Text23.Text = " "
Text24.Text = " "
Text25.Text = " "

End Sub

Private Sub Command2_Click()
MASTER.Recordset.Edit
MASTER.Recordset.Fields(1) = Text2.Text
MASTER.Recordset.Fields(2) = Text14.Text
MASTER.Recordset.Fields(3) = Text3.Text
MASTER.Recordset.Fields(4) = Text15.Text
MASTER.Recordset.Fields(5) = Text4.Text
MASTER.Recordset.Fields(6) = Text16.Text
MASTER.Recordset.Fields(7) = Text5.Text
MASTER.Recordset.Fields(8) = Text17.Text
MASTER.Recordset.Fields(9) = Text6.Text
MASTER.Recordset.Fields(10) = Text18.Text
MASTER.Recordset.Fields(11) = Text7.Text
MASTER.Recordset.Fields(12) = Text19.Text
MASTER.Recordset.Fields(13) = Text10.Text
MASTER.Recordset.Fields(14) = Text22.Text
MASTER.Recordset.Fields(15) = Text9.Text
MASTER.Recordset.Fields(16) = Text21.Text
MASTER.Recordset.Fields(17) = Text8.Text
MASTER.Recordset.Fields(18) = Text20.Text
MASTER.Recordset.Fields(19) = Text11.Text
MASTER.Recordset.Fields(20) = Text23.Text
MASTER.Recordset.Fields(21) = Text12.Text
MASTER.Recordset.Fields(22) = Text24.Text
MASTER.Recordset.Fields(23) = Text13.Text
MASTER.Recordset.Fields(24) = Text25.Text
 MASTER.Recordset.Update
    MsgBox "Data Succesfully updated"
    MASTER.Refresh
End Sub

Private Sub Command3_Click()
Form3.Hide
Form2.Show

End Sub

Private Sub Command4_Click()
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
Text15.Text = " "
Text16.Text = " "
Text17.Text = " "
Text18.Text = " "
Text19.Text = " "
Text20.Text = " "
Text21.Text = " "
Text22.Text = " "
Text23.Text = " "
Text24.Text = " "
Text25.Text = " "
End Sub
