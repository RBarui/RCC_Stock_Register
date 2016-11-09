VERSION 5.00
Begin VB.Form Form25 
   BackColor       =   &H8000000E&
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835
   LinkTopic       =   "Form25"
   ScaleHeight     =   6000
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "HOME"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "PREPARED BY"
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
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "UNDER SUPERVISION OF"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "RAJIB SAHA (COURSE CO-ORDINATOR)  9830373904 rajibsaha4u@yahoo.com"
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
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   7575
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "SANCHARI DAS   8906580194    sancharid67@gmail.com"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   3720
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "ANKESH PRINCE   7044337498   princesinha141@gmail.com"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   4200
      Width           =   6375
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "    ARNAB MONDAL 7407613296   arnabmondal482@gmail.com"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "   RAJARSHI BARUI     9046490662     rajarshibarui94@gmail.com"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form25.Hide
Form2.Show
End Sub
