VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17205
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   17205
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7440
      Top             =   4680
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "BookMaster"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2160
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Car"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   16560
      MaskColor       =   &H00FF8080&
      Picture         =   "Form1.frx":31BB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7680
      Picture         =   "Form1.frx":6376
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12360
      Picture         =   "Form1.frx":95CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1680
      Top             =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Travel And Tour"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1095
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Width           =   9135
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   2280
      Picture         =   "Form1.frx":C76C
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   16560
      Picture         =   "Form1.frx":F439
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image Image5 
      Height          =   735
      Index           =   0
      Left            =   7200
      Picture         =   "Form1.frx":205AE
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   10920
      Picture         =   "Form1.frx":23ADE
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Label1.Caption = "Bus Forum"
Form2.Label6.Caption = "Bus Name"
Form2.Show

End Sub

Private Sub Command2_Click()
Form2.Label1.Caption = "TRAIN Forum"
Form2.Label6.Caption = "TRAIN Name"
Form2.Show

End Sub

Private Sub Command3_Click()
End

End Sub

Private Sub Command4_Click()
Form3.Show

End Sub

Private Sub e_Click()
End

End Sub


Private Sub Timer1_Timer()
Label2.Caption = Date
Label3.Caption = Time

End Sub

