VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selection Page"
   ClientHeight    =   9630
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   19245
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   19245
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1680
      Top             =   600
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
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   1695
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
      Picture         =   "Form4.frx":31A2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   2055
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
      Picture         =   "Form4.frx":63F6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2295
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
      Picture         =   "Form4.frx":95B1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   2895
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
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   10920
      Picture         =   "Form4.frx":C76C
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
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
   Begin VB.Image Image5 
      Height          =   735
      Left            =   7200
      Picture         =   "Form4.frx":E376
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   16560
      Picture         =   "Form4.frx":118A6
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   1560
      Picture         =   "Form4.frx":22A1B
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2415
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
      TabIndex        =   0
      Top             =   480
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   11400
      Left            =   0
      Picture         =   "Form4.frx":256E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20520
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Label1.Caption = "Bus Forum"
Form5.Label6.Caption = "Bus Name"
Form5.Show

End Sub

Private Sub Command2_Click()
Form5.Label1.Caption = "TRAIN Forum"
Form5.Label6.Caption = "TRAIN Name"
Form5.Show

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
