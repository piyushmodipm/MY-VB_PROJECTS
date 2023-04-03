VERSION 5.00
Begin VB.Form frmaboutus 
   Caption         =   "About VMS"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmaboutus.frx":0000
   ScaleHeight     =   9585
   ScaleWidth      =   15975
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmaboutus.frx":36F68
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   6975
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   14175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VOTING MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   15975
   End
End
Attribute VB_Name = "frmaboutus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
