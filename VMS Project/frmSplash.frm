VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6015
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6090
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11625
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   5280
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   9840
         Top             =   4200
      End
      Begin VB.Image Image1 
         Height          =   6120
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version:6.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   285
         Left            =   10080
         TabIndex        =   1
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform: Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   360
         Left            =   6000
         TabIndex        =   2
         Top             =   2040
         Width           =   2850
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voting Management System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   480
         Left            =   6000
         TabIndex        =   3
         Top             =   1440
         Width           =   5505
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   
End Sub

Private Sub Frame1_Click(index As Integer)
    Unload Me
End Sub

Private Sub lblCopyright_Click()

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 2
If ProgressBar1.Value >= 98 Then
    ProgressBar1.Value = 100
    Timer1.Enabled = False
    frmLogin.Show
End If
Label1.Caption = ProgressBar1.Value & "%"
End Sub
