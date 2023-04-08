VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmnewbook 
   Caption         =   "New Book Entry form"
   ClientHeight    =   8685
   ClientLeft      =   2025
   ClientTop       =   240
   ClientWidth     =   14190
   FillColor       =   &H0080C0FF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmnewbook.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   14190
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10680
      TabIndex        =   23
      Text            =   "2000"
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calculate"
      Height          =   735
      Left            =   8040
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox cmbbooktype 
      DataField       =   "type"
      DataSource      =   "Adodc1"
      Height          =   405
      ItemData        =   "frmnewbook.frx":4A92D
      Left            =   3600
      List            =   "frmnewbook.frx":4A93D
      TabIndex        =   4
      Top             =   3960
      Width           =   3015
   End
   Begin VB.ComboBox cmbsupplierno 
      DataField       =   "guideno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmnewbook.frx":4A984
      Left            =   10680
      List            =   "frmnewbook.frx":4A986
      TabIndex        =   6
      Top             =   1440
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1080
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      TabIndex        =   12
      Top             =   6360
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   11
      Top             =   6360
      Width           =   2235
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      TabIndex        =   10
      Top             =   6360
      Width           =   2235
   End
   Begin VB.TextBox txtpublication 
      DataField       =   "hotel"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10680
      TabIndex        =   7
      Top             =   2280
      Width           =   3000
   End
   Begin VB.TextBox txtedition 
      DataField       =   "tenure"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10680
      TabIndex        =   8
      Top             =   3120
      Width           =   3000
   End
   Begin VB.TextBox txtpurchasedate 
      DataField       =   "Purdate"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   3000
   End
   Begin VB.TextBox txtauthor 
      DataField       =   "goal"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   5
      Top             =   4800
      Width           =   3000
   End
   Begin VB.TextBox txtprice 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10560
      TabIndex        =   9
      Top             =   4920
      Width           =   3000
   End
   Begin VB.TextBox txttitle 
      DataField       =   "Location"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   3
      Top             =   3000
      Width           =   3000
   End
   Begin VB.TextBox txtbookno 
      DataField       =   "Bookno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   2
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Guide no."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   8040
      TabIndex        =   21
      Top             =   1440
      Width           =   2355
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tenure"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   3120
      Width           =   2115
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Hotel "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   2280
      Width           =   2115
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   18
      Top             =   1440
      Width           =   2115
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   17
      Top             =   4800
      Width           =   1995
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   16
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Package type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   3960
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   13
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "   New Booking"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmnewbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbbooktype_Change()
If cmbbooktype.Text = "Golden Package" Then
Text1.Text = "2000"
End If
If cmbbooktype.Text = "Silver Package" Then
Text1.Text = "1000"
End If
If cmbbooktype.Text = "Platinum Package" Then
Text1.Text = "3000"
End If
If cmbbooktype.Text = "Diamond Package" Then
Text1.Text = "5000"
End If
End Sub

Private Sub cmdAdd_Click()
Dim Bookno As Integer
Dim con As New Connection
Dim rs As New Recordset

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\my vb projects\LMS PROJECT\LMSdb.mdb;Persist Security Info=False"
rs.Open "select * from BookMaster", con, adOpenDynamic, adLockOptimistic


If rs.RecordCount = -1 Then
   Bookno = 1
Else
   Bookno = rs.RecordCount + 1
End If

rs.AddNew
rs.Fields(0).Value = txtbookno.Text
rs.Fields(1).Value = txttitle.Text
rs.Fields(2).Value = cmbbooktype.ListIndex
rs.Fields(3).Value = txtauthor.Text
rs.Fields(4).Value = txtprice.Text
rs.Fields(5).Value = txtpurchasedate.Text
rs.Fields(6).Value = txtpublication.Text
rs.Fields(7).Value = txtedition.Text
rs.Fields(8).Value = cmbsupplierno.ListIndex

rs.Update

MsgBox "successful"

End Sub

Private Sub cmdCancel_Click()
ClearData

End Sub

Private Sub cmdclose_Click()
Me.Hide

End Sub

Private Sub Command1_Click()
txtprice = Val(Form2.Text4.Text) + Val(Val(Text1.Text) * Val(txtedition.Text))
End Sub

Function ClearData()

txtpurchasedate.Text = Date
txttitle.Text = ""
txtauthor.Text = ""
txtprice.Text = ""
cmbsupplierno.Text = ""
txtpublication.Text = ""
txtedition.Text = ""
cmbbooktype.Text = ""
End Function





