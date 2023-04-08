VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17910
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   17910
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   13680
      TabIndex        =   27
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "Passenger Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   14
      Text            =   " "
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Source"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   13
      Text            =   "Name"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Destination"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   12
      Text            =   "Name"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      DataField       =   "Date"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   11
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Bus Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   9
      Text            =   "Name"
      Top             =   4320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17880
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "Time"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   7
      Top             =   5160
      Width           =   2535
   End
   Begin VB.ComboBox Combo5 
      DataField       =   "Class"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   6
      Text            =   "Name"
      Top             =   6240
      Width           =   2535
   End
   Begin VB.ComboBox Combo6 
      DataField       =   "Ticket Type"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   5
      Text            =   "Name"
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      DataField       =   "Amount"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   4
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15480
      Picture         =   "Form2.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   17760
      Picture         =   "Form2.frx":3745
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Picture         =   "Form2.frx":6713
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      Picture         =   "Form2.frx":98B5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   1680
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2055
      Left            =   14040
      TabIndex        =   10
      Top             =   2640
      Width           =   3495
      _Version        =   524288
      _ExtentX        =   6165
      _ExtentY        =   3625
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2019
      Month           =   2
      Day             =   7
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Bus  Forum"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8400
      TabIndex        =   26
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Booking no."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6495
      TabIndex        =   25
      Top             =   720
      Width           =   2625
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destination State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6540
      TabIndex        =   24
      Top             =   1560
      Width           =   2520
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destination City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6510
      TabIndex        =   23
      Top             =   2400
      Width           =   2475
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   22
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bus Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6495
      TabIndex        =   21
      Top             =   4320
      Width           =   2505
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   20
      Top             =   5280
      Width           =   2475
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   19
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ticket Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   18
      Top             =   7440
      Width           =   2550
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   17
      Top             =   8400
      Width           =   2475
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   0
      Picture         =   "Form2.frx":C97A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label12 
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
      Left            =   960
      TabIndex        =   16
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   855
      Left            =   3240
      Picture         =   "Form2.frx":FEAA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label13 
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
      Left            =   4320
      TabIndex        =   15
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset


Private Sub Combo3_Click()
If Combo3.Text = "Babu Bus" Then
Text3.Text = "08:00 a.m"
ElseIf Combo3.Text = "Sree Ram Bus" Then
Text3.Text = "10:00 a.m"
ElseIf Combo3.Text = "Sita Ram Bus" Then
Text3.Text = "02:00 p.m"
ElseIf Combo3.Text = "Jai Shree Bus" Then
Text3.Text = "04:00 p.m"
ElseIf Combo3.Text = "Rajasthan Bus" Then
Text3.Text = "07:00 p.m"
ElseIf Combo3.Text = "Shreenath Bus" Then
Text3.Text = "10:00 p.m"
ElseIf Combo3.Text = "Sairam Bus" Then
Text3.Text = "12:00 p.m"
ElseIf Combo3.Text = "DC Bus" Then
Text3.Text = "06:00 a.m"
Else
End If
End Sub



Private Sub Combo5_Click()
If Combo5.Text = "Sleeper" Then
Text4.Text = "500 Rs"
ElseIf Combo5.Text = "A.C" Then
Text4.Text = "1000 Rs"
ElseIf Combo5.Text = "Non A.C" Then
Text4.Text = "2000 Rs"

Else
End If
End Sub

Private Sub Command1_Click()
Text2.Text = Calendar1.Value
End Sub

Private Sub Command2_Click()
frmnewbook.txtbookno.Text = Text1.Text
frmnewbook.txttitle.Text = Combo1.Text + "  " + Combo2.Text

frmnewbook.Show
 MsgBox ("record saved")


End Sub

Private Sub Command3_Click()
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text2.Text = ""
Combo3.Text = ""
Text3.Text = ""
Combo5.Text = ""
Combo6.Text = ""
Text4.Text = ""

End Sub

Private Sub Command4_Click()
If Combo1.Text = "Jammu & Kashmir" Then
Combo2.AddItem "1  Kupwara"
Combo2.AddItem "2  Badgam"
Combo2.AddItem "3  Leh (Ladakh)"
Combo2.AddItem "4  Kargil"
Combo2.AddItem "5  Punch"
Combo2.AddItem "6  Rajouri"
Combo2.AddItem "7  Kathua"
Combo2.AddItem "8  Baramula"
Combo2.AddItem "9  Bandipore"
Combo2.AddItem "10 Srinagar"
Combo2.AddItem "11 Ganderbal"
Combo2.AddItem "12 Pulwama"
Combo2.AddItem "13 Shupiyan"
Combo2.AddItem "14 Anantnag"
Combo2.AddItem "15 Kulgam"
Combo2.AddItem "16 Doda"
Combo2.AddItem "17 Ramban"
Combo2.AddItem "18 Kishtwar"
Combo2.AddItem "19 Udhampur"
Combo2.AddItem "20 Reasi"
Combo2.AddItem "21 Jammu"
Combo2.AddItem "22 Samba"
End If


If Combo1.Text = "Himachal Pradesh" Then
Combo2.AddItem "1  Chamba"
Combo2.AddItem "2  Kangra"
Combo2.AddItem "03 Lahul & Spiti"
Combo2.AddItem "4  Kullu"
Combo2.AddItem "5  Mandi"
Combo2.AddItem "6  Hamirpur"
Combo2.AddItem "7  Una"
Combo2.AddItem "8  Bilaspur"
Combo2.AddItem "9  Solan"
Combo2.AddItem "10 Sirmaur"
Combo2.AddItem "11 Shimla"
Combo2.AddItem "12 Kinnaur"
End If
If Combo1.Text = "Punjab" Then
Combo2.AddItem "1  Gurdaspur"
Combo2.AddItem "2  Kapurthala"
Combo2.AddItem "3  Jalandhar"
Combo2.AddItem "4  Hoshiarpur"
Combo2.AddItem "05 Shahid Bhagat Singh Nagar"
Combo2.AddItem "6  Fatehgarh Sahib"
Combo2.AddItem "7  Ludhiana"
Combo2.AddItem "8  Moga"
Combo2.AddItem "9  Firozpur"
Combo2.AddItem "10 Muktsar"
Combo2.AddItem "11 Faridkot"
Combo2.AddItem "12 Bathinda"
Combo2.AddItem "13 Mansa"
Combo2.AddItem "14 Patiala"
Combo2.AddItem "15 Amritsar"
Combo2.AddItem "16 Tarn Taran"
Combo2.AddItem "17 Rupnagar"
Combo2.AddItem "18 Sahibzada Ajit Singh Nagar"
Combo2.AddItem "19 Sangrur"
Combo2.AddItem "20 Barnala"
End If
If Combo1.Text = "Chandigarh" Then
Combo2.AddItem "01 Chandigarh"
End If
If Combo1.Text = "Uttarakhand" Then
Combo2.AddItem "1  Uttarkashi"
Combo2.AddItem "2  Chamoli"
Combo2.AddItem "3  Rudraprayag"
Combo2.AddItem "4  Tehri Garhwal"
Combo2.AddItem "5  Dehradun"
Combo2.AddItem "6  Garhwal"
Combo2.AddItem "7  Pithoragarh"
Combo2.AddItem "8  Bageshwar"
Combo2.AddItem "9  Almora"
Combo2.AddItem "10 Champawat"
Combo2.AddItem "11 Nainital"
Combo2.AddItem "12 Udham Singh Nagar"
Combo2.AddItem "13 Hardwar"
End If
If Combo1.Text = "Haryana" Then
Combo2.AddItem "1  Panchkula"
Combo2.AddItem "2  Ambala"
Combo2.AddItem "3  Yamunanagar"
Combo2.AddItem "4  Kurukshetra"
Combo2.AddItem "5  Kaithal"
Combo2.AddItem "6  Karnal"
Combo2.AddItem "7  Panipat"
Combo2.AddItem "8  Sonipat"
Combo2.AddItem "9  Jind"
Combo2.AddItem "10 Fatehabad"
Combo2.AddItem "11 Sirsa"
Combo2.AddItem "12 Hisar"
Combo2.AddItem "13 Bhiwani"
Combo2.AddItem "14 Rohtak"
Combo2.AddItem "15 Jhajjar"
Combo2.AddItem "16 Mahendragarh"
Combo2.AddItem "17 Rewari"
Combo2.AddItem "18 Gurgaon"
Combo2.AddItem "19 Mewat"
Combo2.AddItem "20 Faridabad"
Combo2.AddItem "21 Palwal"
End If
If Combo1.Text = "NCT of Delhi" Then
Combo2.AddItem "01 North West"
Combo2.AddItem "2  North"
Combo2.AddItem "3  North East"
Combo2.AddItem "4  East"
Combo2.AddItem "05 New Delhi"
Combo2.AddItem "6  Central"
Combo2.AddItem "7  West"
Combo2.AddItem "8  South West"
Combo2.AddItem "9  South"
End If
If Combo1.Text = "Rajasthan" Then
Combo2.AddItem "01 Ganganagar"
Combo2.AddItem "2  Hanumangarh"
Combo2.AddItem "3  Bikaner"
Combo2.AddItem "4  Churu"
Combo2.AddItem "5  Jhunjhunun"
Combo2.AddItem "6  Alwar"
Combo2.AddItem "7  Bharatpur"
Combo2.AddItem "8  Dhaulpur"
Combo2.AddItem "9  Karauli"
Combo2.AddItem "10 Sawai Madhopur"
Combo2.AddItem "11 Dausa"
Combo2.AddItem "12 Jaipur"
Combo2.AddItem "13 Sikar"
Combo2.AddItem "14 Nagaur"
Combo2.AddItem "15 Jodhpur"
Combo2.AddItem "16 Jaisalmer"
Combo2.AddItem "17 Barmer"
Combo2.AddItem "18 Jalor"
Combo2.AddItem "19 Sirohi"
Combo2.AddItem "20 Pali"
Combo2.AddItem "21 Ajmer"
Combo2.AddItem "22 Tonk"
Combo2.AddItem "23 Bundi"
Combo2.AddItem "24 Bhilwara"
Combo2.AddItem "25 Rajsamand"
Combo2.AddItem "26 Dungarpur"
Combo2.AddItem "27 Banswara"
Combo2.AddItem "28 Chittaurgarh"
Combo2.AddItem "29 Kota"
Combo2.AddItem "30 Baran"
Combo2.AddItem "31 Jhalawar"
Combo2.AddItem "32 Udaipur"
Combo2.AddItem "33 Pratapgarh"
End If
If Combo1.Text = "Uttar Pradesh" Then
Combo2.AddItem "01 Saharanpur"
Combo2.AddItem "2  Muzaffarnagar"
Combo2.AddItem "3  Bijnor"
Combo2.AddItem "4  Moradabad"
Combo2.AddItem "5  Rampur"
Combo2.AddItem "06 Jyotiba Phule Nagar"
Combo2.AddItem "7  Meerut"
Combo2.AddItem "8  Baghpat"
Combo2.AddItem "9  Ghaziabad"
Combo2.AddItem "10 Gautam Buddha Nagar"
Combo2.AddItem "11 Bulandshahar"
Combo2.AddItem "12 Aligarh"
Combo2.AddItem "13 Mahamaya Nagar"
Combo2.AddItem "14 Mathura"
Combo2.AddItem "15 Agra"
Combo2.AddItem "16 Firozabad"
Combo2.AddItem "17 Mainpuri"
Combo2.AddItem "18 Budaun"
Combo2.AddItem "19 Bareilly"
Combo2.AddItem "20 Pilibhit"
Combo2.AddItem "21 Shahjahanpur"
Combo2.AddItem "22 Kheri"
Combo2.AddItem "23 Sitapur"
Combo2.AddItem "24 Hardoi"
Combo2.AddItem "25 Unnao"
Combo2.AddItem "26 Lucknow"
Combo2.AddItem "27 Rae Bareli"
Combo2.AddItem "28 Farrukhabad"
Combo2.AddItem "29 Kannauj"
Combo2.AddItem "30 Etawah"
Combo2.AddItem "31 Auraiya"
Combo2.AddItem "32 Kanpur Dehat"
Combo2.AddItem "33 Kanpur Nagar"
Combo2.AddItem "34 Jalaun"
Combo2.AddItem "35 Jhansi"
Combo2.AddItem "36 Lalitpur"
Combo2.AddItem "37 Hamirpur"
Combo2.AddItem "38 Mahoba"
Combo2.AddItem "39 Banda"
Combo2.AddItem "40 Chitrakoot"
Combo2.AddItem "41 Fatehpur"
Combo2.AddItem "42 Pratapgarh"
Combo2.AddItem "43 Kaushambi"
Combo2.AddItem "44 Allahabad"
Combo2.AddItem "45 Bara Banki"
Combo2.AddItem "46 Faizabad"
Combo2.AddItem "47 Ambedkar Nagar"
Combo2.AddItem "48 Sultanpur"
Combo2.AddItem "49 Bahraich"
Combo2.AddItem "50 Shrawasti"
Combo2.AddItem "51 Balrampur"
Combo2.AddItem "52 Gonda"
Combo2.AddItem "53 Siddharthnagar"
Combo2.AddItem "54 Basti"
Combo2.AddItem "55 Sant Kabir Nagar"
Combo2.AddItem "56 Mahrajganj"
Combo2.AddItem "57 Gorakhpur"
Combo2.AddItem "58 Kushinagar"
Combo2.AddItem "59 Deoria"
Combo2.AddItem "60 Azamgarh"
Combo2.AddItem "61 Mau"
Combo2.AddItem "62 Ballia"
Combo2.AddItem "63 Jaunpur"
Combo2.AddItem "64 Ghazipur"
Combo2.AddItem "65 Chandauli"
Combo2.AddItem "66 Varanasi"
Combo2.AddItem "67 Sant Ravidas Nagar(Bhadohi)"
Combo2.AddItem "68 Mirzapur"
Combo2.AddItem "69 Sonbhadra"
Combo2.AddItem "70 Etah"
Combo2.AddItem "71 Kanshiram Nagar"
End If
If Combo1.Text = "Bihar" Then
Combo2.AddItem "01 Pashchim Champaran"
Combo2.AddItem "2  Purba Champaran"
Combo2.AddItem "3  Sheohar"
Combo2.AddItem "4  Sitamarhi"
Combo2.AddItem "5  Madhubani"
Combo2.AddItem "6  Supaul"
Combo2.AddItem "7  Araria"
Combo2.AddItem "8  Kishanganj"
Combo2.AddItem "9  Purnia"
Combo2.AddItem "10 Katihar"
Combo2.AddItem "11 Madhepura"
Combo2.AddItem "12 Saharsa"
Combo2.AddItem "13 Darbhanga"
Combo2.AddItem "14 Muzaffarpur"
Combo2.AddItem "15 Gopalganj"
Combo2.AddItem "16 Siwan"
Combo2.AddItem "17 Saran"
Combo2.AddItem "18 Vaishali"
Combo2.AddItem "19 Samastipur"
Combo2.AddItem "20 Begusarai"
Combo2.AddItem "21 Khagaria"
Combo2.AddItem "22 Bhagalpur"
Combo2.AddItem "23 Banka"
Combo2.AddItem "24 Munger"
Combo2.AddItem "25 Lakhisarai"
Combo2.AddItem "26 Sheikhpura"
Combo2.AddItem "27 Nalanda"
Combo2.AddItem "28 Patna"
Combo2.AddItem "29 Bhojpur"
Combo2.AddItem "30 Buxar"
Combo2.AddItem "31 Kaimur (Bhabua)"
Combo2.AddItem "32 Rohtas"
Combo2.AddItem "33 Aurangabad"
Combo2.AddItem "34 Gaya"
Combo2.AddItem "35 Nawada"
Combo2.AddItem "36 Jamui"
Combo2.AddItem "37 Jehanabad"
Combo2.AddItem "38 Arwal"
End If
If Combo1.Text = "Manipur" Then
Combo2.AddItem "1  Senapati"
Combo2.AddItem "2  Tamenglong"
Combo2.AddItem "3  Churachandpur"
Combo2.AddItem "4  Bishnupur"
Combo2.AddItem "5  Thoubal"
Combo2.AddItem "6  Imphal West"
Combo2.AddItem "7  Imphal East"
Combo2.AddItem "8  Ukhrul"
Combo2.AddItem "9  Chandel"
End If
If Combo1.Text = "Mizoram" Then
Combo2.AddItem "1  Mamit"
Combo2.AddItem "2  Kolasib"
Combo2.AddItem "3  Aizawl"
Combo2.AddItem "4  Champhai"
Combo2.AddItem "5  Serchhip"
Combo2.AddItem "6  Lunglei"
Combo2.AddItem "7  Lawngtlai"
Combo2.AddItem "8  Saiha"
End If
If Combo1.Text = "Meghalaya" Then
Combo2.AddItem "01 West Garo Hills"
Combo2.AddItem "02 East Garo Hills"
Combo2.AddItem "03 South Garo Hills"
Combo2.AddItem "04 West Khasi Hills"
Combo2.AddItem "5  Ri Bhoi"
Combo2.AddItem "06 East Khasi Hills"
Combo2.AddItem "7  Jaintia Hills"
End If
If Combo1.Text = "Assam" Then
Combo2.AddItem "1  Kokrajhar"
Combo2.AddItem "2  Dhubri"
Combo2.AddItem "3  Goalpara"
Combo2.AddItem "4  Barpeta"
Combo2.AddItem "5  Morigaon"
Combo2.AddItem "6  Nagaon"
Combo2.AddItem "7  Sonitpur"
Combo2.AddItem "8  Lakhimpur"
Combo2.AddItem "9  Dhemaji"
Combo2.AddItem "10 Tinsukia"
Combo2.AddItem "11 Dibrugarh"
Combo2.AddItem "12 Sivasagar"
Combo2.AddItem "13 Jorhat"
Combo2.AddItem "14 Golaghat"
Combo2.AddItem "15 Karbi Anglong"
Combo2.AddItem "16 Dima Hasao"
Combo2.AddItem "17 Cachar"
Combo2.AddItem "18 Karimganj"
Combo2.AddItem "19 Hailakandi"
Combo2.AddItem "20 Bongaigaon"
Combo2.AddItem "21 Chirang"
Combo2.AddItem "22 Kamrup"
Combo2.AddItem "23 Kamrup Metropolitan"
Combo2.AddItem "24 Nalbari"
Combo2.AddItem "25 Baksa"
Combo2.AddItem "26 Darrang"
Combo2.AddItem "27 Udalguri"
End If
If Combo1.Text = "West Bengal" Then
Combo2.AddItem "1  Darjiling"
Combo2.AddItem "2  Jalpaiguri"
Combo2.AddItem "3  Koch Bihar"
Combo2.AddItem "4  Uttar Dinajpur"
Combo2.AddItem "5  Dakshin Dinajpur"
Combo2.AddItem "6  Maldah"
Combo2.AddItem "7  Murshidabad"
Combo2.AddItem "8  Birbhum"
Combo2.AddItem "9  Barddhaman"
Combo2.AddItem "10 Nadia"
Combo2.AddItem "11 North Twenty Four"
Combo2.AddItem "Parganas"
Combo2.AddItem "12 Hugli"
Combo2.AddItem "13 Bankura"
Combo2.AddItem "14 Puruliya"
Combo2.AddItem "15 Haora"
Combo2.AddItem "16 Kolkata"
Combo2.AddItem "17 South Twenty Four Parganas"
Combo2.AddItem "18 Paschim Medinipur"
Combo2.AddItem "19 Purba Medinipur"
End If
If Combo1.Text = "Goa" Then
Combo2.AddItem "1  North Goa"
Combo2.AddItem "2  South Goa"
End If
If Combo1.Text = "Madhya Pradesh" Then
Combo2.AddItem "01 Sheopur"
Combo2.AddItem "2  Morena"
Combo2.AddItem "3  Bhind"
Combo2.AddItem "4  Gwalior"
Combo2.AddItem "5  Datia"
Combo2.AddItem "6  Shivpuri"
Combo2.AddItem "7  Tikamgarh"
Combo2.AddItem "8  Chhatarpur"
Combo2.AddItem "9  Panna"
Combo2.AddItem "10 Sagar"
Combo2.AddItem "11 Damoh"
Combo2.AddItem "12 Satna"
Combo2.AddItem "13 Rewa"
Combo2.AddItem "14 Umaria"
Combo2.AddItem "15 Neemuch"
Combo2.AddItem "16 Mandsaur"
Combo2.AddItem "17 Ratlam"
Combo2.AddItem "18 Ujjain"
Combo2.AddItem "19 Shajapur"
Combo2.AddItem "20 Dewas"
Combo2.AddItem "21 Dhar"
Combo2.AddItem "22 Indore"
Combo2.AddItem "23 West Nimar"
Combo2.AddItem "24 Barwani"
Combo2.AddItem "25 Rajgarh"
Combo2.AddItem "26 Vidisha"
Combo2.AddItem "27 Bhopal"
Combo2.AddItem "28 Sehore"
Combo2.AddItem "29 Raisen"
Combo2.AddItem "30 Betul"
Combo2.AddItem "31 Harda"
Combo2.AddItem "32 Hoshangabad"
Combo2.AddItem "33 Katni"
Combo2.AddItem "34 Jabalpur"
Combo2.AddItem "35 Narsimhapur"
Combo2.AddItem "36 Dindori"
Combo2.AddItem "37 Mandla"
Combo2.AddItem "38 Chhindwara"
Combo2.AddItem "39 Seoni"
Combo2.AddItem "40 Balaghat"
Combo2.AddItem "41 Guna"
Combo2.AddItem "42 Ashoknagar"
Combo2.AddItem "43 Shahdol"
Combo2.AddItem "44 Anuppur"
Combo2.AddItem "45 Sidhi"
Combo2.AddItem "46 Singrauli"
Combo2.AddItem "47 Jhabua"
Combo2.AddItem "48 Alirajpur"
Combo2.AddItem "49 East Nimar"
Combo2.AddItem "50 Burhanpur"
End If
If Combo1.Text = "Jharkhand" Then
End If
If Combo1.Text = "Chhattisgarh" Then
End If
If Combo1.Text = "Gujarat" Then
End If
If Combo1.Text = "Maharashtra" Then
Combo2.AddItem "1  Nandurbar"
Combo2.AddItem "2  Dhule"
Combo2.AddItem "3  Jalgaon"
Combo2.AddItem "4  Buldana"
Combo2.AddItem "5  Akola"
Combo2.AddItem "6  Washim"
Combo2.AddItem "7  Amravati"
Combo2.AddItem "8  Wardha"
Combo2.AddItem "9  Nagpur"
Combo2.AddItem "10 Bhandara"
Combo2.AddItem "11 Gondiya"
Combo2.AddItem "12 Gadchiroli"
Combo2.AddItem "13 Chandrapur"
Combo2.AddItem "14 Yavatmal"
Combo2.AddItem "15 Nanded"
Combo2.AddItem "16 Hingoli"
Combo2.AddItem "17 Parbhani"
Combo2.AddItem "18 Jalna"
Combo2.AddItem "19 Aurangabad"
Combo2.AddItem "20 Nashik"
Combo2.AddItem "21 Thane"
Combo2.AddItem "22 Mumbai Suburban"
Combo2.AddItem "23 Mumbai"
Combo2.AddItem "24 Raigarh"
Combo2.AddItem "25 Pune"
Combo2.AddItem "26 Ahmadnagar"
Combo2.AddItem "27 Bid"
Combo2.AddItem "28 Latur"
Combo2.AddItem "29 Osmanabad"
Combo2.AddItem "30 Solapur"
Combo2.AddItem "31 Satara"
Combo2.AddItem "32 Ratnagiri"
Combo2.AddItem "33 Sindhudurg"
Combo2.AddItem "34 Kolhapur"
Combo2.AddItem "35 Sangli"
End If
If Combo1.Text = "Karnataka" Then
End If
If Combo1.Text = "Kerala" Then
End If
If Combo1.Text = "Tamil Nadu" Then
End If

End Sub

Private Sub Command5_Click()
 MsgBox ("record deleted")


End Sub

Private Sub Command6_Click()
Form4.Show

End Sub

Private Sub Command7_Click()
End

End Sub

Private Sub Form_Load()


Combo1.AddItem "Jammu & Kashmir"
Combo1.AddItem "Himachal Pradesh"
Combo1.AddItem "Punjab"
Combo1.AddItem "Chandigarh"
Combo1.AddItem "Uttarakhand"
Combo1.AddItem "Haryana"
Combo1.AddItem "NCT of Delhi"
Combo1.AddItem "Rajasthan"
Combo1.AddItem "Uttar Pradesh"
Combo1.AddItem "Bihar"
Combo1.AddItem "Ganganagar"
Combo1.AddItem "Manipur"
Combo1.AddItem "Mizoram"
Combo1.AddItem "Tripura"
Combo1.AddItem "Meghalaya"
Combo1.AddItem "Assam"
Combo1.AddItem "West Bengal"
Combo1.AddItem "Jharkhand"
Combo1.AddItem "Chhattisgarh"
Combo1.AddItem "Madhya Pradesh"
Combo1.AddItem "Gujarat"
Combo1.AddItem "Maharashtra"
Combo1.AddItem "Karnataka"
Combo1.AddItem "Andhra Pradesh"
Combo1.AddItem "Goa"
Combo1.AddItem "Kerala"
Combo1.AddItem "Tamil Nadu"








Combo3.AddItem "Babu Bus"
Combo3.AddItem "Sree Ram Bus"
Combo3.AddItem "Sita Ram Bus"
Combo3.AddItem "Jai Shree Bus"
Combo3.AddItem "Rajasthan Bus"
Combo3.AddItem "Shreenath Bus"
Combo3.AddItem "Sairam Bus"
Combo3.AddItem "DC Bus"






Combo5.AddItem "Sleeper"
Combo5.AddItem "A.C"
Combo5.AddItem "Non A.C"

Combo6.AddItem "Single"
Combo6.AddItem "Return"
End Sub

Private Sub Timer1_Timer()
Label12.Caption = Date
Label13.Caption = Time

End Sub

