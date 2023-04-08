VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bus Forum"
   ClientHeight    =   10950
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   20250
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form5"
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
Combo4.Text = ""
Combo5.Text = ""
Combo6.Text = ""
Text4.Text = ""

End Sub

Private Sub Command4_Click()
Form7.Show

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
Combo1.AddItem ""
Combo1.AddItem "Udaipur"
Combo1.AddItem "Bhilwara"
Combo1.AddItem "Alwar"
Combo1.AddItem "Bharatpur"
Combo1.AddItem "Sikar"
Combo1.AddItem "Ganganagar"
Combo1.AddItem "Pali"
Combo1.AddItem "Chittorgarh"
Combo1.AddItem "Tonk"
Combo1.AddItem "Kishangarh"
Combo1.AddItem "Bewar"
Combo1.AddItem "Hanumangarh"
Combo1.AddItem "Dholpur"
Combo1.AddItem "Gangapur city"
Combo1.AddItem "Sawai Madhopur"
Combo1.AddItem "Churu"
Combo1.AddItem "Pali"
Combo1.AddItem "Nagpur"







Combo2.AddItem "Jaipur"
Combo2.AddItem "Jodhpur"
Combo2.AddItem "Kota"
Combo2.AddItem "Bikaner"
Combo2.AddItem "Ajmer"
Combo2.AddItem "Udaipur"
Combo2.AddItem "Bhilwara"
Combo2.AddItem "Alwar"
Combo2.AddItem "Bharatpur"
Combo2.AddItem "Sikar"
Combo2.AddItem "Ganganagar"
Combo2.AddItem "Pali"
Combo2.AddItem "Chittorgarh"
Combo2.AddItem "Tonk"
Combo2.AddItem "Kishangarh"
Combo2.AddItem "Bewar"
Combo2.AddItem "Hanumangarh"
Combo2.AddItem "Dholpur"
Combo2.AddItem "Gangapur city"
Combo2.AddItem "Sawai Madhopur"
Combo2.AddItem "Churu"
Combo2.AddItem "Pali"
Combo2.AddItem "Nagpur"






Combo3.AddItem "Babu Bus"
Combo3.AddItem "Sree Ram Bus"
Combo3.AddItem "Sita Ram Bus"
Combo3.AddItem "Jai Shree Bus"
Combo3.AddItem "Rajasthan Bus"
Combo3.AddItem "Shreenath Bus"
Combo3.AddItem "Sairam Bus"
Combo3.AddItem "DC Bus"






Combo4.AddItem "1"
Combo4.AddItem "2"
Combo4.AddItem "3"
Combo4.AddItem "4"
Combo4.AddItem "5"
Combo4.AddItem "6"
Combo4.AddItem "7"
Combo4.AddItem "8"
Combo4.AddItem "9"
Combo4.AddItem "10"
Combo4.AddItem "11"
Combo4.AddItem "12"
Combo4.AddItem "13"
Combo4.AddItem "14"
Combo4.AddItem "15"
Combo4.AddItem "16"
Combo4.AddItem "17"
Combo4.AddItem "18"
Combo4.AddItem "19"
Combo4.AddItem "20"
Combo4.AddItem "21"
Combo4.AddItem "22"
Combo4.AddItem "23"
Combo4.AddItem "24"
Combo4.AddItem "25"
Combo4.AddItem "26"
Combo4.AddItem "27"
Combo4.AddItem "28"
Combo4.AddItem "29"
Combo4.AddItem "30"



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
