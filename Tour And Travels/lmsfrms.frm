VERSION 5.00
Begin VB.MDIForm lmsfrms 
   BackColor       =   &H8000000C&
   Caption         =   "Library Management System"
   ClientHeight    =   8115
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15765
   LinkTopic       =   "MDIForm1"
   Picture         =   "lmsfrms.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnubookmaster 
      Caption         =   "New Entry"
      Begin VB.Menu minewbook 
         Caption         =   "New Booking"
      End
      Begin VB.Menu mibookdetails 
         Caption         =   "Booking Details"
      End
   End
   Begin VB.Menu mnumembermaster 
      Caption         =   "Member Master"
      Begin VB.Menu minewmember 
         Caption         =   "New Member Registration"
      End
      Begin VB.Menu mimemberdetails 
         Caption         =   "Member Details"
      End
      Begin VB.Menu mimemberupdation 
         Caption         =   "Member Updation"
      End
   End
   Begin VB.Menu mnusuppliermaster 
      Caption         =   "Guide"
      Begin VB.Menu minewsupplier 
         Caption         =   "New Guide"
      End
      Begin VB.Menu misupplierupdation 
         Caption         =   "Guide Updation"
      End
      Begin VB.Menu misupplierdetails 
         Caption         =   "Guide Details"
      End
   End
   Begin VB.Menu mnutransactionmaster 
      Caption         =   "Transaction "
      Begin VB.Menu miissue 
         Caption         =   "Payment"
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
      Begin VB.Menu miaboutus 
         Caption         =   "About Us"
      End
      Begin VB.Menu miexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "lmsfrms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub miaboutus_Click()
frmaboutus.Show
End Sub

Private Sub mibookdetails_Click()
frmbookdetails.Show

End Sub

Private Sub mibooksearch_Click()
frmbooksearch.Show

End Sub

Private Sub mibookupdation_Click()
frmbookupdation.Show

End Sub

Private Sub micancelbook_Click()
frmcancelbook.Show

End Sub

Private Sub miexit_Click()
End

End Sub

Private Sub miissue_Click()
frmissue.Show
End Sub

Private Sub mimemberdetails_Click()
frmmemberdetails.Show

End Sub

Private Sub mimembersearch_Click()
frmmembersearch.Show

End Sub

Private Sub mimemberupdation_Click()
frmmemberupdation.Show

End Sub

Private Sub minewbook_Click()
Form1.Show


End Sub

Private Sub minewmember_Click()
frmnewmemberregistration.Show

End Sub

Private Sub minewsupplier_Click()
frmnewsupplier.Show
End Sub

Private Sub mireturn_Click()
frmreturn.Show

End Sub

Private Sub misupplierdetails_Click()
frmsupplierdetails.Show

End Sub

Private Sub misupplierupdation_Click()
frmsupplierupdation.Show
End Sub


