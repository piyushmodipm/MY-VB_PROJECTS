VERSION 5.00
Begin VB.MDIForm vmsfrms 
   BackColor       =   &H8000000C&
   Caption         =   "Voting Management System"
   ClientHeight    =   8115
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15765
   Icon            =   "vmsfrms.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "vmsfrms.frx":0832
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuvotermaster 
      Caption         =   "Voter Master"
      Begin VB.Menu minewvoter 
         Caption         =   "New Voter"
      End
      Begin VB.Menu mivoterdetails 
         Caption         =   "Voter Details"
      End
      Begin VB.Menu mivoterupdate 
         Caption         =   "Voter Update"
      End
   End
   Begin VB.Menu mnucandidatemaster 
      Caption         =   "Candidate Master"
      Begin VB.Menu minewcandidate 
         Caption         =   "New Candidate Registration"
      End
      Begin VB.Menu micandidatedetails 
         Caption         =   "Candidate Details"
      End
      Begin VB.Menu micandidateupdation 
         Caption         =   "candidate Updation"
      End
   End
   Begin VB.Menu mnuVotingmaster 
      Caption         =   "Voting Master"
      Begin VB.Menu minvote 
         Caption         =   "Vote "
      End
      Begin VB.Menu mivotesdetails 
         Caption         =   "Votes Details"
      End
   End
   Begin VB.Menu mnucomplainmaster 
      Caption         =   "Complaints Master"
      Begin VB.Menu micomplain 
         Caption         =   "Complain"
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
Attribute VB_Name = "vmsfrms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub micandidatedetails_Click()
frmCandidatedetails.Show
End Sub

Private Sub micandidateupdation_Click()
frmcandidateupdation.Show
End Sub

Private Sub miexit_Click()
Unload Me
End Sub

Private Sub minewcandidate_Click()
frmnewCandidateregistration.Show
End Sub

Private Sub minewvoter_Click()
frmnewvoter.Show
End Sub

Private Sub minvote_Click()
frmvote.Show
End Sub

Private Sub mivoterdetails_Click()
frmvoterdetails.Show
End Sub


Private Sub mivoterupdate_Click()
frmvoterUpdate.Show
End Sub

Private Sub mivotesdetails_Click()
frmVotesdetails.Show
End Sub
