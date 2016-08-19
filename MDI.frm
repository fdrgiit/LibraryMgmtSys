VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   11700
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13725
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu NEW 
      Caption         =   "FILE"
      Begin VB.Menu ADD 
         Caption         =   "ADDING"
         Shortcut        =   ^A
      End
      Begin VB.Menu UPDATE 
         Caption         =   "UPDATING"
         Shortcut        =   ^U
      End
      Begin VB.Menu DELETE 
         Caption         =   "DELETING"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu DISPLAYING 
         Caption         =   "DISPLAY"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADD_Click()
Me.Hide
Form1.Show
End Sub

Private Sub EXIT_Click()
End
End Sub
