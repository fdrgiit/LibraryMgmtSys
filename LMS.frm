VERSION 5.00
Begin VB.MDIForm LIBRARY 
   BackColor       =   &H8000000B&
   Caption         =   "LIBRARY MANAGEMENT SYSTEM"
   ClientHeight    =   12885
   ClientLeft      =   1995
   ClientTop       =   750
   ClientWidth     =   18960
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   12855
      Left            =   0
      Picture         =   "LMS.frx":0000
      ScaleHeight     =   12795
      ScaleWidth      =   18900
      TabIndex        =   0
      Top             =   0
      Width           =   18960
   End
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
Attribute VB_Name = "LIBRARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADD_Click()
Me.Hide
ADDING.Show
End Sub

Private Sub DELETE_Click()
Me.Hide
DELETING.Show
End Sub

Private Sub DISPLAYING_Click()
Me.Hide
DISPLAY.Show
End Sub

Private Sub EXIT_Click()
MsgBox "THANK YOU FOR COMING TO LIBRARY OF CONGRESS.", vbOKOnly, "LIBRARY MANAGEMENT SYSTEM"
End
End Sub

Private Sub UPDATE_Click()
Me.Hide
UPDATING.Show
End Sub
