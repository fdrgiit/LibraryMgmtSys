VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   LinkTopic       =   "Form2"
   ScaleHeight     =   11670
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6480
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   5280
      Width           =   3375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      TabIndex        =   6
      Top             =   8400
      Width           =   6135
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7680
      TabIndex        =   5
      Top             =   6720
      Width           =   5895
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   4
      Top             =   6720
      Width           =   6255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      TabIndex        =   3
      Top             =   9960
      Width           =   7455
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "BRANCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "BOOK NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "BOOK ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdExit_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub cmdSave_Click()
    rs.Fields(1) = Text2.Text
    rs.Fields(2) = Text3.Text
    rs.Update
    MsgBox "The record has been saved successfully.", , "ADD"
        
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Combo1.Locked = True
    
    cmdSearch.Enabled = True
    cmdUpdate.Enabled = True
    cmdSave.Visible = False
    cmdExit.Enabled = True
           
End Sub

Private Sub cmdSearch_Click()

    Dim key As String, str As String
    key = InputBox("ENTER THE BOOK-ID WHOSE DETAILS YOU WANT TO KNOW : ")
    If IsNumeric(key) Then
        key = CInt(key)
        Set rs = Nothing
        str = "select * from library where book_id=" & key
        rs.Open str, adoconn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
        Text1.Text = rs.Fields(0)
        Text2.Text = rs.Fields(1)
        Text3.Text = rs.Fields(2)
        Combo1.Text = rs.Fields(3)
        Else
            MsgBox "NO RECORDS WHERE FOUND FOR THAT BOOK ID."
        End If
    Else
        MsgBox "BOOK ID MUST BE NUMERIC."
    End If
    
    Set rs = Nothing
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    
End Sub

Private Sub cmdUpdate_Click()
    Dim ans As String
    ans = MsgBox("Do you really want to modify the current record?", vbExclamation + vbYesNo, "DELETE")
    If ans = vbYes Then
        Text1.Locked = True
        Text2.Locked = False
        Text3.Locked = False
        Combo1.Locked = False
        rs.Update
        
        cmdSearch.Enabled = False
        cmdUpdate.Enabled = False
        cmdSave.Visible = True
        cmdExit.Enabled = False
                
    End If
End Sub

Private Sub Form_Load()

    Combo1.AddItem ("IT")
    Combo1.AddItem ("COMPS")
    Combo1.AddItem ("EXTC")
    Combo1.AddItem ("MECH")

    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Combo1.Locked = True
    
    cmdSave.Visible = False
    Dim str As String
    Me.Caption = "LIBRARY DATABASE"
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Administrator\My Documents\Database11.mdb;Persist Security Info=False"
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
        Text1.Text = rs.Fields(0)
        Text2.Text = rs.Fields(1)
        Text3.Text = rs.Fields(2)
        Combo1.Text = rs.Fields(3)
        
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
Exit Sub
Else
MsgBox "Invalid Value"
KeyAscii = False
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
Exit Sub
Else
MsgBox "Invalid Value"
KeyAscii = False
End If
End Sub

