VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   11715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13800
   LinkTopic       =   "Form3"
   ScaleHeight     =   11715
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
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
      Left            =   3840
      TabIndex        =   16
      Top             =   6840
      Width           =   6135
   End
   Begin VB.CommandButton Command8 
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
      Left            =   10320
      TabIndex        =   15
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Command7 
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
      Left            =   10320
      TabIndex        =   14
      Top             =   4920
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DELETE"
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
      Left            =   10320
      TabIndex        =   13
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
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
      Left            =   3960
      TabIndex        =   12
      Top             =   10320
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10680
      TabIndex        =   11
      Top             =   8400
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7200
      TabIndex        =   10
      Top             =   8400
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   9
      Top             =   8400
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   8
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   6360
      TabIndex        =   1
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1440
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
      Left            =   2760
      TabIndex        =   7
      Top             =   5520
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
      Left            =   2760
      TabIndex        =   6
      Top             =   4200
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
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()

Command2.Enabled = True
Command3.Enabled = False
rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)

End Sub



Private Sub Command2_Click()

Command2.Enabled = True
If rs.EOF = True Then
MsgBox "Last Record"
Command3.Enabled = False
Else
rs.MoveNext
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
End If

End Sub

Private Sub Command3_Click()

'Set rs = New ADODB.Recordset
'rs.Open "select * from library", cn, adOpenDynamic, adLockOptimistic
'rs.MovePrevious
If rs.BOF = True Then
Command2.Enabled = False
Else
rs.MovePrevious
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
End If

End Sub

Private Sub Command4_Click()

Command3.Enabled = True
Command2.Enabled = False
rs.MoveLast
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)

End Sub

Private Sub Command5_Click()

End

End Sub

Private Sub Command6_Click()
    Dim ans As String, str As String
    ans = MsgBox("Do you really want to delete the current record?", vbExclamation + vbYesNo, "DELETE")
    If ans = vbYes Then
        adoconn.Execute ("delete from library where book_id=" & Text1.Text)
        MsgBox ("The record has been deleted successfully.")
        Set rs = Nothing
        str = "select * from library"
        rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
        rs.MoveFirst
        Text1.Text = rs.Fields(0)
        Text2.Text = rs.Fields(1)
        Text3.Text = rs.Fields(2)
        Text4.Text = rs.Fields(3)
        
    End If
End Sub

Private Sub Command7_Click()
    Dim ans As String
    ans = MsgBox("Do you really want to modify the current record?", vbExclamation + vbYesNo, "DELETE")
    If ans = vbYes Then
        rs.Update
        
        Command1.Enabled = False
        Command4.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command6.Enabled = False
        Command8.Enabled = False
        Command7.Enabled = False
        Command9.Visible = True
        Command5.Enabled = False
      
        
    End If
End Sub

Private Sub Command8_Click()
    Dim key As String, str As String
    key = InputBox("ENTER THE BOOK ID WHOSE DETAILS YOU WANT TO KNOW: ")
    If IsNumeric(key) Then
        key = CInt(key)
        Set rs = Nothing
        str = "select * from library where BOOK_ID=" & key
        rs.Open str, adoconn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
        Text1.Text = rs.Fields(0)
        Text2.Text = rs.Fields(1)
        Text3.Text = rs.Fields(2)
        Text4.Text = rs.Fields(3)
        Else
            MsgBox "No records found for that ID."
        End If
    Else
        MsgBox "Book id. must be numeric."
    End If
    
    Set rs = Nothing
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Command9_Click()

    rs.Fields(1) = Text2.Text
    rs.Fields(2) = Text3.Text
    rs.Fields(3) = Text4.Text
    rs.Update
    MsgBox "The record has been saved successfully.", , "ADD"
        Command1.Enabled = True
        Command4.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = True
        Command6.Enabled = True
        Command8.Enabled = True
        Command7.Enabled = True
        Command9.Visible = False
        Command5.Enabled = True
       
End Sub

Private Sub Form_Load()

    Dim str As String
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Administrator\My Documents\Database11.mdb;Persist Security Info=False"
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
    Text1.Text = rs.Fields(0)
    Text2.Text = rs.Fields(1)
    Text3.Text = rs.Fields(2)
    Text4.Text = rs.Fields(3)


End Sub

