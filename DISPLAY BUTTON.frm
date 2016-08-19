VERSION 5.00
Begin VB.Form DISPLAY 
   BackColor       =   &H00C00000&
   Caption         =   "PRESENTATION OF BOOKS"
   ClientHeight    =   11715
   ClientLeft      =   1995
   ClientTop       =   2550
   ClientWidth     =   13800
   LinkTopic       =   "Form3"
   ScaleHeight     =   11715
   ScaleWidth      =   13800
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      Picture         =   "DISPLAY BUTTON.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6720
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      Picture         =   "DISPLAY BUTTON.frx":7BF5
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10320
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11040
      Picture         =   "DISPLAY BUTTON.frx":BB9C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      Picture         =   "DISPLAY BUTTON.frx":C4D7
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5400
      Picture         =   "DISPLAY BUTTON.frx":CF29
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      Picture         =   "DISPLAY BUTTON.frx":D98E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   9840
      Picture         =   "DISPLAY BUTTON.frx":E2C4
      Top             =   1440
      Width           =   6000
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "DISPLAYING  BOOKS  IN  LMS"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   14
      Top             =   360
      Width           =   12135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "BRANCH"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   7
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "BOOK NAME"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "BOOK ID"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "DISPLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()

rs.MoveFirst
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)

End Sub



Private Sub Command2_Click()

rs.MoveNext
If rs.EOF = True Then
MsgBox "THIS IS THE LAST RECORD.", vbExclamation, "NOTE IT..."
rs.MoveLast
End If
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)


End Sub

Private Sub Command3_Click()

rs.MovePrevious
If rs.BOF = True Then
MsgBox "THIS IS THE FIRST RECORD.", vbExclamation, "NOTE IT..."
rs.MoveFirst
End If
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)

End Sub

Private Sub Command4_Click()

rs.MoveLast
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)

End Sub

Private Sub Command5_Click()
Me.Hide
LIBRARY.Show
End Sub

Private Sub Command8_Click()
    Dim key As String, str As String
    key = InputBox("ENTER THE BOOK ID WHOSE DETAILS YOU WANT TO KNOW: ")
    If IsNumeric(key) Then
        key = CInt(key)
        Set rs = Nothing
        str = "select * from library where id_no=" & key
        rs.Open str, adoconn, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)
        Else
            MsgBox "NO RECORDS WHERE FOUND FOR THAT BOOK ID.", vbOKOnly, "PRESENTATION OF BOOKS"
        End If
    Else
        MsgBox "BOOK ID MUST BE NUMERIC.", vbExclamation, "PRESENTATION OF BOOKS"
    End If
    
    Set rs = Nothing
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Form_Load()

    Dim str As String
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database11.mdb;Persist Security Info=False"
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)

End Sub


