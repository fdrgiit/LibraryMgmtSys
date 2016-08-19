VERSION 5.00
Begin VB.Form DELETING 
   BackColor       =   &H00404000&
   Caption         =   "REMOVING BOOKS"
   ClientHeight    =   11580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   11580
   ScaleWidth      =   13410
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLast 
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12120
      Picture         =   "DELETE BUTTON.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   12480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrevious 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9000
      Picture         =   "DELETE BUTTON.frx":093B
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   12480
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Picture         =   "DELETE BUTTON.frx":138D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   12480
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirst 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2040
      Picture         =   "DELETE BUTTON.frx":1DF2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   12480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5880
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      Picture         =   "DELETE BUTTON.frx":2728
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10320
      Width           =   6135
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   1320
      MaskColor       =   &H000000FF&
      Picture         =   "DELETE BUTTON.frx":66CF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   5535
   End
   Begin VB.CommandButton cmdSearch 
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   8520
      Picture         =   "DELETE BUTTON.frx":15509
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   5220
      Left            =   10080
      Picture         =   "DELETE BUTTON.frx":1D0FE
      Top             =   1800
      Width           =   7500
   End
   Begin VB.Label Label5 
      Caption         =   "DELETING AND SEARCHING BOOKS IN LMS"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   16695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      Caption         =   "BOOK ID"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      Caption         =   "BOOK NAME"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   2640
      TabIndex        =   9
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   3120
      TabIndex        =   8
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      Caption         =   "BRANCH"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   2880
      TabIndex        =   7
      Top             =   5880
      Width           =   3015
   End
End
Attribute VB_Name = "DELETING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdDelete_Click()

    Dim ans As String, str As String
    ans = MsgBox("DO YOU REALLY WANT TO DELETE THE CURRENT RECORD?", vbExclamation + vbYesNo, "DELETE")
    If ans = vbYes Then
        adoconn.Execute ("delete from library where id_no=" & Text1.Text)
        MsgBox ("THE RECORD HAS BEEN DELETED SUCCESSFULLY."), vbOKOnly, "REMOVING BOOKS"
        Set rs = Nothing
        str = "select * from library"
        rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
        rs.MoveFirst
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)
    End If
    
End Sub

Private Sub cmdExit_Click()
Me.Hide
LIBRARY.Show
End Sub

Private Sub cmdFirst_Click()
    rs.MoveFirst
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)
End Sub

Private Sub cmdLast_Click()
    rs.MoveLast
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Text4.Text = rs.Fields(2)
End Sub

Private Sub cmdNext_Click()
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

Private Sub cmdPrevious_Click()
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

Private Sub cmdSearch_Click()

    Dim key As String, str As String
    key = InputBox("ENTER THE BOOK-ID WHOSE DETAILS YOU WANT TO KNOW : ")
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
            MsgBox "NO RECORDS WHERE FOUND FOR THAT BOOK ID.", vbOKOnly, "REMOVING BOOKS"
        End If
    Else
        MsgBox "BOOK ID MUST BE NUMERIC.", vbExclamation, "REMOVING BOOKS"
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

