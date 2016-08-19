VERSION 5.00
Begin VB.Form UPDATING 
   BackColor       =   &H00800080&
   Caption         =   "MODIFYING ENTRIES"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   LinkTopic       =   "Form2"
   ScaleHeight     =   13890
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   4440
      Picture         =   "UPDATE BUTTON.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9120
      Width           =   1215
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
      Left            =   8040
      Picture         =   "UPDATE BUTTON.frx":0936
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9120
      Width           =   1335
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
      Left            =   11400
      Picture         =   "UPDATE BUTTON.frx":139B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9120
      Width           =   1335
   End
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
      Left            =   14520
      Picture         =   "UPDATE BUTTON.frx":1DED
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9120
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4920
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   6000
      Width           =   3375
   End
   Begin VB.CommandButton cmdSave 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   9720
      Picture         =   "UPDATE BUTTON.frx":2728
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7260
      Width           =   7815
   End
   Begin VB.CommandButton cmdUpdate 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   960
      Picture         =   "UPDATE BUTTON.frx":5A1C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   8415
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   56.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6960
      Picture         =   "UPDATE BUTTON.frx":19FA3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10695
      Width           =   6135
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
      Left            =   4920
      MaxLength       =   50
      TabIndex        =   2
      Top             =   4560
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
      Left            =   4920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3240
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
      Left            =   4920
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      Caption         =   "UPDATING BOOKS IN LMS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   2520
      TabIndex        =   11
      Top             =   600
      Width           =   9615
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   8760
      Picture         =   "UPDATE BUTTON.frx":1DF4A
      Top             =   1920
      Width           =   4800
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800080&
      Caption         =   "BRANCH"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1320
      TabIndex        =   9
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800080&
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1320
      TabIndex        =   8
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800080&
      Caption         =   "BOOK NAME"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1320
      TabIndex        =   7
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "BOOK ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
End
Attribute VB_Name = "UPDATING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdExit_Click()
Me.Hide
LIBRARY.Show
End Sub

Private Sub cmdFirst_Click()
    rs.MoveFirst
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Combo1.Text = rs.Fields(2)
End Sub


Private Sub cmdLast_Click()
    rs.MoveLast
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Combo1.Text = rs.Fields(2)
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
        Combo1.Text = rs.Fields(2)
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
        Combo1.Text = rs.Fields(2)
End Sub

Private Sub cmdSave_Click()
    
        
        rs.Fields(0) = Text2.Text
        rs.Fields(1) = Text3.Text
        rs.Fields(2) = Combo1.Text
    rs.UPDATE

    MsgBox "THE RECORD HAS BEEN SAVED SUCCESSFULLY.", , "ADD"
        
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Combo1.Locked = True
    
  
    cmdUpdate.Enabled = True
    cmdSave.Visible = False
    cmdExit.Enabled = True
    cmdFirst.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    cmdPrevious.Enabled = True
           
End Sub


Private Sub cmdUpdate_Click()
    Dim ans As String
    ans = MsgBox("Do you really want to modify the current record?", vbExclamation + vbYesNo, "ADD")
    If ans = vbYes Then
        Text1.Locked = True
        Text2.Locked = False
        Text3.Locked = False
        Combo1.Locked = False
        rs.UPDATE
        
        
        cmdUpdate.Enabled = False
        cmdSave.Visible = True
        cmdExit.Enabled = False
        cmdFirst.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
        cmdPrevious.Enabled = False
                
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
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database11.mdb;Persist Security Info=False"
    str = "select * from library"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
        Text1.Text = rs.Fields(3)
        Text2.Text = rs.Fields(0)
        Text3.Text = rs.Fields(1)
        Combo1.Text = rs.Fields(2)
        
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii >= 80 And KeyAscii <= 89 Then
Exit Sub
Else
MsgBox "INVALID VALUE", vbCritical, "MODIFYING ENTRIES"
KeyAscii = False
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Then
Exit Sub
Else
MsgBox "INVALID VALUE", vbCritical, "MODIFYING ENTRIES"
KeyAscii = False
End If
End Sub



