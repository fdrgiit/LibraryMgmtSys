VERSION 5.00
Begin VB.Form ADDING 
   BackColor       =   &H000040C0&
   Caption         =   "POPULATING BOOKS"
   ClientHeight    =   11820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   11820
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
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
      ItemData        =   "ADD BUTTON.frx":0000
      Left            =   3480
      List            =   "ADD BUTTON.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Frame fam1 
      BackColor       =   &H8000000B&
      Caption         =   "DATA ENTRY"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox txtid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4560
         Width           =   2355
      End
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox t2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H00FF80FF&
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         MaskColor       =   &H000080FF&
         Picture         =   "ADD BUTTON.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   7320
         Width           =   5295
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         Caption         =   "ID                               ::"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lbl2 
         BackColor       =   &H8000000B&
         Caption         =   "BOOK NAME        ::"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lbl3 
         BackColor       =   &H8000000B&
         Caption         =   "AUTHOR NAME   ::"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lbl10 
         BackColor       =   &H8000000B&
         Caption         =   "BRANCH                 ::"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   6000
         Width           =   2415
      End
   End
   Begin VB.Frame fam2 
      BackColor       =   &H8000000B&
      Caption         =   "DATA STORED"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   12480
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame fam3 
         BackColor       =   &H8000000B&
         Caption         =   "GENERAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   480
         TabIndex        =   12
         Top             =   960
         Width           =   5535
         Begin VB.Label lbl13 
            BackColor       =   &H8000000B&
            Caption         =   "BOOK NAME         ::"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label l2 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   17
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label l1 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   16
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lbl14 
            BackColor       =   &H8000000B&
            Caption         =   "AUTHOR NAME  ::"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000B&
            Caption         =   "DATE OF BIRTH  ::"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   4200
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000007&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   495
            Left            =   2520
            TabIndex        =   13
            Top             =   4320
            Width           =   2655
         End
      End
      Begin VB.Frame fam4 
         BackColor       =   &H8000000B&
         Caption         =   "ADMINISTRATIVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   480
         TabIndex        =   7
         Top             =   5400
         Width           =   5535
         Begin VB.Label l7 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   11
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lbl19 
            BackColor       =   &H8000000B&
            Caption         =   "BOOK ID                 ::"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label lbl20 
            BackColor       =   &H8000000B&
            Caption         =   "BRANCH                 ::"
            BeginProperty Font 
               Name            =   "Perpetua"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label l8 
            BackColor       =   &H8000000B&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   8
            Top             =   1920
            Width           =   2655
         End
      End
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Picture         =   "ADD BUTTON.frx":662A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10440
      Width           =   6255
   End
   Begin VB.CommandButton cmd3 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12360
      Picture         =   "ADD BUTTON.frx":CB45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   10440
      Width           =   6375
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Caption         =   "ADDING BOOKS IN LMS"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   23
      Top             =   120
      Width           =   9975
   End
   Begin VB.Image Image1 
      Height          =   5025
      Left            =   6840
      Picture         =   "ADD BUTTON.frx":10AEC
      Top             =   3000
      Width           =   7500
   End
End
Attribute VB_Name = "ADDING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmd1_Click()
Dim str As String
If (t1.Text = "" Or t2.Text = "" Or Combo1.Text = "") Then
MsgBox "KINDLY ENTER COMPLETE DATA", vbCritical, "POPULATING BOOKS"

Else
rs.AddNew
rs.Fields(0) = t1.Text
rs.Fields(1) = t2.Text
rs.Fields(2) = Combo1.Text
rs.UPDATE
cmd1.Enabled = False
fam2.Visible = True
t1.Enabled = False
t2.Enabled = False
txtid.Enabled = False
Combo1.Enabled = False

l1.Caption = t1.Text
l2.Caption = t2.Text
l7.Caption = txtid.Text
l8.Caption = Combo1.Text
    
End If

End Sub

Private Sub cmd2_Click()
cmd1.Enabled = True
t1.Text = ""
t2.Text = ""
rs.MoveLast
txtid.Text = rs.Fields(3) + 1
t1.Enabled = True
t2.Enabled = True
txtid.Enabled = True
Combo1.Enabled = True
fam2.Visible = False

End Sub

Private Sub cmd3_Click()
Me.Hide
LIBRARY.Show
End Sub

Private Sub Form_Load()

Combo1.AddItem ("IT")
Combo1.AddItem ("COMPS")
Combo1.AddItem ("EXTC")
Combo1.AddItem ("MECH")


Dim str As String
Set adoconn = Nothing
adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database11.mdb;Persist Security Info=False"
str = "select * from library"
rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
rs.MoveLast
txtid.Text = rs.Fields(3) + 1

End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
Exit Sub
Else
MsgBox "INVALID VALUE", vbCritical, "POPULATING BOOKS"
KeyAscii = False
End If
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 39 Then
Exit Sub
Else
MsgBox "INVALID VALUE", vbCritical, "POPULATING BOOKS"
KeyAscii = False
End If
End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
Exit Sub
Else
MsgBox "INVALID VALUE", vbCritical, "POPULATING BOOKS"
KeyAscii = False
End If
End Sub




