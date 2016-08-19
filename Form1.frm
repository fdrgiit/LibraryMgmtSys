VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   13725
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3960
      List            =   "Form1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6840
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
      Left            =   840
      TabIndex        =   19
      Top             =   840
      Width           =   6255
      Begin VB.TextBox t1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
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
            Name            =   "MS Sans Serif"
            Size            =   12
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
      Begin VB.TextBox t5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   5
         Top             =   4800
         Width           =   2295
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   0
         Top             =   7920
         Width           =   2055
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
         TabIndex        =   23
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lbl8 
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
         Left            =   480
         TabIndex        =   22
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label lbl3 
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
         Left            =   480
         TabIndex        =   21
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
         TabIndex        =   20
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
      Left            =   7440
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame fam3 
         BackColor       =   &H8000000B&
         Caption         =   "PERSONAL"
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
         TabIndex        =   13
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
            TabIndex        =   24
            Top             =   360
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
            TabIndex        =   18
            Top             =   1080
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   960
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
            TabIndex        =   15
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
            TabIndex        =   14
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
         TabIndex        =   8
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   960
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
            TabIndex        =   9
            Top             =   960
            Width           =   2655
         End
      End
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   9960
      Width           =   1695
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   9960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As ADODB.Connection
Public RS As ADODB.Recordset
Private Sub cmd1_Click()
If (t1.Text = "" Or t2.Text = "" Or t5.Text = "" Or Combo1.Text = "") Then
MsgBox "Kindly Enter Complete Data"
Else
Set RS = New ADODB.Recordset
RS.Open "select * from library", cn, adOpenDynamic, adLockOptimistic
RS.AddNew
RS(0) = Val(t5.Text)

'MsgBox "value of rs1 is :" & rs(0).Value

RS(1) = t1.Text
RS(2) = t2.Text
RS(3) = Combo1.Text
RS.Update
cmd1.Enabled = False
fam2.Visible = True
t1.Enabled = False
t2.Enabled = False
t5.Enabled = False
Combo1.Enabled = False

l1.Caption = t1.Text
l2.Caption = t2.Text
l7.Caption = t5.Text
l8.Caption = Combo1.Text

End If
End Sub

Private Sub cmd2_Click()
cmd1.Enabled = True
t1.Text = ""
t2.Text = ""
t5.Text = ""
t1.Enabled = True
t2.Enabled = True
t5.Enabled = True
Combo1.Enabled = True
fam2.Visible = False

End Sub

Private Sub cmd3_Click()
End
End Sub

Private Sub Form_Load()

Combo1.AddItem ("IT")
Combo1.AddItem ("COMPS")
Combo1.AddItem ("EXTC")
Combo1.AddItem ("MECH")


Set cn = New ADODB.Connection
Set RS = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\Administrator\My Documents\Database11.mdb;Persist Security Info=False"
'rs.Open "select * from library", cn, adOpenDynamic, adLockOptimistic

End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
Exit Sub
Else
MsgBox "Invalid Value"
KeyAscii = False
End If
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 122 Or KeyAscii = 8 Or KeyAscii = 32 Then
Exit Sub
Else
MsgBox "Invalid Value"
KeyAscii = False
End If
End Sub




Private Sub t5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
Exit Sub
Else
MsgBox "Invalid Value"
KeyAscii = False
End If
End Sub



