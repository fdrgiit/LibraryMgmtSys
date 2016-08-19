VERSION 5.00
Begin VB.Form SPLASH 
   BackColor       =   &H00000000&
   Caption         =   "STARTUP"
   ClientHeight    =   11745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13830
   LinkTopic       =   "Form5"
   ScaleHeight     =   13890
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   4035
      Left            =   7920
      TabIndex        =   3
      Top             =   5280
      Width           =   7680
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6360
         Top             =   2880
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2880
         Width           =   6855
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Constantia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   2880
         Width           =   6735
      End
      Begin VB.Image imgLogo 
         Height          =   1905
         Left            =   240
         Picture         =   "SplashScreen.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "LIBRARY MANAGEMENT system"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   4620
      End
      Begin VB.Shape Shape1 
         Height          =   1815
         Left            =   2880
         Top             =   600
         Width           =   4575
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Left            =   240
         Top             =   600
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   -120
         X2              =   9360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   6975
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   6495
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   10320
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   8520
      TabIndex        =   1
      Top             =   10320
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800080&
      Caption         =   $"SplashScreen.frx":1991
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   5055
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   10575
   End
   Begin VB.Image Image1 
      Height          =   9030
      Left            =   240
      Picture         =   "SplashScreen.frx":1AE9
      Top             =   120
      Width           =   7425
   End
End
Attribute VB_Name = "SPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shape4.Width = 0
Timer1.Enabled = True
Label1.Caption = Date
Label2.Caption = Time
End Sub

Private Sub Image2_Click()

End Sub

Private Sub Timer1_Timer()
Timer1.Interval = Rnd * 200
Shape4.Width = Shape4.Width + 80
Select Case Shape4.Width
    Case 0 To 800
        Label4.Caption = "INITIALIZING APPLICATION . . ."
    Case 801 To 1600
        Label4.Caption = "CHECKING DATABASE . . ."
    Case 1601 To 4000
        Label4.Caption = "CONNECTING DATABASE . . ."
    Case 4001 To 5200
        Label4.Caption = "LOADING FORMS . . ."
    Case 5201 To 6735
        Label4.Caption = "FINISHING . . ."
End Select
If Shape4.Width = Shape3.Width Then
    Timer1.Enabled = False
LIBRARY.Show
MsgBox "WELCOME TO LIBRARY OF CONGRESS.", vbOKOnly, "LIBRARY MANAGEMENT SYSTEM"
Unload Me
End If
End Sub
