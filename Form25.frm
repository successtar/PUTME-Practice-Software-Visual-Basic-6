VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PUTME practice cbt questions"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11220
   Icon            =   "Form25.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   11220
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Achiever"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Walestar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   7215
         Begin VB.CommandButton Command1 
            Caption         =   "START"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3000
            TabIndex        =   6
            Top             =   8160
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Quotes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   495
            Left            =   2400
            TabIndex        =   13
            Top             =   120
            Width           =   5055
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form25.frx":014A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1335
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   5400
            Width           =   6495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "   The hardest tests in life is the patience to wait for the right moment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   4440
            Width           =   6495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "    Failures are necessary on the part of success.Just learn from it and move on."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Index           =   3
            Left            =   360
            TabIndex        =   10
            Top             =   3360
            Width           =   6495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form25.frx":0205
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Index           =   2
            Left            =   360
            TabIndex        =   9
            Top             =   1920
            Width           =   6495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "  A man is not finished when he is defeated.He is finished when he quits."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Index           =   1
            Left            =   480
            TabIndex        =   8
            Top             =   7080
            Width           =   6495
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form25.frx":0313
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   975
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   840
            Width           =   6495
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6000
         TabIndex        =   4
         Top             =   8520
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LAUNCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   8520
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form25.frx":03E6
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   7815
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SUCCESSTAR !!!!!!!"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1800
      TabIndex        =   14
      Top             =   0
      Width           =   9735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -4920
      Picture         =   "Form25.frx":05D0
      Top             =   -120
      Width           =   24000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "You have 30 minutes to answer 50 questions. Goodluck"
Me.Hide
Form4.Show
Form3.Show
Form2.Show

End Sub

Private Sub Command2_Click()
Frame2.Visible = True
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Label1.Caption = "1"


End Sub

Private Sub Timer1_Timer()
If Label1.Caption = 0 Then
Frame1.Visible = True
Else
Label1.Caption = Label1.Caption - 1
End If

End Sub
