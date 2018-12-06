VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14130
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   8580
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   11520
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   7560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Width           =   7815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SUCCESSTAR nonI!!!"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   7800
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "  cORRECTION"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " 01E 02A 03E 04E 05A 06D 07E 08B 09B 10B 11A 12B 13A 14B 15B 16B 17D 18D 19C 20D 21B 22A 23C 24D 25B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7935
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " 26C 27A 28C 29B 30B 31A 32B 33B 34C 35D 36B 37B 38B 39D 40E 41B 42B 43D 44C 45B 46D 47B 48B 49B 50C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   7935
      Index           =   0
      Left            =   13200
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_click()
Form2.Show
Form3.Show

End Sub

Private Sub Label1_Click(Index As Integer)
Form2.Show
Form3.Show
End Sub

Private Sub Label2_Click(Index As Integer)
Form2.Show
Form3.Show
End Sub

Private Sub Label4_Click()
Form2.Show
Form3.Show
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Time
End Sub
