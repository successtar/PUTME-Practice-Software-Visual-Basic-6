VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFF00&
   Caption         =   "Result!!!!!"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11805
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
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
      Height          =   735
      Left            =   8640
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CORRECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "google.com/+osanyinpejuhammed www.facebook.com/successtar1 www.facebook.com/am.goal @successtar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   5880
      TabIndex        =   12
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "+2347061855688"
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
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " Developed by :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Osanyinpeju Hammed Olalekan  Successtar         a.k.a          Walestar  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1815
      Left            =   2280
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   240
      Picture         =   "Form5.frx":014A
      Top             =   3360
      Width           =   1920
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   8415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   6615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "out of 100"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Label1.Caption < 60 Then
intresponse = MsgBox("You have to go over it again, your score is too low to view the correction. You must score at least 60, before you can view the correction.", vbOK + vbQuestion)
Else
Me.Hide
Form3.Command41.Visible = True
Form6.Show
Form2.Show
Form3.Show
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Form4.Hide
o = a1 + a2 + a3 + a4 + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13 + a14 + a15 + a16 + a17 + a18 + a19 + a20 + a21 + a22 + a23 + a24 + a25 + a26 + a27 + a28 + a29 + a30 + a31 + a32 + a33 + a34 + a35 + a36 + a37 + a38 + a39 + a40 + a41 + a42 + a43 + a44 + a45 + a46 + a47 + a48 + a49 + a50

Label1.Caption = o
If o >= 70 Then
Label5.Caption = "EXCELLENT"
Label6.Caption = "Wow!!! what a great result,keep it up."
ElseIf o < 70 Then
Label5.Caption = "GOOD"
Label6.Caption = "Nice result but you can still do better. You may go over it again before going for the correction for you desire."
End If
If o <= 60 Then
Label5.Caption = "CREDIT"
Label6.Caption = "Average result, work harder. Ensure you go over it again before going for the correction."
End If
If o < 50 Then
Label5.Caption = "PASS"
Label6.Caption = "Poor result but you can still do better.  Ensure you go over it again before going for the correction."
End If
If o < 40 Then
Label5.Caption = "FAIL"
Label5.ForeColor = vbRed
Label1.ForeColor = vbRed
Label6.ForeColor = vbRed
Label3.ForeColor = vbRed
Label6.Caption = "Very poor but don't give up,you can still do better. Ensure you go over it again before going for the correction."

End If






End Sub

Private Sub Label12_Click()
Me.Hide
Form3.Command41.Visible = True
Form6.Show
Form2.Show
Form3.Show
End Sub
