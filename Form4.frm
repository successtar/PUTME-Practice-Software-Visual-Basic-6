VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11865
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   8460
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   480
      Top             =   4680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
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
      Left            =   9840
      TabIndex        =   3
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   360
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   6480
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   7680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "2"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "click"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2175
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "59"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMINATION IN PROGRESS !!!"
      BeginProperty Font 
         Name            =   "Xirod"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
intresponse = MsgBox("You are about to end the Exam.Press Ok to continue or press Cancel to return!!", vbOKCancel + vbQuestion, "End PUTME practice question")
If intresponse = vbOK Then
Timer2.Enabled = False
Form3.Hide
Form2.Hide
Form5.Show
ElseIf intresponse = vbCancel Then
Form2.Show
End If


End Sub

Private Sub Form_Load()
Form4.BackColor = vbYellow


End Sub

Private Sub Form_click()
Form3.Show
Form2.Show



End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time

End Sub

Private Sub Timer2_Timer()

Label3.Caption = Label3.Caption - 1

If Label3.Caption = -1 Then
Label6.Caption = Label6.Caption - 1
Label3.Caption = 59
Label8.Visible = False
End If

If Label6.Caption = 9 Then
Label9.Visible = True
End If
If Label3.Caption = 9 Then
Label8.Visible = True

End If


If (Label6.Caption = 4 And Label3.Caption = 59) Then
Label6.ForeColor = &HFF&
Label5.ForeColor = &HFF&
Label3.ForeColor = &HFF&
Label8.ForeColor = &HFF&
Label9.ForeColor = &HFF&
MsgBox ("You have less than 5 minutes for the test !!")
End If


If Label6.Caption = -1 Then
Timer2.Enabled = False
Label3.Caption = "00"
Label6.Caption = "0"
MsgBox ("Your time is up. You have used 30min for this exam")
MsgBox "You have now completed the exam, press Ok to go to the homepage", vbOKOnly + vbInformation, "PUTME prractice exam completed"
Form3.Hide
Form2.Hide
Form5.Show

End If

End Sub

Private Sub Timer3_Timer()
Label7.Caption = Label7.Caption - 1
If Label7.Caption = 0 Then
Label7.Caption = 2
End If


If Label6.Caption = 0 Then


If Label7.Caption = 1 Then

Label3.Visible = False
Label5.Visible = False
Label6.Visible = False

Label8.Visible = False
Label9.Visible = False

Else

Label3.Visible = True
Label5.Visible = True
Label6.Visible = True

If Label3.Caption <= 9 Then
Label8.Visible = True
End If

If Label6.Caption <= 9 Then
Label9.Visible = True
End If


End If
End If
End Sub
