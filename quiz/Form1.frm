VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quiz"
   ClientHeight    =   1545
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Begin"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   480
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Answer"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu highscores 
      Caption         =   "Highscores"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim question As String
Dim answer As String
Dim amountanswered As Double
Dim correctanswers As Double
Private Sub Command1_Click()
checkanswer 'Calls the checkanswer function
End Sub
Private Sub Command2_Click()
amountanswered = 0
correctanswers = 0
generatequestion 'calls the generatequestion function
Label2.Caption = "60"
Timer1.Enabled = True
Command2.Visible = False
End Sub
Private Sub Command4_Click()
End
End Sub
Private Sub Form_Load()
App.TaskVisible = False
Label2.Caption = "60"
Randomize 'Need this so that numbers produced are different each time the program is loaded
End Sub
Private Sub highscores_Click()
Form2.Show
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKeyReturn 'If return key is pressed then question is answered and checked
checkanswer
End Select
End Sub
Private Sub Timer1_Timer()
If Label2.Caption = 0 Then
MsgBox "Well done you scored " & correctanswers & " out of " & amountanswered & " in 60 seconds"
Timer1.Enabled = False
score = correctanswers
Command2.Visible = True
Form2.Show
Else
Label2.Caption = Label2.Caption - 1
End If
End Sub
Public Function checkanswer()
If Text1.Text = "" Then
MsgBox "At least make an attempt at an answer", vbOKOnly, "Help"
Else
Text1.Text = UCase(Text1.Text) 'Makes all text entered in the textbox upper case
If Text1.Text = answer Then 'checks if the text in the text box = the answer selected from answers.mezz
correctanswers = correctanswers + 1 'adds 1 to the amount of correctanswers
Else
End If
Text1.Text = ""
generatequestion 'calls the generatequestion function
amountanswered = amountanswered + 1
End If
End Function
Public Function generatequestion()
Dim R As Integer
Open App.Path & "\Questions.mezz" For Input As #1 'opens questions.mezz to use as input
Open App.Path & "\answers.mezz" For Input As #2 'opens answers.mezz to use as input
R = Int(Rnd * 72) 'calculates a random number this is then used to select which question and which question is generated from the files
For i = 0 To R
Line Input #1, question 'enters a question from questions.mezz
Line Input #2, answer
Next i
Label1.Caption = question
Close #1
Close #2
End Function
