VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Highscores"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Highscores"
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.Label Label6 
         BackColor       =   &H80000007&
         Caption         =   "Name"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000007&
         Caption         =   "Name"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         Caption         =   "Name"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000007&
         Caption         =   "0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempname1 As String
Dim tempname2 As String
Dim tempname3 As String

Private Sub Command1_Click()
Unload Form2
End Sub

Private Sub Form_Load()
Open App.Path & "\score.mezz" For Input As #1
Input #1, tempname1, tempscore1, tempname2, tempscore2, tempname3, tempscore3
If score > tempscore1 Then
tempscore1 = score
tempname1 = InputBox("Please enter your name:", "Name")
End If
If score < tempscore1 And score > tempscore2 Then
tempscore2 = score
tempname2 = InputBox("Please enter your name:", "Name")
End If
If score < tempscore1 And score < tempscore2 And score > tempscore3 Then
tempscore3 = score
tempname3 = InputBox("Please enter your name:", "Name")
End If
Label1.Caption = tempscore1
Label2.Caption = tempscore2
Label3.Caption = tempscore3
Label4.Caption = tempname1
Label5.Caption = tempname2
Label6.Caption = tempname3
Close #1
Open App.Path & "\score.mezz" For Output As #1
Write #1, tempname1, tempscore1, tempname2, tempscore2, tempname3, tempscore3
Close #1
End Sub
