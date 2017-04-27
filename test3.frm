VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gameplay Options"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "test3.frx":0000
      Left            =   2400
      List            =   "test3.frx":0013
      TabIndex        =   1
      Text            =   "Standard (450)"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Difficulty setting (lower numbers are more difficult):"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Form3.Combo1.Text
Case "Easy Az Hell (6000)"
tgwidth = 6000
Case "Easy (3000)"
tgwidth = 3000
Case "Standard (450)"
tgwidth = 450
Case "Hard (200)"
tgwidth = 200
Case "Hard Az Hell (45)"
tgwidth = 45
Case Else
tgwidth = Abs(Form3.Combo1.Text)
End Select
Form3.Visible = False
Form1.Image2.Width = tgwidth
End Sub

Private Sub Command2_Click()
Form3.Visible = False
End Sub
