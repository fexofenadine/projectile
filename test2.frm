VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Physical Constants"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "test2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "Unavailable"
      Top             =   820
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "test2.frx":030A
      Left            =   3000
      List            =   "test2.frx":0314
      TabIndex        =   1
      Text            =   "Earth - 9.8"
      Top             =   190
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Wind opposing projectile (negative for other direction):"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Air viscosity:"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Acceleration due to gravity (m / s^2):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Select Case Form2.Combo1.Text
Case "Earth - 9.8"
g = 9.8
Case "Mars - 3.4"
g = 3.4
Case Else
g = (Form2.Combo1.Text)
End Select
'visc = Abs(Form2.Text1.Text)
wind = Form2.Text2.Text
Form2.Visible = False
If wind >= 0 Then Form1.Label4.Caption = "Wind = " & wind & " <<<<" Else Form1.Label4.Caption = "Wind = " & Abs(wind) & " >>>>"
Form1.Label7.Caption = "Gravity = " & g
'Form1.Label8.Caption = "Viscosity = " & visc
End Sub

Private Sub Command2_Click()
Form2.Visible = False
End Sub

