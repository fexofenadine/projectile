VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "About Projectile Motion"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Version 1.0.0"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.Label Label1 
         Caption         =   "Technopanda's Projectile Motion Demonstrator.  "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         Caption         =   $"test5.frx":0000
         ForeColor       =   &H00000000&
         Height          =   930
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   3885
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "K"
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   120
      Picture         =   "test5.frx":00B4
      Top             =   2280
      Width           =   3000
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Visible = False
End Sub
