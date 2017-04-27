VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Our Savior!!!!!!!!!!!!!!!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      Picture         =   "test4.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Cancel          =   -1  'True
         Caption         =   "k"
         Height          =   375
         Left            =   1320
         MaskColor       =   &H0000FF00&
         TabIndex        =   2
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "YOU DESTROYED THE BOMB AND SAVED THE EARTH FROM A HORRIBLE FATE!!!! HOW CAN WE EVER REPAY YOU?"
         ForeColor       =   &H0000C000&
         Height          =   975
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "YOU KICKED ASS!!!!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Visible = False
End Sub
