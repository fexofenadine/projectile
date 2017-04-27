VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projectile Motion"
   ClientHeight    =   6330
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7725
   Icon            =   "test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Kalkulayshunz"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   7455
      Begin VB.Label Label5 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   570
         Width           =   7215
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit  : ("
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   5280
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "45"
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Text            =   "200"
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   120
      Picture         =   "test.frx":030A
      ScaleHeight     =   3555
      ScaleWidth      =   7395
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Viscosity"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gravity"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Wind"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5400
         TabIndex        =   11
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   450
         Left            =   3000
         Picture         =   "test.frx":55D74
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   450
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start  : )"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Tries : 0"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   3890
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Hit the bomb with your diffusion grenade to prevent the earth from being destroyed!!!!!!"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Velocity (m/s):"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Angle (degrees):"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4080
      Picture         =   "test.frx":5687E
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnuPhysics 
      Caption         =   "&Physical Constants"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Gameplay Options"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call shoot
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Call immediate
End Sub

Private Sub mnuAbout_Click()
Form5.Visible = True
End Sub

Private Sub mnuOptions_Click()
Form3.Visible = True
End Sub

Private Sub mnuPhysics_Click()
Form2.Visible = True
End Sub
