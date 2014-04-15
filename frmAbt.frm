VERSION 5.00
Begin VB.Form frmAbt 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ASCII Artist"
   ClientHeight    =   3090
   ClientLeft      =   5160
   ClientTop       =   4530
   ClientWidth     =   5115
   Icon            =   "frmAbt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      Picture         =   "frmAbt.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Freeware, GPL"
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sanjul;s ASCII Artist"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.0"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by,"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sanjul.A.S"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ssanjul@gmail.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label wsiteLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "www.sanjulsoft.co.cc"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image bg 
      Height          =   3135
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Sanjul.A.S
'www.sanjulsoft.co.cc
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
imgIcon.Picture = frmMain.Icon
bg.Picture = frmMain.bg.Picture
'bg.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

