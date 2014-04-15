VERSION 5.00
Begin VB.Form FrmCharSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Set"
   ClientHeight    =   3120
   ClientLeft      =   3645
   ClientTop       =   3645
   ClientWidth     =   3240
   Icon            =   "FrmCharSet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSet 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Picture         =   "FrmCharSet.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Character                    ASCII Value"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image bg 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4455
      Left            =   0
      Picture         =   "FrmCharSet.frx":652C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "FrmCharSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Sanjul.A.S
'www.sanjulsoft.co.cc
Private Sub cmdInsert_Click()
If InStr(currentForm.txtLib.Text, Left(lstSet.Text, 1)) <> 0 Then
 MsgBox "Same character cannot represent different brightness", vbInformation, ""
Else
 currentForm.txtLib.SelText = Left(lstSet.Text, 1)
End If
End Sub

Private Sub Form_Load()

Dim i As Integer
For i = 1 To 255
 lstSet.AddItem Chr(i) & vbTab & Chr(179) & "  " & i
Next
End Sub



Private Sub lstSet_DblClick()
cmdInsert_Click
End Sub
