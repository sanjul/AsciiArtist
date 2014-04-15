VERSION 5.00
Begin VB.Form frmAddText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Text"
   ClientHeight    =   2265
   ClientLeft      =   3990
   ClientTop       =   4335
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4335
   Begin VB.CommandButton decSize 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Picture         =   "frmAddText.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton incSize 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Picture         =   "frmAddText.frx":6520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      Picture         =   "frmAddText.frx":CA40
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAddText.frx":12F60
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image bg 
      Height          =   2535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmAddText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Sanjul.A.S
'www.sanjulsoft.co.cc
Dim dest As Object
Dim Preview As Object


Private Sub cmbFont_Click()
txt.Font = cmbFont.Text
txt_Change
End Sub

Private Sub cmdOK_Click()
dest.CurrentX = currX
dest.CurrentY = currY
dest.Print txt.Text

Unload Me
Preview.Caption = ""
Preview.Visible = False
frmDraw.convert
frmDraw.MPointer.Visible = True
End Sub



Private Sub decSize_Click()
txt.FontSize = txt.FontSize - 5
txt_Change
End Sub

Private Sub Form_Load()
Set dest = frmDraw.Paper
Set Preview = frmDraw.preCapt
For i = 0 To Screen.FontCount - 1
 cmbFont.AddItem Screen.Fonts(i), i
Next
  cmbFont.Text = txt.FontName
  
bg.Picture = frmMain.bg.Picture
txt_Change
txt.SelStart = 0
txt.SelLength = Len(txt.Text)

End Sub

Private Sub incSize_Click()
txt.FontSize = txt.FontSize + 5
txt_Change
End Sub



Public Sub txt_Change()

dest.Font = txt.FontName
dest.FontSize = txt.FontSize
dest.FontBold = txt.FontBold
dest.FontItalic = txt.FontItalic
Preview.ForeColor = dest.ForeColor
Preview.Font = txt.FontName
Preview.FontBold = txt.FontBold
Preview.FontItalic = txt.FontItalic
Preview.FontSize = txt.FontSize
Preview.Caption = txt.Text
Preview.Move currX, currY
End Sub
