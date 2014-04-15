VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDraw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   2010
   ClientTop       =   2880
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11700
   Begin VB.CheckBox chkMonospaced 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Monospaced"
      Height          =   255
      Left            =   3120
      TabIndex        =   21
      Top             =   240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   720
      ScaleHeight     =   2175
      ScaleWidth      =   3855
      TabIndex        =   12
      Top             =   4200
      Width           =   3855
      Begin VB.TextBox txtLib 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdInsertASCII 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   4.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Picture         =   "frmDraw.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1560
         Width           =   375
      End
      Begin VB.HScrollBar ResSlid 
         Height          =   255
         Left            =   1680
         Max             =   10
         TabIndex        =   16
         Top             =   1080
         Value           =   8
         Width           =   2055
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         Picture         =   "frmDraw.frx":6520
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.HScrollBar brushWSlid 
         Height          =   255
         LargeChange     =   2
         Left            =   1680
         Max             =   20
         Min             =   1
         TabIndex        =   13
         Top             =   480
         Value           =   4
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Precision"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ASCII Library"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Brush Width :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.PictureBox mind 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   1
      Left            =   5640
      ScaleHeight     =   237
      ScaleMode       =   0  'User
      ScaleWidth      =   237
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox mind 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Index           =   0
      Left            =   5280
      ScaleHeight     =   237
      ScaleMode       =   0  'User
      ScaleWidth      =   237
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   615
      Begin VB.OptionButton OptPick 
         Height          =   375
         Left            =   120
         Picture         =   "frmDraw.frx":CA40
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton optText 
         Height          =   375
         Left            =   120
         Picture         =   "frmDraw.frx":D35E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox clr 
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   0
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   3600
         Width           =   495
      End
      Begin MSComDlg.CommonDialog cmdlgclr 
         Left            =   120
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton optFill 
         Height          =   375
         Left            =   120
         Picture         =   "frmDraw.frx":DA30
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   375
      End
      Begin VB.OptionButton optLine 
         Height          =   375
         Left            =   120
         Picture         =   "frmDraw.frx":E252
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   375
      End
      Begin VB.OptionButton optEraser 
         Height          =   375
         Left            =   120
         Picture         =   "frmDraw.frx":EA4C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
      Begin VB.OptionButton optBrush 
         Height          =   375
         Left            =   120
         Picture         =   "frmDraw.frx":F34E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   375
      End
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.PictureBox Paper 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   10
      FillStyle       =   0  'Solid
      Height          =   3615
      Left            =   840
      MousePointer    =   2  'Cross
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      Begin VB.Label preCapt 
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         Height          =   10455
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   10215
      End
      Begin VB.Line ln 
         Visible         =   0   'False
         X1              =   48
         X2              =   136
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Shape MPointer 
         Height          =   375
         Left            =   2160
         Shape           =   2  'Oval
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ASCII Art:"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Make your drawings here :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image bg 
      Height          =   495
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu mnuConvMode 
         Caption         =   "Conversion mode"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuSaveASCII 
         Caption         =   "Save ASCII Art"
      End
      Begin VB.Menu mnuCopyASCII 
         Caption         =   "Copy ASCII Art"
      End
      Begin VB.Menu mnuTransNotePad 
         Caption         =   "Transfer to notepad"
      End
      Begin VB.Menu mnuSavePict 
         Caption         =   "Save Picture"
      End
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Sanjul.A.S
'www.sanjulsoft.co.cc
Private Declare Sub _
  ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
                            ByVal y As Long, ByVal crColor As Long, _
                            ByVal wFillType As Long)


Private asc As New PictToASCII
Public Sub convert()
asc.Monospaced = chkMonospaced.value
 asc.pixelwidth = (11 - ResSlid.value)
  asc.initAsciiLibrary txtLib.Text
  asc.convert
  txt.Text = asc.PictureASCII
End Sub

Private Sub chkMonospaced_Click()
convert
End Sub

Private Sub clr_Click()
cmdlgclr.ShowColor
clr.BackColor = cmdlgclr.Color
Paper.FillColor = cmdlgclr.Color
Paper.ForeColor = cmdlgclr.Color
End Sub

Private Sub cmdClear_Click()
Paper.Cls
convert
End Sub

Private Sub cmdInsertASCII_Click()
FrmCharSet.Show
Set currentForm = Me
End Sub

Private Sub Form_Load()
Me.Caption = "Drawing pad : " & frmMain.Caption
Me.Icon = frmMain.Icon
bg.Picture = frmMain.bg.Picture
bg.Move 0, 0
asc.Picture = Paper
txtLib.Text = asc.AsciiLibrary
asc.initAsciiLibrary frmMain.txtLib.Text
asc.imageFooter = frmMain.txtFooter.Text
asc.imageHeader = frmMain.txtHeader.Text
convert
End Sub

Private Sub Form_Resize()
bg.Width = ScaleWidth
bg.Height = ScaleHeight
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub



Private Sub mnuConvMode_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub mnuCopyASCII_Click()
asc.copyToClipboard
End Sub

Private Sub mnuRedo_Click()
Paper.PaintPicture mind(1).Image, 0, 0
convert
mnuRedo.Enabled = False
mnuUndo.Enabled = True
End Sub

Private Sub mnuSaveASCII_Click()
With frmMain.cmdlgSave
   .Filter = "Text File|*.txt"
   .ShowSave
   If asc.saveASCIIimage(.FileName) Then _
MsgBox "The ASCII Art is saved!", vbInformation, "ASCII Artist"
End With

End Sub

Private Sub mnuSavePict_Click()
With frmMain.cmdlgSave
   .Filter = "JPEG image|*.JPEG"
   .ShowSave
   If asc.saveSourcePicture(.FileName) = True Then
    MsgBox "Picture saved!", vbInformation, "ASCII Artist"
   End If
End With
End Sub

Private Sub mnuTransNotePad_Click()
asc.TransferToNotepad "terminal", 5
End Sub

Private Sub mnuUndo_Click()
mind(1).PaintPicture Paper.Image, 0, 0
Paper.PaintPicture mind(0).Image, 0, 0
convert
mnuUndo.Enabled = False
mnuRedo.Enabled = True
End Sub









Private Sub Paper_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  mind(0).Cls
  mind(0).PaintPicture Paper.Image, 0, 0
  mnuUndo.Enabled = True
  mnuRedo.Enabled = False
If optLine = True And Button = 1 Then
  ln.Visible = True
  ln.BorderColor = clr.BackColor
  MPointer.Visible = False
  ln.X1 = x
  ln.Y1 = y
  ln.X2 = x
  ln.Y2 = y
ElseIf optFill.value = True Then
 ExtFloodFill Paper.hdc, x, y, Paper.Point(x, y), 1
End If


End Sub

Private Sub Paper_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Paper.DrawWidth = brushWSlid.value
MPointer.Move x - (MPointer.Width / 2), y - (MPointer.Height / 2)
Dim colr As Long
If optEraser.value = True Then
  colr = vbWhite
Else
  colr = clr.BackColor
End If
If Button = 1 Then
   If optBrush = True Or optEraser = True Then
       Paper.Line -(x, y), colr
   ElseIf optLine = True Then
       ln.BorderWidth = Paper.DrawWidth
       ln.X2 = x
       ln.Y2 = y
   ElseIf optText = True Then
       currX = x
       currY = y
       frmAddText.txt_Change
   End If
Else
  Paper.CurrentX = x
  Paper.CurrentY = y

End If
End Sub


Private Sub Paper_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

 If ln.Visible = True Then
  ln.Visible = False
  Paper.Line (ln.X1, ln.Y1)-(ln.X2, ln.Y2)
  MPointer.Visible = True
 ElseIf optText = True Then
  currX = x
  currY = y
  preCapt.Visible = True
  preCapt.ForeColor = Paper.ForeColor
  MPointer.Visible = False
  frmAddText.Show
 ElseIf OptPick = True Then
  clr.BackColor = Paper.Point(x, y)
  Paper.FillColor = clr.BackColor
  Paper.ForeColor = clr.BackColor
  optBrush = True
 End If
 
  convert

End Sub

Private Sub ResSlid_Change()
txt.SetFocus
convert
End Sub


Private Sub txtLib_Change()
convert
End Sub
