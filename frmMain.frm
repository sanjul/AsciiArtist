VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "ASCII Artist 1.0"
   ClientHeight    =   7905
   ClientLeft      =   855
   ClientTop       =   1725
   ClientWidth     =   13050
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   395.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   652.5
   Begin VB.Frame frameMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.TextBox ResultTXT 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   3840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   600
         Width           =   9135
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   375
         Left            =   3840
         Picture         =   "frmMain.frx":FA8A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy to Clipboard"
         Height          =   375
         Left            =   5640
         Picture         =   "frmMain.frx":15FAA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.HScrollBar ResSlid 
         Height          =   255
         Left            =   1080
         Max             =   10
         TabIndex        =   16
         Top             =   4320
         Value           =   8
         Width           =   1935
      End
      Begin VB.CheckBox chkMonospaced 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monospaced"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Timer FlickerTMR 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2760
         Top             =   120
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Browse"
         Height          =   375
         Left            =   120
         Picture         =   "frmMain.frx":1C4CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox Container 
         BackColor       =   &H8000000C&
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2355
         ScaleWidth      =   2475
         TabIndex        =   5
         Top             =   720
         Width           =   2535
         Begin VB.VScrollBar VScroll 
            Height          =   1815
            Left            =   2160
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.HScrollBar HScroll 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2040
            Width           =   2055
         End
         Begin VB.PictureBox pict 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   0
            ScaleHeight     =   143
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   143
            TabIndex        =   6
            Top             =   0
            Width           =   2175
            Begin VB.Shape Area 
               BorderWidth     =   2
               Height          =   495
               Left            =   360
               Top             =   360
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.Image tempimg 
            Height          =   255
            Left            =   120
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
      End
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
         Left            =   1080
         TabIndex        =   4
         Top             =   4800
         Width           =   1575
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
         Left            =   2640
         Picture         =   "frmMain.frx":229EA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4800
         Width           =   375
      End
      Begin VB.TextBox txtHeader 
         Height          =   1095
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmMain.frx":28F0A
         Top             =   5280
         Width           =   1935
      End
      Begin VB.TextBox txtFooter 
         Height          =   1095
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmMain.frx":28F14
         Top             =   6480
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog cmdlgSave 
         Left            =   2280
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cmdlgOpen 
         Left            =   1800
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label selMSG 
         BackStyle       =   0  'Transparent
         Caption         =   "Click and drang on the picture to select a region"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Footer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Header"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ASCII Library"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Precision"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   975
      End
      Begin VB.Image bg 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   7935
         Left            =   3720
         Picture         =   "frmMain.frx":28F4C
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   9975
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu mnuDrawMode 
         Caption         =   "Drawing Mode"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuBrowse 
         Caption         =   "Browse..."
      End
      Begin VB.Menu mnuSelectRgn 
         Caption         =   "Select region"
      End
      Begin VB.Menu mnuDeselect 
         Caption         =   "Deselect"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save ASCII Art"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy ASCII Art"
      End
      Begin VB.Menu mnuTransNotePad 
         Caption         =   "Transfer to notepad"
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Sanjul.A.S
'www.sanjulsoft.co.cc
Public ASCII As New PictToASCII
Public Sub convert()
ASCII.Picture = pict
ASCII.initAsciiLibrary txtLib.Text
ASCII.pixelwidth = (11 - ResSlid.value)
ASCII.imageHeader = txtHeader.Text
ASCII.imageFooter = txtFooter.Text
ASCII.Monospaced = chkMonospaced.value
If Area.Visible = False Then
  ASCII.convert
Else
  ASCII.convert Area.Top, Area.Left, Area.Width, Area.Height
End If
ResultTXT.Text = ASCII.PictureASCII
End Sub
Public Sub openPicture(FileName As String)
tempimg.Picture = LoadPicture(FileName)
Dim wid As Double, ratio As Double
ratio = tempimg.Height / pict.ScaleHeight
wid = tempimg.Width / ratio
pict.ScaleWidth = wid
If wid < 1 Then wid = 1 / wid
pict.Cls
pict.PaintPicture tempimg.Picture, 0, 0, wid, pict.ScaleHeight
End Sub





Private Sub chkMonospaced_Click()
convert
End Sub

Private Sub cmdBrowse_Click()
'To open a picture
On Error Resume Next
With cmdlgOpen
 .Filter = "All types|*.*"
 .ShowOpen
End With
If cmdlgOpen.FileName <> "" Then
openPicture cmdlgOpen.FileName
scrollAlign
convert
End If
End Sub
Sub scrollAlign()
With VScroll
    .Left = Container.ScaleWidth - .Width
    .Top = 0
    .Height = Container.Height - HScroll.Height
    .Max = Container.ScaleHeight
    .value = 0
End With
With HScroll
    .Top = Container.ScaleHeight - .Height
    .Left = 0
    .Width = Container.ScaleWidth - VScroll.Width
    .Max = Container.ScaleWidth
    .value = 0
End With
End Sub
Sub scrollmove()
pict.Top = -VScroll
pict.Left = -HScroll
End Sub


Private Sub FlickerTMR_Timer()
Area.BorderColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
End Sub

Private Sub Form_Load()
On Error Resume Next
openPicture App.Path & "\ASCIIArtist.bmp"
scrollAlign
convert
txtLib.Text = ASCII.AsciiLibrary
End Sub

Private Sub cmdInsertASCII_Click()
FrmCharSet.Show
Set currentForm = Me
End Sub

Private Sub cmdRefresh_Click()

End Sub

Private Sub cmdSave_Click()
With cmdlgSave
  .Filter = "Text File|*.txt"
  .ShowSave
End With
If ASCII.saveASCIIimage(cmdlgSave.FileName) Then _
MsgBox "The ASCII Art is saved!", vbInformation, "ASCII Artist"
End Sub
Private Sub cmdCopy_Click()
ASCII.copyToClipboard
MsgBox "The ASCII Art is copied to the clip board", vbInformation, "ASCII Artist"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim frHei As Integer, frWid As Integer
frHei = frmMain.Height
frWid = frmMain.Width
frameMain.Height = frHei
frameMain.Width = frWid
bg.Height = frHei - 10
bg.Width = frWid - bg.Left - 10
ResultTXT.Height = frHei - ResultTXT.Top - 1000
ResultTXT.Width = frWid - ResultTXT.Left - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll_Scroll()
scrollmove
End Sub



Private Sub mnuAbout_Click()
frmAbt.Show
End Sub


Private Sub mnuBrowse_Click()
cmdBrowse_Click
End Sub

Private Sub mnuCopy_Click()
cmdCopy_Click
End Sub

Private Sub mnuDeselect_Click()
Area.Visible = False
FlickerTMR.Enabled = False
selMSG.Visible = False
convert
End Sub

Private Sub mnuDrawMode_Click()
frmDraw.Show
Me.Hide
End Sub

Private Sub mnuRefresh_Click()
convert
End Sub

Private Sub mnuSave_Click()
cmdSave_Click
End Sub

Private Sub mnuSelectRgn_Click()
FlickerTMR.Enabled = True
Area.Visible = True
selMSG.Visible = True
Area.Move 0, 0, pict.ScaleWidth, pict.ScaleHeight
End Sub

Private Sub mnuTransNotePad_Click()
ASCII.TransferToNotepad "terminal", 5
End Sub

Private Sub pict_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuImage
ElseIf Button = 1 And Area.Visible = True Then
Area.Move x, y, 1, 1
End If
End Sub

Private Sub pict_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And Area.Visible = True Then
Area.Width = Abs(x - Area.Left)
Area.Height = Abs(y - Area.Top)
End If
End Sub

Private Sub pict_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
convert
End Sub

Private Sub ResSlid_Change()
ResultTXT.SetFocus
convert
End Sub



Private Sub txtFooter_Change()
convert
End Sub

Private Sub txtHeader_Change()
convert
End Sub

Private Sub txtLib_Change()
convert
End Sub

Private Sub vScroll_Scroll()
scrollmove
End Sub


