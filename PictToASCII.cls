VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictToASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'code by Sanjul.A.S
'www.sanjulsoft.co.cc
'Global CharBright As String 'Library for characters
Private ASCIIimage As String
Private pict As PictureBox
Private PixWidth As Integer
Private Header As String
Private Footer As String
Private CharBright As String
Private btwChar As String
Public Monospaced As Boolean
Private Function GetRed(colour As Long) As Long
'Get the Red component from a color
GetRed = (colour) Mod 256
End Function
Private Function getGreen(colour As Long) As Long
'Get the Green component from a color
getGreen = (colour \ 256) Mod 256
End Function
Private Function Getblue(colour As Long) As Long
'Get the Green component from a color
Getblue = (colour \ 65536) Mod 256
End Function

Public Sub initAsciiLibrary(Optional str As String)
'Assign a  character library to the
'CharBright variable in the order of brightness
If str <> "" Then
  CharBright = str
Else
              'Chr(178) & Chr(177) & Chr(176) &
  CharBright = "#@O=~-,." & Chr(255)
End If
End Sub
Public Property Get AsciiLibrary() As String
If CharBright = "" Then initAsciiLibrary
AsciiLibrary = CharBright

End Property
Private Function getCBright(ByVal value As Long) As String
'To retrive the desired character from the CharBright
value = value + 1
getCBright = Mid(CharBright, value, 1)
End Function

Private Function setChar(colour As Long) As String
'Returns the character after calculating the brightness of color

Dim avg As Integer, dv As Integer, char As String * 1
dv = 255 / Len(CharBright) 'It is to get a devider for the brightness
avg = (GetRed(colour) + Getblue(colour) + getGreen(colour)) / 3
    'Avg of Red,Green, and Blue component colors gives the brightness
char = getCBright(Int(avg / dv))   'Avg/dv gives the region of brightness
setChar = char 'returns result
End Function

Public Property Let Picture(ByRef pic As PictureBox)
'Link with the picture box
Set pict = pic
If CharBright = "" Then initAsciiLibrary
If pixelwidth = 0 Then pixelwidth = 1
End Property

Public Property Get PictureASCII()
'Return resultant ascii image
PictureASCII = ASCIIimage
End Property
Public Property Let pixelwidth(ByVal Pixelwid As Integer)
'pixel width determines the resolution of picture
If Pixelwid = 0 Then
 PixWidth = 1
Else
 PixWidth = Pixelwid
End If
End Property
Public Property Get pixelwidth() As Integer
pixelwidth = PixWidth
End Property

Public Property Let imageHeader(ByVal hed As String)
Header = hed
End Property
Public Property Get imageHeader() As String
imageHeader = Header
End Property
Public Property Let imageFooter(ByVal foot As String)
Footer = foot
End Property
Public Property Get imageFooter() As String
imageFooter = Footer
End Property

Public Sub convert(Optional Top As Long, Optional Left As Long, Optional Width As Long, Optional Height As Long)
If Width = 0 Then Width = pict.ScaleWidth - Left
If Height = 0 Then Height = pict.ScaleHeight - Top
If CharBright = "" Then initAsciiLibrary
If pixelwidth = 0 Then pixelwidth = 1
If Monospaced = True Then
 btwChar = ""
Else
 btwChar = "_"
End If
ASCIIimage = ""
ASCIIimage = Header & vbCrLf
Dim x As Long, y As Long
For y = Top To Height - 1 Step PixWidth
   DoEvents
   For x = Left To Width - 1 Step PixWidth
      ASCIIimage = ASCIIimage & btwChar & setChar(pict.Point(x, y))
   Next
   ASCIIimage = ASCIIimage & vbCrLf
Next
ASCIIimage = ASCIIimage & Footer
End Sub
Public Function saveASCIIimage(FileName As String) As Boolean
If FileName <> "" Then
 Open FileName For Output As #1
 Print #1, ASCIIimage
 Close #1
 saveASCIIimage = True
Else
 saveASCIIimage = False
End If
End Function
Public Function saveSourcePicture(FileName As String) As Boolean
If FileName <> "" Then
  SavePicture pict.Image, FileName
  saveSourcePicture = True
Else
  saveSourcePicture = False
End If
End Function
Public Sub copyToClipboard()
Clipboard.Clear
Clipboard.SetText ASCIIimage
End Sub
Public Sub TransferToNotepad(fntname As String, FontSize As Integer)
copyToClipboard
Shell "notepad.exe", vbNormalFocus
SendKeys "^v"
SendKeys "%of"
SendKeys fntname
SendKeys vbTab
SendKeys "bold"
SendKeys vbTab
SendKeys CStr(FontSize)
SendKeys "~"
End Sub
