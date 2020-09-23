VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Filter: Cartoon = 1    Photo ~ 100             "
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   0
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin MSComCtl2.FlatScrollBar scrllFilter 
         Height          =   375
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   $"Form1.frx":0000
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         LargeChange     =   50
         Max             =   500
         Orientation     =   8323073
         SmallChange     =   10
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Speed"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   3135
      Begin VB.PictureBox picProgress 
         Height          =   135
         Left            =   1800
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   77
         TabIndex        =   6
         ToolTipText     =   "Progress Bar"
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Speed of display. May be changed during transition."
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         LargeChange     =   3000
         Max             =   30000
         Orientation     =   8323073
         SmallChange     =   300
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "      Paint Options"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdGo 
         Caption         =   "No Y"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "'Paint picture, pixel selection according to bandpass filter, two passes, No Luinence, Yes Luminemce"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "BW+C"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   10
         ToolTipText     =   "'Paint picture, pixel selection according to bandpass filter, three passes, B&W+B, B&W+GB, B&W+RGB"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Color"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "'Paint picture, pixel selection according to bandpass filter"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "RGB"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   3
         ToolTipText     =   "'Paint picture, pixel selection according to bandpass filter, three passes, R, RG, RGB"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "BW"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "'Paint picture, pixel selection according to bandpass filter, two passes, B&W, Color"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9600
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   840
      Width           =   615
      Begin VB.Image Image1 
         Height          =   255
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuBackGround 
         Caption         =   "White Background"
         Index           =   0
      End
      Begin VB.Menu mnuBackGround 
         Caption         =   "Gray Background"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuBackGround 
         Caption         =   "Black Background"
         Index           =   2
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'**********************************************************
'**********************************************************
'***                         *                        ***
'***                         *                        ***
'***                    * * * * * *                   ***
'***                         *                        ***
'***                         *                        ***
'***                         *                        ***
'***                         *                        ***
'***                         *                        ***
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************


'**********************************************************
'*****************Summary**********************************
'**********************************************************
'Basically what this program does is paint a picture
'with pixel selection according to a user entered filter.
'Also, when we do the actual painting we do it with different
'techniques.
'For example let's say our user has opened a picture and
'entered a value of one for the filter. A one is what works
'best for cartoon type pictures with solid colors. For
'photos, larger numbers work better.
'Ok, first our program will select a pixel at random from
'the picture in the array memory bank. It will paint this
'pixel at the proper location on our screen.
'It will paint it either the same color or some variation
'depending on which paint option/button was clicked.
'I have dedicated a seperate module for each option to
'simplify tracking the program flow, with the main structure
'difference in each being a tiny function. This function is
'what will determine how the painting differences will be done.
'Let's continue with the summary;
'It will then
'test the eight immediately adjacent pixels to see if any
'are the exact same color and paint them too. Then those
'ones just painted will be tested of their surrounding
'pixels and so forth. This will continue until no adjacent
'pixels match the filter criteria.
'Now another pixel is selected at random, tested to see if
'this pixel has already been painted and if not go through
'the adjacent pixel checking routine. This will continue
'until each and every pixel has been checked.
'On larger than one filter settings instead of tracing an
'exact color the program will trace the exact color and
'colors close to this color according to the filter setting.
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
Option Explicit
     
Private Sub Form_Load()
'Set all dimensions
Dim z As Long
'cmdGo and Open Menu enabled only when necessary
mnuFileOpen.Enabled = True
cmdGo(0).Enabled = False
cmdGo(1).Enabled = False
cmdGo(2).Enabled = False
cmdGo(3).Enabled = False
cmdGo(4).Enabled = False
'Set all dimensions
frmMain.ScaleMode = 3
frmMain.Height = 7215
frmMain.Width = 8900
picMain.ScaleMode = 3
picMain.Left = 8
picMain.Picture = picMain.Image
picMain.AutoRedraw = True
    'Speed
    FlatScrollBar1.Value = 15000 'Center
    z = 30000 - FlatScrollBar1.Value
    lngDelay = z * 2 'Set painting speed according to ScrollBar
    Frame3.Caption = "Speed " & Str(Int(100 - (z / 300))) & "%"
'Filter
scrllFilter.Value = 100
Text1.Text = Str(scrllFilter.Value)
intColorFilter = Val(Text1.Text)

End Sub

Private Sub mnuBackGround_Click(Index As Integer)
'Set our PictureBox's background
Dim intIndex As Integer
For intIndex = 0 To 2
mnuBackGround(intIndex).Checked = False
Next intIndex

Select Case Index
    Case 0
    picMain.BackColor = vbWhite
    picMain.Picture = LoadPicture("")
    Case 1
    picMain.BackColor = vbButtonFace
    picMain.Picture = LoadPicture("")
    Case 2
    picMain.BackColor = vbBlack
    picMain.Picture = LoadPicture("")
End Select
mnuBackGround(Index).Checked = True
End Sub

Private Sub scrllFilter_Change()
'User Filter Selection
Text1.Text = Str(scrllFilter.Value)
intColorFilter = Val(Text1.Text)
End Sub

Private Sub Text1_Change()
'User Filter Selection
If Val(Text1.Text) > 500 Then Text1.Text = "500"
If Val(Text1.Text) < 0 Then Text1.Text = "0"
scrllFilter.Value = Val(Text1.Text)
intColorFilter = Val(Text1.Text)
End Sub
Private Sub FlatScrollBar1_Change()
'Speed
Dim z As Long
z = 30000 - FlatScrollBar1.Value
lngDelay = z * 2
Frame3.Caption = "Speed " & Str(Int(100 - (z / 300))) & "%"
End Sub

Private Sub cmdGo_Click(Index As Integer)
'Houston, we have a go!
Dim X As Integer, Y As Integer, w As Long
'cmdGo and Open Menu enabled only when necessary
mnuFileOpen.Enabled = False
cmdGo(0).Enabled = False
cmdGo(1).Enabled = False
cmdGo(2).Enabled = False
cmdGo(3).Enabled = False
cmdGo(4).Enabled = False
frmMain.Refresh
    'If new picture then do everything
    If strFileNameBkp <> strFileName Then
    rtnLoadArrays
    Else
    'If same picture then don't repeat array loading
    GoSub lblGetPixelColors:
    picMain.Picture = LoadPicture("") 'Clear picture box
    End If
    
'Let's do it
Select Case Index
    
    Case 0
    'Paint picture, pixel selection according to bandpass filter
    rtnFilterWithPaintOption1
    
    Case 1
    'Paint picture, pixel selection according to bandpass filter,
    'three passes, R,RG,RGB
    rOnOff = 1: gOnOff = 0: bOnOff = 0:
    rtnFilterWithPaintOption2
    GoSub lblGetPixelColors:
    rOnOff = 1: gOnOff = 1: bOnOff = 0:
    rtnFilterWithPaintOption2
    GoSub lblGetPixelColors:
    rOnOff = 1: gOnOff = 1: bOnOff = 1:
    rtnFilterWithPaintOption2
    
    Case 2
    'Paint picture, pixel selection according to bandpass filter,
    'two passes, B&W, Color
    rtnFilterWithPaintOption3
    GoSub lblGetPixelColors:
    rtnFilterWithPaintOption1
    
    Case 3
    'Paint picture, pixel selection according to bandpass filter,
    'three passes, B&W+B, B&W+GB, B&W+RGB
    rOnOff = 0: gOnOff = 0: bOnOff = 1:
    rtnFilterWithPaintOption4
    GoSub lblGetPixelColors:
    rOnOff = 0: gOnOff = 1: bOnOff = 1:
    rtnFilterWithPaintOption4
    GoSub lblGetPixelColors:
    rOnOff = 1: gOnOff = 1: bOnOff = 1:
    rtnFilterWithPaintOption4

    Case 4
    'Paint picture, pixel selection according to bandpass filter,
    'two passes, Color minus Luminence, Color plus Luminence
    rtnFilterWithPaintOption5
    GoSub lblGetPixelColors:
    rtnFilterWithPaintOption1

End Select
picMain.AutoRedraw = True
picMain.Picture = picFilename
picMain.Picture = picMain.Image

'We use strFileNameBkp to compare with strFileName at cmdGo
'to avoid redundant array loading
'if painting the same picture again
strFileNameBkp = strFileName

'cmdGo and Open Menu enabled only when necessary
mnuFileOpen.Enabled = True
cmdGo(0).Enabled = True
cmdGo(1).Enabled = True
cmdGo(2).Enabled = True
cmdGo(3).Enabled = True
cmdGo(4).Enabled = True
Exit Sub

lblGetPixelColors:
        'Load our backed up colors to avoid redundant
        'color processing and array loading
        For Y = 0 To picMain.Height - 1
        For X = 0 To picMain.Width - 1
        w = w + 1
        lngPixelColors(X, Y) = lngPixelColorsBkp(X, Y)
        Next X
        Next Y
        Return
End Sub

Private Sub mnuExit_Click()
'hmmmm, wonder if I should comment here
Unload frmMain
End
End Sub

Private Sub mnuFileOpen_Click()
                           'Common Dialog Window
  '***************************************************************************
'cmdGo and Open Menu enabled only when necessary
cmdGo(0).Enabled = False
cmdGo(1).Enabled = False
cmdGo(2).Enabled = False
cmdGo(3).Enabled = False
cmdGo(4).Enabled = False
  CommonDialog1.CancelError = True           'Enable on error or cancel GoTo
    On Error GoTo cancelPressed
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
  CommonDialog1.DialogTitle = "Open Picture" 'Title displayed
  'CommonDialog1.InitDir = App.Path           'Start Directory
  CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg )|*.bmp;*.gif;*.jpg| All (*.*)|*.*"
  CommonDialog1.FileName = ""
  CommonDialog1.ShowOpen
  strFileName = CommonDialog1.FileName
  '***************************************************************************
'Now paint picture here and resize picture box accordingly
picMain.AutoRedraw = True
Image1.Picture = LoadPicture(strFileName)
picMain.Width = Image1.Width
picMain.Height = Image1.Height
picMain.PaintPicture Image1.Picture, 0, 0
picMain.Picture = picMain.Image
Set picFilename = picMain.Picture 'For final repainting of picture
cancelPressed:
    
    'Whether or not File selection was canceled,
    'as long as we have
    'a picture in memory then enable controls
    If strFileName <> "" Then
    cmdGo(0).Enabled = True
    cmdGo(1).Enabled = True
    cmdGo(2).Enabled = True
    cmdGo(3).Enabled = True
    cmdGo(4).Enabled = True
    End If
End Sub

Private Sub rtnLoadArrays()
Dim X As Long, Y As Long, w As Long, z As Long, prog As Long, prog2 As Long

'Ok, here we will load our arrays
X = picMain.Width
Y = picMain.Height
ReDim arrAdjPxlsToChk(X * Y, 1) 'Holds XYs of adjacent pixels not yet checked
ReDim arrAdjPxlsBeingChkd(X * Y, 1) As Integer 'Holds XYs of adjacent pixels being checked
lngTotalPixelCount = picMain.Width * picMain.Height
ReDim lngPixelColors(X, Y)
ReDim lngPixelColorsBkp(X, Y)
ReDim intRandomXYs(1, lngTotalPixelCount)
w = 0

'First we load with each pixel's color and X, Y location
For Y = 0 To picMain.Height - 1
    For X = 0 To picMain.Width - 1
    w = w + 1
    lngPixelColors(X, Y) = picMain.Point(X, Y)
    intRandomXYs(cX, w) = X
    intRandomXYs(cY, w) = Y
    Next X
prog = Y * 50 / (picMain.Height - 1)           'Progress Bar
picProgress.Line (0, 0)-(prog, 10), vbBlue, BF 'Progress Bar
Next Y

Dim tmp As Integer
'Now we will randomize the X,Y order so that our painting will not
'follow a sequential order
Randomize
For w = 1 To lngTotalPixelCount
'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
z = Int((lngTotalPixelCount - 1 + 1) * Rnd + 1)
X = intRandomXYs(cX, z)
Y = intRandomXYs(cY, z)
intRandomXYs(cX, z) = intRandomXYs(cX, w)
intRandomXYs(cY, z) = intRandomXYs(cY, w)
intRandomXYs(cX, w) = X
intRandomXYs(cY, w) = Y
prog2 = w * 10 / lngTotalPixelCount
    If w Mod 20 = 0 Then                                   'Progress Bar
    picProgress.Line (0, 0)-(prog + prog2, 10), vbBlue, BF 'Progress Bar
    End If                                                 'Progress Bar
Next w
prog = prog + prog2


'Finally we place a border (cAlreadyPainted) around picture's
'perimeter (in array) to keep our painting within this border

'Top and Bottom
For X = 0 To picMain.Width - 1
lngPixelColors(X, 0) = cAlreadyPainted
lngPixelColors(X, picMain.Height - 1) = cAlreadyPainted
prog2 = X * 2.5 / (picMain.Width - 1)                  'Progress Bar
picProgress.Line (0, 0)-(prog + prog2, 10), vbBlue, BF 'Progress Bar
Next X
prog = prog + prog2                                    'Progress Bar

'Sides
picProgress.Line (0, 0)-(prog, 10), vbBlue, BF
For Y = 0 To picMain.Height - 1
lngPixelColors(0, Y) = cAlreadyPainted
lngPixelColors(picMain.Width - 1, Y) = cAlreadyPainted
prog2 = Y * 2.5 / (picMain.Width - 1)                  'Progress Bar
picProgress.Line (0, 0)-(prog + prog2, 10), vbBlue, BF 'Progress Bar
Next Y
prog = prog + prog2                                    'Progress Bar

'Backup everything to avoid redundant processing and array loading
'when painting the same picture in memory again
For Y = 0 To picMain.Height - 1
    For X = 0 To picMain.Width - 1
    lngPixelColorsBkp(X, Y) = lngPixelColors(X, Y)
    Next X
prog2 = Y * 10 / (picMain.Height - 1)                  'Progress Bar
picProgress.Line (0, 0)-(prog + prog2, 10), vbBlue, BF 'Progress Bar
Next Y

picMain.Picture = LoadPicture("") 'Clear picture box
picProgress.Picture = LoadPicture("") 'Clear           'Progress Bar
End Sub


