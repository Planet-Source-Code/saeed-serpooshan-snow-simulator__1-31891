VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSnow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   Caption         =   "Snow"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FormSnow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextNParticle 
      Height          =   285
      Left            =   7020
      MaxLength       =   4
      TabIndex        =   19
      Text            =   "250"
      Top             =   5040
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Pictures|*.Jpg;*.Gif;*.bmp|All Files|*.*"
   End
   Begin VB.CommandButton CommandLoadPic 
      Caption         =   "Load Picture"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   5760
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00C0C0C0&
      Height          =   3240
      Left            =   5880
      TabIndex        =   14
      Top             =   420
      Width           =   1875
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   3
      LargeChange     =   4
      Left            =   4980
      Max             =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5100
      Value           =   4
      Width           =   1755
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      LargeChange     =   20
      Left            =   4980
      Max             =   100
      Min             =   1
      SmallChange     =   2
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4740
      Value           =   10
      Width           =   1755
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      LargeChange     =   4
      Left            =   2340
      Max             =   10
      Min             =   -10
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5100
      Value           =   3
      Width           =   1755
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      LargeChange     =   4
      Left            =   2340
      Max             =   10
      Min             =   -10
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4740
      Value           =   2
      Width           =   1755
   End
   Begin VB.CommandButton CommandStopSnow 
      BackColor       =   &H00800000&
      Caption         =   "Stop Fall"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5220
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Form"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   4980
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5760
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Dir List"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5760
      Width           =   795
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Picture Box"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5760
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   1440
   End
   Begin VB.CommandButton CommandFallSnow 
      BackColor       =   &H00800000&
      Caption         =   "Fall Snow"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4740
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   480
      Picture         =   "FormSnow.frx":0442
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   0
      Top             =   420
      Width           =   5250
   End
   Begin VB.PictureBox PictureTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1320
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   18
      Top             =   5700
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TextBorder 
      BackColor       =   &H00C0C0C0&
      Height          =   4020
      Left            =   435
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   375
      Width           =   5325
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000EFEF&
      Height          =   270
      Left            =   7140
      MouseIcon       =   "FormSnow.frx":8C39
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5820
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Timer:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   4440
      TabIndex        =   20
      Top             =   4800
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You can fall snow in any window using it's 'hwnd' property! "
      ForeColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   5940
      TabIndex        =   15
      Top             =   3720
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Random:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   5
      Left            =   4260
      TabIndex        =   13
      Top             =   5100
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Weight:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   8
      Top             =   5100
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Wind-x:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1740
      TabIndex        =   7
      Top             =   4740
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Particles:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   7020
      TabIndex        =   5
      Top             =   4800
      Width           =   660
   End
End
Attribute VB_Name = "FormSnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------
'
'    « In The Name Of The Most High »
'
'  Snow Simulator 1.0
'  - The program choose n particle initially
'    in the screen and their properties such
'    as Vx,Vy and Color. then animate them in
'    each timer interval.
'
'  By: Saeed Serpooshan - Mechanical Engineer
'  All comments are welcome
'  Email: SSerpooshan@Yahoo.Com
'  Iran - 2002
'
' --------------------------------------------

DefLng A-Z 'define Long type as default declaration of variables.
Option Explicit

'API Declaretions:
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc, ByVal x, ByVal y) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'-- Device Context functions
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'dwRop in BitBlt (and StretchBlt):
Private Const SRCPAINT = &HEE0086   'dest = source OR dest
Private Const SRCERASE = &H440328   'dest = source AND (NOT dest )
Private Const SRCAND = &H8800C6     'dest = source AND dest
Private Const SRCCOPY = &HCC0020    'dest = source
Private Const SRCINVERT = &H660046  'dest = source XOR dest

'Custom Declarations:
Dim nSnow As Long
Dim VxMinSnow As Single, VxMaxSnow As Single, VyMinSnow As Single, VyMaxSnow As Single
Dim VxAddMin As Single, VxAddMax As Single, VyAddMin As Single, VyAddMax As Single
Dim WidthWindowSnow, HeightWindowSnow
Dim xSnow() As Single, ySnow() As Single, VxSnow() As Single, VySnow() As Single
Dim ColPrevSnow(), ColSnow() As Long
Dim hdcSnow As Long, HwndSnow As Long
Dim StopSnow As Integer, DontClearParticles As Boolean
Dim IsInAnimateSnow As Boolean, IsInCmdFall As Boolean


Private Sub CommandFallSnow_Click()

If IsInCmdFall Then Exit Sub
IsInCmdFall = True

Timer1.Enabled = False
StopSnow = False

ClearSnowParticles
DontClearParticles = False 'the value True means that don't clear particles in 'ClearSnowParticles' sub

nSnow = Val(TextNParticle.Text)
If nSnow < 0 Then nSnow = 250

Dim i

SetInitialSnowPositions
'this loop is to reach to steady state motion!
For i = 1 To 50
 AnimateSnow False
 DoEvents
Next

If -1 = True Then
 'this loop cause to see begin of falling and particles do not appear instantly
 For i = 1 To nSnow
  ySnow(i) = ySnow(i) - (HeightWindowSnow + 1) * Sgn(VySnow(i))
 Next
End If

AnimateSnow
Timer1.Enabled = True
IsInCmdFall = False


End Sub

Sub SetInitialSnowPositions(Optional DrawInitialParticles As Boolean = False)

ReDim xSnow(nSnow), ySnow(nSnow), VxSnow(nSnow), VySnow(nSnow), ColPrevSnow(nSnow), ColSnow(nSnow)
Dim w, h, hdc, i, x, y, c

hdc = hdcSnow
w = WidthWindowSnow: h = HeightWindowSnow

'set a x,y position , velocity(x,y) and a color for each particle:
For i = 1 To nSnow
 x = Rnd * w: y = Rnd * h
 xSnow(i) = x: ySnow(i) = y
 If Rnd < 0.3 Then
  c = &HFFFFFF ' &HFFEFEF
 Else
  c = 150 + Rnd * (260 - 150): If c > 255 Then c = 255
  c = GetRealNearestColor(hdc, RGB(c, c, c))
 End If
 ColSnow(i) = c
 ColPrevSnow(i) = GetPixel(hdc, x, y)
 VxSnow(i) = VxMinSnow + Rnd * (VxMaxSnow - VxMinSnow)
 VySnow(i) = VyMinSnow + Rnd * (VyMaxSnow - VyMinSnow)
 If DrawInitialParticles Then SetPixelV hdc, x, y, c
Next

End Sub

Sub SetSpeed()

Dim vx As Single, vy As Single, r As Single

'Vx,yMin,MaxSnow: determine Min,Max value of absolute speed
'VxMinSnow = -1.5: VxMaxSnow = 1.5
'VyMinSnow = -1: VyMaxSnow = 2

'Vx,yAddMin,Max: determine Min,Max of rate of change in speed (i.e. acceleration)
VxAddMin = -0.1: VxAddMax = 0.1
VyAddMin = -0.1: VyAddMax = 0.1

vx = HScroll1(1).Value / 2: vy = HScroll1(2).Value / 2
r = HScroll1(3).Value / 4
VxMinSnow = vx - r / 2: VxMaxSnow = vx + r / 2
VyMinSnow = vy - r / 2: VyMaxSnow = vy + r / 2

End Sub

Sub AnimateSnow(Optional DrawParticles As Boolean = -1)

If IsInAnimateSnow Then Exit Sub
IsInAnimateSnow = True

Dim w, h, hdc, i, x As Single, y As Single, vx As Single, vy As Single, c
Dim j, jx, c2

hdc = hdcSnow
w = WidthWindowSnow: h = HeightWindowSnow

If DrawParticles Then 'clear old pixels:
 For i = nSnow To 1 Step -1
  c = ColPrevSnow(i)
  If c <> -1 Then SetPixelV hdc, xSnow(i), ySnow(i), c
 Next
End If

For i = 1 To nSnow
 x = xSnow(i): y = ySnow(i)
 vx = VxSnow(i) + VxAddMin + Rnd * (VxAddMax - VxAddMin)
 vy = VySnow(i) + VyAddMin + Rnd * (VyAddMax - VyAddMin)
 SetValueInRange vx, VxMinSnow, VxMaxSnow
 SetValueInRange vy, VyMinSnow, VyMaxSnow
 VxSnow(i) = vx: VySnow(i) = vy
 x = x + vx: y = y + vy
 If Not StopSnow Then
   'SetValueInRange y, 0, h, True
   If y > h And vy >= 0 Then
    y = 0
   Else
    If y < 0 And vy <= 0 Then y = h
   End If
 End If
 SetValueInRange x, 0, w, True
 
 c = GetPixel(hdc, x, y) 'return (-1) if x,y be out of real part of window
 xSnow(i) = x: ySnow(i) = y
 ColPrevSnow(i) = c
 If DrawParticles And c <> -1 Then
   SetPixelV hdc, x, y, ColSnow(i)
 End If
 

Next

IsInAnimateSnow = False

End Sub

Sub ClearSnowParticles()
 'clear current particles
 Dim hdc, i, c
 If DontClearParticles Then Exit Sub 'particles are cleared before and must be clear again (because hdc may be changed!)

 hdc = hdcSnow
 
 For i = nSnow To 1 Step -1
  c = ColPrevSnow(i)
  If c <> -1 Then SetPixelV hdc, xSnow(i), ySnow(i), c
 Next
End Sub

Sub SetValueInRange(v As Variant, ByVal RangeMin As Variant, ByVal RangeMax As Variant, Optional SwapMaxMin As Boolean = False)
 If SwapMaxMin Then 'swapMaxMin=True:
  If v < RangeMin Then v = RangeMax Else If v > RangeMax Then v = RangeMin
 Else 'default (swapmaxmin=false)
  If v < RangeMin Then v = RangeMin Else If v > RangeMax Then v = RangeMax
 End If
End Sub

Private Sub CommandLoadPic_Click()
On Error Resume Next
CD1.ShowOpen
If Err Then Err.Clear: Exit Sub
ClearSnowParticles
DontClearParticles = True
Dim w, h
With PictureTmp
 .Picture = LoadPicture(CD1.FileName)
 w = Picture1.ScaleWidth: h = Picture1.ScaleHeight
 Picture1.PaintPicture .Picture, 0, 0, w, h, 0, 0, .ScaleWidth, .ScaleHeight
End With
Option1_Click (0)
CommandFallSnow_Click

End Sub

Private Sub CommandStopSnow_Click()
 StopSnow = True
End Sub

Private Sub Form_Load()
 Picture1.ScaleMode = vbPixels
 Me.Show: DoEvents
 Call SetSpeed
 'ColSnow = &HFFFFFF
 'ColSnow(i) = GetRealNearestColor(hdc, ColSnow(i))
 Option1_Click (0)
 CommandFallSnow_Click
End Sub

Function GetRealNearestColor(ByVal hdc1 As Long, ByVal Col As Long) As Long
 Dim c As Long
 'to get real color (using getnearestcolor function may have some problems!
 c = GetPixel(hdc1, 1, 1)
 If c <> -1 Then
  GetRealNearestColor = SetPixel(hdc1, 1, 1, Col)
  SetPixelV hdc1, 1, 1, c
 Else
  GetRealNearestColor = Col 'faild to test
 End If

End Function

Private Sub HScroll1_Change(Index As Integer)
 If Index = 0 Then 'timer
  Timer1.Interval = HScroll1(0).Value
 Else 'wind,weight or random
  Call SetSpeed
 End If
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
 HScroll1_Change Index
End Sub

Private Sub Label5_Click()
MsgBox "Snow Simulator (First Release)" + vbCrLf + vbCrLf + "SSerpooshan@" + Chr(89) + "ahoo" + ".c" + Chr(111) + "m" + vbCrLf + "By:S.Serpooshan" + Space(20) + "Iran - 2002" + Space(4) + vbCrLf, , "About..."
End Sub

Private Sub Option1_Click(Index As Integer)
 If Option1(Index).Value = 0 Then Option1(Index).Value = 1
 
 Dim obj(3) As Object, Cobj As Object
 Set obj(1) = Picture1: Set obj(2) = Dir1: Set obj(3) = Me
 
 DeleteUsedSnowDC
 DontClearParticles = True
 
 If Index <> 1 And Dir1.Enabled = False Then Dir1.Enabled = True
 
 Set Cobj = obj(Index + 1)
 With Cobj
  If Index = 1 Then
   .Enabled = False
  Else
   WidthWindowSnow = .ScaleWidth - 1:  HeightWindowSnow = .ScaleHeight - 1
  End If
  HwndSnow = .Hwnd
 End With
 
 hdcSnow = GetDC(HwndSnow)
 
 If Timer1.Enabled Then CommandFallSnow_Click
 
End Sub

Private Sub TextNParticle_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then KeyAscii = 0: CommandFallSnow_Click
End Sub

Private Sub Timer1_Timer()
 AnimateSnow
End Sub

Private Sub Form_Unload(Cancel As Integer)
 DeleteUsedSnowDC
End Sub

Sub DeleteUsedSnowDC()
 If hdcSnow <> 0 Then
  ClearSnowParticles
  ReleaseDC HwndSnow, hdcSnow
 End If
End Sub
