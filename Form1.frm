VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   930
   ClientTop       =   990
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Index           =   4
      Left            =   5925
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   5340
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1500
      Left            =   30
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   5
      Top             =   4785
      Width           =   1500
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         Height          =   150
         Left            =   0
         Top             =   0
         Width           =   150
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Index           =   3
      Left            =   4770
      Picture         =   "Form1.frx":7572
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   4
      Top             =   5310
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Index           =   2
      Left            =   3555
      Picture         =   "Form1.frx":EAE4
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   3
      Top             =   5445
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Index           =   1
      Left            =   2025
      Picture         =   "Form1.frx":16056
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   5445
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Index           =   0
      Left            =   1380
      Picture         =   "Form1.frx":1D5C8
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   5775
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox Picture1 
      Height          =   4140
      Left            =   1620
      ScaleHeight     =   272
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   330
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Replace these images with your own things and see them on a tile engine. Email Mikepote@mailcity.com"
      Height          =   480
      Left            =   1695
      TabIndex        =   7
      Top             =   4890
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCPAINT = &HEE0086    ' dest = source OR dest
Private Grid(0 To 100, 0 To 100) As Integer
Public Down As Boolean, Px, Py, CX, CY, Vx, Vy, Size, I As Integer, VeX, VeY

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
For XX = VeX To VeX + Int(Picture1.Width / Size)
For YY = VeY - 1 To -Int(Vy / Size) + Int(Picture1.Height / Size)
BitBlt Picture1.hdc, (XX * Size) + Vx, (YY * Size) + Vy, Size, Size, Picture2(Grid(XX, YY)).hdc, 0, 0, SRCCOPY
Shape1.Left = -Int(Vx / Size)
Shape1.Top = -Int(Vy / Size)
Next
Next
End Sub

Private Sub Form_Load()


'This code sets up the map.
' This is a completely random map so you can make it load from a file or something.
Dim J As Integer
For XX = 0 To 100
For YY = 0 To 100
J = J + 1
If J >= 5 Then J = 0
Let Grid(XX, YY) = J
Next
Next
DrawMap
Let Vx = -500
Let Vy = -500
End Sub

Private Sub Form_Resize()
Picture1.Width = ScaleWidth - Picture1.Left - 10
Picture1.Height = ScaleHeight - Picture1.Top - 10
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = True
CX = X
CY = Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Down Then
Px = CX
Py = CY
CX = X
CY = Y
Vx = Vx + (CX - Px)
Vy = Vy + (CY - Py)
If Vx >= 0 Then Vx = 0
If Vy >= 0 Then Vy = 0
If Vx <= -9000 Then Vx = -9000
If Vy <= -9000 Then Vy = -9000
'Vx and Vy are the vaules representing where the screen is on the map.
' they are negative.
End If
Dim XX, YY
Size = 100
On Error Resume Next
For XX = -Int(Vx / Size) - 1 To -Int(Vx / Size) + Int(Picture1.Width / Size)
For YY = -Int(Vy / Size) - 1 To -Int(Vy / Size) + Int(Picture1.Height / Size)
' This the most important line of this whole project.
' If you jont understand Bitblt think about downloading
' Mike Canjeros Skinning Example.
BitBlt Picture1.hdc, (XX * Size) + Vx, (YY * Size) + Vy, Size, Size, Picture2(Grid(XX, YY)).hdc, 0, 0, SRCCOPY
Shape1.Left = -Int(Vx / Size)
Shape1.Top = -Int(Vy / Size)
Next
Next
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = False
End Sub

Sub DrawMap()
For XX = 0 To 100
For YY = 0 To 100
' This just draws the automap
Picture3.PSet (XX, YY), Picture2(Grid(XX, YY)).Point(50, 50)
Next
Next
Picture3.Refresh
End Sub
