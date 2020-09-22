VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reader"
   ClientHeight    =   4710
   ClientLeft      =   3375
   ClientTop       =   1995
   ClientWidth     =   8100
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   72
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   Begin VB.CommandButton Command5 
      Caption         =   "How it Works?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   24
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   13
      Left            =   6720
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   12
      Left            =   5640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   11
      Left            =   4560
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   10
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   18
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   9
      Left            =   2400
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   8
      Left            =   1320
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   7
      Left            =   240
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6000
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   6
      Left            =   6720
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   5
      Left            =   5640
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   4
      Left            =   4560
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   3
      Left            =   3480
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   2
      Left            =   2400
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   1
      Left            =   1320
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicDrewLet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   0
      Left            =   240
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox PicLetter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Index           =   0
      Left            =   120
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   3
      Top             =   3240
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyze"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox PicDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   0
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   539
      TabIndex        =   0
      Top             =   0
      Width           =   8085
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1560
      TabIndex        =   23
      Top             =   3240
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "**Important Note**"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Letter List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Dim DrawLetter As Boolean
Dim Letter(25) As String
Dim LetNum As Integer
Dim NextStartX As Integer
Dim ShowLetter As Integer

Private Sub Command1_Click()
Dim x, y, i, z As Integer
Dim LetTotal(25) As Integer
Dim CutLimit, NumMade, GotIn(25) As Integer
Dim Finish As Boolean

Form1.MousePointer = 11
lblLetter.Caption = ""

For z = 0 To LetNum - 1
    For i = 0 To 25
        LetTotal(i) = 0
    Next i
    For i = 0 To 25
        For x = 0 To 67
            For y = 0 To 69
                If PicDrewLet(z).Point(x, y) = PicLetter(i).Point(x, y) Then
                    LetTotal(i) = LetTotal(i) + 1
                End If
            Next y
        Next x
    Next i
    
    CutLimit = 2000
    For i = 0 To 25
        GotIn(i) = 0
    Next i
    NumMade = 0
    Finish = False
    
    Do
        For i = 0 To 25
            If LetTotal(i) > CutLimit Then
                GotIn(NumMade) = i
                NumMade = NumMade + 1
            End If
        Next i
        If NumMade = 1 Then
            lblLetter.Caption = lblLetter.Caption & Letter(GotIn(0))
            Finish = True
        End If
        If CutLimit > 10000 Then
            lblLetter.Caption = lblLetter.Caption & "?"
            Finish = True
        End If
        NumMade = 0
        CutLimit = CutLimit + 25
    Loop Until Finish = True
Next z

Form1.MousePointer = 0
End Sub

Private Sub Command2_Click()
ShowLetter = ShowLetter - 1
If ShowLetter = -1 Then
    ShowLetter = 25
End If

PicLetter(ShowLetter).ZOrder 0
End Sub

Private Sub Command3_Click()
ShowLetter = ShowLetter + 1
If ShowLetter = 26 Then
    ShowLetter = 0
End If

PicLetter(ShowLetter).ZOrder 0
End Sub

Private Sub Command4_Click()
PicDraw.Cls
LetNum = 0
NextStartX = 0
lblLetter.Caption = ""
End Sub

Private Sub Command5_Click()
MsgBox "   Firstly i got this idea from another project posted on the internet.  The program was " & vbNewLine _
       & "Quickwrite by some who called himself ((VBMASTER)).  But i only got the idea from him " & vbNewLine _
       & "nothing more. " & vbNewLine & vbNewLine _
       & "   The program has all 26 letter stored in picture boxes in the array PicLetter(). This " & vbNewLine _
       & "is the same array in the Letter List.  When you draw a letter and release the mouse " & vbNewLine _
       & "the program works in from the left right top and bottom to find the rectangle that the " & vbNewLine _
       & "letter is in using the functions: FindBottom, FindTop, FindLeft, FindRight.  Then it " & vbNewLine _
       & "stretches that rectangle into the PicDrewLet() array (each letter for a picturebox." & vbNewLine _
       & "Then when the user click Analyze it compares the image to each letter.  If two " & vbNewLine _
       & "pixels match that letter gets a point.  Then it compares each letters points and the " & vbNewLine _
       & "letter with the highest points is the letter that the image is.", vbOKOnly, "Read"
End Sub

Private Sub Form_Load()
Dim i As Integer

Letter(0) = "A"
Letter(1) = "B"
Letter(2) = "C"
Letter(3) = "D"
Letter(4) = "E"
Letter(5) = "F"
Letter(6) = "G"
Letter(7) = "H"
Letter(8) = "I"
Letter(9) = "J"
Letter(10) = "K"
Letter(11) = "L"
Letter(12) = "M"
Letter(13) = "N"
Letter(14) = "O"
Letter(15) = "P"
Letter(16) = "Q"
Letter(17) = "R"
Letter(18) = "S"
Letter(19) = "T"
Letter(20) = "U"
Letter(21) = "V"
Letter(22) = "W"
Letter(23) = "X"
Letter(24) = "Y"
Letter(25) = "Z"

For i = 1 To 25
    Load PicLetter(i)
    PicLetter(i).Visible = True
Next i

For i = 0 To 25
    PicLetter(i).Picture = LoadPicture(App.Path & "\" & Letter(i) & ".bmp")
Next i
End Sub

Private Sub PicDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DrawLetter = True
End Sub

Private Sub PicDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DrawLetter = True Then
    PicDraw.Circle (x, y), 6, vbBlack
End If
End Sub

Private Sub PicDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, x2, y1, y2 As Integer

Form1.MousePointer = 11
DrawLetter = False

y1 = FindTop
y2 = FindBottom
x1 = FindLeft
x2 = FindRight

StretchBlt PicDrewLet(LetNum).hdc, 0, 0, 67, 69, PicDraw.hdc, x1, y1, x2 - x1, y2 - y1, SRCCOPY
LetNum = LetNum + 1
Form1.MousePointer = 0
End Sub

Function FindTop() As Integer
Dim x, y As Integer

For y = 0 To PicDraw.Height
    For x = 0 + NextStartX To PicDraw.Width
        If PicDraw.Point(x, y) = vbBlack Then
            FindTop = y
            Exit Function
        End If
    Next x
Next y


End Function

Function FindLeft() As Integer
Dim x, y As Integer

For x = 0 + NextStartX To PicDraw.Width
    For y = 0 To PicDraw.Height
        If PicDraw.Point(x, y) = vbBlack Then
            FindLeft = x
            Exit Function
        End If
    Next y
Next x



End Function

Function FindRight() As Integer
Dim x, y As Integer

For x = PicDraw.Width To 0 Step -1
    For y = 0 To PicDraw.Height
        If PicDraw.Point(x, y) = vbBlack Then
            FindRight = x
            NextStartX = x + 1
            Exit Function
        End If
    Next y
Next x


End Function

Function FindBottom() As Integer
Dim x, y As Integer

For y = PicDraw.Height To 0 Step -1
    For x = 0 + NextStartX To PicDraw.Width
        If PicDraw.Point(x, y) = vbBlack Then
            FindBottom = y
            Exit Function
        End If
    Next x
Next y


End Function

