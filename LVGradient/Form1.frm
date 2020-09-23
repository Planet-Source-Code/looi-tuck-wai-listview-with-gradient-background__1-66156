VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listview Gradient Background Demostration"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkCircular 
      Caption         =   "Circular"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CheckBox chkMiddleOut 
      Caption         =   "Middle-Out"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   5040
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   615
      Left            =   1080
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF0000&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   120
      ScaleHeight     =   205
      ScaleMode       =   0  'User
      ScaleWidth      =   355
      TabIndex        =   2
      Top             =   6480
      Width           =   5355
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Grid"
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   3240
      Width           =   1155
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3105
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5477
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog C1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog C2 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider sldAngle 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   1
      Max             =   360
      SelStart        =   116
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   116
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient Option"
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   5415
      Begin VB.PictureBox Picture4 
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   615
         Left            =   960
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   1200
         Width           =   615
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   615
         Left            =   2640
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkFour 
         Caption         =   "Four Corner"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Colour 3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Colour 4"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblGradAngle 
         Alignment       =   2  'Center
         Caption         =   "Angle = 116"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Degree :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Colour 2"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Colour 1"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog C3 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog C4 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'================================================================================'
'Author : Looi Tuck Wai                                                          '
'Date   : 31/07/06                                                               '
'================================================================================'
'Credits goes to these VBGurus :
'Deming Shang     === LV Background Painting Module
'Carles P.V       === Original Writer Of mGradient Module
'Matthew R. Usner === Enchancement FormGradient Module & OmniGradient Writer
'Karl E. Peterson === CStopWatch Class Module Writer
Private CorM As Boolean
Private lvHgt As Single
Private GradientAngle As Single
Private MiddleOut As Boolean
Private bmiheader                     As BITMAPINFOHEADER
Private BG_lBits()                  As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiheader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type cRGB
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0

'*******************************************************************************
' DrawGradient (FUNCTION)
'
' DESCRIPTION:
' This function is used to draw gradients with four colours
'
' Arguments:
' hDC - The device to draw on
' Top - Distance in pixels, from top
' Left - Distance in pixels, from left
' Width - In pixels
' Height - In pixels
' colourTopLeft - The colour of the top-left corner
' colourTopRight - The colour of the top-right corner
' colourBottomLeft - The colour of the bottom-left corner
' colourBottomRight - The colour of the bottom-right corner
'*******************************************************************************
Public Function DrawGradient(hDC As Long, Left As Long, Top As Long, Width As Long, Height As Long, colourTopLeft As Long, colourTopRight As Long, colourBottomLeft As Long, colourBottomRight As Long)
    Dim bi24BitInfo     As BITMAPINFO
    Dim bBytes()        As Byte
    Dim LeftGrads()     As cRGB
    Dim RightGrads()    As cRGB
    Dim MiddleGrads()   As cRGB
    Dim TopLeft         As cRGB
    Dim TopRight        As cRGB
    Dim BottomLeft      As cRGB
    Dim BottomRight     As cRGB
    Dim iLoop           As Long
    Dim bytesWidth      As Long
    
    With TopLeft
        .Red = Red(colourTopLeft)
        .Green = Green(colourTopLeft)
        .Blue = Blue(colourTopLeft)
    End With
    
    With TopRight
        .Red = Red(colourTopRight)
        .Green = Green(colourTopRight)
        .Blue = Blue(colourTopRight)
    End With
    
    With BottomLeft
        .Red = Red(colourBottomLeft)
        .Green = Green(colourBottomLeft)
        .Blue = Blue(colourBottomLeft)
    End With
    
    With BottomRight
        .Red = Red(colourBottomRight)
        .Green = Green(colourBottomRight)
        .Blue = Blue(colourBottomRight)
    End With
    
    GradateColours LeftGrads, Height, TopLeft, BottomLeft
    GradateColours RightGrads, Height, TopRight, BottomRight
    
    With bi24BitInfo.bmiheader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiheader)
        .biWidth = Width
        .biHeight = 1
    End With
    
    ReDim bBytes(1 To bi24BitInfo.bmiheader.biWidth * bi24BitInfo.bmiheader.biHeight * 3) As Byte
    
    bytesWidth = (Width) * 3
    
    For iLoop = 0 To Height - 1
        GradateColours MiddleGrads, Width, LeftGrads(iLoop), RightGrads(iLoop)
        CopyMemory bBytes(1), MiddleGrads(0), bytesWidth
        SetDIBitsToDevice hDC, Left, Top + iLoop, bi24BitInfo.bmiheader.biWidth, bi24BitInfo.bmiheader.biHeight, 0, 0, 0, bi24BitInfo.bmiheader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    Next iLoop
    
    
End Function

'*******************************************************************************
' GradateColours (FUNCTION)
'
' DESCRIPTION:
' This function is to blend colour1 to colour2
'*******************************************************************************
Private Function GradateColours(cResults() As cRGB, Length As Long, Colour1 As cRGB, Colour2 As cRGB)
    Dim fromR   As Integer
    Dim toR     As Integer
    Dim fromG   As Integer
    Dim toG     As Integer
    Dim fromB   As Integer
    Dim toB     As Integer
    Dim stepR   As Single
    Dim stepG   As Single
    Dim stepB   As Single
    Dim iLoop   As Long
    
    ReDim cResults(0 To Length)
    
    fromR = Colour1.Red
    fromG = Colour1.Green
    fromB = Colour1.Blue
    
    toR = Colour2.Red
    toG = Colour2.Green
    toB = Colour2.Blue
    
    stepR = Divide(toR - fromR, Length)
    stepG = Divide(toG - fromG, Length)
    stepB = Divide(toB - fromB, Length)
    
    For iLoop = 0 To Length
        cResults(iLoop).Red = fromR + (stepR * iLoop)
        cResults(iLoop).Green = fromG + (stepG * iLoop)
        cResults(iLoop).Blue = fromB + (stepB * iLoop)
    Next iLoop
End Function

'*******************************************************************************
' Blue (FUNCTION)
'
' DESCRIPTION:
' Retrieve Blue from Long
'*******************************************************************************
Private Function Blue(Colour As Long) As Long
    Blue = (Colour And &HFF0000) / &H10000
End Function

'*******************************************************************************
' Green (FUNCTION)
'
' DESCRIPTION:
' Retrieve Green as long
'*******************************************************************************
Private Function Green(Colour As Long) As Long
    Green = (Colour And &HFF00&) / &H100
End Function

'*******************************************************************************
' Red (FUNCTION)
'
' DESCRIPTION:
' Retrieve Red from Long
'*******************************************************************************
Private Function Red(Colour As Long) As Long
    Red = (Colour And &HFF&)
End Function

'*******************************************************************************
' Divide (FUNCTION)
'
' DESCRIPTION:
' Division function to avoid division by 0 error
'*******************************************************************************
Private Function Divide(Numerator, Denominator) As Single
    If Numerator = 0 Or Denominator = 0 Then
        Divide = 0
    Else
        Divide = Numerator / Denominator
    End If
End Function

Private Sub Check1_Click()
  lv.GridLines = Check1.Value
End Sub
Private Sub chkCircular_Click()

   If chkCircular.Value = vbChecked Then
      chkMiddleOut.Value = 0
      sldAngle.Enabled = False
      lblGradAngle.Enabled = False
      CorM = True
   Else
      chkMiddleOut.Value = 1
      sldAngle.Enabled = True
      lblGradAngle.Enabled = True
      CorM = False
   End If
   DisplayGradient
End Sub

Private Sub chkFour_Click()
If CorM = True Then
   chkCircular.Value = 1
   chkMiddleOut.Value = 0
Else
   chkCircular.Value = 0
   chkMiddleOut.Value = 1
End If
   If chkFour.Value = vbChecked Then
      chkMiddleOut.Enabled = False
      chkCircular.Enabled = False
      sldAngle.Enabled = False
      lblGradAngle.Enabled = False
      Label3.Enabled = False
      Picture4.Enabled = True
      Picture5.Enabled = True
      Label4.Enabled = True
      Label5.Enabled = True
      lvHgt = lv.ListItems(1).Height
    With picBg
      .Width = lv.Width
      .Height = lvHgt * (lv.ListItems.Count) + lvHgt
      
      DrawGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor
      
    End With
   Else
      chkCircular.Enabled = True
      chkMiddleOut.Enabled = True

      sldAngle.Enabled = True
      lblGradAngle.Enabled = True
      Picture4.Enabled = False
      Picture5.Enabled = False
      Label4.Enabled = False
      Label5.Enabled = False
      Label3.Enabled = True
      DisplayGradient
   End If
End Sub

Private Sub chkMiddleOut_Click()
   MiddleOut = (chkMiddleOut.Value = vbChecked)
   CorM = False
   DisplayGradient
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mRow As ListItem
    
    Me.ScaleMode = vbTwips
    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.ColumnHeaders.Add , , "No."
    lv.ColumnHeaders.Add , , "Description"
    For i = 0 To 40
      Set mRow = lv.ListItems.Add(, , CStr(i))
      mRow.SubItems(1) = "This is Item " & i
    Next
    
    picBg.BackColor = lv.BackColor
    picBg.BorderStyle = vbBSNone
    picBg.AutoRedraw = True
    picBg.Visible = False
    GradientAngle = 116
    MiddleOut = True
    CorM = False
DisplayGradient
End Sub
Private Sub DisplayGradient()

   Dim CircularFlag As Boolean
   Dim Clr1 As Long
   Dim Clr2 As Long
   Dim SwpClr As Long
   Dim sw As New CStopWatch

   Clr1 = Picture2.BackColor
   Clr2 = Picture3.BackColor
   CircularFlag = (chkCircular.Value = vbChecked)

   If CircularFlag Then
      SwpClr = Clr1
      Clr1 = Clr2
      Clr2 = SwpClr
   End If
   lvHgt = lv.ListItems(1).Height
    picBg.Width = lv.Width
    picBg.Height = lvHgt * (lv.ListItems.Count) + lvHgt
   With picBg
    .Cls
      sw.Reset
      CalcGradient .ScaleWidth, .ScaleHeight, Clr1, Clr2, GradientAngle, MiddleOut, bmiheader, BG_lBits(), CircularFlag
      sw.Reset
      mGradient.PaintGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0, .ScaleWidth, .ScaleHeight, BG_lBits(), bmiheader
      .Refresh
       lv.Picture = picBg.Image
   End With
End Sub

Private Sub sldAngle_Scroll()
   GradientAngle = sldAngle.Value
   lblGradAngle.Caption = "Angle = " & CStr(GradientAngle)
   DisplayGradient
End Sub

Private Sub Picture2_Click()
C1.ShowColor
Picture2.BackColor = C1.Color
If chkFour.Value = vbChecked Then
      With picBg
      .Width = lv.Width
      .Height = lvHgt * (lv.ListItems.Count) + lvHgt
      DrawGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor
    End With
Else
Picture2.BackColor = C1.Color
DisplayGradient
End If
End Sub

Private Sub Picture3_Click()
C2.ShowColor
Picture3.BackColor = C2.Color
If chkFour.Value = vbChecked Then
      With picBg
      .Width = lv.Width
      .Height = lvHgt * (lv.ListItems.Count) + lvHgt
      DrawGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor
    End With
Else
Picture3.BackColor = C2.Color
DisplayGradient
End If
End Sub

Private Sub Picture4_Click()
On Error GoTo clrerr:
C3.ShowColor
Picture4.BackColor = C3.Color
clrerr:
Picture4.BackColor = C3.Color
    With picBg
      .Width = lv.Width
      .Height = lvHgt * (lv.ListItems.Count) + lvHgt
      DrawGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor
    End With
End Sub

Private Sub Picture5_Click()
On Error GoTo clrerr:
C4.ShowColor
Picture5.BackColor = C4.Color
clrerr:
Picture5.BackColor = C4.Color
    With picBg
      .Width = lv.Width
      .Height = lvHgt * (lv.ListItems.Count) + lvHgt
      DrawGradient .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Picture2.BackColor, Picture3.BackColor, Picture4.BackColor, Picture5.BackColor
    End With
End Sub
