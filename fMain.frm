VERSION 5.00
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Digital Clock"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   179
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal lnYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal bf As Long) As Boolean
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim mDC As Long  ' Memory hDC
Dim mainBitmap As Long ' Memory Bitmap
Dim blendFunc32bpp As BLENDFUNCTION
Dim token As Long ' Needed to close GDI+
Dim oldBitmap As Long
Dim imgBack As Long, imgGloss As Long
Dim imgNum(0 To 9) As Long, imgCN As Long

Private Sub Form_Initialize()
   ' Start up GDI+
   Dim GpInput As GdiplusStartupInput
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> 0 Then
     MsgBox "Error loading GDI+!", vbCritical
     Unload Me
   End If
End Sub

Private Sub Form_Load()
'   Dim lngHeight As Long, lngWidth As Long
'   Dim img As Long
'   Dim graphics As Long
'
'   ' GDI Initializations
'   Call GdipCreateFromHDC(hdc, graphics)
'   Call GdipLoadImageFromFile(StrConv(App.Path & "\images\background.png", vbUnicode), img)  ' Load Png
'   Call GdipGetImageHeight(img, lngHeight)
'   Call GdipGetImageWidth(img, lngWidth)
'   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)
'
'   Refresh
  Call Initialize
  MakeTrans
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cleanup everything
    Call CleanUp
    Call GdiplusShutdown(token)
End Sub

Private Sub Initialize()
  Dim i As Integer
  Call GdipLoadImageFromFile(StrConv(App.Path & "\images\background.png", vbUnicode), imgBack)  ' Load Png
  Call GdipLoadImageFromFile(StrConv(App.Path & "\images\gloss.png", vbUnicode), imgGloss)   ' Load Png
  Call GdipLoadImageFromFile(StrConv(App.Path & "\images\cn.png", vbUnicode), imgCN)  ' Load Png
  For i = 0 To 9
    Call GdipLoadImageFromFile(StrConv(App.Path & "\images\" & i & ".png", vbUnicode), imgNum(i))  ' Load Png
  Next
End Sub

Private Sub CleanUp()
  Dim i As Integer
  Call GdipDisposeImage(imgBack)
  Call GdipDisposeImage(imgGloss)
  Call GdipDisposeImage(imgCN)
  For i = 0 To 9
    Call GdipDisposeImage(imgNum(i))
  Next
End Sub

Private Function MakeTrans() As Boolean
  Dim tempBI As BITMAPINFO
  Dim tempBlend As BLENDFUNCTION      ' Used to specify what kind of blend we want to perform
  Dim lngHeight As Long, lngWidth As Long
  Dim curWinLong As Long
  Dim img As Long, pngPath As String
  Dim graphics As Long
  Dim winSize As Size
  Dim srcPoint As POINTAPI
  
  With tempBI.bmiHeader
    .biSize = Len(tempBI.bmiHeader)
    .biBitCount = 32    ' Each pixel is 32 bit's wide
    .biHeight = Me.ScaleHeight  ' Height of the form
    .biWidth = Me.ScaleWidth    ' Width of the form
    .biPlanes = 1   ' Always set to 1
    .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8) ' This is the number of bytes that the bitmap takes up. It is equal to the Width*Height*ByteCount (bitCount/8)
  End With
  mDC = CreateCompatibleDC(Me.hdc)
  mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
  oldBitmap = SelectObject(mDC, mainBitmap)   ' Select the new bitmap, track the old that was selected
  
  Dim Tmp As String
  Tmp = Format(Now, "hh:mm")
  ' GDI Initializations
  Call GdipCreateFromHDC(mDC, graphics)
  Call GdipDrawImageRect(graphics, imgBack, 0, 0, 130, 93)
  Call GdipDrawImageRect(graphics, imgNum(Int(Left(Tmp, 1))), 20, 29, 16, 26)
  Call GdipDrawImageRect(graphics, imgNum(Int(Mid(Tmp, 2, 1))), 40, 29, 16, 26)
  Call GdipDrawImageRect(graphics, imgCN, 60, 32, 8, 21)
  Call GdipDrawImageRect(graphics, imgNum(Int(Mid(Tmp, 4, 1))), 72, 29, 16, 26)
  Call GdipDrawImageRect(graphics, imgNum(Int(Mid(Tmp, 5, 1))), 92, 29, 16, 26)
  Call GdipDrawImageRect(graphics, imgGloss, -17, 18, 130, 50)
  
  ' Change windows extended style to be used by updatelayeredwindow
  curWinLong = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
  ' Accidently did This line below which flipped entire form, it's neat so I left it in
  ' Comment out the line above and uncomment line below.
  'curWinLong = GetWindowLong(Me.hwnd, GWL_STYLE)
  SetWindowLong Me.hwnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED
  
  ' Make the window a top-most window so we can always see the cool stuff
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  
  ' Needed for updateLayeredWindow call
  srcPoint.x = 0
  srcPoint.y = 0
  winSize.cx = 130
  winSize.cy = 93
  
  With blendFunc32bpp
    .AlphaFormat = AC_SRC_ALPHA ' 32 bit
    .BlendFlags = 0
    .BlendOp = AC_SRC_OVER
    .SourceConstantAlpha = 255
  End With
  
  Call GdipDisposeImage(img)
  Call GdipDeleteGraphics(graphics)
  Call UpdateLayeredWindow(Me.hwnd, Me.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Function

Private Sub Timer1_Timer()
  Dim Tmp As String
  Static lTmp As String
  Tmp = Format(Now, "hh:mm")
  If Tmp <> lTmp Then
    MakeTrans
    lTmp = Tmp
  End If
End Sub
