Attribute VB_Name = "Module1"
'Module1: RRPublics.bas

' Mainly to hold Publics

Option Base 1
DefLng A-W
DefSng X-Z

' APIs for getting DIB bits to PalBGR

Public Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal HDC As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
(ByVal HDC As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
(ByVal HDC As Long) As Long

'---------------------------------------------------------------

'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type

'NB PICTURE STORED IN MEMORY UPSIDE DOWN
'WITH INCREASING MEMORY GOING UP THE PICTURE
'bmp.bmBits points to the bottom left of the picture

Public bmp As BITMAP
'------------------------------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   'bmiH As RGBTRIPLE            'NB Palette NOT NEEDED for 16,24 & 32-bit
End Type
Public bm As BITMAPINFO

' For transferring BGRA array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal HDC As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcX As Long, ByVal SrcY As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long
'------------------------------------------------------------------------------

'To shift cursor out of the way
'Public Declare Sub SetCursorPos Lib "user32" (ByVal IX As Long, ByVal IY As Long)

'Copy one array to another of same number of bytes
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetPixel Lib "gdi32" _
(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long

'Public Declare Function SetPixelV Lib "gdi32" _
'(ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'Public Declare Function StretchBlt Lib "gdi32" _
'(ByVal HDC As Long, _
'ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
'ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nPICWidth As Long, ByVal nPICHeight As Long, ByVal dwRop As Long) As Long

'Used to extract small bitmap from a large one and show shrunken bitmap
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
'ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
'ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'------------------------------------------------------------------------------


'-----------------------------------------------------------------------

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long

Public IFilterMC() As Byte     ' Array to hold machine code
Public ptMC, ptrStruc          ' Ptrs to Machine Code & Structure

' MCode Structure
Public Type MCodeStruc
   PICW As Long
   PICH As Long
   zmf1 As Single
   zmf2 As Single
   zdf As Single
   zf As Single
   ckQB As Long
   ckABS As Long
   QBLongColor As Long
   PtrMat As Long
   PtrPalBGR As Long
End Type
Public MC As MCodeStruc
'-----------------------------------------------------------------------

Public Done As Boolean      ' For LOOPING
Public ASM As Boolean       ' VB <-> ASM

Public SysBPP               ' 16 or 32 (24) bits/pixel

Public PICW, PICH           ' Display picbox Width & Height (pixels)
Public PalBGR() As Byte     ' To hold 3 full palettes (12 x PICW x PICH)
Public PicFrameW, PicFrameH ' Size of PIC frame container


' General byte RGBs
Public QBRed As Byte, QBGreen As Byte, QBBlue As Byte
Public QBLongColor   '= RGB(QBRed, QBGreen, QBBlue)

Public PalBGRPtr            ' Pointer to PalBGR(1,1,1,1)
Public PalSize              ' Size of 1 palette (4 x PICW x PICH)

Public Mat(), MatIndex     ' To hold matrix multipliers (Long)
Public zmf1, zmf2, zdf, zf ' Multiplying factors (Single)

' Unlimited Realities Filters
Public URName$()
Public zURFactors(), URMatrix()
' Manuel Santos Filters
Public MSName$()
Public zMSFactors(), MSMatrix()
' User Filters
Public USName$()
Public zUSFactors(), USMatrix()
Public NumUserFilters

Public Const pi# = 3.1415926535898
Public Const d2r# = pi# / 180

' For frmMsg   Form Message Box response
Public ynYES, ynNO, ynOK

