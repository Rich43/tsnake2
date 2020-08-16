Attribute VB_Name = "Module1"
 Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
 ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
 ByVal ySrc As Long, ByVal dwRop As Long) As Long

 Public Const SRCCOPY = &HCC0020
 Public Const SRCPAINT = &HEE0086
 Public Const SRCAND = &H8800C6

