Attribute VB_Name = "Blit"
'*****************************************************************************
' This code was written for VB Planet. You may use this code freely. However
' displaying any of this code on a webpage or distributing it in
' uncompiled form is strictly prohibited. Thank you.
'          2000 VB Planet
'*****************************************************************************
' API CALLS
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
' Bit Blit operations Enumeration
Enum SRCOps
 SRCCOPY = &HCC0020   ' Copies image exactly
 SRCAND = &H8800C6    ' Used for mask
 SRCPAINT = &HEE0086  ' Used for Transparencies
 SRCINVERT = &H660046 ' XOR
 SRCERASE = &H440328  ' dest = source AND (NOT dest )
End Enum
'@========================================================
' BlitPicture: Blits from one picture box to another
' Parameters:
'     SrcPic = Source PictureBox
'     DestPic = Destination PictureBox
'     [SrcOp]= Operation to perform
'=========================================================
Sub BlitPicture(ByVal SrcPic As PictureBox, ByVal DestPic As PictureBox, _
                Optional ByVal Srcop As SRCOps = SRCCOPY)
 Dim RET As Long ' Return Value from API Call
 ' Blit from source to Dest
 RET = BitBlt(DestPic.hdc, 0, 0, _
   DestPic.Width / Screen.TwipsPerPixelX, _
   DestPic.Height / Screen.TwipsPerPixelY, _
   SrcPic.hdc, 0, 0, Srcop)
 ' refresh Destination picture
 DestPic.Refresh
End Sub
'@========================================================
' StretchPicture: Blits from one picture box to another
' Parameters:
'     SrcPic = Source PictureBox
'     DestPic = Destination PictureBox
'     [SrcOp]= Operation to perform
'=========================================================
Sub StretchPicture(ByVal SrcPic As PictureBox, ByVal DestPic As PictureBox, _
                   Optional ByVal Srcop As SRCOps = SRCCOPY)
 Dim RET As Long ' Return Value from API Call
 ' Blit from source to Dest
 BlitPicture SrcPic, DestPic, Srcop
 ' refresh Destination picture
 DestPic.Refresh
 ' now we stretch
 RET = StretchBlt(DestPic.hdc, 0, 0, _
  DestPic.Width / Screen.TwipsPerPixelX, DestPic.Height / Screen.TwipsPerPixelY, _
  DestPic.hdc, 0, 0, _
  SrcPic.Width / Screen.TwipsPerPixelX, _
  SrcPic.Height / Screen.TwipsPerPixelY, _
  Srcop)
 ' refresh again
 DestPic.Refresh
End Sub
'@=============================================================
' CreateMask: Creates a mask from a picture box
'     SrcPic = Source PictureBox
'     DestPic = Destination PictureBox
'     [TransColor] = Color to treat as transparent
'     [ForeColor] = Color to use for foreground
'     [BackColor] = Color to use for background
'================================================================
Sub CreateMask(ByVal SrcPic As PictureBox, ByVal DestPic As PictureBox, _
               Optional ByVal TransColor = vbBlack, _
               Optional ByVal ForeColor = vbBlack, _
               Optional ByVal BackColor = vbWhite)
 Dim X As Long
 Dim Y As Long
 ' check every pixel in srcpic
  For X = 0 To SrcPic.Width Step Screen.TwipsPerPixelX
   For Y = 0 To SrcPic.Height Step Screen.TwipsPerPixelY
     ' if we se transparant color set to backcolor
    If SrcPic.Point(X, Y) = TransColor Then
     DestPic.PSet (X, Y), BackColor
    Else ' we set to forecolor
     DestPic.PSet (X, Y), ForeColor
    End If
   Next Y
  Next X
End Sub

