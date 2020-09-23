VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureStretched 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3195
      Left            =   2550
      ScaleHeight     =   3135
      ScaleWidth      =   4755
      TabIndex        =   10
      Top             =   3750
      Width           =   4815
   End
   Begin VB.CommandButton CmdDoIt 
      Caption         =   "&Do It"
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   6240
      Width           =   2385
   End
   Begin VB.PictureBox PictureResult 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2400
      Left            =   90
      ScaleHeight     =   2340
      ScaleWidth      =   2340
      TabIndex        =   6
      Top             =   3750
      Width           =   2400
   End
   Begin VB.PictureBox PictureBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2400
      Left            =   5010
      ScaleHeight     =   2340
      ScaleWidth      =   2340
      TabIndex        =   4
      Top             =   1020
      Width           =   2400
   End
   Begin VB.PictureBox PictureMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2400
      Left            =   2550
      ScaleHeight     =   2340
      ScaleWidth      =   2340
      TabIndex        =   2
      Top             =   1020
      Width           =   2400
   End
   Begin VB.PictureBox PictureSource 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2400
      Left            =   90
      ScaleHeight     =   2340
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   1020
      Width           =   2400
   End
   Begin VB.Label Label6 
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "StretchBlt Demonstrated"
      Height          =   225
      Left            =   2580
      TabIndex        =   8
      Top             =   3480
      Width           =   4755
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Transparency Demonstrated"
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   3480
      Width           =   2355
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Background Picture"
      Height          =   255
      Left            =   5010
      TabIndex        =   5
      Top             =   750
      Width           =   2355
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Mask Picture"
      Height          =   255
      Left            =   2550
      TabIndex        =   3
      Top             =   750
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Source Picture"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   2385
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
'*****************************************************************************
' This code was written for VB Planet. You may use this code freely. However
' displaying any of this code on a webpage or distributing it in
' uncompiled form is strictly prohibited. Thank you.
'          2000 VB Planet
'*****************************************************************************
' API CALLS
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' COnstants for BITBLT/STRETCHBLT
Private Const SRCCOPY = &HCC0020   ' Copies image exactly
Private Const SRCAND = &H8800C6    ' Used for mask
Private Const SRCPAINT = &HEE0086  ' Used for Transparencies
Private Const SRCINVERT = &H660046 ' XOR
Private Const SRCERASE = &H440328  ' dest = source AND (NOT dest )
'@=================================================
' CmdDoIt_Click:
'  This is where everything is done for our Demo
'==================================================
Private Sub CmdDoIt_Click()
 Dim X As Single ' used for positioning drawing methods
 Dim Y As Single ' used for positioning drawing methods
 Dim Radius As Single ' use for circle method
 Dim RET As Long ' Return value from API Calls
 Dim StretchX As Long
 Dim StretchY As Long
' Clear pictures
 PictureSource.Cls
 PictureBack.Cls
 PictureMask.Cls
 PictureResult.Cls
' ******** Step 1
' first we draw a picture in source picture
 MsgBox ("Step 1. load a picture into the source picture.")
 PictureSource.Picture = LoadPicture(App.Path + "\fgrid.bmp")
'********* Step 2
' next we create a mask in picturemask
 MsgBox ("Step 2. Create a mask for the picture.")
 Blit.CreateMask PictureSource, PictureMask
'******* Step 3
' now we create a background picture
 MsgBox ("Step 3. Load a background picture")
 PictureBack.Picture = LoadPicture(App.Path + "\back.bmp")
' ****** Step 4
' We copy the background
 MsgBox ("Step 4. Copy the background using BitBlt with SRCCOPY.")
 Blit.BlitPicture PictureBack, PictureResult
' ***** Step 5
' Now blit the mask
 MsgBox ("Step 5. Put the mask ontop of the background using BitBlt with SRCAND.")
 Blit.BlitPicture PictureMask, PictureResult, SRCAND
' ***** Step 6
' Now we put the source picture ontop of the mask
 MsgBox ("Step 6. Put the source picture ontop of the mask using BitBlt and SRCPAINT.")
 Blit.BlitPicture PictureSource, PictureResult, SRCPAINT
'****** Step 7
' Now we stretch the blit into picturestretched
 MsgBox ("Step 7. For fun we stretch the blit using Stretchblt API call")
 Blit.StretchPicture PictureResult, PictureStretched
End Sub
