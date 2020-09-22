VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6660
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3086.957
      ScaleMode       =   0  'User
      ScaleWidth      =   5915.045
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Keeps me from being sloppy :-)
Dim XPos, YPos As Integer 'Current X and Y positions of the "laser"
Dim Color As Long 'The color that the "laser" is currentlly drawing

Dim vLeft As Boolean, hLeft As Boolean

Private Enum LaserDrawModes
    LaserCorner
    PrinterScan
    WierdDraw
    WierdDrawSlow
End Enum

'LaserDraw
'**** Description ***********
'Copies a picture from one picture box (or form) to another, with an animated "laser" effect
'**** Usage ***************
'LaserDraw PictureToDraw, DrawSurface, LaserOriginX, LaserOriginY, BackColor
'**** Inputs ***************
'0 PictureToDraw - Picturebox containing the picture to be copied
'0 DrawSurface - Picturebox or Form which the picture should be copied to
'0 LaserOriginX - Optional; The x coordinate of where the laser should come from.
'   Default is the width of the PictureToDraw picture box
'0 LaserOriginY - Optional; The y coordinate of where the laser should come from.
'   Default is the height of the PictureToDraw picture box
'0 BackColor - Optional; The background color of the DrawSurface
'   Default is the current background color of DrawSurface
'0 LaserDrawMode - Optional; The style of the laser draw
'   LaserCorner - Original mode, draws the picture, one line at a time, as if from a laser in a corner
'   PrinterScan - Draws the picture as if a printer were going along and drawing each dot
'   WierdDraw - Wierd draw mode, similar to PrinterScan. Try it for yourself :-)
'   Default is LaserCorner
'**** Outputs *****************
'None

Private Sub LaserDraw(PictureToDraw As PictureBox, DrawSurface As Object, Optional LaserOriginX = -1, Optional LaserOriginY = -1, Optional BackColor As ColorConstants = -1, Optional LaserDrawMode As LaserDrawModes = LaserCorner)
    'Set up the DrawSurface picture box
        DrawSurface.ScaleMode = vbPixels 'Set the scale mode of the "canvas" to pixels
        If BackColor <> -1 Then 'Background color specified
            DrawSurface.BackColor = BackColor 'Set the background color of the "canvas" to the desired background color
        End If
    'Set up the PictureToDraw picture box
        PictureToDraw.ScaleMode = vbPixels 'Set the scale mode of the picturebox containing the picture to be drawn to pixels
        PictureToDraw.AutoRedraw = True 'Set the autoredraw property of the picturebox containing the picture to be drawn to true
        PictureToDraw.Visible = False 'Hide the picturebox containing the picture to be drawn
    'Set up the X and Y coordinates of the "laser"
        If LaserOriginX = -1 Then 'No X coordinate of the "laser" is specified
            LaserOriginX = PictureToDraw.ScaleWidth 'Set it to the width of the picturebox containing the picture to be drawn
        End If
        If LaserOriginY = -1 Then 'No Y coordinate of the "laser" is specified
            LaserOriginY = PictureToDraw.ScaleHeight 'Set it to the height of the picturebox containing the picture to be drawn
        End If
    'Start the "Laser" effect
        For XPos = 0 To PictureToDraw.ScaleWidth 'Move the "laser" horizantally along the "canvas"
            DoEvents 'Allow input to be prosessed
            For YPos = 0 To PictureToDraw.ScaleHeight 'Move the "laser" verticlly along the "canvas"
                Color = PictureToDraw.Point(XPos, YPos) 'Determine the color of the pixel to be drawn
                If LaserDrawMode = LaserCorner Then 'Normal Drawing
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, LaserOriginY), Color 'Draw a line from the origin coordinates to the coordinates of the pixel to be drawn
                ElseIf LaserDrawMode = PrinterScan Then '"Printer Scanning" mode
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, YPos), Color 'Draw a straight line from the pixel to LaserOrginX
                    DrawSurface.Line (XPos + 1, YPos - 1)-(LaserOriginX, YPos - 1), BackColor 'Erase the last position of the "laser"
                    DoEvents 'Alow input to be prosessed
                ElseIf LaserDrawMode = WierdDrawSlow Then '"Weird Draw Slow" mode
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, YPos), Color 'Draw a straight line from the pixel to LaserOrginX
                    DoEvents 'Alow input to be prosessed
                Else '"Wierd Draw" mode
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, YPos), Color 'Draw a straight line from the pixel to LaserOrginX
                End If
            Next
        Next
End Sub


Private Sub Form_Load()
    Me.Show
    'While True
        'LaserDraw Picture1, Me, Me.ScaleWidth, Me.ScaleHeight, vbBlack, LaserCorner
        'Me.Cls
        LaserDraw Picture1, Me, Me.ScaleWidth, Me.ScaleHeight, vbBlack, WierdDraw
        'Me.Cls
        'LaserDraw Picture1, Me, Me.ScaleWidth, Me.ScaleHeight, vbBlack, WierdDrawSlow
        'Me.Cls
        'LaserDraw Picture1, Me, Me.ScaleWidth, Me.ScaleHeight, vbBlack, PrinterScan
    'Wend
End Sub
