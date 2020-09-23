<div align="center">

## Add Splitter Bars to your app \*Revised\*


</div>

### Description

Add vertical and horizontal splitter bars to your application. This

code creates a form with three panes, and two splitter bars. I originally

found the code in Planet Source Code last year, but found a couple of

bugs, so I rebuilt it from scratch, and added the horizontal splitter.

There are no DLL's or references to worry about, very easy to use.
 
### More Info
 
It is easy to integrate into your applications. You simply paste whatever

objects that you want to use into one of the three picture boxes used as

frames. Then in the _Resize event of each, change the .Move statement to

your object (or control).

Copy the code and paste it into a text editor (like notepad). Save the file

as Form1.frm. You can always change it later.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eugene](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eugene.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eugene-add-splitter-bars-to-your-app-revised__1-5098/archive/master.zip)





### Source Code

```
VERSION 5.00
Begin VB.Form Form1
 Caption = "Form1"
 ClientHeight = 6180
 ClientLeft = 210
 ClientTop = 1800
 ClientWidth = 7575
 LinkTopic = "Form1"
 ScaleHeight = 6180
 ScaleWidth = 7575
 Begin VB.PictureBox picOuterFrame
 Appearance = 0 'Flat
 ForeColor = &H80000008&
 Height = 5535
 Left = 120
 ScaleHeight = 5505
 ScaleWidth = 7065
 TabIndex = 0
 Top = 120
 Width = 7095
 Begin VB.PictureBox spltVertical
 Appearance = 0 'Flat
 CausesValidation= 0 'False
 ClipControls = 0 'False
 FillColor = &H8000000F&
 FillStyle = 0 'Solid
 ForeColor = &H8000000F&
 Height = 4935
 Left = 3480
 MousePointer = 9 'Size W E
 ScaleHeight = 4905
 ScaleWidth = 225
 TabIndex = 1
 Top = 0
 Width = 255
 End
 Begin VB.PictureBox picRight
 Appearance = 0 'Flat
 BackColor = &H80000005&
 ForeColor = &H80000008&
 Height = 4815
 Left = 3840
 ScaleHeight = 4785
 ScaleWidth = 2985
 TabIndex = 2
 Top = 240
 Width = 3015
 End
 Begin VB.PictureBox picLeft
 Appearance = 0 'Flat
 ForeColor = &H80000008&
 Height = 4575
 Left = 0
 ScaleHeight = 4545
 ScaleWidth = 3345
 TabIndex = 3
 Top = 240
 Width = 3375
 Begin VB.PictureBox spltHorizontal
 Appearance = 0 'Flat
 FillColor = &H8000000F&
 FillStyle = 0 'Solid
 ForeColor = &H8000000F&
 Height = 255
 Left = 480
 MousePointer = 7 'Size N S
 ScaleHeight = 225
 ScaleWidth = 2385
 TabIndex = 4
 Top = 2160
 Width = 2415
 End
 Begin VB.PictureBox picTopLeft
 Appearance = 0 'Flat
 BackColor = &H80000005&
 ForeColor = &H80000008&
 Height = 1815
 Left = 480
 ScaleHeight = 1785
 ScaleWidth = 2025
 TabIndex = 6
 Top = 120
 Width = 2055
 End
 Begin VB.PictureBox picBottomLeft
 Appearance = 0 'Flat
 BackColor = &H80000005&
 ForeColor = &H80000008&
 Height = 1815
 Left = 600
 ScaleHeight = 1785
 ScaleWidth = 2025
 TabIndex = 5
 Top = 2520
 Width = 2055
 End
 End
 End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SPLT_WDTH As Long = 80 'width of the spltter bar
Private Const MIN_WINDOW As Long = 10 'Minimum size for any frame created by splitter bars
Private Sub Form_Load()
 '**** Splitter Code ****
 'No Borders, they are for development and debugging
 spltVertical.BorderStyle = 0
 spltHorizontal.BorderStyle = 0
 picOuterFrame.BorderStyle = 0
 picLeft.BorderStyle = 0
 picTopLeft.BorderStyle = 0
 picBottomLeft.BorderStyle = 0
 picRight.BorderStyle = 0
 '**** End Splitter Code ****
End Sub
Private Sub picRight_Resize()
 'Resize your object to the inside of the frame
 'YourObject.Move 0, 0, picRight.Width, picRight.Height
End Sub
Private Sub picTopLeft_Resize()
 'Resize your object to the inside of the frame
 'YourObject.Move 0, 0, picTopLeft.Width, picTopLeft.Height
End Sub
Private Sub picBottomLeft_Resize()
 'Resize your object to the inside of the frame
 'YourObject.Move 0, 0, picBottomLeft.Width, picBottomLeft.Height
End Sub
Private Sub Form_Resize()
 'For this example, I chose to reside all the frames, depending on the size of the
 ' form. You may choose to put this whole assembly in another sub-frame.
 '**** Splitter Code ****
 'Resize the outer frame
 Dim height1 As Long, width1 As Long
 height1 = ScaleHeight - (2 * SPLT_WDTH)
 If height1 < 0 Then height1 = 0
 width1 = ScaleWidth - (2 * SPLT_WDTH)
 If width1 < 0 Then width1 = 0
 picOuterFrame.Move SPLT_WDTH, SPLT_WDTH, width1, height1
 '**** End Splitter Code ****
End Sub
'**** Splitter Code ****
Private Sub spltVertical_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = vbLeftButton Then
 spltVertical.Move (spltVertical.Left - (SPLT_WDTH \ 2)) + x, 0, SPLT_WDTH, picOuterFrame.ScaleHeight
 spltVertical.BackColor = vbButtonShadow 'change the splitter colour
 End If
End Sub
Private Sub spltVertical_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If spltVertical.BackColor = vbButtonShadow Then
 spltVertical.Move (spltVertical.Left - (SPLT_WDTH \ 2)) + x, 0, SPLT_WDTH, picOuterFrame.ScaleHeight
 End If
End Sub
Private Sub spltVertical_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If spltVertical.BackColor = vbButtonShadow Then
 spltVertical.BackColor = vbButtonFace 'restore splitter colour
 spltVertical.Move (spltVertical.Left - (SPLT_WDTH \ 2)) + x, 0, SPLT_WDTH, picOuterFrame.ScaleHeight
 'Set the absolute Boundaries
 Dim lAbsLeft As Long
 Dim lAbsRight As Long
 lAbsLeft = MIN_WINDOW
 lAbsRight = picOuterFrame.ScaleWidth - (SPLT_WDTH + MIN_WINDOW)
 Select Case spltVertical.Left
 Case Is < lAbsLeft 'the pane is too thin
 spltVertical.Move lAbsLeft, 0, SPLT_WDTH, picOuterFrame.ScaleHeight
 Case Is > lAbsRight 'the pane is too wide
 spltVertical.Move lAbsRight, 0, SPLT_WDTH, picOuterFrame.ScaleHeight
 End Select
 'reposition both frames, and the spltVertical bar
 picOuterFrame_Resize
 End If
End Sub
Private Sub spltHorizontal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = vbLeftButton Then
 spltHorizontal.BackColor = vbButtonShadow 'change the splitter colour
 spltHorizontal.Move 0, (spltHorizontal.Top - (SPLT_WDTH \ 2)) + y, picLeft.ScaleWidth, SPLT_WDTH
 End If
End Sub
Private Sub spltHorizontal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If spltHorizontal.BackColor = vbButtonShadow Then
 spltHorizontal.Move 0, (spltHorizontal.Top - (SPLT_WDTH \ 2)) + y, picLeft.ScaleWidth, SPLT_WDTH
 End If
End Sub
Private Sub splthorizontal_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If spltHorizontal.BackColor = vbButtonShadow Then
 spltHorizontal.BackColor = vbButtonFace 'restore splitter colour
 spltHorizontal.Move 0, (spltHorizontal.Top - (SPLT_WDTH \ 2)) + y, picLeft.ScaleWidth, SPLT_WDTH
 'Set the absolute Boundaries
 Dim lAbsTop As Long
 Dim lAbsBottom As Long
 lAbsTop = MIN_WINDOW
 lAbsBottom = picLeft.ScaleHeight - (SPLT_WDTH + MIN_WINDOW)
 Select Case spltHorizontal.Top
 Case Is < lAbsTop 'the pane is too short
 spltHorizontal.Move 0, lAbsTop, picLeft.ScaleWidth, SPLT_WDTH
 Case Is > lAbsBottom 'the pane is too tall
 spltHorizontal.Move 0, lAbsBottom, picLeft.ScaleWidth, SPLT_WDTH
 End Select
 'reposition both sub-frames, and the spltHorizontal bar
 picLeft_Resize
 End If
End Sub
Private Sub picOuterFrame_Resize()
 Dim x1 As Long
 Dim x2 As Long
 Dim y1 As Long
 On Error Resume Next
 y1 = picOuterFrame.ScaleHeight
 x1 = spltVertical.Left
 x2 = x1 + SPLT_WDTH + 1
 picLeft.Move 0, 0, x1 - 1, y1
 spltVertical.Move x1, 0, SPLT_WDTH, y1
 picRight.Move x2, 0, picOuterFrame.ScaleWidth - x2, y1
 'Force a refresh on the left side
 picLeft_Resize
End Sub
Private Sub picLeft_Resize()
 'Resize the internal stuff. Only the width's
 Dim x1 As Long
 Dim y1 As Long
 Dim y2 As Long
 Dim y3 as Long
 x1 = picLeft.Width
 y1 = spltHorizontal.Top
 y2 = y1 + SPLT_WDTH + 1
 'We have to make sure that we do not size any windows to a negative dimension
 y3 = y1 - 1
 If y3 < MIN_WINDOW Then
 y3 = MIN_WINDOW
 End If
 picTopLeft.Move 0, 0, x1, y3
 spltHorizontal.Move 0, y1, x1, SPLT_WDTH
 y3 = picLeft.ScaleHeight - y2
 If y3 < MIN_WINDOW Then
 y3 = MIN_WINDOW
 End If
 picBottomLeft.Move 0, y2, x1, y3
End Sub
'**** End Splitter Code ****
```

