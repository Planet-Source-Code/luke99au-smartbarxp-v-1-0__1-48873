VERSION 5.00
Begin VB.UserControl Clock 
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   Picture         =   "Clock.ctx":0000
   ScaleHeight     =   1245
   ScaleWidth      =   1830
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Option Explicit
Dim second As Long, minute As Long, hour As Long
Dim second_Prv As Long, minute_Prv As Long, hour_Prv As Long
Dim Pos_x(1 To 60) As Double
Dim Pos_y(1 To 60) As Double
Dim Minute_to_x As Double, Minute_to_y As Double
Dim Hour_to_x As Double, Hour_to_y As Double
Dim swp As Long, shp As Long
Dim fwp As Long, fhp As Long
Dim i As Long, j As Long, pi As Double, c As Long

Private Sub Timer1_Timer()
second = Val(Right(Time$, 2))
second_Prv = second - 1
minute = Val(Mid(Time$, 4, 2))
minute_Prv = minute - 1
hour = Val(Left(Time$, 2)) Mod 12
hour_Prv = hour - 1
If hour = 0 Then hour = 12: hour_Prv = 11
If minute = 0 Then minute = 60: minute_Prv = 59
If second = 0 Then second = 60: second_Prv = 59
If hour = 1 Then hour_Prv = 0
If minute = 1 Then minute_Prv = 60
If second = 1 Then second_Prv = 60
UserControl.Line (51, 50)-(51 + Pos_x(second_Prv), 50 + Pos_y(second_Prv)), vbWhite
Minute_to_x = 51 + 0.9 * Pos_x(minute_Prv)
Minute_to_y = 50 + 0.9 * Pos_y(minute_Prv)
draw_minute False
If second = 60 Then
   If minute < 12 Or minute = 60 Then
      Hour_to_x = 51 + 0.7 * Pos_x(hour_Prv * 5 + Int(59 / 12))
      Hour_to_y = 50 + 0.7 * Pos_y(hour_Prv * 5 + Int(59 / 12))
   Else
      If minute = 12 Then
         Hour_to_x = 51 + 0.7 * Pos_x(hour * 5 + Int(minute_Prv / 12))
         Hour_to_y = 50 + 0.7 * Pos_y(hour * 5 + Int(minute_Prv / 12))
      Else
         If hour = 12 Then hour = 0
         Hour_to_x = 51 + 0.7 * Pos_x(hour * 5 + Int(minute_Prv / 12))
         Hour_to_y = 50 + 0.7 * Pos_y(hour * 5 + Int(minute_Prv / 12))
         If hour = 0 Then hour = 12
      End If
   End If
   draw_hour False
End If
Draw_Second
Minute_to_x = 51 + 0.9 * Pos_x(minute)
Minute_to_y = 50 + 0.9 * Pos_y(minute)
draw_minute True
If hour = 12 Then hour = 0
If minute = 60 Then minute = 0
If hour = 0 And minute < 12 Then hour = 12
If hour = 0 And minute = 0 Then hour = 12
Hour_to_x = 51 + 0.7 * Pos_x(hour * 5 + Int(minute / 12))
Hour_to_y = 50 + 0.7 * Pos_y(hour * 5 + Int(minute / 12))
draw_hour True
End Sub

Sub Draw_Second()
UserControl.Line (51, 50)-(51 + Pos_x(second), 50 + Pos_y(second)), RGB(128, 128, 128)
End Sub
Sub draw_minute(Draw_Or_Clear As Boolean)
   Dim Draw_Color As ColorConstants
   If Draw_Or_Clear = True Then
      Draw_Color = RGB(128, 128, 128)
   Else
      Draw_Color = vbWhite
   End If
   For i = 50 To 52
      For j = 49 To 51
         UserControl.Line (i, j)-(Minute_to_x, Minute_to_y), Draw_Color
      Next j
   Next i
End Sub

Sub draw_hour(Draw_Or_Clear As Boolean)
   Dim Draw_Color As ColorConstants
   If Draw_Or_Clear = True Then
      Draw_Color = RGB(128, 128, 128)
   Else
      Draw_Color = vbWhite
   End If
   For i = 49 To 53
      For j = 48 To 52
         UserControl.Line (i, j)-(Hour_to_x, Hour_to_y), Draw_Color
      Next j
   Next i
End Sub

Private Sub UserControl_Initialize()
pi = 3.1415629
swp = Screen.Width / Screen.TwipsPerPixelX
shp = Screen.Height / Screen.TwipsPerPixelY
fwp = 105
fhp = 105
c = CreateEllipticRgn(0, 0, 103, 102)
SetWindowRgn UserControl.hWnd, c, True
With UserControl
   .BackColor = vbBlack
End With
For i = 1 To 60
Pos_x(i) = 35 * Cos(-(pi / 2) + (i * (pi / 30)))
Pos_y(i) = 35 * Sin(-(pi / 2) + (i * (pi / 30)))
Next i
UserControl.ScaleMode = vbPixels
End Sub
