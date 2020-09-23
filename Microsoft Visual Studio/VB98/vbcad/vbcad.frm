VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "VBCAD"
   ClientHeight    =   4068
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   10
   ScaleMode       =   0  'User
   ScaleWidth      =   10
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   216
      Left            =   2892
      Max             =   5
      TabIndex        =   2
      Top             =   3804
      Value           =   1
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Circle"
      Height          =   216
      Left            =   1236
      TabIndex        =   1
      Top             =   3816
      Width           =   984
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Snap to Grid"
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   3792
      Value           =   1  'Checked
      Width           =   1416
   End
   Begin VB.Label Label1 
      Caption         =   "Zoom"
      Height          =   204
      Left            =   2412
      TabIndex        =   3
      Top             =   3804
      Width           =   828
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X, Y, TX, TY, SX, SY, ZH, ZW
Dim xx(1000), yy(1000), xy(1000), yx(1000) ' save points for line redraw
Dim cx(1000), cy(1000), cr(1000) 'save center and radius of circles
Public cxx, cyy
Public drawcircle As Boolean
Public linenumber As Double
Public circlenumber As Double
Public MouseCounter As Boolean
Public Zoom
Public Snap
Public ended As Boolean

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then Snap = True Else Snap = False ' check for snap to grid
End Sub

Private Sub Command1_Click()
drawcircle = True
Command1.Enabled = False
End Sub

Private Sub Form_DblClick()
If Snap = False Then
Line (TX, TY)-(SX, SY) 'end shape
linenumber = linenumber + 1
xx(linenumber) = TX
yy(linenumber) = TY
xy(linenumber) = SX
yx(linenumber) = SY
Else
Line (round(TX), round(TY))-(round(SX), round(SY)) 'end shape
linenumber = linenumber + 1
xx(linenumber) = round(TX)
yy(linenumber) = round(TY)
xy(linenumber) = round(SX)
yx(linenumber) = round(SY)

End If
MouseCounter = True
End Sub

Private Sub Form_Load()
ZW = 10
ZH = 11
Form_Resize ' INIT GRID
Zoom = 0     ' INIT ZOOM
Snap = True  ' INIT SETTINGS
MouseCounter = True
ended = True
linenumber = 0
drawcircle = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not drawcircle Then
Select Case Button
Case 1
If MouseCounter Then
TX = X 'set temp variables
TY = Y
SX = X
SY = Y
ended = False
MouseCounter = False
Else
If ended = True Then GoTo ened ' if it ended dont draw line!
If Snap = False Then
Line (TX, TY)-(X, Y)
linenumber = linenumber + 1
xx(linenumber) = TX
yy(linenumber) = TY
xy(linenumber) = X
yx(linenumber) = Y
Else
Line (round(TX), round(TY))-(round(X), round(Y)) 'round to closest point
linenumber = linenumber + 1
xx(linenumber) = round(TX)
yy(linenumber) = round(TY)
xy(linenumber) = round(X)
yx(linenumber) = round(Y)
End If
If ended Then 'check to see if still drawing shape
SX = X
SY = Y
MouseCounter = True
Else
ended = False
MouseCounter = False
TX = X
TY = Y
End If
End If
Case 2
ended = True
MouseCounter = True
End Select
Else ' draw the circle
circlenumber = circlenumber + 1
cxx = X
cyy = Y
newradius = InputBox("Please Enter The Radius", "VBCAD CIRCLE", "0")
Form1.Circle (cxx, cyy), newradius
cx(circlenumber) = cxx
cy(circlenumber) = cyy
cr(circlenumber) = newradius
drawcircle = False
Command1.Enabled = True
End If
Exit Sub
ened:
SX = X
SY = Y
MouseCounter = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Caption = "VBCAD - X" & X & ", Y" & Y ' Change Titlebar
End Sub

Private Sub Form_Resize()
'On Error GoTo ender
GoTo noender
ender:
End
noender:
Command1.Top = Form1.ScaleHeight - 0.8 'move ctrls
Command1.Height = 0.8
Check1.Top = Form1.ScaleHeight - 0.8
Check1.Height = 0.8
Label1.Top = Form1.ScaleHeight - 0.8
Label1.Height = 0.8
VScroll1.Top = Form1.ScaleHeight - 0.8
VScroll1.Height = 0.8
Form1.Refresh 'Remove Old Lines
Form1.ScaleHeight = ZH 'Re-establish Scaleheights and Widths
Form1.ScaleWidth = ZW
Form1.ForeColor = &H808080 'setforecolor to dark grey
Form1.DrawStyle = 2 ' make lines dotted
For i = 0 To Form1.ScaleHeight
If Not i = round(Form1.ScaleHeight / 2) Then
Line (0, i)-(Form1.ScaleHeight, i) ' draw grid
Else
Form1.ForeColor = &H909090 'setforecolor to darkergrey
Form1.DrawStyle = 1 ' make lines dashed
Line (0, i)-(Form1.ScaleHeight, i)
Form1.ForeColor = &H808080 'setforecolor to dark grey
Form1.DrawStyle = 2 ' make lines dotted
End If
Next
For i = 0 To Form1.ScaleWidth
If Not i = (Form1.ScaleWidth / 2) Then
Line (i, 0)-(i, Form1.ScaleWidth)
Else
Form1.ForeColor = &H909090 'setforecolor to darkergrey
Form1.DrawStyle = 1 ' make lines dashed
Line (i, 0)-(i, Form1.ScaleWidth)
Form1.ForeColor = &H808080 'setforecolor to dark grey
Form1.DrawStyle = 2 ' make lines dotted
End If
Next
Form1.DrawStyle = 0 ' return to original values
Form1.ForeColor = &H0
For i = 0 To linenumber 'redraw old lines
Line (xx(i), yy(i))-(xy(i), yx(i))
Next
For i = 0 To circlenumber ' redraw old circles
Circle (cx(i), cy(i)), cr(i) 'redraw old circles
Next

End Sub
Function round(ins) As Double
If Len(ins) = 1 Then
round = Int(ins)
Exit Function
End If
' if is a whole number, no need to round
temp = Mid(ins, 3, 1)
If temp > 5 Then
round = Int(ins) + 1 'round up
Else
round = Int(ins) 'round down
End If

End Function

Private Sub VScroll1_Change()
ZH = 11 + (10 * VScroll1.Value)
ZW = 10 + (10 * VScroll1.Value)
Form_Resize
End Sub
