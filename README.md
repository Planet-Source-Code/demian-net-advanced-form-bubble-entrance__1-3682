<div align="center">

## ADVANCED Form Bubble Entrance


</div>

### Description

Makes Your Form Have A Bubble Entrance!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Demian Net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/demian-net.md)
**Level**          |Advanced
**User Rating**    |2.6 (18 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/demian-net-advanced-form-bubble-entrance__1-3682/archive/master.zip)

### API Declarations

```
Public Declare Function CreateEllipticRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "User32" _
(ByVal Hwnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long
```


### Source Code

```
'Make Your Form Name frm
Private Sub Form_Load()
frm.Show
Dim a As Integer
Dim b As Integer
Dim C As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim w As Integer
Dim X As Integer
Dim Y As Integer
Dim z As Integer
Call frm.Move(0, 0)
w = frm.Height
X = frm.Width
Y = frm.Top
z = frm.Left
a = 0
b = 0
C = w
d = X
e = Y
f = z
Do While a < frm.Height / 15 Or b < frm.Width / 15
a = a + 25
b = b + 25
e = e + 70
f = f + 70
If a > frm.Height / 15 Then a = a - 24
If b > frm.Width / 15 Then b = b - 24
Call frm.Move(f, e, d, C)
current = Timer
Do While Timer - current < 0.01
DoEvents
Loop
Call SetWindowRgn(frm.Hwnd, CreateEllipticRgn(0, 0, b, a), True)
Loop
current = Timer
Do While Timer - current < 1
DoEvents
Loop
Call SetWindowRgn(frm.Hwnd, CreateEllipticRgn(0, 0, 0, 0), True)
End Sub
```

