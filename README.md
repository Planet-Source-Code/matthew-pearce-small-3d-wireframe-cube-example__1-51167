<div align="center">

## Small 3D Wireframe Cube Example


</div>

### Description

Creates a 3D cube on the form which can be manipulated using the left, right, up and down arrow keys.
 
### More Info
 
This code uses Sine and Cosine to spin the cube around.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Pearce](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-pearce.md)
**Level**          |Intermediate
**User Rating**    |4.7 (42 globes from 9 users)
**Compatibility**  |VB 5\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-pearce-small-3d-wireframe-cube-example__1-51167/archive/master.zip)





### Source Code

```
' Name: 3D Cube Program
' Author: Matthew Pearce
' Date: Tuesday, 20th January 2004
' Purpose: To display an interactive 3D cube on the screen using Sin and Cos
Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, pitch As Double
Private Sub Form_Activate()
  i1 = degtorad(0)
  i2 = degtorad(90)
  i3 = degtorad(180)
  i4 = degtorad(270)
  pitch = 0.5
  UpdateCube
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case vbKeyLeft:
    i1 = i1 - 10
    i2 = i2 - 10
    i3 = i3 - 10
    i4 = i4 - 10
  Case vbKeyRight:
    i1 = i1 + 10
    i2 = i2 + 10
    i3 = i3 + 10
    i4 = i4 + 10
  Case vbKeyUp:
    If pitch > -1 Then
      pitch = pitch - 0.1
    End If
  Case vbKeyDown:
    If pitch < 1 Then
      pitch = pitch + 0.1
    End If
  End Select
  UpdateCube
End Sub
Function degtorad(deg As Integer)
  deg = deg / 360 * 3.14 * 200
  degtorad = deg
End Function
Function LineCalc(inum1, inum2, inum3, inum4, add1, add2)
  ' Calculate verteces
  LX1 = LineCalcX(inum1)
  LX2 = LineCalcX(inum2)
  LY1 = LineCalcY(inum3, add1)
  LY2 = LineCalcY(inum4, add2)
  ' Draw lines between verteces
  Line (LX1, LY1)-(LX2, LY2)
End Function
Function LineCalcX(inum)
  LineCalcX = Sin(inum / 100) * 1000 + 2000
End Function
Function LineCalcY(inum, a1)
  LineCalcY = Cos(inum / 100) * 400 * pitch + a1
End Function
Function UpdateCube()
  Refresh
  ' Draw Top Face
  LineCalc i1, i2, i1, i2, 1000, 1000
  LineCalc i2, i3, i2, i3, 1000, 1000
  LineCalc i3, i4, i3, i4, 1000, 1000
  LineCalc i4, i1, i4, i1, 1000, 1000
  ' Draw Sides
  LineCalc i1, i2, i1, i2, 2000, 2000
  LineCalc i2, i3, i2, i3, 2000, 2000
  LineCalc i3, i4, i3, i4, 2000, 2000
  LineCalc i4, i1, i4, i1, 2000, 2000
  ' Draw Bottom Face
  LineCalc i1, i1, i1, i1, 1000, 2000
  LineCalc i2, i2, i2, i2, 1000, 2000
  LineCalc i3, i3, i3, i3, 1000, 2000
  LineCalc i4, i4, i4, i4, 1000, 2000
End Function
```

