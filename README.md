<div align="center">

## Percent Bar \- Transforms a PictureBox into a Percent Bar\!


</div>

### Description

I thought this code might help those struggling to make a percent bar. This code transforms a normal picturebox into a Percent Bar (progress bar with % in the middle). Choose from 2 different borders.
 
### More Info
 
pic As PictureBox, ByVal Percent As Integer, Optional ByVal Flat As Boolean = False


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\_andy\_](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andy.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andy-percent-bar-transforms-a-picturebox-into-a-percent-bar__1-48265/archive/master.zip)





### Source Code

```
Public Sub UpdateProgress(pic As PictureBox, ByVal Percent As Single, Optional ByVal Flat As Boolean = False)
 With pic
 'Configure PictureBox
 .AutoRedraw = True
 .Appearance = Flat + 1
 .ScaleWidth = 100
 .ForeColor = vbHighlight
 .BackColor = vbButtonFace
 .DrawMode = vbNotXorPen
 'Clear the PictureBox
 .Cls
 'Draw the text
 .CurrentX = (.ScaleWidth - .TextWidth(Int(Percent) & "%")) \ 2
 .CurrentY = (.ScaleHeight - .TextHeight(Int(Percent) & "%")) \ 2
 pic.Print Int(Percent) & "%"
 'Draw the progress
 pic.Line (0, 0)-(Percent, .ScaleHeight), , BF
 End With
End Sub
```

