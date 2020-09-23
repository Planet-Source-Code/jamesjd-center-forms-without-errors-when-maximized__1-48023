<div align="center">

## \[ Center Forms Without Errors When Maximized \]


</div>

### Description

This will center forms and when you maximize the form the Command Button which you click to re-center it will be disabled.
 
### More Info
 
You will need: One Command Button (you just need to keep there default names)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[JamesJD](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jamesjd.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jamesjd-center-forms-without-errors-when-maximized__1-48023/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
If Form1.WindowState = 0 Then
Form1.Left = Screen.Width / 2 - Form1.Width / 2
Form1.Top = Screen.Height / 2 - Form1.Height / 2
Else
Command1.Enabled = False
End If
End Sub
Private Sub Form_Load()
Form1.Left = Screen.Width / 2 - Form1.Width / 2
Form1.Top = Screen.Height / 2 - Form1.Height / 2
End Sub
Private Sub Form_Resize()
If Form1.WindowState = 2 Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub
```

