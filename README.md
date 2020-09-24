<div align="center">

## Add Commands to Menus in VB \(AddIn Creator\)


</div>

### Description

Hey, you always have seen some people adding some cool addin to VB which esase their programming right? now you can do it too...You can add commands, command buttons, to menus with any command you like...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VbNick](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vbnick.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vbnick-add-commands-to-menus-in-vb-addin-creator__1-11796/archive/master.zip)





### Source Code

```
' by Kayhan Tanriseven  The Benchmarker®
'
' Example code of to add menu items to VB's popup Menus
'
' If needed, I will post a sample zipped project also..
' for this reason, please feedback..
' create all the user interface items
On Error GoTo CreateMenuItems_Error
' create the menu items in the code window and code break window
With VBInstance.CommandBars("Code Window").Controls
	Set MenuItem1 = .Add(msoControlButton)
	MenuItem1.Caption = "&Append To Clipboard"
	MenuItem1.BeginGroup = True
	Set MenuHandler1 = 	VBInstance.Events.CommandBarEvents(MenuItem1)
	Set MenuItem2 = .Add(msoControlButton)
	MenuItem2.Caption = "Clipboard &History"
	Set MenuHandler2 = 	VBInstance.Events.CommandBarEvents(MenuItem2)
End With
With VBInstance.CommandBars("Code Window (Break)").Controls
	Set MenuItem3 = .Add(msoControlButton)
	MenuItem1.Caption = "&Append To Clipboard"
	MenuItem1.BeginGroup = True
	Set MenuHandler3 = 	VBInstance.Events.CommandBarEvents(MenuItem3)
	Set MenuItem4 = .Add(msoControlButton)
	MenuItem4.Caption = "Clipboard &History"
	Set MenuHandler4 = 	VBInstance.Events.CommandBarEvents(MenuItem4)
End With
Exit Sub
CreateMenuItems_Error:
MsgBox "Unable To create necessary menu items", vbCritical
```

