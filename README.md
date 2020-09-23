<div align="center">

## Another way to place images within menu


</div>

### Description

Places images in the menu
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sean Street](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-street.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sean-street-another-way-to-place-images-within-menu__1-7942/archive/master.zip)

### API Declarations

```
Declare Function GetMenu Lib "user32" _
(ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
ByVal hBitmapChecked As Long) As Long
Public Const MF_BITMAP = &H4&
Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type
Declare Function GetMenuItemCount Lib "user32" _
(ByVal hMenu As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" _
Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
ByVal un As Long, ByVal b As Boolean, _
lpMenuItemInfo As MENUITEMINFO) As Boolean
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&
```


### Source Code

```
'	Add one form to the project.
'	Add a picturebox (Autosize = True) with a bitmap (not an icon!!!), max. 13X13
'	Add a commandbutton with following code:
Private Sub Form_Load()
hMenu& = GetMenu(Form1.hwnd)
hSubMenu& = GetSubMenu(hMenu&, 0)
hID& = GetMenuItemID(hSubMenu&, 0)
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
Picture1.Picture, _
Picture1.Picture
End Sub
```

