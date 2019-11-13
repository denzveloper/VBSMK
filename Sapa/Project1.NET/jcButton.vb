Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class jcbutton
	Inherits System.Windows.Forms.UserControl
	Public Event PicturePushOnHoverChange()
	Public Event DropDownSymbolChange()
	Public Event MaskColorChange()
	Public Event ToolTipTypeChange()
	Public Event PictureOpacityOnOverChange()
	Public Event PictureEffectOnOverChange()
	Public Event ColorSchemeChange()
	Public Event ValueChange()
	Public Event CaptionAlignChange()
	Public Event PictureAlignChange()
	Public Event PictureEffectOnDownChange()
	Public Event UseMaskColorChange()
	Public Event CaptionEffectsChange()
	Public Event RightToLeftChange()
	Public Event MouseIconChange()
	Public Event MousePointerChange()
	Public Event DisabledPictureModeChange()
	Public Event PictureHotChange()
	Public Event BackColorChange()
	Public Event ForeColorChange()
	Public Event TooltipBackcolorChange()
	Public Event ShowFocusRectChange()
	Public Event PictureShadowChange()
	Public Event ToolTipChange()
	Public Event CaptionChange()
	Public Event ButtonStyleChange()
	Public Event HandPointerChange()
	Public Event FontChange()
	Public Event PictureDownChange()
	Public Event PictureNormalChange()
	Public Event PictureOpacityChange()
	Public Event DropDownSeparatorChange()
	Public Event TooltipIconChange()
	Public Event ForeColorHoverChange()
	Public Event TooltipTitleChange()
	Public Event EnabledChange()
	Public Event ModeChange()
	
	
	'***************************************************************************
	'*  Title:      JC button
	'*  Function:   An ownerdrawn multistyle button
	'*  Author:     Juned Chhipa
	'*  Created:    November 2008
	'*  Dedicated:  To my Parents and my Teachers :-)
	'*  Contact me: juned.chhipa@yahoo.com
	'*
	'*  Copyright © 2008-2009 Juned Chhipa. All rights reserved.
	'****************************************************************************
	'* This control can be used as an alternative to Command Button. It is      *
	'* a lightweight button control which will emulate new button styles.       *
	'*                                                                          *
	'* This control uses self-subclassing routines of Paul Caton.               *
	'* Feel free to use this control. Please read Documentation.chm             *
	'* Please send comments/suggestions/bug reports to mail address stated above*
	'****************************************************************************
	'*
	'* - CREDITS:
	'* - Dana Seaman :-  Worked much for this control (Thanks a million)
	'* - Paul Caton  :-  Self-Subclass Routines
	'* - Noel Dacara :-  Inspiration for DropDown menu support
	'* - Tuan Hai    :-  Numerous Suggestions and appreciating me ;)
	'* - Fred.CPP    :-  For the amazing Aqua Style and for flexible tooltips
	'* - Gonkuchi    :-  For his sub TransBlt to make grayscale pictures
	'* - Carles P.V. :-  For fastest gradient routines
	'*
	'* I have tested this control painstakingly and tried my best to make
	'* it work as a real command button.
	'*
	'****************************************************************************
	'* This software is provided "as-is" without any express/implied warranty.  *
	'* In no event shall the author be held liable for any damages arising      *
	'* from the use of this software.                                           *
	'* If you do not agree with these terms, do not install "JCButton". Use     *
	'* of the program implicitly means you have agreed to these terms.          *
	'*                                                                          *
	'* Permission is granted to anyone to use this software for any purpose,    *
	'* including commercial use, and to alter and redistribute it, provided     *
	'* that the following conditions are met:                                   *
	'*                                                                          *
	'* 1.All redistributions of source code files must retain all copyright     *
	'*   notices that are currently in place, and this list of conditions       *
	'*   without any modification.                                              *
	'*                                                                          *
	'* 2.All redistributions in binary form must retain all occurrences of      *
	'*   above copyright notice and web site addresses that are currently in    *
	'*   place (for example, in the About boxes).                               *
	'*                                                                          *
	'* 3.Modified versions in source or binary form must be plainly marked as   *
	'*   such, and must not be misrepresented as being the original software.   *
	'****************************************************************************
	
	'* N'joy ;)
	
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure POINT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByRef lpPoint As POINT) As Integer
	'UPGRADE_WARNING: Structure BITMAPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Integer, ByVal hBitmap As Integer, ByVal nStartScan As Integer, ByVal nNumScans As Integer, ByRef lpBits As Any, ByRef lpbi As BITMAPINFO, ByVal wUsage As Integer) As Integer
	'UPGRADE_WARNING: Structure BITMAPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal Scan As Integer, ByVal NumScans As Integer, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal wSrcWidth As Integer, ByVal wSrcHeight As Integer, ByRef lpBits As Any, ByRef lpBitsInfo As Any, ByVal wUsage As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Integer) As Integer
	Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Integer) As Integer
	Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Integer) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer
	Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Integer) As Integer
	Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
	Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Integer) As Integer
	Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Integer, ByVal crColor As Integer) As Integer
	Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Integer, ByVal hPalette As Integer, ByRef pccolorref As Integer) As Integer
	Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Integer, ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure tLogFont may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateFontIndirect Lib "gdi32"  Alias "CreateFontIndirectA"(ByRef lpLogFont As tLogFont) As Integer
	'UPGRADE_NOTE: GetObject was upgraded to GetObject_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetObject_Renamed Lib "gdi32.dll"  Alias "GetObjectA"(ByVal hObject As Integer, ByVal nCount As Integer, ByRef lpObject As Any) As Integer
	
	'User32 Declares
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OffsetRect Lib "user32" (ByRef lpRect As RECT, ByVal X As Integer, ByVal Y As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CopyRect Lib "user32" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Integer, ByRef qrc As RECT, ByVal edge As Integer, ByVal grfFlags As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Integer, ByRef lpRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FrameRect Lib "user32" (ByVal hDC As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	Private Declare Function TransparentBlt Lib "MSIMG32.dll" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal nSrcWidth As Integer, ByVal nSrcHeight As Integer, ByVal crTransparent As Integer) As Boolean
	Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal nSrcWidth As Integer, ByVal nSrcHeight As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Integer, ByVal yPoint As Integer) As Integer
	'UPGRADE_WARNING: Structure POINT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Integer
	Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Integer, ByVal hrgn As Integer, ByVal bRedraw As Boolean) As Integer
	Private Declare Function LoadCursor Lib "user32.dll"  Alias "LoadCursorA"(ByVal hInstance As Integer, ByVal lpCursorName As Integer) As Integer
	Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Integer) As Integer
	Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	Private Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	
	' --for tooltips
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function CreateWindowEx Lib "user32"  Alias "CreateWindowExA"(ByVal dwExStyle As Integer, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hWndParent As Integer, ByVal hMenu As Integer, ByVal hInstance As Integer, ByRef lpParam As Any) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function GetClassLong Lib "user32"  Alias "GetClassLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
	Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
	Private Declare Function SetClassLong Lib "user32"  Alias "SetClassLongA"(ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	
	' --Theme Stuff
	Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Integer, ByVal pszClassList As Integer) As Integer
	Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Integer, ByVal lhDC As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pRect As RECT, ByRef pClipRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetThemeBackgroundRegion Lib "uxtheme.dll" (ByVal hTheme As Integer, ByVal hDC As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pRect As RECT, ByRef pRegion As Integer) As Integer
	Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Integer
	
	Private Declare Function ReleaseCapture Lib "user32.dll" () As Integer
	Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function GetCapture Lib "user32.dll" () As Integer
	
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawText Lib "user32"  Alias "DrawTextA"(ByVal hDC As Integer, ByVal lpStr As String, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Integer, ByVal lpStr As Integer, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Integer
	Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Integer, ByVal xLeft As Integer, ByVal yTop As Integer, ByVal hIcon As Integer, ByVal cxWidth As Integer, ByVal cyWidth As Integer, ByVal istepIfAniCur As Integer, ByVal hbrFlickerFreeDraw As Integer, ByVal diFlags As Integer) As Integer
	Private Declare Function SetLayout Lib "gdi32" (ByVal hDC As Integer, ByVal dwLayout As Integer) As Integer
	
	'==========================================================================================================================================================================================================================================================================================
	' Subclassing Declares
	Private Enum MsgWhen
		MSG_AFTER = 1 'Message calls back after the original (previous) WndProc
		MSG_BEFORE = 2 'Message calls back before the original (previous) WndProc
		MSG_BEFORE_AND_AFTER = MsgWhen.MSG_AFTER Or MsgWhen.MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
	End Enum
	
	Private Enum TRACKMOUSEEVENT_FLAGS
		TME_HOVER = &H1s
		TME_LEAVE = &H2s
		TME_QUERY = &H40000000
		TME_CANCEL = &H80000000
	End Enum
	
	'Windows Messages
	Private Const WM_MOUSELEAVE As Integer = &H2A3s
	Private Const WM_THEMECHANGED As Integer = &H31As
	Private Const WM_SYSCOLORCHANGE As Integer = &H15s
	Private Const WM_MOVING As Integer = &H216s
	Private Const WM_NCACTIVATE As Integer = &H86s
	Private Const WM_ACTIVATE As Integer = &H6s
	
	Private Const ALL_MESSAGES As Integer = -1 'All messages added or deleted
	Private Const GMEM_FIXED As Integer = 0 'Fixed memory GlobalAlloc flag
	Private Const GWL_WNDPROC As Integer = -4 'Get/SetWindow offset to the WndProc procedure address
	Private Const PATCH_04 As Integer = 88 'Table B (before) address patch offset
	Private Const PATCH_05 As Integer = 93 'Table B (before) entry count patch offset
	Private Const PATCH_08 As Integer = 132 'Table A (after) address patch offset
	Private Const PATCH_09 As Integer = 137 'Table A (after) entry count patch offset
	
	Private Structure TRACKMOUSEEVENT_STRUCT
		Dim cbSize As Integer
		Dim dwFlags As TRACKMOUSEEVENT_FLAGS
		Dim hwndTrack As Integer
		Dim dwHoverTime As Integer
	End Structure
	
	'for subclass
	Private Structure SubClassDatatype
		Dim hwnd As Integer
		Dim nAddrSclass As Integer
		Dim nAddrOrig As Integer
		Dim nMsgCountA As Integer
		Dim nMsgCountB As Integer
		Dim aMsgTabelA() As Integer
		Dim aMsgTabelB() As Integer
	End Structure
	
	'for subclass
	Private SubclassData() As SubClassDatatype 'Subclass data array
	Private TrackUser32 As Boolean
	
	'Kernel32 declares used by the Subclasser
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "KERNEL32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub RtlMoveMemory Lib "KERNEL32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Private Declare Function GetModuleHandleA Lib "KERNEL32" (ByVal lpModuleName As String) As Integer
	Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Integer) As Integer
	Private Declare Function LoadLibraryA Lib "KERNEL32" (ByVal lpLibFileName As String) As Integer
	'UPGRADE_WARNING: Structure TRACKMOUSEEVENT_STRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Integer
	'UPGRADE_WARNING: Structure TRACKMOUSEEVENT_STRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function TrackMouseEventComCtl Lib "Comctl32"  Alias "_TrackMouseEvent"(ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Integer
	Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
	Private Declare Function GetModuleHandle Lib "KERNEL32"  Alias "GetModuleHandleA"(ByVal lpModuleName As String) As Integer
	Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
	Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Integer) As Integer
	Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	'UPGRADE_WARNING: Structure OSVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetVersionEx Lib "KERNEL32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFO) As Integer
	
	'  End of Subclassing Declares
	'==========================================================================================================================================================================================================================================================================================================
	
	'[Enumerations]
	Public Enum enumButtonStlyes
		eStandard '1) Standard VB Button
		eFlat '2) Standard Toolbar Button
		eWindowsXP '3) Famous Win XP Button
		eVistaAero '5) The New Vista Aero Button
		eOfficeXP
		eOffice2003 '13) Office 2003 Style
		eXPToolbar '4) XP Toolbar
		eVistaToolbar '9) Vista Toolbar Button
		eOutlook2007 '8) Office 2007 Outlook Button
		eInstallShield '7) InstallShield?!?~?
		eGelButton '11) Gel Button
		e3DHover '13) 3D Hover Button
		eFlatHover '12) Flat Hover Button
		eWindowsTheme
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private eStandard, eFlat, eVistaAero, eVistaToolbar, eInstallShield, eFlatHover, eOffice2003
	Private eWindowsXP, eXPToolbar, e3DHover, eGelButton, eOutlook2007, eOfficeXP, eWindowsTheme
#End If
	
	Public Enum enumButtonModes
		ebmCommandButton
		ebmCheckBox
		ebmOptionButton
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private ebmCommandButton, ebmCheckBox, ebmOptionButton
#End If
	
	Public Enum enumButtonStates
		eStateNormal 'Normal State
		eStateOver 'Hover State
		eStateDown 'Down State
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	'A trick to preserve casing when typing in IDE
	Private eStateNormal, eStateOver, eStateDown, eStateFocused
#End If
	
	Public Enum enumCaptionAlign
		ecLeftAlign
		ecCenterAlign
		ecRightAlign
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	'A trick to preserve casing when typing in IDE
	Private ecLeftAlign, ecCenterAlign, ecRightAlign
#End If
	
	Public Enum enumPictureAlign
		epLeftEdge
		epLeftOfCaption
		epRightEdge
		epRightOfCaption
		epBackGround
		epTopEdge
		epTopOfCaption
		epBottomEdge
		epBottomOfCaption
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private epLeftEdge, epRightEdge, epRightOfCaption, epLeftOfCaption, epBackGround
	Private epTopEdge, epTopOfCaption, epBottomEdge, epBottomOfCaption
#End If
	
	' --Tooltip Icons
	Public Enum enumIconType
		TTNoIcon
		TTIconInfo
		TTIconWarning
		TTIconError
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private TTNoIcon, TTIconInfo, TTIconWarning, TTIconError
#End If
	
	' --Tooltip [ Balloon / Standard ]
	Public Enum enumTooltipStyle
		TooltipStandard
		TooltipBalloon
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private TooltipStandard, TooltipBalloon
#End If
	
	' --Caption effects
	Public Enum enumCaptionEffects
		eseNone
		eseEmbossed
		eseEngraved
		eseShadowed
		eseOutline
		eseCover
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private eseNone, eseEmbossed, eseEngraved, eseShadowed, eseOutline, eseCover
#End If
	
	Public Enum enumPicEffect
		epeNone
		epeLighter
		epeDarker
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private epeNone, epeLighter, epeDarker, epePushUp
#End If
	
	' --For dropdown symbols
	Public Enum enumSymbol
		ebsNone
		ebsArrowUp = 5
		ebsArrowDown = 6
		ebsArrowRight = 4
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private ebsArrowUp, ebsArrowDown, ebsNone
#End If
	
	Public Enum enumXPThemeColors
		ecsBlue = 0
		ecsOliveGreen = 1
		ecsSilver = 2
		ecsCustom = 3
	End Enum
	
	' --A trick to preserve casing of enums while typing in IDE
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private ecsBlue, ecsOliveGreen, ecsSilver, ecsCustom
#End If
	
	Public Enum enumDisabledPicMode
		edpBlended
		edpGrayed
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private edpBlended, edpGrayed
#End If
	
	' --For gradient subs
	Public Enum GradientDirectionCts
		gdHorizontal = 0
		gdVertical = 1
		gdDownwardDiagonal = 2
		gdUpwardDiagonal = 3
	End Enum
	
	' --A trick to preserve casing of enums when typing in IDE
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private gdHorizontal, gdVertical, gdDownwardDiagonal, gdUpwardDiagonal
#End If
	
	Public Enum enumMenuAlign
		edaBottom = 0
		edaTop = 1
		edaLeft = 2
		edaRight = 3
		edaTopLeft = 4
		edaBottomLeft = 5
		edaTopRight = 6
		edaBottomRight = 7
	End Enum
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private edaBottom, edaTop, edaTopLeft, edaBottomLeft, edaTopRight, edaBottomRight
#End If
	
	'  used for Button colors
	Private Structure tButtonColors
		Dim tBackColor As Integer
		Dim tDisabledColor As Integer
		Dim tForeColor As Integer
		Dim tForeColorOver As Integer
		Dim tGreyText As Integer
	End Structure
	
	'  used to define various graphics areas
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Top_Renamed As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		'UPGRADE_NOTE: Bottom was upgraded to Bottom_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Bottom_Renamed As Integer
	End Structure
	
	''Tooltip Window Types
	Private Structure TOOLINFO
		Dim lSize As Integer
		Dim lFlags As Integer
		Dim lhWnd As Integer
		Dim lId As Integer
		Dim lpRect As RECT
		Dim hInstance As Integer
		Dim lpStr As String
		Dim lParam As Integer
	End Structure
	
	''Tooltip Window Types [for UNICODE support]
	Private Structure TOOLINFOW
		Dim lSize As Integer
		Dim lFlags As Integer
		Dim lhWnd As Integer
		Dim lId As Integer
		Dim lpRect As RECT
		Dim hInstance As Integer
		Dim lpStrW As Integer
		Dim lParam As Integer
	End Structure
	
	Private Structure POINT
		Dim X As Integer
		Dim Y As Integer
	End Structure
	
	' --Used for creating a drop down symbol
	' --I m using Marlett Font to create that symbol
	Private Structure tLogFont
		Dim lfHeight As Integer
		Dim lfWidth As Integer
		Dim lfEscapement As Integer
		Dim lfOrientation As Integer
		Dim lfWeight As Integer
		Dim lfItalic As Byte
		Dim lfUnderline As Byte
		Dim lfStrikeOut As Byte
		Dim lfCharSet As Byte
		Dim lfOutPrecision As Byte
		Dim lfClipPrecision As Byte
		Dim lfQuality As Byte
		Dim lfPitchAndFamily As Byte
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(32),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=32)> Public lfFaceName() As Char
	End Structure
	
	'  RGB Colors structure
	Private Structure RGBColor
		Dim R As Single
		Dim G As Single
		Dim B As Single
	End Structure
	
	Private Structure BITMAP
		Dim bmType As Integer
		Dim bmWidth As Integer
		Dim bmHeight As Integer
		Dim bmWidthBytes As Integer
		Dim bmPlanes As Short
		Dim bmBitsPixel As Short
		Dim bmBits As Integer
	End Structure
	
	'  for gradient painting and bitmap tiling
	Private Structure BITMAPINFOHEADER
		Dim biSize As Integer
		Dim biWidth As Integer
		Dim biHeight As Integer
		Dim biPlanes As Short
		Dim biBitCount As Short
		Dim biCompression As Integer
		Dim biSizeImage As Integer
		Dim biXPelsPerMeter As Integer
		Dim biYPelsPerMeter As Integer
		Dim biClrUsed As Integer
		Dim biClrImportant As Integer
	End Structure
	
	Private Structure ICONINFO
		Dim fIcon As Integer
		Dim xHotspot As Integer
		Dim yHotspot As Integer
		Dim hbmMask As Integer
		Dim hbmColor As Integer
	End Structure
	
	Private Structure RGBTRIPLE
		Dim rgbBlue As Byte
		Dim rgbGreen As Byte
		Dim rgbRed As Byte
	End Structure
	
	Private Structure RGBQUAD
		Dim rgbBlue As Byte
		Dim rgbGreen As Byte
		Dim rgbRed As Byte
		Dim rgbAlpha As Byte
	End Structure
	
	Private Structure BITMAPINFO
		Dim bmiHeader As BITMAPINFOHEADER
		Dim bmiColors As RGBTRIPLE
	End Structure
	
	Private Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformId As Integer
		'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char '* Maintenance string for PSS usage.
	End Structure
	
	' --constants for unicode support
	Private Const VER_PLATFORM_WIN32_NT As Short = 2
	
	' --constants for  Flat Button
	Private Const BDR_RAISEDINNER As Integer = &H4s
	
	' --constants for Win 98 style buttons
	Private Const BDR_SUNKEN95 As Integer = &HAs
	Private Const BDR_RAISED95 As Integer = &H5s
	
	Private Const BF_LEFT As Integer = &H1s
	Private Const BF_TOP As Integer = &H2s
	Private Const BF_RIGHT As Integer = &H4s
	Private Const BF_BOTTOM As Integer = &H8s
	Private Const BF_RECT As Integer = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
	
	' --System Hand Pointer
	Private Const IDC_HAND As Integer = 32649
	
	' --Color Constant
	Private Const COLOR_BTNFACE As Integer = 15
	Private Const COLOR_BTNHIGHLIGHT As Integer = 20
	Private Const COLOR_BTNSHADOW As Integer = 16
	Private Const COLOR_HIGHLIGHT As Integer = 13
	Private Const COLOR_GRAYTEXT As Integer = 17
	Private Const CLR_INVALID As Integer = &HFFFFs
	Private Const DIB_RGB_COLORS As Integer = 0
	
	' --Windows Messages
	Private Const WM_USER As Integer = &H400s
	Private Const GWL_STYLE As Integer = -16
	Private Const WS_CAPTION As Integer = &HC00000
	Private Const WS_THICKFRAME As Integer = &H40000
	Private Const WS_MINIMIZEBOX As Integer = &H20000
	Private Const SWP_REFRESH As Integer = (&H1s Or &H2s Or &H4s Or &H20s)
	Private Const SWP_NOACTIVATE As Integer = &H10s
	Private Const SWP_NOMOVE As Integer = &H2s
	Private Const SWP_NOSIZE As Integer = &H1s
	Private Const SWP_SHOWWINDOW As Integer = &H40s
	Private Const HWND_TOPMOST As Integer = -&H1s
	Private Const CW_USEDEFAULT As Integer = &H80000000
	
	''Tooltip Window Constants
	Private Const TTS_NOPREFIX As Integer = &H2s
	Private Const TTF_TRANSPARENT As Integer = &H100s
	Private Const TTF_IDISHWND As Integer = &H1s
	Private Const WS_EX_LAYOUTRTL As Integer = &H400000
	Private Const TTF_CENTERTIP As Integer = &H2s
	Private Const TTM_ADDTOOLA As Integer = (WM_USER + 4)
	Private Const TTM_ADDTOOLW As Integer = (WM_USER + 50)
	Private Const TTM_ACTIVATE As Integer = WM_USER + 1
	Private Const TTM_UPDATETIPTEXTA As Integer = (WM_USER + 12)
	Private Const TTM_SETMAXTIPWIDTH As Integer = (WM_USER + 24)
	Private Const TTM_SETTIPBKCOLOR As Integer = (WM_USER + 19)
	Private Const TTM_SETTIPTEXTCOLOR As Integer = (WM_USER + 20)
	Private Const TTM_SETTITLE As Integer = (WM_USER + 32)
	Private Const TTM_SETTITLEW As Integer = (WM_USER + 33)
	Private Const TTS_BALLOON As Integer = &H40s
	Private Const TTS_ALWAYSTIP As Integer = &H1s
	Private Const TTF_SUBCLASS As Integer = &H10s
	Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"
	
	' --Formatting Text Consts
	Private Const DT_CALCRECT As Integer = &H400s
	Private Const DT_CENTER As Integer = &H1s
	Private Const DT_WORDBREAK As Integer = &H10s
	Private Const DT_RTLREADING As Integer = &H20000 ' Right to left
	Private Const DT_DRAWFLAG As Integer = DT_CENTER Or DT_WORDBREAK
	
	' --drawing Icon Constants
	Private Const DI_NORMAL As Integer = &H3s
	
	' --Property Variables:
	
	Private m_ButtonStyle As enumButtonStlyes 'Choose your Style
	Private m_Buttonstate As enumButtonStates 'Normal / Over / Down
	
	Private m_bIsDown As Boolean 'Is button is pressed?
	Private m_bMouseInCtl As Boolean 'Is Mouse in Control
	Private m_bHasFocus As Boolean 'Has focus?
	Private m_bHandPointer As Boolean 'Use Hand Pointer
	Private m_lCursor As Integer
	Private m_bDefault As Boolean 'Is Default?
	Private m_DropDownSymbol As enumSymbol
	Private m_bDropDownSep As Boolean
	Private m_ButtonMode As enumButtonModes 'Command/Check/Option button
	Private m_CaptionEffects As enumCaptionEffects
	Private m_bValue As Boolean 'Value (Checked/Unchekhed)
	Private m_bShowFocus As Boolean 'Bool to show focus
	Private m_bParentActive As Boolean 'Parent form Active or not
	Private m_lParenthWnd As Integer 'Is parent active?
	Private m_WindowsNT As Integer 'OS Supports Unicode?
	Private m_bEnabled As Boolean 'Enabled/Disabled
	Private m_Caption As String 'String to draw caption
	Private m_CaptionAlign As enumCaptionAlign
	Private m_bColors As tButtonColors 'Button Colors
	Private m_bUseMaskColor As Boolean 'Transparent areas
	Private m_lMaskColor As Integer 'Set Transparent color
	Private m_lButtonRgn As Integer 'Button Region
	Private m_bIsSpaceBarDown As Boolean 'Space bar down boolean
	Private m_ButtonRect As RECT 'Button Position
	Private m_FocusRect As RECT
	Private mFont As System.Drawing.Font
	Private m_lXPColor As enumXPThemeColors
	Private m_bIsThemed As Boolean
	Private m_bHasUxTheme As Boolean
	
	Private m_lDownButton As Short 'For click/Dblclick events
	Private m_lDShift As Short 'A flag for dblClick
	Private m_lDX As Single
	Private m_lDY As Single
	
	' --Popup menu variables
	Private m_bPopupEnabled As Boolean 'Popus is enabled
	Private m_bPopupShown As Boolean 'Popupmenu is shown
	Private m_bPopupInit As Boolean 'Flag to prevent WM_MOUSLEAVE to redraw the button
	Private DropDownMenu As System.Windows.Forms.ToolStripMenuItem 'Popupmenu to be shown
	Private MenuAlign As enumMenuAlign 'PopupMenu Alignments
	Private MenuFlags As Integer 'PopupMenu Flags
	Private DefaultMenu As System.Windows.Forms.ToolStripMenuItem 'Default menu in the popupmenu
	
	' --Tooltip variables
	Private m_sTooltipText As String
	Private m_sTooltiptitle As String
	Private m_lToolTipIcon As enumIconType
	Private m_lTooltipType As enumTooltipStyle
	Private m_lttBackColor As Integer
	Private m_lttCentered As Boolean
	Private m_bttRTL As Boolean 'Right to Left reading
	Private m_lttHwnd As Integer
	Private m_hMode As Integer 'Added this, as tooltips
	'were not displayed in
	'compiled exe. (Thanks to Jim Jose)
	' --Caption variables
	Private CaptionW As Integer 'Width of Caption
	Private CaptionH As Integer 'Height of Caption
	Private CaptionX As Integer 'Left of Caption
	Private CaptionY As Integer 'Top of Caption
	Private lpSignRect As RECT 'Drop down Symbol rect
	Private m_bRTL As Boolean
	Private m_TextRect As RECT 'Caption drawing area
	
	' --Picture variables
	Private m_Picture As System.Drawing.Image
	Private m_PictureHot As System.Drawing.Image
	Private m_PictureDown As System.Drawing.Image
	Private m_PictureOpacity As Byte
	Private m_PicOpacityOnOver As Byte
	Private m_PicDisabledMode As enumDisabledPicMode
	Private m_PictureShadow As Boolean
	Private m_PictureAlign As enumPictureAlign 'Picture Alignments
	Private m_PicEffectonOver As enumPicEffect
	Private m_PicEffectonDown As enumPicEffect
	Private m_bPicPushOnHover As Boolean
	Private PicH As Integer
	Private PicW As Integer
	Private aLighten(255) As Byte 'Light Picture
	Private aDarken(255) As Byte 'Dark Picture
	
	Private tmppic As System.Drawing.Image = New System.Drawing.Bitmap(1, 1) 'Temp picture
	Private PicX As Integer 'X position of picture
	Private PicY As Integer 'Y Position of Picture
	Private m_PicRect As RECT 'Picture drawing area
	Private lH As Integer 'ScaleHeight of button
	Private lW As Integer 'ScaleWidth of button
	
	'  Events
	Public Shadows Event Click(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Event DblClick(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event MouseEnter(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event MouseLeave(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event MouseMove(ByVal Sender As System.Object, ByVal e As MouseMoveEventArgs)
	Public Shadows Event MouseUp(ByVal Sender As System.Object, ByVal e As MouseUpEventArgs)
	Public Shadows Event MouseDown(ByVal Sender As System.Object, ByVal e As MouseDownEventArgs)
	Public Shadows Event KeyDown(ByVal Sender As System.Object, ByVal e As KeyDownEventArgs)
	Public Shadows Event KeyUp(ByVal Sender As System.Object, ByVal e As KeyUpEventArgs)
	Public Shadows Event KeyPress(ByVal Sender As System.Object, ByVal e As KeyPressEventArgs)
	
	Private Sub DrawLineApi(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Color As Integer)
		
		'****************************************************************************
		'*  draw lines
		'****************************************************************************
		
		Dim pt As POINT
		Dim hPen As Integer
		Dim hPenOld As Integer
		
		hPen = CreatePen(0, 1, Color)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		hPenOld = SelectObject(hDC, hPen)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		MoveToEx(hDC, X1, Y1, pt)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		LineTo(hDC, X2, Y2)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SelectObject(hDC, hPenOld)
		DeleteObject(hPen)
		
	End Sub
	
	Private Function BlendColorEx(ByRef Color1 As Integer, ByRef Color2 As Integer, Optional ByRef Percent As Integer = 0) As Integer
		
		'Combines two colors together by how many percent.
		
		Dim g1, r1, b1 As Integer
		Dim g2, r2, b2 As Integer
		Dim g3, r3, b3 As Integer
		
		If Percent <= 0 Then Percent = 0
		If Percent >= 100 Then Percent = 100
		
		r1 = Color1 And 255
		g1 = (Color1 \ 256) And 255
		b1 = (Color1 \ 65536) And 255
		
		r2 = Color2 And 255
		g2 = (Color2 \ 256) And 255
		b2 = (Color2 \ 65536) And 255
		
		r3 = r1 + (r1 - r2) * Percent \ 100
		g3 = g1 + (g1 - g2) * Percent \ 100
		b3 = b1 + (b1 - b2) * Percent \ 100
		
		BlendColorEx = r3 + 256 * g3 + 65536 * b3
		
	End Function
	
	Private Function BlendColors(ByVal lBackColorFrom As Integer, ByVal lBackColorTo As Integer) As Integer
		
		'***************************************************************************
		'*  Combines (mix) two colors                                              *
		'*  This is another method in which you can't specify percentage
		'***************************************************************************
		
		BlendColors = RGB(CShort(CShort(lBackColorFrom And &HFFs) + CShort(lBackColorTo And &HFFs)) / 2, CShort(CShort((lBackColorFrom \ &H100s) And &HFFs) + CShort((lBackColorTo \ &H100s) And &HFFs)) / 2, CShort(CShort((lBackColorFrom \ &H10000) And &HFFs) + CShort((lBackColorTo \ &H10000) And &HFFs)) / 2)
		
	End Function
	
	Private Sub DrawRectangle(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Color As Integer)
		
		'****************************************************************************
		'*  Draws a rectangle specified by coords and color of the rectangle        *
		'****************************************************************************
		
		Dim bRect As RECT
		Dim hBrush As Integer
		Dim ret As Integer
		
		With bRect
			.Left_Renamed = X
			.Top_Renamed = Y
			.Right_Renamed = X + Width
			.Bottom_Renamed = Y + Height
		End With
		
		hBrush = CreateSolidBrush(Color)
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		ret = FrameRect(hDC, bRect, hBrush)
		
		ret = DeleteObject(hBrush)
		
	End Sub
	
	Private Sub DrawFocusRectangle(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer)
		
		'****************************************************************************
		'*  Draws a Focus Rectangle inside button if m_bShowFocus property is True  *
		'****************************************************************************
		
		Dim bRect As RECT
		Dim RetVal As Integer
		
		With bRect
			.Left_Renamed = X
			.Top_Renamed = Y
			.Right_Renamed = X + Width
			.Bottom_Renamed = Y + Height
		End With
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		RetVal = DrawFocusRect(hDC, bRect)
		
	End Sub
	
	Private Sub TransBlt(ByVal DstDC As Integer, ByVal DstX As Integer, ByVal DstY As Integer, ByVal DstW As Integer, ByVal DstH As Integer, ByVal SrcPic As System.Drawing.Image, Optional ByVal TransColor As Integer = -1, Optional ByVal BrushColor As Integer = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False)
		
		'****************************************************************************
		'* Routine : To make transparent and grayscale images
		'* Author  : Gonkuchi
		'
		'* Modified by Dana Seaman
		'****************************************************************************
		
		Dim i, H, B, F, newW As Integer
		Dim TmpBmp, TmpDC, TmpObj As Integer
		Dim Sr2Bmp, Sr2DC, Sr2Obj As Integer
		Dim DataDest() As RGBTRIPLE
		Dim DataSrc() As RGBTRIPLE
		Dim Info As BITMAPINFO
		Dim BrushRGB As RGBTRIPLE
		Dim gCol As Integer
		Dim hOldOb As Integer
		Dim PicEffect As enumPicEffect
		Dim tObj, SrcDC, ttt As Integer
		Dim bDisOpacity As Byte
		Dim OverOpacity As Byte
		Dim a2 As Integer
		Dim a1 As Integer
		
		If DstW = 0 Or DstH = 0 Then Exit Sub
		If SrcPic Is Nothing Then Exit Sub
		
		If m_Buttonstate = enumButtonStates.eStateOver Then
			PicEffect = m_PicEffectonOver
		ElseIf m_Buttonstate = enumButtonStates.eStateDown Then 
			PicEffect = m_PicEffectonDown
		End If
		
		If Not m_bEnabled Then
			Select Case m_PicDisabledMode
				Case enumDisabledPicMode.edpBlended
					bDisOpacity = 52
				Case enumDisabledPicMode.edpGrayed
					bDisOpacity = m_PictureOpacity * 0.75
					isGreyscale = True
			End Select
		End If
		
		If m_Buttonstate = enumButtonStates.eStateOver Then
			OverOpacity = m_PicOpacityOnOver
		End If
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SrcDC = CreateCompatibleDC(hDC)
		
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property SrcPic.Width was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If DstW < 0 Then DstW = MyBase.ScaleX(SrcPic.Width, 8, MyBase.ScaleMode)
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property SrcPic.Height was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If DstH < 0 Then DstH = MyBase.ScaleY(SrcPic.Height, 8, MyBase.ScaleMode)
		
		'UPGRADE_ISSUE: Constant vbPicTypeBitmap was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Dim hBrush As Integer
		If SrcPic.Type = vbPicTypeBitmap Then 'check if it's an icon or a bitmap
			tObj = SelectObject(SrcDC, CInt(CObj(SrcPic)))
		Else
			tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
			hBrush = CreateSolidBrush(TransColor)
			'UPGRADE_ISSUE: Picture property SrcPic.Handle was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawIconEx(SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, DI_NORMAL)
			DeleteObject(hBrush)
		End If
		
		TmpDC = CreateCompatibleDC(SrcDC)
		Sr2DC = CreateCompatibleDC(SrcDC)
		TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		TmpObj = SelectObject(TmpDC, TmpBmp)
		Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
		ReDim DataDest(DstW * DstH * 3 - 1)
		ReDim DataSrc(UBound(DataDest))
		With Info.bmiHeader
			.biSize = Len(Info.bmiHeader)
			.biWidth = DstW
			.biHeight = DstH
			.biPlanes = 1
			.biBitCount = 24
		End With
		
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy)
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy)
		'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDIBits(TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object DataSrc(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDIBits(Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0)
		
		If BrushColor > 0 Then
			BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100s
			BrushRGB.rgbGreen = (BrushColor \ &H100s) Mod &H100s
			BrushRGB.rgbRed = BrushColor And &HFFs
		End If
		
		' --No Maskcolor to use
		If Not m_bUseMaskColor Then TransColor = -1
		
		newW = DstW - 1
		
		For H = 0 To DstH - 1
			F = H * DstW
			For B = 0 To newW
				i = F + B
				If m_Buttonstate = enumButtonStates.eStateOver Then
					a1 = OverOpacity
				Else
					a1 = IIf(m_bEnabled, m_PictureOpacity, bDisOpacity)
				End If
				a2 = 255 - a1
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If GetNearestColor(hDC, CInt(DataSrc(i).rgbRed) + 256 * DataSrc(i).rgbGreen + 65536 * DataSrc(i).rgbBlue) <> TransColor Then
					With DataDest(i)
						If BrushColor > -1 Then
							If MonoMask Then
								'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(i). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If (CInt(DataSrc(i).rgbRed) + DataSrc(i).rgbGreen + DataSrc(i).rgbBlue) <= 384 Then DataDest(i) = BrushRGB
							Else
								If a1 = 255 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(i). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									DataDest(i) = BrushRGB
								ElseIf a1 > 0 Then 
									.rgbRed = (a2 * .rgbRed + a1 * BrushRGB.rgbRed) \ 256
									.rgbGreen = (a2 * .rgbGreen + a1 * BrushRGB.rgbGreen) \ 256
									.rgbBlue = (a2 * .rgbBlue + a1 * BrushRGB.rgbBlue) \ 256
								End If
							End If
						Else
							If isGreyscale Then
								gCol = CInt(DataSrc(i).rgbRed * 0.3) + DataSrc(i).rgbGreen * 0.59 + DataSrc(i).rgbBlue * 0.11
								If a1 = 255 Then
									.rgbRed = gCol : .rgbGreen = gCol : .rgbBlue = gCol
								ElseIf a1 > 0 Then 
									.rgbRed = (a2 * .rgbRed + a1 * gCol) \ 256
									.rgbGreen = (a2 * .rgbGreen + a1 * gCol) \ 256
									.rgbBlue = (a2 * .rgbBlue + a1 * gCol) \ 256
								End If
							Else
								If a1 = 255 Then
									If PicEffect = enumPicEffect.epeLighter Then
										.rgbRed = aLighten(DataSrc(i).rgbRed)
										.rgbGreen = aLighten(DataSrc(i).rgbGreen)
										.rgbBlue = aLighten(DataSrc(i).rgbBlue)
									ElseIf PicEffect = enumPicEffect.epeDarker Then 
										.rgbRed = aDarken(DataSrc(i).rgbRed)
										.rgbGreen = aDarken(DataSrc(i).rgbGreen)
										.rgbBlue = aDarken(DataSrc(i).rgbBlue)
									Else
										'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(i). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										DataDest(i) = DataSrc(i)
									End If
								ElseIf a1 > 0 Then 
									If (PicEffect = enumPicEffect.epeLighter) Then
										.rgbRed = (a2 * .rgbRed + a1 * aLighten(DataSrc(i).rgbRed)) \ 256
										.rgbGreen = (a2 * .rgbGreen + a1 * aLighten(DataSrc(i).rgbGreen)) \ 256
										.rgbBlue = (a2 * .rgbBlue + a1 * aLighten(DataSrc(i).rgbBlue)) \ 256
									ElseIf PicEffect = enumPicEffect.epeDarker Then 
										.rgbRed = (a2 * .rgbRed + a1 * aDarken(DataSrc(i).rgbRed)) \ 256
										.rgbGreen = (a2 * .rgbGreen + a1 * aDarken(DataSrc(i).rgbGreen)) \ 256
										.rgbBlue = (a2 * .rgbBlue + a1 * aDarken(DataSrc(i).rgbBlue)) \ 256
									Else
										.rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
										.rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
										.rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
									End If
								End If
							End If
						End If
					End With
				End If
			Next B
		Next H
		
		' /--Paint it!
		'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetDIBitsToDevice(DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0)
		
		Erase DataDest
		Erase DataSrc
		DeleteObject(SelectObject(TmpDC, TmpObj))
		DeleteObject(SelectObject(Sr2DC, Sr2Obj))
		'UPGRADE_ISSUE: Constant vbPicTypeIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If SrcPic.Type = vbPicTypeIcon Then DeleteObject(SelectObject(SrcDC, tObj))
		DeleteDC(TmpDC)
		DeleteDC(Sr2DC)
		DeleteObject(tObj)
		DeleteDC(SrcDC)
		
	End Sub
	
	Private Sub TransBlt32(ByVal DstDC As Integer, ByVal DstX As Integer, ByVal DstY As Integer, ByVal DstW As Integer, ByVal DstH As Integer, ByVal SrcPic As System.Drawing.Image, Optional ByVal BrushColor As Integer = -1, Optional ByVal isGreyscale As Boolean = False)
		
		'****************************************************************************
		'* Routine : Renders 32 bit Bitmap                                          *
		'* Author  : Dana Seaman                                                    *
		'****************************************************************************
		
		Dim i, H, B, F, newW As Integer
		Dim TmpBmp, TmpDC, TmpObj As Integer
		Dim Sr2Bmp, Sr2DC, Sr2Obj As Integer
		Dim DataDest() As RGBQUAD
		Dim DataSrc() As RGBQUAD
		Dim Info As BITMAPINFO
		Dim BrushRGB As RGBQUAD
		Dim gCol As Integer
		Dim hOldOb As Integer
		Dim PicEffect As enumPicEffect
		Dim tObj, SrcDC, ttt As Integer
		Dim bDisOpacity As Byte
		Dim OverOpacity As Byte
		Dim a2 As Integer
		Dim a1 As Integer
		
		If DstW = 0 Or DstH = 0 Then Exit Sub
		If SrcPic Is Nothing Then Exit Sub
		
		If m_Buttonstate = enumButtonStates.eStateOver Then
			PicEffect = m_PicEffectonOver
		ElseIf m_Buttonstate = enumButtonStates.eStateDown Then 
			PicEffect = m_PicEffectonDown
		End If
		
		If Not m_bEnabled Then
			Select Case m_PicDisabledMode
				Case enumDisabledPicMode.edpBlended
					bDisOpacity = 52
				Case enumDisabledPicMode.edpGrayed
					bDisOpacity = m_PictureOpacity * 0.75
					isGreyscale = True
			End Select
		End If
		
		If m_Buttonstate = enumButtonStates.eStateOver Then
			OverOpacity = m_PicOpacityOnOver
		End If
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SrcDC = CreateCompatibleDC(hDC)
		
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property SrcPic.Width was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If DstW < 0 Then DstW = MyBase.ScaleX(SrcPic.Width, 8, MyBase.ScaleMode)
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Picture property SrcPic.Height was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If DstH < 0 Then DstH = MyBase.ScaleY(SrcPic.Height, 8, MyBase.ScaleMode)
		
		tObj = SelectObject(SrcDC, CInt(CObj(SrcPic)))
		
		TmpDC = CreateCompatibleDC(SrcDC)
		Sr2DC = CreateCompatibleDC(SrcDC)
		
		TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		TmpObj = SelectObject(TmpDC, TmpBmp)
		Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
		
		With Info.bmiHeader
			.biSize = Len(Info.bmiHeader)
			.biWidth = DstW
			.biHeight = DstH
			.biPlanes = 1
			.biBitCount = 32
			.biSizeImage = 4 * ((DstW * .biBitCount + 31) \ 32) * DstH
		End With
		ReDim DataDest(Info.bmiHeader.biSizeImage - 1)
		ReDim DataSrc(UBound(DataDest))
		
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy)
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		BitBlt(Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy)
		'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDIBits(TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object DataSrc(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetDIBits(Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0)
		
		If BrushColor <> -1 Then
			BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100s
			BrushRGB.rgbGreen = (BrushColor \ &H100s) Mod &H100s
			BrushRGB.rgbRed = BrushColor And &HFFs
		End If
		
		newW = DstW - 1
		
		For H = 0 To DstH - 1
			F = H * DstW
			For B = 0 To newW
				i = F + B
				If m_bEnabled Then
					If m_Buttonstate = enumButtonStates.eStateOver Then
						a1 = (CInt(DataSrc(i).rgbAlpha) * OverOpacity) \ 255
					Else
						a1 = (CInt(DataSrc(i).rgbAlpha) * m_PictureOpacity) \ 255
					End If
				Else
					a1 = (CInt(DataSrc(i).rgbAlpha) * bDisOpacity) \ 255
				End If
				a2 = 255 - a1
				With DataDest(i)
					If BrushColor <> -1 Then
						If a1 = 255 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(i). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							DataDest(i) = BrushRGB
						ElseIf a1 > 0 Then 
							.rgbRed = (a2 * .rgbRed + a1 * BrushRGB.rgbRed) \ 256
							.rgbGreen = (a2 * .rgbGreen + a1 * BrushRGB.rgbGreen) \ 256
							.rgbBlue = (a2 * .rgbBlue + a1 * BrushRGB.rgbBlue) \ 256
						End If
					Else
						If isGreyscale Then
							gCol = CInt(DataSrc(i).rgbRed * 0.3) + DataSrc(i).rgbGreen * 0.59 + DataSrc(i).rgbBlue * 0.11
							If a1 = 255 Then
								.rgbRed = gCol : .rgbGreen = gCol : .rgbBlue = gCol
							ElseIf a1 > 0 Then 
								.rgbRed = (a2 * .rgbRed + a1 * gCol) \ 256
								.rgbGreen = (a2 * .rgbGreen + a1 * gCol) \ 256
								.rgbBlue = (a2 * .rgbBlue + a1 * gCol) \ 256
							End If
						Else
							If a1 = 255 Then
								If (PicEffect = enumPicEffect.epeLighter) Then
									.rgbRed = aLighten(DataSrc(i).rgbRed)
									.rgbGreen = aLighten(DataSrc(i).rgbGreen)
									.rgbBlue = aLighten(DataSrc(i).rgbBlue)
								ElseIf PicEffect = enumPicEffect.epeDarker Then 
									.rgbRed = aDarken(DataSrc(i).rgbRed)
									.rgbGreen = aDarken(DataSrc(i).rgbGreen)
									.rgbBlue = aDarken(DataSrc(i).rgbBlue)
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(i). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									DataDest(i) = DataSrc(i)
								End If
							ElseIf a1 > 0 Then 
								If (PicEffect = enumPicEffect.epeLighter) Then
									.rgbRed = (a2 * .rgbRed + a1 * aLighten(DataSrc(i).rgbRed)) \ 256
									.rgbGreen = (a2 * .rgbGreen + a1 * aLighten(DataSrc(i).rgbGreen)) \ 256
									.rgbBlue = (a2 * .rgbBlue + a1 * aLighten(DataSrc(i).rgbBlue)) \ 256
								ElseIf PicEffect = enumPicEffect.epeDarker Then 
									.rgbRed = (a2 * .rgbRed + a1 * aDarken(DataSrc(i).rgbRed)) \ 256
									.rgbGreen = (a2 * .rgbGreen + a1 * aDarken(DataSrc(i).rgbGreen)) \ 256
									.rgbBlue = (a2 * .rgbBlue + a1 * aDarken(DataSrc(i).rgbBlue)) \ 256
								Else
									.rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
									.rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
									.rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
								End If
							End If
						End If
					End If
				End With
			Next B
		Next H
		
		' /--Paint it!
		'UPGRADE_WARNING: Couldn't resolve default property of object DataDest(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SetDIBitsToDevice(DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0)
		
		Erase DataDest
		Erase DataSrc
		DeleteObject(SelectObject(TmpDC, TmpObj))
		DeleteObject(SelectObject(Sr2DC, Sr2Obj))
		'UPGRADE_ISSUE: Constant vbPicTypeIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If SrcPic.Type = vbPicTypeIcon Then DeleteObject(SelectObject(SrcDC, tObj))
		DeleteDC(TmpDC)
		DeleteDC(Sr2DC)
		DeleteObject(tObj)
		DeleteDC(SrcDC)
		
	End Sub
	
	' --By Dana Seaman
	Private Function Lighten(ByVal Color As Byte) As Byte
		
		Dim lColor As Integer
		lColor = (293 * Color) \ 255
		If lColor > 255 Then
			Lighten = 255
		Else
			Lighten = lColor
		End If
		
	End Function
	
	' --By Dana Seaman
	Private Function Darken(ByVal Color As Byte) As Byte
		
		Darken = (217 * Color) \ 255
		
	End Function
	
	Private Sub DrawGradientEx(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Color1 As Integer, ByVal Color2 As Integer, ByVal GradientDirection As GradientDirectionCts)
		
		'****************************************************************************
		'* Draws very fast Gradient in four direction.                              *
		'* Author: Carles P.V (Gradient Master)                                     *
		'* This routine works as a heart for this control.                          *
		'* Thank you so much Carles.                                                *
		'****************************************************************************
		
		Dim uBIH As BITMAPINFOHEADER
		Dim lBits() As Integer
		Dim lGrad() As Integer
		
		Dim r1 As Integer
		Dim g1 As Integer
		Dim b1 As Integer
		Dim r2 As Integer
		Dim g2 As Integer
		Dim b2 As Integer
		Dim dR As Integer
		Dim dG As Integer
		Dim dB As Integer
		
		Dim Scan As Integer
		Dim i As Integer
		Dim iEnd As Integer
		Dim iOffset As Integer
		Dim j As Integer
		Dim jEnd As Integer
		Dim iGrad As Integer
		
		'-- A minor check
		
		'If (Width < 1 Or Height < 1) Then Exit Sub
		If (Width < 1 Or Height < 1) Then
			Exit Sub
		End If
		
		'-- Decompose colors
		Color1 = Color1 And &HFFFFFF
		r1 = Color1 Mod &H100
		Color1 = Color1 \ &H100
		g1 = Color1 Mod &H100
		Color1 = Color1 \ &H100
		b1 = Color1 Mod &H100
		Color2 = Color2 And &HFFFFFF
		r2 = Color2 Mod &H100
		Color2 = Color2 \ &H100
		g2 = Color2 Mod &H100
		Color2 = Color2 \ &H100
		b2 = Color2 Mod &H100
		
		'-- Get color distances
		dR = r2 - r1
		dG = g2 - g1
		dB = b2 - b1
		
		'-- Size gradient-colors array
		Select Case GradientDirection
			Case GradientDirectionCts.gdHorizontal
				ReDim lGrad(Width - 1)
			Case GradientDirectionCts.gdVertical
				ReDim lGrad(Height - 1)
			Case Else
				ReDim lGrad(Width + Height - 2)
		End Select
		
		'-- Calculate gradient-colors
		iEnd = UBound(lGrad)
		If (iEnd = 0) Then
			'-- Special case (1-pixel wide gradient)
			lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
		Else
			For i = 0 To iEnd
				lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (g1 + (dG * i) \ iEnd) + 65536 * (r1 + (dR * i) \ iEnd)
			Next i
		End If
		
		'-- Size DIB array
		ReDim lBits(Width * Height - 1)
		iEnd = Width - 1
		jEnd = Height - 1
		Scan = Width
		
		'-- Render gradient DIB
		Select Case GradientDirection
			
			Case GradientDirectionCts.gdHorizontal
				
				For j = 0 To jEnd
					For i = iOffset To iEnd + iOffset
						lBits(i) = lGrad(i - iOffset)
					Next i
					iOffset = iOffset + Scan
				Next j
				
			Case GradientDirectionCts.gdVertical
				
				For j = jEnd To 0 Step -1
					For i = iOffset To iEnd + iOffset
						lBits(i) = lGrad(j)
					Next i
					iOffset = iOffset + Scan
				Next j
				
			Case GradientDirectionCts.gdDownwardDiagonal
				
				iOffset = jEnd * Scan
				For j = 1 To jEnd + 1
					For i = iOffset To iEnd + iOffset
						lBits(i) = lGrad(iGrad)
						iGrad = iGrad + 1
					Next i
					iOffset = iOffset - Scan
					iGrad = j
				Next j
				
			Case GradientDirectionCts.gdUpwardDiagonal
				
				iOffset = 0
				For j = 1 To jEnd + 1
					For i = iOffset To iEnd + iOffset
						lBits(i) = lGrad(iGrad)
						iGrad = iGrad + 1
					Next i
					iOffset = iOffset + Scan
					iGrad = j
				Next j
		End Select
		
		'-- Define DIB header
		With uBIH
			.biSize = 40
			.biPlanes = 1
			.biBitCount = 32
			.biWidth = Width
			.biHeight = Height
		End With
		
		'-- Paint it!
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_WARNING: Couldn't resolve default property of object uBIH. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		StretchDIBits(hDC, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
		
	End Sub
	
	Private Function TranslateColor(ByVal clrColor As System.Drawing.Color, Optional ByRef hPalette As Integer = 0) As Integer
		
		'****************************************************************************
		'*  System color code to long rgb                                           *
		'****************************************************************************
		
		If OleTranslateColor(System.Drawing.ColorTranslator.ToOle(clrColor), hPalette, TranslateColor) Then
			TranslateColor = CLR_INVALID
		End If
		
	End Function
	
	Private Sub RedrawButton()
		
		'****************************************************************************
		'*  The main routine of this usercontrol. Everything is drawn here.         *
		'****************************************************************************
		
		'UPGRADE_ISSUE: UserControl method UserControl.Cls was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		MyBase.Cls() 'Clears usercontrol
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		
		SetRect(m_ButtonRect, 0, 0, lW, lH) 'Sets the button rectangle
		
		If (m_ButtonMode <> enumButtonModes.ebmCommandButton) Then 'If Checkboxmode True
			If Not (m_ButtonStyle = enumButtonStlyes.eStandard Or m_ButtonStyle = enumButtonStlyes.eXPToolbar) Then
				If m_bValue Then m_Buttonstate = enumButtonStates.eStateDown
			End If
		End If
		
		Select Case m_ButtonStyle
			
			Case enumButtonStlyes.eStandard
				DrawStandardButton(m_Buttonstate)
			Case enumButtonStlyes.e3DHover
				DrawStandardButton(m_Buttonstate)
			Case enumButtonStlyes.eFlat
				DrawStandardButton(m_Buttonstate)
			Case enumButtonStlyes.eFlatHover
				DrawStandardButton(m_Buttonstate)
			Case enumButtonStlyes.eWindowsXP
				DrawWinXPButton(m_Buttonstate)
			Case enumButtonStlyes.eXPToolbar
				DrawXPToolbar(m_Buttonstate)
			Case enumButtonStlyes.eGelButton
				DrawGelButton(m_Buttonstate)
			Case enumButtonStlyes.eOfficeXP
				DrawOfficeXP(m_Buttonstate)
			Case enumButtonStlyes.eInstallShield
				DrawInstallShieldButton(m_Buttonstate)
			Case enumButtonStlyes.eVistaAero
				DrawVistaButton(m_Buttonstate)
			Case enumButtonStlyes.eVistaToolbar
				DrawVistaToolbarStyle(m_Buttonstate)
			Case enumButtonStlyes.eOutlook2007
				DrawOutlook2007(m_Buttonstate)
			Case enumButtonStlyes.eOffice2003
				DrawOffice2003(m_Buttonstate)
			Case enumButtonStlyes.eWindowsTheme
				If IsThemed Then
					' --Theme can be applied
					WindowsThemeButton(m_Buttonstate)
				Else
					' --Fallback to ownerdraw WinXP Button
					m_ButtonStyle = enumButtonStlyes.eWindowsXP
					m_lXPColor = enumXPThemeColors.ecsBlue
					SetThemeColors()
					DrawWinXPButton(m_Buttonstate)
				End If
		End Select
		
	End Sub
	
	Private Sub CreateRegion()
		
		'***************************************************************************
		'*  Create region everytime you redraw a button.                           *
		'*  Because some settings may have changed the button regions              *
		'***************************************************************************
		
		If m_lButtonRgn Then DeleteObject(m_lButtonRgn)
		Select Case m_ButtonStyle
			Case enumButtonStlyes.eWindowsXP, enumButtonStlyes.eVistaAero, enumButtonStlyes.eVistaToolbar, enumButtonStlyes.eInstallShield
				m_lButtonRgn = CreateRoundRectRgn(0, 0, lW + 1, lH + 1, 3, 3)
			Case enumButtonStlyes.eGelButton, enumButtonStlyes.eXPToolbar
				m_lButtonRgn = CreateRoundRectRgn(0, 0, lW + 1, lH + 1, 4, 4)
			Case Else
				m_lButtonRgn = CreateRectRgn(0, 0, MyBase.ClientRectangle.Width, VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height))
		End Select
		SetWindowRgn(MyBase.Handle.ToInt32, m_lButtonRgn, True) 'Set Button Region
		DeleteObject(m_lButtonRgn) 'Free memory
		
	End Sub
	
	Private Sub DrawSymbol(ByVal eArrow As enumSymbol)
		
		Dim hOldFont As Integer
		Dim hNewFont As Integer
		Dim sSign As String
		Dim BtnSymbol As enumSymbol
		
		hNewFont = BuildSymbolFont(14)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		hOldFont = SelectObject(hDC, hNewFont)
		
		sSign = CStr(eArrow)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		DrawText(hDC, sSign, 1, lpSignRect, DT_WORDBREAK) '!!
		DeleteObject(hNewFont)
		
	End Sub
	
	Private Function BuildSymbolFont(ByRef lFontSize As Integer) As Integer
		
		Const SYMBOL_CHARSET As Short = 2
		Dim lpFont As tLogFont
		
		With lpFont
			.lfFaceName = "Marlett" & vbNullChar 'Standard Marlett Font
			.lfHeight = lFontSize 'I was using Webdings first,
			.lfCharSet = SYMBOL_CHARSET 'but I am not sure whether
		End With 'it is installed in every machine!
		'Still Im not sure about Marlet :)
		BuildSymbolFont = CreateFontIndirect(lpFont) 'I got inspirations from
		'Light Templer's Project
	End Function
	
	Private Sub DrawPicwithCaption()
		
		'****************************************************************************
		' Calculate Caption rects and draw the pictures and caption                 *
		'****************************************************************************
		Dim lpRect As RECT 'RECT to draw caption
		Dim pRect As RECT
		Dim lShadowClr As Integer
		Dim lPixelClr As Integer
		lW = ClientRectangle.Width 'ScaleHeight of Button
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height) 'ScaleWidth of Button
		
		If (m_Buttonstate = enumButtonStates.eStateDown Or (m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue = True)) Then
			'-- Mouse down
			If Not m_PictureDown Is Nothing Then
				tmppic = m_PictureDown
			Else
				If Not m_PictureHot Is Nothing Then
					tmppic = m_PictureHot
				Else
					tmppic = m_Picture
				End If
			End If
		ElseIf (m_Buttonstate = enumButtonStates.eStateOver) Then 
			'-- Mouse in (over)
			If Not m_PictureHot Is Nothing Then
				tmppic = m_PictureHot
			Else
				tmppic = m_Picture
			End If
		Else
			'-- Mouse out (normal)
			tmppic = m_Picture
		End If
		
		' --Adjust Picture Sizes
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbHimetric was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Picture property tmppic.Height was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method jcbutton.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		PicH = ScaleX(tmppic.Height, vbHimetric, vbPixels)
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant vbHimetric was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Picture property tmppic.Width was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: UserControl method jcbutton.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		PicW = ScaleX(tmppic.Width, vbHimetric, vbPixels)
		
		' --Get the drawing area of caption
		If m_DropDownSymbol <> enumSymbol.ebsNone Or m_bDropDownSep Then
			If m_PictureAlign = enumPictureAlign.epRightEdge Or m_PictureAlign = enumPictureAlign.epRightOfCaption Then
				SetRect(m_TextRect, 0, 0, lW - 24, lH)
			Else
				SetRect(m_TextRect, 0, 0, lW - 16, lH)
			End If
		Else
			SetRect(m_TextRect, 0, 0, lW - 8, lH)
		End If
		
		' --Calc rects for multiline
		If m_WindowsNT Then
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawTextW(hDC, StrPtr(m_Caption), -1, m_TextRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTL, DT_RTLREADING, 0))
		Else
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(hDC, m_Caption, -1, m_TextRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTL, DT_RTLREADING, 0))
		End If
		
		' --Copy rect into temp var
		CopyRect(lpRect, m_TextRect)
		
		' --Move the caption area according to Caption alignments
		Select Case m_CaptionAlign
			Case enumCaptionAlign.ecLeftAlign
				OffsetRect(lpRect, 2, (lH - lpRect.Bottom_Renamed) \ 2)
				
			Case enumCaptionAlign.ecCenterAlign
				OffsetRect(lpRect, (lW - lpRect.Right_Renamed + PicW) \ 2, (lH - lpRect.Bottom_Renamed) \ 2)
				If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
					OffsetRect(lpRect, -8, 0)
				End If
				If m_PictureAlign = enumPictureAlign.epBottomEdge Or m_PictureAlign = enumPictureAlign.epBottomOfCaption Or m_PictureAlign = enumPictureAlign.epTopOfCaption Or m_PictureAlign = enumPictureAlign.epTopEdge Then
					OffsetRect(lpRect, -(PicW \ 2), 0)
				End If
				
			Case enumCaptionAlign.ecRightAlign
				OffsetRect(lpRect, lW - lpRect.Right_Renamed - 4, (lH - lpRect.Bottom_Renamed) \ 2)
				
		End Select
		
		With lpRect
			
			If Not m_Picture Is Nothing Then
				Select Case m_PictureAlign
					Case enumPictureAlign.epLeftEdge, enumPictureAlign.epLeftOfCaption
						If m_CaptionAlign <> enumCaptionAlign.ecCenterAlign Then
							If .Left_Renamed < PicW + 4 Then
								.Left_Renamed = PicW + 4 : .Right_Renamed = .Right_Renamed + PicW + 4
							End If
						End If
						
					Case enumPictureAlign.epRightEdge, enumPictureAlign.epRightOfCaption
						If .Right_Renamed > lW - PicW - 4 Then
							.Right_Renamed = lW - PicW - 4 : .Left_Renamed = .Left_Renamed - PicW - 4
						End If
						If m_CaptionAlign = enumCaptionAlign.ecCenterAlign Then
							OffsetRect(lpRect, -12, 0)
						End If
						
					Case enumPictureAlign.epTopOfCaption, enumPictureAlign.epTopEdge
						OffsetRect(lpRect, 0, PicH \ 2)
						
					Case enumPictureAlign.epBottomOfCaption, enumPictureAlign.epBottomEdge
						OffsetRect(lpRect, 0, -PicH \ 2)
						
					Case enumPictureAlign.epBackGround
						If m_CaptionAlign = enumCaptionAlign.ecCenterAlign Then
							OffsetRect(lpRect, -16, 0)
						End If
				End Select
			End If
			
			If m_CaptionAlign = enumCaptionAlign.ecRightAlign Then
				If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
					OffsetRect(lpRect, -16, 0)
				End If
			End If
			
			' --For themed style, we are not able to draw borders
			' --after drawing the caption. i mean the whole button is painted at once.
			If m_ButtonStyle = enumButtonStlyes.eWindowsTheme Then
				If .Left_Renamed < 4 Then .Left_Renamed = 4
				If .Right_Renamed > ClientRectangle.Width - 4 Then .Right_Renamed = ClientRectangle.Width - 4
				If .Top_Renamed < 4 Then .Top_Renamed = 4
				If .Bottom_Renamed > VB6.PixelsToTwipsY(ClientRectangle.Height) - 4 Then .Bottom_Renamed = VB6.PixelsToTwipsY(ClientRectangle.Height) - 4
			End If
		End With
		
		' --Save the caption rect
		CopyRect(m_TextRect, lpRect)
		
		' --Calculate Pictures positions once we have caption rects
		CalcPicRects()
		
		' --Calculate rects with the dropdown symbol
		If m_DropDownSymbol <> enumSymbol.ebsNone Then
			' --Drawing area for dropdown symbol  (the symbol is optional;)
			SetRect(lpSignRect, lW - 15, lH / 2 - 7, lW, lH / 2 + 8)
		End If
		
		If m_bDropDownSep Then
			If m_PictureAlign <> enumPictureAlign.epRightEdge Or m_PictureAlign <> enumPictureAlign.epRightOfCaption Then
				If m_TextRect.Right_Renamed < ClientRectangle.Width - 8 Then
					'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawLineApi(lW - 16, 3, lW - 16, lH - 3, ShiftColor(GetPixel(hDC, 7, 7), -0.1))
					'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawLineApi(lW - 15, 3, lW - 15, lH - 3, ShiftColor(GetPixel(hDC, 7, 7), 0.1))
				End If
			ElseIf m_PictureAlign = enumPictureAlign.epRightEdge Or m_PictureAlign = enumPictureAlign.epRightOfCaption Then 
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawLineApi(lW - 16, 3, lW - 16, lH - 3, ShiftColor(GetPixel(hDC, 7, 7), -0.1))
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawLineApi(lW - 15, 3, lW - 15, lH - 3, ShiftColor(GetPixel(hDC, 7, 7), 0.1))
			End If
		End If
		
		' --Some styles on down state donot change their text positions
		' --See your XP and Vista buttons ;)
		If m_Buttonstate = enumButtonStates.eStateDown Then
			If m_ButtonStyle = enumButtonStlyes.e3DHover Or m_ButtonStyle = enumButtonStlyes.eFlat Or m_ButtonStyle = enumButtonStlyes.eFlatHover Or m_ButtonStyle = enumButtonStlyes.eGelButton Or m_ButtonStyle = enumButtonStlyes.eOffice2003 Or m_ButtonStyle = enumButtonStlyes.eXPToolbar Or m_ButtonStyle = enumButtonStlyes.eVistaToolbar Or m_ButtonStyle = enumButtonStlyes.eStandard Then
				OffsetRect(m_TextRect, 1, 1)
				OffsetRect(m_PicRect, 1, 1)
				OffsetRect(lpSignRect, 1, 1)
			End If
		End If
		
		' --Draw Pictures
		If m_bPicPushOnHover And m_Buttonstate = enumButtonStates.eStateOver Then
			lShadowClr = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC0C0C0))
			DrawPicture(m_PicRect, lShadowClr)
			CopyRect(pRect, m_PicRect)
			OffsetRect(pRect, -2, -2)
			DrawPicture(pRect)
		Else
			DrawPicture(m_PicRect)
		End If
		
		If m_PictureShadow Then
			If Not (m_bPicPushOnHover And m_Buttonstate = enumButtonStates.eStateOver) Then
				DrawPicShadow()
			End If
		End If
		
		' --Text Effects
		If m_CaptionEffects <> enumCaptionEffects.eseNone Then
			DrawCaptionEffect()
		End If
		
		' --At Last, draw the Captions
		If m_bEnabled Then
			If m_Buttonstate = enumButtonStates.eStateOver Then
				DrawCaptionEx(m_TextRect, TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColorOver)), 0, 0)
			Else
				DrawCaptionEx(m_TextRect, TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColor)), 0, 0)
			End If
		Else
			DrawCaptionEx(m_TextRect, GetSysColor(COLOR_GRAYTEXT), 0, 0)
		End If
		
		If m_DropDownSymbol <> enumSymbol.ebsNone Then
			
			If m_ButtonStyle = enumButtonStlyes.eStandard Or m_ButtonStyle = enumButtonStlyes.e3DHover Or m_ButtonStyle = enumButtonStlyes.eFlat Or m_ButtonStyle = enumButtonStlyes.eFlatHover Or m_ButtonStyle = enumButtonStlyes.eVistaToolbar Or m_ButtonStyle = enumButtonStlyes.eXPToolbar Then
				' --move the symbol downwards for some button style on mouse down
				If m_Buttonstate = enumButtonStates.eStateDown Then
					OffsetRect(lpSignRect, 1, 1)
				End If
			End If
			
			If m_bEnabled Then
				MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColor)))
			Else
				MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_GRAYTEXT))
			End If
			DrawSymbol(m_DropDownSymbol)
		End If
		
	End Sub
	
	Private Sub CalcPicRects()
		
		'****************************************************************************
		' Calculate the rects for positioning pictures                              *
		'****************************************************************************
		
		If m_Picture Is Nothing Then Exit Sub
		
		With m_PicRect
			
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If LenB(m_Caption) > 0 And m_PictureAlign <> enumPictureAlign.epBackGround Then
				
				Select Case m_PictureAlign
					
					Case enumPictureAlign.epLeftEdge
						.Left_Renamed = 3
						.Top_Renamed = (lH - PicH) \ 2
						If m_PicRect.Left_Renamed < 0 Then
							OffsetRect(m_PicRect, PicW, 0)
							OffsetRect(m_TextRect, PicW, 0)
						End If
						
					Case enumPictureAlign.epLeftOfCaption
						.Left_Renamed = m_TextRect.Left_Renamed - PicW - 4
						.Top_Renamed = (lH - PicH) \ 2
						
					Case enumPictureAlign.epRightEdge
						.Left_Renamed = lW - PicW - 3
						.Top_Renamed = (lH - PicH) \ 2
						' --If picture overlaps text
						If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
							OffsetRect(m_PicRect, -16, 0)
						End If
						If .Left_Renamed < m_TextRect.Right_Renamed + 2 Then
							.Left_Renamed = m_TextRect.Right_Renamed + 2
						End If
						
					Case enumPictureAlign.epRightOfCaption
						.Left_Renamed = m_TextRect.Right_Renamed + 4
						.Top_Renamed = (lH - PicH) \ 2
						If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
							OffsetRect(m_PicRect, -16, 0)
						End If
						' --If picture overlaps text
						If .Left_Renamed < m_TextRect.Right_Renamed + 2 Then
							.Left_Renamed = m_TextRect.Right_Renamed + 2
						End If
						
					Case enumPictureAlign.epTopOfCaption
						.Left_Renamed = (lW - PicW) \ 2
						.Top_Renamed = m_TextRect.Top_Renamed - PicH - 2
						If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
							OffsetRect(m_PicRect, -8, 0)
						End If
						
					Case enumPictureAlign.epTopEdge
						.Left_Renamed = (lW - PicW) \ 2
						.Top_Renamed = 4
						If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
							OffsetRect(m_PicRect, -8, 0)
						End If
						
					Case enumPictureAlign.epBottomOfCaption
						.Left_Renamed = (lW - PicW) \ 2
						.Top_Renamed = m_TextRect.Bottom_Renamed + 2
						If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
							OffsetRect(m_PicRect, -8, 0)
						End If
						
					Case enumPictureAlign.epBottomEdge
						.Left_Renamed = (lW - PicW) \ 2
						.Top_Renamed = lH - PicH - 4
						If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
							OffsetRect(m_PicRect, -8, 0)
						End If
						
				End Select
			Else
				.Left_Renamed = (lW - PicW) \ 2
				.Top_Renamed = (lH - PicH) \ 2
				If m_bDropDownSep Or m_DropDownSymbol <> enumSymbol.ebsNone Then
					OffsetRect(m_PicRect, -8, 0)
				End If
			End If
			
			' --Set the height and width
			.Right_Renamed = .Left_Renamed + PicW
			.Bottom_Renamed = .Top_Renamed + PicH
			
		End With
		
	End Sub
	
	Private Sub DrawPicture(ByRef lpRect As RECT, Optional ByRef lBrushColor As Integer = -1)
		
		'****************************************************************************
		' draw the picture by calling the TransBlt routines                         *
		'****************************************************************************
		
		Dim tmpMaskColor As Integer
		
		' --Draw picture
		'UPGRADE_ISSUE: Constant vbPicTypeIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Picture property tmppic.Type was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If tmppic.Type = vbPicTypeIcon Then
			tmpMaskColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC0C0C0))
		Else
			tmpMaskColor = m_lMaskColor
		End If
		
		If Is32BitBMP(tmppic) Then
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			TransBlt32(hDC, lpRect.Left_Renamed, lpRect.Top_Renamed, PicW, PicH, tmppic, lBrushColor)
		Else
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			TransBlt(hDC, lpRect.Left_Renamed, lpRect.Top_Renamed, PicW, PicH, tmppic, tmpMaskColor, lBrushColor)
		End If
		
	End Sub
	
	Private Sub DrawPicShadow()
		
		'  Still not satisfied results for picture shadows
		
		Dim lShadowClr As Integer
		Dim lPixelClr As Integer
		Dim lpRect As RECT
		
		If m_bPicPushOnHover And m_Buttonstate = enumButtonStates.eStateOver Then
			OffsetRect(m_PicRect, -2, -2)
		End If
		
		lShadowClr = BlendColors(TranslateColor(System.Drawing.ColorTranslator.FromOle(&H808080)), TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)))
		CopyRect(lpRect, m_PicRect)
		
		OffsetRect(lpRect, 2, 2)
		DrawPicture(lpRect, ShiftColor(lShadowClr, 0.05))
		OffsetRect(lpRect, -1, -1)
		DrawPicture(lpRect, ShiftColor(lShadowClr, -0.1))
		
		DrawPicture(m_PicRect)
		
	End Sub
	
	Private Sub DrawCaptionEffect()
		
		'****************************************************************************
		'* Draws the caption with/without unicode along with the special effects    *
		'****************************************************************************
		
		Dim bColor As Integer 'BackColor
		
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		
		' --Set new colors according to effects
		Select Case m_CaptionEffects
			Case enumCaptionEffects.eseEmbossed
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, 0.14), -1, -1)
			Case enumCaptionEffects.eseEngraved
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, 0.14), 1, 1)
			Case enumCaptionEffects.eseShadowed
				DrawCaptionEx(m_TextRect, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC0C0C0)), 1, 1)
			Case enumCaptionEffects.eseOutline
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, 0.1), 1, 1)
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, 0.1), 1, -1)
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, 0.1), -1, 1)
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, 0.1), -1, -1)
			Case enumCaptionEffects.eseCover
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, -0.1), 1, 1)
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, -0.1), 1, -1)
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, -0.1), -1, 1)
				DrawCaptionEx(m_TextRect, ShiftColor(bColor, -0.1), -1, -1)
				
		End Select
		
		If m_bEnabled Then
			DrawCaptionEx(m_TextRect, TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColor)), 0, 0)
		Else
			DrawCaptionEx(m_TextRect, GetSysColor(COLOR_GRAYTEXT), 0, 0)
		End If
		
	End Sub
	
	Private Sub DrawCaptionEx(ByRef lpRect As RECT, ByRef lColor As Integer, ByRef OffsetX As Integer, ByRef OffsetY As Integer)
		
		Dim tRect As RECT
		Dim lOldForeColor As Integer
		
		' --Get current forecolor
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		lOldForeColor = GetTextColor(hDC)
		
		CopyRect(tRect, lpRect)
		OffsetRect(tRect, OffsetX, OffsetY)
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetTextColor(hDC, lColor)
		
		If m_WindowsNT Then
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawTextW(hDC, StrPtr(m_Caption), -1, tRect, DT_DRAWFLAG Or IIf(m_bRTL, DT_RTLREADING, 0))
		Else
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawText(hDC, m_Caption, -1, tRect, DT_DRAWFLAG Or IIf(m_bRTL, DT_RTLREADING, 0))
		End If
		
		' --Restore previous forecolor
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetTextColor(hDC, lOldForeColor)
		
	End Sub
	
	Private Sub UncheckAllValues()
		
		' --Many Thanks to Morgan Haueisen
		
		Dim objButton As Object
		' Check all controls in parent
		'UPGRADE_WARNING: Control property .Parent was upgraded to .FindForm which has a new behavior. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
		For	Each objButton In FindForm.Controls
			' Is it a jcbutton?
			If TypeOf objButton Is jcbutton Then
				' Is the button in the same container?
				'UPGRADE_ISSUE: UserControl property UserControl.ContainerHwnd was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object objButton.Container. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If objButton.Container.hwnd = MyBase.ContainerHwnd Then
					' is the button type Option?
					'UPGRADE_WARNING: Couldn't resolve default property of object objButton.Mode. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If objButton.Mode = enumButtonModes.ebmOptionButton Then
						' is it not this button
						'UPGRADE_WARNING: Couldn't resolve default property of object objButton.hwnd. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If Not objButton.hwnd = MyBase.Handle.ToInt32 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object objButton.value. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							objButton.value = False
						End If
					End If
				End If
			End If
		Next objButton
		
	End Sub
	
	Private Sub SetAccessKey()
		
		Dim i As Integer
		
		'UPGRADE_ISSUE: UserControl property UserControl.AccessKeys was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		MyBase.AccessKeys = vbNullString
		If Len(m_Caption) > 1 Then
			i = InStr(1, m_Caption, "&", CompareMethod.Text)
			If (i < Len(m_Caption)) And (i > 0) Then
				If Mid(m_Caption, i + 1, 1) <> "&" Then
					'UPGRADE_ISSUE: UserControl property UserControl.AccessKeys was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					AccessKeys = LCase(Mid(m_Caption, i + 1, 1))
				Else
					i = InStr(i + 2, m_Caption, "&", CompareMethod.Text)
					If Mid(m_Caption, i + 1, 1) <> "&" Then
						'UPGRADE_ISSUE: UserControl property UserControl.AccessKeys was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						AccessKeys = LCase(Mid(m_Caption, i + 1, 1))
					End If
				End If
			End If
		End If
		
	End Sub
	
	Private Sub DrawCorners(ByRef Color As Integer)
		
		'****************************************************************************
		'* Draws four Corners of the button specified by Color                      *
		'****************************************************************************
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixel(hDC, 1, 1, Color)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixel(hDC, 1, lH - 2, Color)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixel(hDC, lW - 2, 1, Color)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SetPixel(hDC, lW - 2, lH - 2, Color)
		
	End Sub
	
	Private Sub DrawStandardButton(ByVal vState As enumButtonStates)
		
		'****************************************************************************
		' Draws  four different styles in one procedure                             *
		' Makes reading the code difficult, but saves much space!! ;)               *
		'****************************************************************************
		
		Dim FocusRect As RECT
		Dim tmpRect As RECT
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		SetRect(m_ButtonRect, 0, 0, lW, lH)
		
		If Not m_bEnabled Then
			' --Draws raised edge border
			If m_ButtonStyle = enumButtonStlyes.eStandard Then
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawEdge(hDC, m_ButtonRect, BDR_RAISED95, BF_RECT)
			ElseIf m_ButtonStyle = enumButtonStlyes.eFlat Then 
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawEdge(hDC, m_ButtonRect, BDR_RAISEDINNER, BF_RECT)
			End If
			DrawPicwithCaption()
			Exit Sub
		End If
		
		If m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue Then
			PaintRect(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), 0.02), m_ButtonRect)
			DrawPicwithCaption()
			If m_ButtonStyle <> enumButtonStlyes.eFlatHover Then
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawEdge(hDC, m_ButtonRect, BDR_SUNKEN95, BF_RECT)
				If m_bShowFocus And m_bHasFocus And m_ButtonStyle = enumButtonStlyes.eStandard Then
					DrawRectangle(4, 4, lW - 7, lH - 7, TranslateColor(System.Drawing.SystemColors.AppWorkspace))
				End If
			End If
			Exit Sub
		End If
		
		Select Case vState
			Case enumButtonStates.eStateNormal
				CreateRegion()
				PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), m_ButtonRect)
				DrawPicwithCaption()
				Select Case m_ButtonStyle
					Case enumButtonStlyes.eStandard
						'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawEdge(hDC, m_ButtonRect, BDR_RAISED95, BF_RECT)
					Case enumButtonStlyes.eFlat
						'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawEdge(hDC, m_ButtonRect, BDR_RAISEDINNER, BF_RECT)
				End Select
			Case enumButtonStates.eStateOver
				PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), m_ButtonRect)
				DrawPicwithCaption()
				Select Case m_ButtonStyle
					Case enumButtonStlyes.eFlatHover, enumButtonStlyes.eFlat
						' --Draws flat raised edge border
						'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawEdge(hDC, m_ButtonRect, BDR_RAISEDINNER, BF_RECT)
					Case Else
						' --Draws 3d raised edge border
						'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawEdge(hDC, m_ButtonRect, BDR_RAISED95, BF_RECT)
				End Select
				
			Case enumButtonStates.eStateDown
				PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), m_ButtonRect)
				DrawPicwithCaption()
				Select Case m_ButtonStyle
					Case enumButtonStlyes.eStandard
						DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H99A8AC)))
						DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.Color.Black))
					Case enumButtonStlyes.e3DHover
						'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						DrawEdge(hDC, m_ButtonRect, BDR_SUNKEN95, BF_RECT)
					Case enumButtonStlyes.eFlatHover, enumButtonStlyes.eFlat
						' --Draws flat pressed edge
						DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.Color.White))
						DrawRectangle(0, 0, lW + 1, lH + 1, TranslateColor(System.Drawing.SystemColors.GrayText))
				End Select
		End Select
		
		' --Button has focus but not downstate Or button is Default
		
		If m_bHasFocus Or m_bDefault Then
			On Error Resume Next
			If m_bShowFocus And Not DesignMode Then
				If m_ButtonStyle = enumButtonStlyes.e3DHover Or m_ButtonStyle = enumButtonStlyes.eStandard Then
					SetRect(FocusRect, 4, 4, lW - 4, lH - 4)
				Else
					SetRect(FocusRect, 3, 3, lW - 3, lH - 3)
				End If
				If m_bParentActive Then
					'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawFocusRect(hDC, FocusRect)
				End If
			End If
			If vState <> enumButtonStates.eStateDown And m_ButtonStyle = enumButtonStlyes.eStandard Then
				SetRect(tmpRect, 0, 0, lW - 1, lH - 1)
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawEdge(hDC, tmpRect, BDR_RAISED95, BF_RECT)
				DrawRectangle(0, 0, lW - 1, lH - 1, TranslateColor(System.Drawing.SystemColors.AppWorkspace))
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.Color.Black))
			End If
		End If
		
	End Sub
	
	Private Sub DrawXPToolbar(ByVal vState As enumButtonStates)
		
		Dim lpRect As RECT
		Dim bColor As Integer
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		
		If vState = enumButtonStates.eStateDown Then
			m_bColors.tForeColor = TranslateColor(System.Drawing.Color.White)
		Else
			m_bColors.tForeColor = TranslateColor(System.Drawing.SystemColors.ControlText)
		End If
		
		If m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue Then
			If m_bIsDown Then vState = enumButtonStates.eStateDown
		End If
		
		If m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue And vState <> enumButtonStates.eStateDown Then
			SetRect(lpRect, 0, 0, lW, lH)
			PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFEFEFE)), lpRect)
			m_bColors.tForeColor = TranslateColor(System.Drawing.SystemColors.ControlText)
			DrawPicwithCaption()
			DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HAF987A)))
			DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC1B3A0)), -0.2))
			If vState = enumButtonStates.eStateOver Then
				DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEDF0F2))) 'Right Line
				DrawLineApi(2, lH - 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HD8DEE4))) 'Bottom
				DrawLineApi(1, lH - 3, lW - 1, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE8ECEF))) 'Bottom
				DrawLineApi(1, lH - 4, lW - 1, lH - 4, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF8F9FA))) 'Bottom
			End If
			' --Necessary to redraw text & pictures 'coz we are painting usercontrol agaon
			Exit Sub
		End If
		
		Select Case vState
			Case enumButtonStates.eStateNormal
				CreateRegion()
				PaintRect(bColor, m_ButtonRect)
				DrawPicwithCaption()
			Case enumButtonStates.eStateOver
				DrawGradientEx(0, 0, lW, lH / 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFDFEFE)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEEF4F4)), GradientDirectionCts.gdVertical)
				DrawGradientEx(0, lH / 2, lW, lH / 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEEF4F4)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEAF1F1)), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE0E7EA))) 'right line
				DrawLineApi(lW - 3, 2, lW - 3, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEAF0F0)))
				DrawLineApi(0, lH - 4, lW, lH - 4, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE5EDEE))) 'Bottom
				DrawLineApi(0, lH - 3, lW, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HD6E1E4))) 'Bottom
				DrawLineApi(0, lH - 2, lW, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC6D2D7))) 'Bottom
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC3CECE)))
				DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC9D4D4)), -0.05))
			Case enumButtonStates.eStateDown
				PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDDE4E5)), m_ButtonRect) 'Paint with Darker color
				DrawPicwithCaption()
				DrawLineApi(1, 1, lW - 2, 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HD1DADC)), -0.02)) 'Topmost Line
				DrawLineApi(1, 2, lW - 2, 2, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDAE1E3)), -0.02)) 'A lighter top line
				DrawLineApi(1, lH - 3, lW - 2, lH - 3, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDEE5E6)), 0.02)) 'Bottom Line
				DrawLineApi(1, lH - 2, lW - 2, lH - 2, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE5EAEB)), 0.02))
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H929D9D)))
				DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HABB4B5)), -0.2))
		End Select
		
	End Sub
	
	Private Sub DrawWinXPButton(ByVal vState As enumButtonStates)
		
		'****************************************************************************
		'* Windows XP Button                                                        *
		'* Totally written from Scratch and coded by Me!!  hehe                     *
		'****************************************************************************
		
		Dim lpRect As RECT
		Dim bColor As Integer
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		SetRect(m_ButtonRect, 0, 0, lW, lH)
		
		If Not m_bEnabled Then
			CreateRegion()
			PaintRect(BlendColors(GetSysColor(COLOR_BTNFACE), ShiftColor(bColor, 0.1)), m_ButtonRect)
			DrawPicwithCaption()
			DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.1))
			DrawCorners(ShiftColor(bColor, -0.1))
			Exit Sub
		End If
		
		Select Case vState
			
			Case enumButtonStates.eStateNormal
				CreateRegion()
				Select Case m_lXPColor
					Case enumXPThemeColors.ecsBlue, enumXPThemeColors.ecsOliveGreen, enumXPThemeColors.ecsCustom
						' --mimic the XP styles
						DrawGradientEx(0, 0, lW, lH, ShiftColor(bColor, 0.07), bColor, GradientDirectionCts.gdVertical)
						DrawGradientEx(0, 0, lW, 4, ShiftColor(bColor, 0.1), ShiftColor(bColor, 0.08), GradientDirectionCts.gdVertical)
						DrawPicwithCaption()
						DrawLineApi(1, lH - 2, lW - 2, lH - 2, ShiftColor(bColor, -0.09)) 'BottomMost line
						DrawLineApi(1, lH - 3, lW - 2, lH - 3, ShiftColor(bColor, -0.05)) 'Bottom Line
						DrawLineApi(1, lH - 4, lW - 2, lH - 4, ShiftColor(bColor, -0.01)) 'Bottom Line
						DrawLineApi(lW - 2, 2, lW - 2, lH - 2, ShiftColor(bColor, -0.08)) 'Right Line
						DrawLineApi(1, 1, 1, lH - 2, BlendColors(TranslateColor(System.Drawing.Color.White), bColor)) 'Left Line
					Case enumXPThemeColors.ecsSilver
						' --mimic the Silver XP style
						DrawGradientEx(0, 0, lW, lH / 2, ShiftColor(bColor, 0.22), bColor, GradientDirectionCts.gdVertical)
						DrawGradientEx(0, lH / 2, lW, lH / 2, ShiftColor(bColor, -0.01), ShiftColor(bColor, -0.15), GradientDirectionCts.gdVertical)
						DrawPicwithCaption()
						DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.Color.White)) 'Right Line
						DrawLineApi(1, 1, 1, lH - 2, TranslateColor(System.Drawing.Color.White)) 'Left Line
				End Select
				
			Case enumButtonStates.eStateOver
				Select Case m_lXPColor
					Case enumXPThemeColors.ecsBlue, enumXPThemeColors.ecsOliveGreen, enumXPThemeColors.ecsCustom
						DrawGradientEx(0, 0, lW, lH, ShiftColor(bColor, 0.07), bColor, GradientDirectionCts.gdVertical)
						DrawGradientEx(0, 0, lW, 4, ShiftColor(bColor, 0.1), ShiftColor(bColor, 0.08), GradientDirectionCts.gdVertical)
						DrawPicwithCaption()
					Case enumXPThemeColors.ecsSilver
						DrawGradientEx(0, 0, lW, lH / 2, ShiftColor(bColor, 0.22), bColor, GradientDirectionCts.gdVertical)
						DrawGradientEx(0, lH / 2, lW, lH / 2, ShiftColor(bColor, -0.01), ShiftColor(bColor, -0.15), GradientDirectionCts.gdVertical)
						DrawPicwithCaption()
				End Select
				' --Draw the ORANGE border lines....
				DrawLineApi(1, 2, lW - 2, 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H89D8FD))) 'uppermost inner hover
				DrawLineApi(1, 1, lW - 2, 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HCFF0FF))) 'uppermost outer hover
				DrawLineApi(1, 1, 1, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H49BDF9))) 'Leftmost Line
				DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H49BDF9))) 'Rightmost Line
				DrawLineApi(2, 2, 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H7AD2FC))) 'Left Line
				DrawLineApi(lW - 3, 3, lW - 3, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H7AD2FC))) 'Right Line
				DrawLineApi(2, lH - 3, lW - 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H30B3F8))) 'BottomMost Line
				DrawLineApi(2, lH - 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H97E5))) 'Bottom Line
				
			Case enumButtonStates.eStateDown
				Select Case m_lXPColor
					Case enumXPThemeColors.ecsBlue, enumXPThemeColors.ecsOliveGreen, enumXPThemeColors.ecsCustom
						PaintRect(ShiftColor(bColor, -0.05), m_ButtonRect) 'Paint with Darker color
						DrawPicwithCaption()
						DrawLineApi(1, 1, lW - 2, 1, ShiftColor(bColor, -0.16)) 'Topmost Line
						DrawLineApi(1, 2, lW - 2, 2, ShiftColor(bColor, -0.1)) 'A lighter top line
						DrawLineApi(1, lH - 2, lW - 2, lH - 2, ShiftColor(bColor, 0.01)) 'Bottom Line
						DrawLineApi(1, 1, 1, lH - 2, ShiftColor(bColor, -0.16)) 'Leftmost Line
						DrawLineApi(2, 2, 2, lH - 2, ShiftColor(bColor, -0.1)) 'Left1 Line
						DrawLineApi(lW - 2, 2, lW - 2, lH - 2, ShiftColor(bColor, 0.04)) 'Right Line
					Case enumXPThemeColors.ecsSilver
						DrawGradientEx(0, 0, lW, lH - 6, ShiftColor(bColor, -0.2), ShiftColor(bColor, 0.05), GradientDirectionCts.gdVertical)
						DrawGradientEx(0, lH - 6, lW, lH - 1, ShiftColor(bColor, 0.08), TranslateColor(System.Drawing.Color.White), GradientDirectionCts.gdVertical)
						DrawPicwithCaption()
						DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.Color.White))
				End Select
		End Select
		
		If m_bParentActive Then
			If (m_bHasFocus Or m_bDefault) And (vState <> enumButtonStates.eStateDown And vState <> enumButtonStates.eStateOver) Then
				Select Case m_lXPColor
					Case enumXPThemeColors.ecsBlue, enumXPThemeColors.ecsCustom
						DrawLineApi(1, 2, lW - 2, 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF6D4BC))) 'uppermost inner hover
						DrawLineApi(1, 1, lW - 2, 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFE7CE))) 'uppermost outer hover
						DrawLineApi(1, 1, 1, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE6AF8E))) 'Leftmost Line
						DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE6AF8E))) 'Rightmost Line
						DrawLineApi(2, 2, 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF4D1B8))) 'Left Line
						DrawLineApi(lW - 3, 3, lW - 3, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF4D1B8))) 'Right Line
						DrawLineApi(2, lH - 3, lW - 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE4AD89))) 'BottomMost Line
						DrawLineApi(2, lH - 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEE8269))) 'Bottom Line
					Case enumXPThemeColors.ecsOliveGreen
						DrawLineApi(1, 2, lW - 2, 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H8FD1C2))) 'uppermost inner hover
						DrawLineApi(1, 1, lW - 2, 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H80CBB1))) 'uppermost outer hover
						DrawLineApi(1, 1, 1, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H68C8A0))) 'Leftmost Line
						DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H68C8A0))) 'Rightmost Line
						DrawLineApi(2, 2, 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H68C8A0))) 'Left Line
						DrawLineApi(lW - 3, 3, lW - 3, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H68C8A0))) 'Right Line
						DrawLineApi(2, lH - 3, lW - 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H68C8A0))) 'Bottom Line
						DrawLineApi(2, lH - 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H66A7A8))) 'BottomMost Line
					Case enumXPThemeColors.ecsSilver
						DrawLineApi(1, 2, lW - 2, 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF6D4BC))) 'uppermost inner hover
						DrawLineApi(1, 1, lW - 2, 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFE7CE))) 'uppermost outer hover
						DrawLineApi(1, 1, 1, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE6AF8E))) 'Leftmost Line
						DrawLineApi(lW - 2, 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE6AF8E))) 'Rightmost Line
						DrawLineApi(2, 2, 2, lH - 3, TranslateColor(System.Drawing.Color.White)) 'Left Line
						DrawLineApi(lW - 3, 3, lW - 3, lH - 3, TranslateColor(System.Drawing.Color.White)) 'Right Line
						DrawLineApi(2, lH - 3, lW - 2, lH - 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE4AD89))) 'BottomMost Line
						DrawLineApi(2, lH - 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEE8269))) 'Bottom Line
				End Select
			End If
		End If
		
		On Error Resume Next 'Some times error occurs that Client site not available
		If m_bParentActive Then 'I mean some times ;)
			If m_bShowFocus And m_bParentActive And (m_bHasFocus Or m_bDefault) Then 'show focusrect at runtime only
				SetRect(lpRect, 2, 2, lW - 2, lH - 2) 'I don't like this ugly focusrect!!
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawFocusRect(hDC, lpRect)
			End If
		End If
		
		Select Case m_lXPColor
			Case enumXPThemeColors.ecsBlue, enumXPThemeColors.ecsSilver, enumXPThemeColors.ecsCustom
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H743C00)))
				DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&H743C00)), 0.3))
			Case enumXPThemeColors.ecsOliveGreen
				DrawRectangle(0, 0, lW, lH, RGB(55, 98, 6))
				DrawCorners(ShiftColor(RGB(55, 98, 6), 0.3))
		End Select
		
	End Sub
	
	Private Sub DrawOfficeXP(ByVal vState As enumButtonStates)
		
		Dim lpRect As RECT
		Dim pRect As RECT
		Dim bColor As Integer
		Dim oColor As Integer
		Dim BorderColor As Integer
		
		lH = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		lW = MyBase.ClientRectangle.Width
		
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		SetRect(lpRect, 0, 0, lW, lH)
		
		Select Case m_lXPColor
			Case enumXPThemeColors.ecsBlue
				oColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HEED2C1))
				BorderColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC56A31))
			Case enumXPThemeColors.ecsSilver
				oColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE3DFE0))
				BorderColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBFB4B2))
			Case enumXPThemeColors.ecsOliveGreen
				oColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBAD6D4))
				BorderColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&H70A093))
			Case enumXPThemeColors.ecsCustom
				oColor = bColor
				BorderColor = ShiftColor(bColor, -0.12)
		End Select
		
		If m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue Then
			PaintRect(ShiftColor(oColor, -0.05), m_ButtonRect)
			DrawRectangle(0, 0, lW, lH, BorderColor)
			If m_bMouseInCtl Then
				PaintRect(ShiftColor(oColor, -0.01), m_ButtonRect)
				DrawRectangle(0, 0, lW, lH, BorderColor)
			End If
			DrawPicwithCaption()
			Exit Sub
		End If
		
		Select Case vState
			Case enumButtonStates.eStateNormal
				PaintRect(bColor, lpRect)
			Case enumButtonStates.eStateOver
				PaintRect(ShiftColor(oColor, 0.03), lpRect)
			Case enumButtonStates.eStateDown
				PaintRect(ShiftColor(oColor, -0.08), lpRect)
		End Select
		
		DrawPicwithCaption()
		
		If m_Buttonstate <> enumButtonStates.eStateNormal Then
			DrawRectangle(0, 0, lW, lH, BorderColor)
		End If
		
	End Sub
	
	Private Sub DrawInstallShieldButton(ByVal vState As enumButtonStates)
		
		'****************************************************************************
		'* I saw this style while installing JetAudio in my PC.                     *
		'* I liked it, so I implemented and gave it a name 'InstallShield'          *
		'* hehe .....
		'****************************************************************************
		
		Dim FocusRect As RECT
		Dim lpRect As RECT
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		
		If Not m_bEnabled Then
			vState = enumButtonStates.eStateNormal 'Simple draw normal state for Disabled
		End If
		
		Select Case vState
			Case enumButtonStates.eStateNormal
				CreateRegion()
				SetRect(m_ButtonRect, 0, 0, lW, lH) 'Maybe have changed before!
				
				' --Draw upper gradient
				DrawGradientEx(0, 0, lW, lH / 2, TranslateColor(System.Drawing.Color.White), TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), GradientDirectionCts.gdVertical)
				' --Draw Bottom Gradient
				DrawGradientEx(0, lH / 2, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				' --Draw Inner White Border
				DrawRectangle(1, 1, lW - 2, lH, TranslateColor(System.Drawing.Color.White))
				' --Draw Outer Rectangle
				DrawRectangle(0, 0, lW, lH, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.2))
				DrawLineApi(2, lH - 1, lW - 2, lH - 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.25))
			Case enumButtonStates.eStateOver
				
				' --Draw upper gradient
				DrawGradientEx(0, 0, lW, lH / 2, TranslateColor(System.Drawing.Color.White), TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), GradientDirectionCts.gdVertical)
				' --Draw Bottom Gradient
				DrawGradientEx(0, lH / 2, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				' --Draw Inner White Border
				DrawRectangle(1, 1, lW - 2, lH, TranslateColor(System.Drawing.Color.White))
				' --Draw Outer Rectangle
				DrawRectangle(0, 0, lW, lH, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.2))
				DrawLineApi(2, lH - 1, lW - 2, lH - 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.25))
			Case enumButtonStates.eStateDown
				
				' --draw upper gradient
				DrawGradientEx(0, 0, lW, lH / 2, TranslateColor(System.Drawing.Color.White), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.1), GradientDirectionCts.gdVertical)
				' --Draw Bottom Gradient
				DrawGradientEx(0, lH / 2, lW, lH, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.1), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.05), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				' --Draw Inner White Border
				DrawRectangle(1, 1, lW - 2, lH, TranslateColor(System.Drawing.Color.White))
				' --Draw Outer Rectangle
				DrawRectangle(0, 0, lW, lH, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.23))
				DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.1))
				DrawLineApi(2, lH - 1, lW - 2, lH - 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), -0.4))
				
		End Select
		
		DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), 0.05))
		
		If m_bParentActive And m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
			SetRect(FocusRect, 3, 3, lW - 3, lH - 3)
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			DrawFocusRect(hDC, FocusRect)
		End If
		
	End Sub
	
	Private Sub DrawGelButton(ByVal vState As enumButtonStates)
		
		'****************************************************************************
		' Draws a Gelbutton                                                         *
		'****************************************************************************
		
		Dim lpRect As RECT 'RECT to fill regions
		Dim bColor As Integer 'Original backcolor
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		
		If Not m_bEnabled Then
			
			' --Fill the button region with background color
			SetRect(lpRect, 0, 0, lW, lH)
			PaintRect(bColor, lpRect)
			
			' --Make a shining Upper Light
			DrawGradientEx(0, 0, lW, 5, ShiftColor(BlendColors(bColor, TranslateColor(System.Drawing.Color.White)), 0.05), bColor, GradientDirectionCts.gdVertical)
			DrawGradientEx(0, 6, lW, lH - 1, ShiftColor(bColor, -0.02), BlendColors(TranslateColor(System.Drawing.Color.White), ShiftColor(bColor, 0.08)), GradientDirectionCts.gdVertical)
			
			DrawPicwithCaption()
			DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.2))
			DrawCorners(ShiftColor(bColor, -0.23))
			
			Exit Sub
		End If
		
		Select Case vState
			
			Case enumButtonStates.eStateNormal 'Normal State
				
				CreateRegion()
				
				' --Fill the button region with background color
				SetRect(lpRect, 0, 0, lW, lH)
				PaintRect(ShiftColor(bColor, -0.03), lpRect)
				
				' --Make a shining Upper Light
				DrawGradientEx(0, 0, lW, 5, ShiftColor(BlendColors(bColor, TranslateColor(System.Drawing.Color.White)), 0.1), bColor, GradientDirectionCts.gdVertical)
				DrawGradientEx(0, 6, lW, lH - 1, ShiftColor(bColor, -0.05), BlendColors(TranslateColor(System.Drawing.Color.White), ShiftColor(bColor, 0.1)), GradientDirectionCts.gdVertical)
				
				DrawPicwithCaption()
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.33))
				
			Case enumButtonStates.eStateOver
				' --Fill the button region with background color
				SetRect(lpRect, 0, 0, lW, lH)
				PaintRect(ShiftColor(bColor, -0.03), lpRect)
				
				' --Make a shining Upper Light
				DrawGradientEx(0, 0, lW, 5, ShiftColor(BlendColors(bColor, TranslateColor(System.Drawing.Color.White)), 0.15), bColor, GradientDirectionCts.gdVertical)
				DrawGradientEx(0, 6, lW, lH - 1, ShiftColor(bColor, -0.05), BlendColors(TranslateColor(System.Drawing.Color.White), ShiftColor(bColor, 0.2)), GradientDirectionCts.gdVertical)
				
				DrawPicwithCaption()
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.28))
				
			Case enumButtonStates.eStateDown
				
				' --fill the button region with background color
				SetRect(lpRect, 0, 0, lW, lH)
				PaintRect(ShiftColor(bColor, -0.03), lpRect)
				
				' --Make a shining Upper Light
				DrawGradientEx(0, 0, lW, 5, ShiftColor(BlendColors(bColor, TranslateColor(System.Drawing.Color.White)), 0.1), bColor, GradientDirectionCts.gdVertical)
				DrawGradientEx(0, 6, lW, lH - 1, ShiftColor(bColor, -0.08), BlendColors(TranslateColor(System.Drawing.Color.White), ShiftColor(bColor, 0.05)), GradientDirectionCts.gdVertical)
				
				DrawPicwithCaption()
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.36))
				
		End Select
		
		DrawCorners(ShiftColor(bColor, -0.36))
		
	End Sub
	
	Private Sub DrawVistaToolbarStyle(ByVal vState As enumButtonStates)
		
		Dim lpRect As RECT
		Dim FocusRect As RECT
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		
		If Not m_bEnabled Then
			' --Draw Disabled button
			PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), m_ButtonRect)
			DrawPicwithCaption()
			DrawCorners(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)))
			Exit Sub
		End If
		
		If vState = enumButtonStates.eStateNormal Then
			CreateRegion()
			' --Set the rect to fill back color
			SetRect(lpRect, 0, 0, lW, lH)
			' --Simply fill the button with one color (No gradient effect here!!)
			PaintRect(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), lpRect)
			DrawPicwithCaption()
		ElseIf vState = enumButtonStates.eStateOver Then 
			
			' --Draws a gradient effect with the folowing colors
			DrawGradientEx(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFDF9F1)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF8ECD0)), GradientDirectionCts.gdVertical)
			
			' --Draws a gradient in half region to give a Light Effect
			DrawGradientEx(1, lH / 1.7, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF8ECD0)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF8ECD0)), GradientDirectionCts.gdVertical)
			
			DrawPicwithCaption()
			
			' --Draw outside borders
			DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HCA9E61)))
			DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.Color.White))
			
		ElseIf vState = enumButtonStates.eStateDown Then 
			
			DrawGradientEx(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF1DEB0)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF9F1DB)), GradientDirectionCts.gdVertical)
			
			DrawPicwithCaption()
			' --Draws outside borders
			DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HCA9E61)))
			DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.Color.White))
			
		End If
		
		If vState = enumButtonStates.eStateDown Or vState = enumButtonStates.eStateOver Then
			DrawCorners(ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HCA9E61)), 0.3))
		End If
		
	End Sub
	
	Private Sub DrawVistaButton(ByVal vState As enumButtonStates)
		
		'*************************************************************************
		'* Draws a cool Vista Aero Style Button                                  *
		'* Use a light background color for best result                          *
		'*************************************************************************
		
		Dim lpRect As RECT 'Used to set rect for drawing rectangles
		Dim Color1 As Integer 'Shifted / Blended color
		Dim bColor As Integer 'Original back Color
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		Color1 = ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)), 0.05)
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		
		If Not m_bEnabled Then
			' --Draw the Disabled Button
			CreateRegion()
			' --Fill the button with disabled color
			SetRect(lpRect, 0, 0, lW, lH)
			PaintRect(ShiftColor(bColor, 0.03), lpRect)
			
			DrawPicwithCaption()
			
			' --Draws outside disabled color rectangle
			DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.25))
			DrawRectangle(1, 1, lW - 2, lH - 2, ShiftColor(bColor, 0.25))
			DrawCorners(ShiftColor(bColor, -0.03))
			Exit Sub
		End If
		
		Select Case vState
			
			Case enumButtonStates.eStateNormal
				
				CreateRegion()
				
				' --Draws a gradient in the full region
				DrawGradientEx(1, 1, lW - 1, lH, Color1, bColor, GradientDirectionCts.gdVertical)
				
				' --Draws a gradient in half region to give a glassy look
				DrawGradientEx(1, lH / 2, lW - 2, lH - 2, ShiftColor(bColor, -0.02), ShiftColor(bColor, -0.15), GradientDirectionCts.gdVertical)
				
				DrawPicwithCaption()
				
				' --Draws border rectangle
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H707070))) 'outer
				DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.Color.White)) 'inner
				
			Case enumButtonStates.eStateOver
				
				' --Make gradient in the upper half region
				DrawGradientEx(1, 1, lW - 2, lH / 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFF7E4)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFF3DA)), GradientDirectionCts.gdVertical)
				
				' --Draw gradient in half button downside to give a glass look
				DrawGradientEx(1, lH / 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFE9C1)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFDE1AE)), GradientDirectionCts.gdVertical)
				
				' --Draws left side gradient effects horizontal
				DrawGradientEx(1, 3, 5, lH / 2 - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFEECD)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFF7E4)), GradientDirectionCts.gdHorizontal) 'Left
				DrawGradientEx(1, lH / 2, 5, lH - (lH / 2) - 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFAD68F)), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFDE1AC)), 0.01), GradientDirectionCts.gdHorizontal) 'Left
				
				' --Draws right side gradient effects horizontal
				DrawGradientEx(lW - 6, 3, 5, lH / 2 - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFF7E4)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFEECD)), GradientDirectionCts.gdHorizontal) 'Right
				DrawGradientEx(lW - 6, lH / 2, 5, lH - (lH / 2) - 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFDE1AC)), 0.01), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFAD68F)), GradientDirectionCts.gdHorizontal) 'Right
				
				DrawPicwithCaption()
				' --Draws border rectangle
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HA77532))) 'outer
				DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.Color.White)) 'inner
				
			Case enumButtonStates.eStateDown
				
				' --Draw a gradent in full region
				DrawGradientEx(1, 1, lW - 1, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF6E4C2)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF6E4C2)), GradientDirectionCts.gdVertical)
				
				' --Draw gradient in half button downside to give a glass look
				DrawGradientEx(1, lH / 2, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF0D29A)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF0D29A)), GradientDirectionCts.gdVertical)
				
				' --Draws down rectangle
				
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H5C411D))) '
				DrawLineApi(1, 1, lW - 1, 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HB39C71))) '\Top Lines
				DrawLineApi(1, 2, lW - 1, 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HD6C6A9))) '/
				DrawLineApi(1, 3, lW - 1, 3, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HECD9B9))) '
				
				DrawLineApi(1, 1, 1, lH / 2 - 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HCFB073))) 'Left upper
				DrawLineApi(1, lH / 2, 1, lH - (lH / 2) - 1, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HC5912B))) 'Left Bottom
				
				' --Draws left side gradient effects horizontal
				DrawGradientEx(1, 3, 5, lH / 2 - 2, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE6C891)), 0.02), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF6E4C2)), -0.01), GradientDirectionCts.gdHorizontal) 'Left
				DrawGradientEx(1, lH / 2, 5, lH - (lH / 2) - 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDCAB4E)), 0.02), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF0D29A)), -0.01), GradientDirectionCts.gdHorizontal) 'Left
				
				' --Draws right side gradient effects horizontal
				DrawGradientEx(lW - 6, 3, 5, lH / 2 - 2, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF6E4C2)), -0.01), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE6C891)), 0.02), GradientDirectionCts.gdHorizontal) 'Right
				DrawGradientEx(lW - 6, lH / 2, 5, lH - (lH / 2) - 1, ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HF0D29A)), -0.01), ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDCAB4E)), 0.02), GradientDirectionCts.gdHorizontal) 'Right
				DrawPicwithCaption()
				
		End Select
		
		' --Draw a focus rectangle if button has focus
		
		If m_bParentActive Then
			If (m_bHasFocus Or m_bDefault) And vState = enumButtonStates.eStateNormal Then
				' --Draw darker outer rectangle
				DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HA77532)))
				' --Draw light inner rectangle
				DrawRectangle(1, 1, lW - 2, lH - 2, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFBD848)))
			End If
			
			If (m_bShowFocus And m_bHasFocus) Then
				SetRect(lpRect, 1.5, 1.5, lW - 2, lH - 2)
				'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				DrawFocusRect(hDC, lpRect)
			End If
		End If
		
		' --Create four corners which will be common to all states
		DrawCorners(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBE965F)))
		
	End Sub
	
	Private Sub DrawOutlook2007(ByVal vState As enumButtonStates)
		
		Dim lpRect As RECT
		Dim bColor As Integer
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height)
		lW = ClientRectangle.Width
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		
		If m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue Then
			DrawGradientEx(0, 0, lW, lH / 2.7, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HA9D9FF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H6FC0FF)), GradientDirectionCts.gdVertical)
			DrawGradientEx(0, lH / 2.7, lW, lH - (lH / 2.7), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H3FABFF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H75E1FF)), GradientDirectionCts.gdVertical)
			DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.34))
			If m_bMouseInCtl Then
				DrawGradientEx(0, 0, lW, lH / 2.7, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H58C1FF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H51AFFF)), GradientDirectionCts.gdVertical)
				DrawGradientEx(0, lH / 2.7, lW, lH - (lH / 2.7), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H468FFF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H5FD3FF)), GradientDirectionCts.gdVertical)
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.34))
			End If
			DrawPicwithCaption()
			Exit Sub
		End If
		
		Select Case vState
			Case enumButtonStates.eStateNormal
				PaintRect(bColor, m_ButtonRect)
				DrawGradientEx(0, 0, lW, lH / 2.7, BlendColors(ShiftColor(bColor, 0.09), TranslateColor(System.Drawing.Color.White)), BlendColors(ShiftColor(bColor, 0.07), bColor), GradientDirectionCts.gdVertical)
				DrawGradientEx(0, lH / 2.7, lW, lH - (lH / 2.7), bColor, ShiftColor(bColor, 0.03), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.34))
			Case enumButtonStates.eStateOver
				DrawGradientEx(0, 0, lW, lH / 2.7, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE1FFFF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&HACEAFF)), GradientDirectionCts.gdVertical)
				DrawGradientEx(0, lH / 2.7, lW, lH - (lH / 2.7), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H67D7FF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H99E4FF)), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.34))
			Case enumButtonStates.eStateDown
				DrawGradientEx(0, 0, lW, lH / 2.7, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H58C1FF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H51AFFF)), GradientDirectionCts.gdVertical)
				DrawGradientEx(0, lH / 2.7, lW, lH - (lH / 2.7), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H468FFF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H5FD3FF)), GradientDirectionCts.gdVertical)
				DrawPicwithCaption()
				DrawRectangle(0, 0, lW, lH, ShiftColor(bColor, -0.34))
		End Select
		
	End Sub
	
	Private Sub DrawOffice2003(ByVal vState As enumButtonStates)
		
		Dim lpRect As RECT
		Dim bColor As Integer
		
		lH = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		lW = MyBase.ClientRectangle.Width
		
		bColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor))
		SetRect(m_ButtonRect, 0, 0, lW, lH)
		
		If m_ButtonMode <> enumButtonModes.ebmCommandButton And m_bValue Then
			If m_bMouseInCtl Then
				DrawGradientEx(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H4E91FE)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H8ED3FF)), GradientDirectionCts.gdVertical)
			Else
				DrawGradientEx(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H8CD5FF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H55ADFF)), GradientDirectionCts.gdVertical)
			End If
			DrawPicwithCaption()
			DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H800000)))
			Exit Sub
		End If
		
		Select Case vState
			
			Case enumButtonStates.eStateNormal
				CreateRegion()
				DrawGradientEx(0, 0, lW, lH / 2, BlendColors(TranslateColor(System.Drawing.Color.White), ShiftColor(bColor, 0.08)), bColor, GradientDirectionCts.gdVertical)
				DrawGradientEx(0, lH / 2, lW, lH / 2 + 1, bColor, ShiftColor(bColor, -0.15), GradientDirectionCts.gdVertical)
			Case enumButtonStates.eStateOver
				DrawGradientEx(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&HCCF4FF)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H91D0FF)), GradientDirectionCts.gdVertical)
			Case enumButtonStates.eStateDown
				DrawGradientEx(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H4E91FE)), TranslateColor(System.Drawing.ColorTranslator.FromOle(&H8ED3FF)), GradientDirectionCts.gdVertical)
		End Select
		
		DrawPicwithCaption()
		
		If m_Buttonstate <> enumButtonStates.eStateNormal Then
			DrawRectangle(0, 0, lW, lH, TranslateColor(System.Drawing.ColorTranslator.FromOle(&H800000)))
		End If
		
	End Sub
	
	Private Sub WindowsThemeButton(ByVal vState As enumButtonStates)
		
		Dim tmpState As Integer
		
		MyBase.BackColor = System.Drawing.ColorTranslator.FromOle(GetSysColor(COLOR_BTNFACE))
		
		If Not m_bEnabled Then
			tmpState = 4
			DrawTheme("Button", 1, tmpState)
			DrawPicwithCaption()
			Exit Sub
		End If
		
		Select Case vState
			Case enumButtonStates.eStateNormal
				tmpState = 1
			Case enumButtonStates.eStateOver
				tmpState = 2
			Case enumButtonStates.eStateDown
				tmpState = 3
		End Select
		
		If m_Buttonstate = enumButtonStates.eStateNormal Then
			If (m_bHasFocus Or m_bDefault) And m_bParentActive Then
				tmpState = 5
			End If
		End If
		
		DrawTheme("Button", 1, tmpState)
		DrawPicwithCaption()
		
	End Sub
	
	Private Function DrawTheme(ByRef sClass As String, ByVal iPart As Integer, ByVal vState As Integer) As Boolean
		
		Dim hTheme As Integer
		Dim lResult As Boolean
		Dim m_btnRect As RECT
		Dim hrgn As Integer
		
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		hTheme = OpenThemeData(MyBase.Handle.ToInt32, StrPtr(sClass))
		If hTheme Then
			' --Necessary for rounded buttons
			SetRect(m_btnRect, m_ButtonRect.Left_Renamed - 1, m_ButtonRect.Top_Renamed - 1, m_ButtonRect.Right_Renamed + 1, m_ButtonRect.Bottom_Renamed + 2)
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			GetThemeBackgroundRegion(hTheme, hDC, iPart, vState, m_btnRect, hrgn)
			SetWindowRgn(hwnd, hrgn, True)
			' --clean up
			DeleteObject(hrgn)
			' --Draw the theme
			'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			lResult = DrawThemeBackground(hTheme, hDC, iPart, vState, m_ButtonRect, m_ButtonRect)
			DrawTheme = lResult
		Else
			DrawTheme = False
		End If
		
	End Function
	
	Private Sub PaintRect(ByVal lColor As Integer, ByRef lpRect As RECT)
		
		'Fills a region with specified color
		
		Dim hOldBrush As Integer
		Dim hBrush As Integer
		
		hBrush = CreateSolidBrush(lColor)
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		hOldBrush = SelectObject(hDC, hBrush)
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		FillRect(hDC, lpRect, hBrush)
		
		'UPGRADE_ISSUE: UserControl property jcbutton.hDC was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		SelectObject(hDC, hOldBrush)
		DeleteObject(hBrush)
		
	End Sub
	
	Private Sub ShowPopupMenu()
		
		'* Shows a popupmenu
		'* Inspired from Noel Dacara's dcbutton
		
		Const TPM_BOTTOMALIGN As Integer = &H20
		
		Dim Menu As System.Windows.Forms.ToolStripMenuItem
		Dim Align As enumMenuAlign
		Dim flags As Integer
		Dim DefaultMenu As System.Windows.Forms.ToolStripMenuItem
		
		Dim X As Integer
		Dim Y As Integer
		
		Menu = DropDownMenu
		Align = MenuAlign
		flags = MenuFlags
		DefaultMenu = DefaultMenu
		
		lH = VB6.PixelsToTwipsY(ClientRectangle.Height) : lW = ClientRectangle.Width
		
		m_bPopupInit = True
		
		' --Set the drop down menu position
		Select Case Align
			Case enumMenuAlign.edaBottom
				Y = lH
				
			Case enumMenuAlign.edaLeft, enumMenuAlign.edaBottomLeft
				'UPGRADE_ISSUE: Constant vbPopupMenuRightAlign was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				MenuFlags = MenuFlags Or vbPopupMenuRightAlign
				If (MenuAlign = enumMenuAlign.edaBottomLeft) Then
					Y = lH
				End If
				
			Case enumMenuAlign.edaRight, enumMenuAlign.edaBottomRight
				X = lW
				If (MenuAlign = enumMenuAlign.edaBottomRight) Then
					Y = lH
				End If
				
			Case enumMenuAlign.edaTop, enumMenuAlign.edaTopRight, enumMenuAlign.edaTopLeft
				MenuFlags = TPM_BOTTOMALIGN
				If (MenuAlign = enumMenuAlign.edaTopRight) Then
					X = lW
				ElseIf (MenuAlign = enumMenuAlign.edaTopLeft) Then 
					'UPGRADE_ISSUE: Constant vbPopupMenuRightAlign was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
					MenuFlags = MenuFlags Or vbPopupMenuRightAlign
				End If
				
			Case Else
				m_bPopupInit = False
				
		End Select
		
		Dim lpPoint As POINT
		If (m_bPopupInit) Then
			
			' /--Show the dropdown menu
			If (DefaultMenu Is Nothing) Then
				'UPGRADE_ISSUE: UserControl method UserControl.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				MyBase.PopupMenu(DropDownMenu, MenuFlags, X, Y)
			Else
				'UPGRADE_ISSUE: UserControl method UserControl.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				MyBase.PopupMenu(DropDownMenu, MenuFlags, X, Y, DefaultMenu)
			End If
			
			GetCursorPos(lpPoint)
			
			If (WindowFromPoint(lpPoint.X, lpPoint.Y) = MyBase.Handle.ToInt32) Then
				m_bPopupShown = True
			Else
				m_bIsDown = False
				m_bMouseInCtl = False
				m_bIsSpaceBarDown = False
				m_Buttonstate = enumButtonStates.eStateNormal
				m_bPopupShown = False
				m_bPopupInit = False
				RedrawButton()
			End If
		End If
		
	End Sub
	
	Private Function ShiftColor(ByRef Color As Integer, ByRef PercentInDecimal As Single) As Integer
		
		'****************************************************************************
		'* This routine shifts a color value specified by PercentInDecimal          *
		'* Function inspired from DCbutton                                          *
		'* All Credits goes to Noel Dacara                                          *
		'* A Littlebit modified by me                                               *
		'****************************************************************************
		
		Dim R As Integer
		Dim G As Integer
		Dim B As Integer
		
		'  Add or remove a certain color quantity by how many percent.
		
		R = Color And 255
		G = (Color \ 256) And 255
		B = (Color \ 65536) And 255
		
		R = R + PercentInDecimal * 255 ' Percent should already
		G = G + PercentInDecimal * 255 ' be translated.
		B = B + PercentInDecimal * 255 ' Ex. 50% -> 50 / 100 = 0.5
		
		'  When overflow occurs, ....
		If (PercentInDecimal > 0) Then ' RGB values must be between 0-255 only
			If (R > 255) Then R = 255
			If (G > 255) Then G = 255
			If (B > 255) Then B = 255
		Else
			If (R < 0) Then R = 0
			If (G < 0) Then G = 0
			If (B < 0) Then B = 0
		End If
		
		ShiftColor = R + 256 * G + 65536 * B ' Return shifted color value
		
	End Function
	
	'UPGRADE_ISSUE: UserControl event UserControl.AccessKeyPress was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_AccessKeyPress(ByRef KeyAscii As Short)
		
		If m_bEnabled Then 'Disabled?? get out!!
			If m_bIsSpaceBarDown Then
				m_bIsSpaceBarDown = False
				m_bIsDown = False
			End If
			If m_ButtonMode = enumButtonModes.ebmCheckBox Then 'Checkbox Mode?
				If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape'
				m_bValue = Not m_bValue 'Change Value (Checked/Unchecked)
				If Not m_bValue Then 'If value unchecked then
					m_Buttonstate = enumButtonStates.eStateNormal 'Normal State
				End If
				RedrawButton()
			ElseIf m_ButtonMode = enumButtonModes.ebmOptionButton Then 
				If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape'
				UncheckAllValues()
				m_bValue = True
				RedrawButton()
			End If
			System.Windows.Forms.Application.DoEvents() 'To remove focus from other button and Do events before click event
			RaiseEvent Click(Me, Nothing) 'Now Raiseevent
		End If
		
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.AmbientChanged was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_AmbientChanged(ByRef PropertyName As String)
		
		'UPGRADE_ISSUE: AmbientProperties property Ambient.DisplayAsDefault was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		m_bDefault = Ambient.DisplayAsDefault
		If PropertyName = "DisplayAsDefault" Then
			RedrawButton()
		End If
		
		If PropertyName = "BackColor" Then
			RedrawButton()
		End If
		
	End Sub
	
	Private Sub jcbutton_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.DoubleClick
		
		If m_bHandPointer Then
			SetCursor(m_lCursor)
		End If
		
		If m_lDownButton = 1 Then 'React to only Left button
			
			SetCapture(hwnd) 'Preserve Hwnd on DoubleClick
			If m_Buttonstate <> enumButtonStates.eStateDown Then m_Buttonstate = enumButtonStates.eStateDown
			RedrawButton()
			jcbutton_MouseDown(Me, New System.Windows.Forms.MouseEventArgs(m_lDownButton * &H100000, 0, VB6.TwipsToPixelsX(m_lDX), VB6.TwipsToPixelsY(m_lDY), 0))
			If Not m_bPopupEnabled Then
				RaiseEvent DblClick(Me, Nothing)
			Else
				If Not m_bPopupShown Then
					ShowPopupMenu()
				End If
			End If
		End If
		
	End Sub
	
	Private Sub jcbutton_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		
		m_bHasFocus = True
		
	End Sub
	
	
	'UPGRADE_ISSUE: UserControl event UserControl.Hide was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_Hide()
		
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyBase.Extender.ToolTipText = m_sTooltipText
		
	End Sub
	
	Private Sub UserControl_Initialize()
		
		Dim i As Integer
		Dim OS As OSVERSIONINFO
		
		'Prebuid Lighten/Darken arrays
		For i = 0 To 255
			aLighten(i) = Lighten(i)
			aDarken(i) = Darken(i)
		Next 
		
		' --Get the operating system version for text drawing purposes.
		m_hMode = LoadLibraryA("shell32.dll")
		OS.dwOSVersionInfoSize = Len(OS)
		GetVersionEx(OS)
		m_WindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
		
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_InitProperties()
		
		'Initialize Properties for User Control
		'Called on designtime everytime a control is added
		
		m_ButtonStyle = enumButtonStlyes.eVistaAero 'As all the commercial buttons initialize with this them, ;)
		m_bShowFocus = True
		m_bEnabled = True
		'UPGRADE_ISSUE: AmbientProperties property Ambient.DisplayName was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		m_Caption = Ambient.DisplayName
		MyBase.Font = VB6.FontChangeName(MyBase.Font, "Tahoma")
		mFont = MyBase.Font
		'UPGRADE_ISSUE: Font event mFont.FontChanged was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
		mFont_FontChanged((vbNullString))
		m_PictureOpacity = 255
		m_PicOpacityOnOver = 255
		m_PictureAlign = enumPictureAlign.epLeftOfCaption
		m_bUseMaskColor = True
		m_lMaskColor = &HE0E0E0
		m_CaptionAlign = enumCaptionAlign.ecCenterAlign
		lH = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
		lW = MyBase.ClientRectangle.Width
		InitThemeColors()
		SetThemeColors()
		Refresh()
		
	End Sub
	
	Private Sub jcbutton_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Select Case KeyCode
			Case 13 'Enter Key
				RaiseEvent Click(Me, Nothing)
			Case 37, 38 'Left and Up Arrows
				System.Windows.Forms.SendKeys.Send("+{TAB}") 'Button should transfer focus to other ctl
			Case 39, 40 'Right and Down Arrows
				System.Windows.Forms.SendKeys.Send("{TAB}") 'Button should transfer focus to other ctl
			Case 32 'SpaceBar held down
				If Shift = 4 Then Exit Sub 'System Menu Should pop up
				If Not m_bIsDown Then
					m_bIsSpaceBarDown = True 'Set space bar as pressed
					
					If (m_ButtonMode = enumButtonModes.ebmCheckBox) Then 'Is CheckBoxMode??
						m_bValue = Not m_bValue 'Toggle Check Value
					ElseIf m_ButtonMode = enumButtonModes.ebmOptionButton Then 
						UncheckAllValues() 'Option Button Mode
						m_bValue = True 'Pressed button Checked
					End If
					
					If m_Buttonstate <> enumButtonStates.eStateDown Then
						m_Buttonstate = enumButtonStates.eStateDown 'Button state should be down
						RedrawButton()
					End If
				Else
					If m_bMouseInCtl Then
						If m_Buttonstate <> enumButtonStates.eStateDown Then
							m_Buttonstate = enumButtonStates.eStateDown
							RedrawButton()
						End If
					Else
						If m_Buttonstate <> enumButtonStates.eStateNormal Then
							m_Buttonstate = enumButtonStates.eStateNormal 'jump button from
							RedrawButton() 'downstate - normal state
						End If 'if mouse button is pressed
					End If
				End If 'when spacebar being held
				
				If (Not GetCapture = MyBase.Handle.ToInt32) Then
					ReleaseCapture()
					SetCapture(MyBase.Handle.ToInt32) 'No other processing until spacebar is released
				End If 'Thanks to APIGuide
				
			Case Else
				If m_bIsSpaceBarDown Then
					m_bIsSpaceBarDown = False
					m_Buttonstate = enumButtonStates.eStateNormal
					RedrawButton()
				End If
		End Select
		
		RaiseEvent KeyDown(Me, New KeyDownEventArgs(KeyCode, Shift))
		
		
	End Sub
	
	Private Sub jcbutton_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		' --Simply raise the event =)
		RaiseEvent KeyPress(Me, New KeyPressEventArgs(KeyAscii))
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub jcbutton_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = System.Windows.Forms.Keys.Space Then
			
			ReleaseCapture() 'Now you can process further
			'as the spacebar is released
			If m_bMouseInCtl And m_bIsDown Then
				If m_Buttonstate <> enumButtonStates.eStateDown Then
					m_Buttonstate = enumButtonStates.eStateDown
					RedrawButton()
				End If
			ElseIf m_bMouseInCtl Then  'If spacebar released over ctl
				If m_Buttonstate <> enumButtonStates.eStateOver Then
					m_Buttonstate = enumButtonStates.eStateOver 'Draw Hover State
					RedrawButton()
				End If
				If Not m_bIsDown And m_bIsSpaceBarDown Then
					RaiseEvent Click(Me, Nothing)
				End If
			Else 'If Spacebar released outside ctl
				If m_Buttonstate <> enumButtonStates.eStateNormal Then
					m_Buttonstate = enumButtonStates.eStateNormal
					RedrawButton()
				End If
				If Not m_bIsDown And m_bIsSpaceBarDown Then
					RaiseEvent Click(Me, Nothing)
				End If
			End If
			
			RaiseEvent KeyUp(Me, New KeyUpEventArgs(KeyCode, Shift))
			m_bIsSpaceBarDown = False
			m_bIsDown = False
		End If
		
	End Sub
	
	Private Sub jcbutton_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		
		m_bHasFocus = False 'No focus
		m_bIsDown = False 'No down state
		m_bIsSpaceBarDown = False 'No spacebar held
		If Not m_bParentActive Then
			If m_Buttonstate <> enumButtonStates.eStateNormal Then
				m_Buttonstate = enumButtonStates.eStateNormal
			End If
		ElseIf m_bMouseInCtl Then 
			If m_Buttonstate <> enumButtonStates.eStateOver Then
				m_Buttonstate = enumButtonStates.eStateOver
			End If
		Else
			If m_Buttonstate <> enumButtonStates.eStateNormal Then
				m_Buttonstate = enumButtonStates.eStateNormal
			End If
		End If
		RedrawButton()
		
		If m_bDefault Then 'If default button,
			RedrawButton() 'Show Focus
		End If
		
	End Sub
	
	Private Sub jcbutton_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		m_lDownButton = Button 'Button pressed for Dblclick
		m_lDX = X
		m_lDY = Y
		m_lDShift = Shift
		
		' --Set HandPointer if any!
		If m_bHandPointer Then
			SetCursor(m_lCursor)
		End If
		
		If Button = VB6.MouseButtonConstants.LeftButton Or m_bPopupShown Then
			m_bHasFocus = True
			m_bIsDown = True
			
			If (Not m_bIsSpaceBarDown) Then
				If m_Buttonstate <> enumButtonStates.eStateDown Then
					m_Buttonstate = enumButtonStates.eStateDown
					RedrawButton()
				End If
			End If
			
			If Not m_bPopupEnabled Then
				RaiseEvent MouseDown(Me, New MouseDownEventArgs(Button, Shift, X, Y))
			Else
				If Not m_bPopupShown Then
					ShowPopupMenu()
				End If
			End If
		End If
		
	End Sub
	
	Private Sub CreateToolTip()
		
		'****************************************************************************
		'* A very nice and flexible sub to create balloon tool tips
		'* Author :- Fred.CPP
		'* Added as requested by many users
		'* Modified by me to support unicode
		'* Thanks Alfredo ;)
		'****************************************************************************
		
		Dim lpRect As RECT
		Dim lWinStyle As Integer
		Dim lPtr As Integer
		Dim ttip As TOOLINFO
		Dim ttipW As TOOLINFOW
		Const CS_DROPSHADOW As Integer = &H20000
		Const GCL_STYLE As Integer = (-26)
		
		' --Dont show tooltips if disabled
		If (Not m_bEnabled) Or m_bPopupShown Or m_Buttonstate = enumButtonStates.eStateDown Then Exit Sub
		
		' --Destroy any previous tooltip
		If m_lttHwnd <> 0 Then
			DestroyWindow(m_lttHwnd)
		End If
		
		lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
		
		''create baloon style if desired
		If m_lTooltipType = enumTooltipStyle.TooltipBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
		
		If m_bttRTL Then
			m_lttHwnd = CreateWindowEx(WS_EX_LAYOUTRTL, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, MyBase.Handle.ToInt32, 0, VB6.GetHInstance.ToInt32, 0)
		Else
			m_lttHwnd = CreateWindowEx(0, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, MyBase.Handle.ToInt32, 0, VB6.GetHInstance.ToInt32, 0)
		End If
		
		SetClassLong(m_lttHwnd, GCL_STYLE, GetClassLong(m_lttHwnd, GCL_STYLE) Or CS_DROPSHADOW)
		
		'make our tooltip window a topmost window
		' This is creating some problems as noted by K-Zero
		'SetWindowPos m_lttHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
		
		''get the rect of the parent control
		GetClientRect(MyBase.Handle.ToInt32, lpRect)
		
		If m_WindowsNT Then
			' --set our tooltip info structure  for UNICODE SUPPORT >> WinNT
			With ttipW
				' --if we want it centered, then set that flag
				If m_lttCentered Then
					.lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
				Else
					.lFlags = TTF_SUBCLASS Or TTF_IDISHWND
				End If
				
				' --set the hwnd prop to our parent control's hwnd
				.lhWnd = MyBase.Handle.ToInt32
				.lId = hwnd
				.lSize = Len(ttipW)
				.hInstance = VB6.GetHInstance.ToInt32
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				.lpStrW = StrPtr(m_sTooltipText)
				'UPGRADE_WARNING: Couldn't resolve default property of object ttipW.lpRect. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.lpRect = lpRect
			End With
			' --add the tooltip structure
			'UPGRADE_WARNING: Couldn't resolve default property of object ttipW. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SendMessage(m_lttHwnd, TTM_ADDTOOLW, 0, ttipW)
		Else
			' --set our tooltip info structure for << WinNT
			With ttip
				''if we want it centered, then set that flag
				If m_lttCentered Then
					.lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
				Else
					.lFlags = TTF_SUBCLASS
				End If
				
				' --set the hwnd prop to our parent control's hwnd
				.lhWnd = MyBase.Handle.ToInt32
				.lId = hwnd
				.lSize = Len(ttip)
				.hInstance = VB6.GetHInstance.ToInt32
				.lpStr = m_sTooltipText
				'UPGRADE_WARNING: Couldn't resolve default property of object ttip.lpRect. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.lpRect = lpRect
			End With
			' --add the tooltip structure
			'UPGRADE_WARNING: Couldn't resolve default property of object ttip. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SendMessage(m_lttHwnd, TTM_ADDTOOLA, 0, ttip)
		End If
		
		'if we want a title or we want an icon
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If LenB(m_sTooltiptitle) > 0 Or m_lToolTipIcon <> enumIconType.TTNoIcon Then
			If m_WindowsNT Then
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				lPtr = StrPtr(m_sTooltiptitle)
				If lPtr Then
					SendMessage(m_lttHwnd, TTM_SETTITLEW, m_lToolTipIcon, lPtr)
				End If
			Else
				SendMessage(m_lttHwnd, TTM_SETTITLE, CInt(m_lToolTipIcon), m_sTooltiptitle)
			End If
			
		End If
		SendMessage(m_lttHwnd, TTM_SETMAXTIPWIDTH, 0, 240) 'for Multiline capability
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not IsNothing(m_lttBackColor) Then
			SendMessage(m_lttHwnd, TTM_SETTIPBKCOLOR, TranslateColor(System.Drawing.ColorTranslator.FromOle(m_lttBackColor)), 0)
		End If
		
	End Sub
	
	Private Sub InitThemeColors()
		
		Select Case m_ButtonStyle
			Case enumButtonStlyes.eStandard, enumButtonStlyes.eFlat, enumButtonStlyes.eVistaToolbar, enumButtonStlyes.eXPToolbar, enumButtonStlyes.eOfficeXP, enumButtonStlyes.eWindowsXP, enumButtonStlyes.eOutlook2007, enumButtonStlyes.eGelButton
				m_lXPColor = enumXPThemeColors.ecsBlue
			Case enumButtonStlyes.eInstallShield, enumButtonStlyes.eVistaAero
				m_lXPColor = enumXPThemeColors.ecsSilver
		End Select
		
	End Sub
	
	Private Sub SetThemeColors()
		
		'Sets a style colors to default colors when button initialized
		'or whenever you change the style of Button
		
		With m_bColors
			
			Select Case m_ButtonStyle
				
				Case enumButtonStlyes.eStandard, enumButtonStlyes.eFlat, enumButtonStlyes.eVistaToolbar, enumButtonStlyes.e3DHover, enumButtonStlyes.eFlatHover, enumButtonStlyes.eXPToolbar, enumButtonStlyes.eOfficeXP
					.tBackColor = GetSysColor(COLOR_BTNFACE)
				Case enumButtonStlyes.eWindowsXP
					Select Case m_lXPColor
						Case enumXPThemeColors.ecsBlue
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE7EBEC))
						Case enumXPThemeColors.ecsOliveGreen
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDBEEF3))
						Case enumXPThemeColors.ecsSilver
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFCF1F0))
					End Select
				Case enumButtonStlyes.eOutlook2007, enumButtonStlyes.eGelButton
					Select Case m_lXPColor
						Case enumXPThemeColors.ecsBlue
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFD1AD))
						Case enumXPThemeColors.ecsOliveGreen
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBAD6D4))
						Case enumXPThemeColors.ecsSilver
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE3DFE0))
					End Select
					.tForeColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&H8B4215))
				Case enumButtonStlyes.eVistaAero
					Select Case m_lXPColor
						Case enumXPThemeColors.ecsBlue
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFDECE0))
						Case enumXPThemeColors.ecsOliveGreen
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HDEEDE8))
						Case enumXPThemeColors.ecsSilver
							.tBackColor = ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HD4D4D4)), 0.06)
					End Select
				Case enumButtonStlyes.eInstallShield
					Select Case m_lXPColor
						Case enumXPThemeColors.ecsBlue
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFFD1AD))
						Case enumXPThemeColors.ecsOliveGreen
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBAD6D4))
						Case enumXPThemeColors.ecsSilver
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HE1D6D5))
					End Select
				Case enumButtonStlyes.eOffice2003
					Select Case m_lXPColor
						Case enumXPThemeColors.ecsBlue
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HFCE1CA))
						Case enumXPThemeColors.ecsOliveGreen
							.tBackColor = TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBAD6D4))
						Case enumXPThemeColors.ecsSilver
							.tBackColor = ShiftColor(TranslateColor(System.Drawing.ColorTranslator.FromOle(&HBA9EA0)), 0.15)
					End Select
			End Select
			
			.tForeColor = TranslateColor(System.Drawing.SystemColors.ControlText)
			If m_ButtonStyle = enumButtonStlyes.eFlat Or m_ButtonStyle = enumButtonStlyes.eInstallShield Or m_ButtonStyle = enumButtonStlyes.eStandard Then
				m_bShowFocus = True
			Else
				m_bShowFocus = False
			End If
			
			If m_ButtonStyle = enumButtonStlyes.eOfficeXP Then
				m_bPicPushOnHover = True
			End If
			
		End With
		
	End Sub
	
	Private Sub jcbutton_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		Dim lp As POINT
		
		GetCursorPos(lp)
		' --Set hand pointer if any!
		If m_bHandPointer Then
			SetCursor(m_lCursor)
		End If
		
		If Not (WindowFromPoint(lp.X, lp.Y) = MyBase.Handle.ToInt32) Then
			' --Mouse yet not entered in the control
			m_bMouseInCtl = False
		Else
			m_bMouseInCtl = True
			' --Check when the Mouse leaves the control
			TrackMouseLeave(hwnd)
			' --Raise a MouseEnter event(it's Same as mouseMove)
			RaiseEvent MouseEnter(Me, Nothing)
		End If
		
		' --Proceed only if spacebar is not pressed
		If m_bIsSpaceBarDown Then Exit Sub
		
		' --We are inside button
		If m_bMouseInCtl Then
			
			' --Mouse button is pressed down
			If m_bIsDown Then
				If m_Buttonstate <> enumButtonStates.eStateDown Then
					m_Buttonstate = enumButtonStates.eStateDown
					RedrawButton()
				End If
			Else
				' --Button should be in hot state if user leaves the button
				' --with mouse button pressed
				If m_Buttonstate <> enumButtonStates.eStateOver Then
					m_Buttonstate = enumButtonStates.eStateOver
					RedrawButton()
					' --Create Tooltip Here
					If m_Buttonstate <> enumButtonStates.eStateDown Then
						CreateToolTip()
					End If
				End If
			End If
			
		Else
			If m_Buttonstate <> enumButtonStates.eStateNormal Then
				m_Buttonstate = enumButtonStates.eStateNormal
				RedrawButton()
			End If
		End If
		
		'RaiseEvent MouseMove(Button, Shift, x, Y)
		
	End Sub
	
	Private Sub jcbutton_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		If m_bHandPointer Then
			SetCursor(m_lCursor)
		End If
		
		' --Popupmenu enabled
		If m_bPopupEnabled Then
			m_bIsDown = False
			m_bPopupShown = False
			m_Buttonstate = enumButtonStates.eStateNormal
			RedrawButton()
			Exit Sub
		End If
		
		' --React only to Left mouse button
		If Button = VB6.MouseButtonConstants.LeftButton Then
			'--Button released
			m_bIsDown = False
			' --If button released in button area
			If (X > 0 And Y > 0) And (X < ClientRectangle.Width And Y < VB6.PixelsToTwipsY(ClientRectangle.Height)) Then
				
				' --If check box mode
				If m_ButtonMode = enumButtonModes.ebmCheckBox Then
					m_bValue = Not m_bValue
					RedrawButton()
					' --If option button mode
				ElseIf m_ButtonMode = enumButtonModes.ebmOptionButton Then 
					UncheckAllValues()
					m_bValue = True
				End If
				
				' --redraw Normal State
				m_Buttonstate = enumButtonStates.eStateNormal
				RedrawButton()
				RaiseEvent Click(Me, Nothing)
			End If
		End If
		
		RaiseEvent MouseUp(Me, New MouseUpEventArgs(Button, Shift, X, Y))
		
	End Sub
	
	Private Sub jcbutton_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		' --At least, a checkbox will also need this much of size!!!!
		If VB6.PixelsToTwipsY(Height) < 220 Then Height = VB6.TwipsToPixelsY(220)
		If VB6.PixelsToTwipsX(Width) < 220 Then Width = VB6.TwipsToPixelsX(220)
		
		' --On resize, create button region again
		CreateRegion()
		RedrawButton() 'then redraw
		
	End Sub
	
	Private Sub jcbutton_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		
		' --this routine typically called by Windows when another window covering
		'   this button is removed, or when the parent is moved/minimized/etc.
		
		RedrawButton()
		
	End Sub
	
	'Load property values from storage
	
	'UPGRADE_ISSUE: VBRUN.PropertyBag type was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	'UPGRADE_WARNING: UserControl event ReadProperties is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="92F3B58C-F772-4151-BE90-09F4A232AEAD"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As Object)
		
		With PropBag
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_ButtonStyle = .ReadProperty("ButtonStyle", enumButtonStlyes.eFlat)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bShowFocus = .ReadProperty("ShowFocusRect", False)
			'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			mFont = .ReadProperty("Font", Ambient.Font)
			MyBase.Font = mFont
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bColors.tBackColor = .ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bEnabled = .ReadProperty("Enabled", True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_Caption = .ReadProperty("Caption", "jcbutton")
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bValue = .ReadProperty("Value", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			MyBase.Cursor = .ReadProperty("MousePointer", 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bHandPointer = .ReadProperty("HandPointer", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			MyBase.MouseIcon = .ReadProperty("MouseIcon", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			m_Picture = .ReadProperty("PictureNormal", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			m_PictureHot = .ReadProperty("PictureHot", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			m_PictureDown = .ReadProperty("PictureDown", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PictureShadow = .ReadProperty("PictureShadow", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PictureOpacity = .ReadProperty("PictureOpacity", 255)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PicOpacityOnOver = .ReadProperty("PictureOpacityOnOver", 255)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PicDisabledMode = .ReadProperty("DisabledPictureMode", enumDisabledPicMode.edpBlended)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bPicPushOnHover = .ReadProperty("PicturePushOnHover", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PicEffectonOver = .ReadProperty("PictureEffectOnOver", enumPicEffect.epeLighter)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PicEffectonDown = .ReadProperty("PictureEffectOnDown", enumPicEffect.epeDarker)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lMaskColor = .ReadProperty("MaskColor", &HE0E0E0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bUseMaskColor = .ReadProperty("UseMaskColor", True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_CaptionEffects = .ReadProperty("CaptionEffects", enumCaptionEffects.eseNone)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_ButtonMode = .ReadProperty("Mode", enumButtonModes.ebmCommandButton)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_PictureAlign = .ReadProperty("PictureAlign", enumPictureAlign.epLeftOfCaption)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_CaptionAlign = .ReadProperty("CaptionAlign", enumCaptionAlign.ecCenterAlign)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bColors.tForeColor = .ReadProperty("ForeColor", TranslateColor(System.Drawing.SystemColors.ControlText))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bColors.tForeColorOver = .ReadProperty("ForeColorHover", TranslateColor(System.Drawing.SystemColors.ControlText))
			MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColor)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bDropDownSep = .ReadProperty("DropDownSeparator", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sTooltiptitle = .ReadProperty("TooltipTitle", vbNullString)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_sTooltipText = .ReadProperty("ToolTip", vbNullString)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lToolTipIcon = .ReadProperty("TooltipIcon", enumIconType.TTNoIcon)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lTooltipType = .ReadProperty("TooltipType", enumTooltipStyle.TooltipStandard)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lttBackColor = .ReadProperty("TooltipBackColor", TranslateColor(System.Drawing.SystemColors.Info))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bRTL = .ReadProperty("RightToLeft", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_bttRTL = .ReadProperty("RightToLeft", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_DropDownSymbol = .ReadProperty("DropDownSymbol", enumSymbol.ebsNone)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_lXPColor = .ReadProperty("ColorScheme", enumXPThemeColors.ecsBlue)
			MyBase.Enabled = m_bEnabled
			SetAccessKey()
			lH = VB6.PixelsToTwipsY(MyBase.ClientRectangle.Height)
			lW = MyBase.ClientRectangle.Width
			'UPGRADE_WARNING: Control property UserControl.Parent was upgraded to UserControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
			m_lParenthWnd = MyBase.FindForm.Handle.ToInt32
		End With
		
		If m_bHandPointer Then
			m_lCursor = LoadCursor(0, IDC_HAND) 'Load System Hand pointer
			m_bHandPointer = (Not m_lCursor = 0)
		End If
		
		jcbutton_Resize(Me, New System.EventArgs())
		
		On Error GoTo H
		If Not DesignMode Then 'If we're not in design mode
			TrackUser32 = IsFunctionSupported("TrackMouseEvent", "User32")
			
			If Not TrackUser32 Then IsFunctionSupported("_TrackMouseEvent", "ComCtl32")
			
			'OS supports mouse leave so subclass for it
			With Me
				'Start subclassing the UserControl
				Subclass_Initialize(.Handle.ToInt32)
				Subclass_Initialize(m_lParenthWnd)
				Subclass_AddMsg(.Handle.ToInt32, WM_MOUSELEAVE, MsgWhen.MSG_AFTER)
				Subclass_AddMsg(.Handle.ToInt32, WM_THEMECHANGED, MsgWhen.MSG_AFTER)
				If IsThemed Then
					Subclass_AddMsg(.Handle.ToInt32, WM_SYSCOLORCHANGE, MsgWhen.MSG_AFTER)
				End If
				On Error Resume Next
				'UPGRADE_WARNING: Control property UserControl.Parent was upgraded to UserControl.FindForm which has a new behavior. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
				'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Parent.MDIChild. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If MyBase.FindForm.MDIChild Then
					Call Subclass_AddMsg(m_lParenthWnd, WM_NCACTIVATE, MsgWhen.MSG_AFTER)
				Else
					Call Subclass_AddMsg(m_lParenthWnd, WM_ACTIVATE, MsgWhen.MSG_AFTER)
				End If
			End With
			
		End If
		
H: 
		
	End Sub
	
	'UPGRADE_ISSUE: UserControl event UserControl.Show was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub UserControl_Show()
		
		'UPGRADE_ISSUE: UserControl property UserControl.Extender was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_ISSUE: Object property UserControl.Extender.ToolTipText was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		'UPGRADE_WARNING: Couldn't resolve default property of object UserControl.Extender.ToolTipText. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		MyBase.Extender.ToolTipText = vbNullString
		
	End Sub
	
	'A nice place to stop subclasser
	
	Private Sub UserControl_Terminate()
		
		On Error GoTo Crash
		If m_lButtonRgn Then DeleteObject(m_lButtonRgn) 'Delete button region
		'UPGRADE_NOTE: Object mFont may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mFont = Nothing 'Clean up Font (StdFont)
		FreeLibrary(m_hMode)
		UnsetPopupMenu()
		If Not DesignMode Then
			Subclass_Terminate()
			Subclass_Terminate()
		End If
Crash: 
		
	End Sub
	
	'Write property values to storage
	
	'UPGRADE_ISSUE: VBRUN.PropertyBag type was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	'UPGRADE_WARNING: UserControl event WriteProperties is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="92F3B58C-F772-4151-BE90-09F4A232AEAD"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As Object)
		
		With PropBag
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("ButtonStyle", m_ButtonStyle, enumButtonStlyes.eFlat)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("ShowFocusRect", m_bShowFocus, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("Enabled", m_bEnabled, True)
			'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("Font", mFont, Ambient.Font)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("BackColor", m_bColors.tBackColor, GetSysColor(COLOR_BTNFACE))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("Caption", m_Caption, "jcbutton1")
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("ForeColor", m_bColors.tForeColor, TranslateColor(System.Drawing.SystemColors.ControlText))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("ForeColorHover", m_bColors.tForeColorOver, TranslateColor(System.Drawing.SystemColors.ControlText))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("Mode", m_ButtonMode, enumButtonModes.ebmCommandButton)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("Value", m_bValue, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("MousePointer", MyBase.Cursor, 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("HandPointer", m_bHandPointer, False)
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("MouseIcon", MyBase.MouseIcon, Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureNormal", m_Picture, Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureHot", m_PictureHot, Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureDown", m_PictureDown, Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureAlign", m_PictureAlign, enumPictureAlign.epLeftOfCaption)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureEffectOnOver", m_PicEffectonOver, enumPicEffect.epeLighter)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureEffectOnDown", m_PicEffectonDown, enumPicEffect.epeDarker)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PicturePushOnHover", m_bPicPushOnHover, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureShadow", m_PictureShadow, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureOpacity", m_PictureOpacity, 255)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("PictureOpacityOnOver", m_PicOpacityOnOver, 255)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("DisabledPictureMode", m_PicDisabledMode, enumDisabledPicMode.edpBlended)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("CaptionEffects", m_CaptionEffects, vbNullString)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("UseMaskColor", m_bUseMaskColor, True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("MaskColor", m_lMaskColor, &HE0E0E0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("CaptionAlign", m_CaptionAlign, enumCaptionAlign.ecCenterAlign)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("ToolTip", m_sTooltipText, vbNullString)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("TooltipType", m_lTooltipType, enumTooltipStyle.TooltipStandard)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("TooltipIcon", m_lToolTipIcon, enumIconType.TTNoIcon)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("TooltipTitle", m_sTooltiptitle, vbNullString)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("TooltipBackColor", m_lttBackColor, TranslateColor(System.Drawing.SystemColors.Info))
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("RightToLeft", m_bRTL, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("DropDownSymbol", m_DropDownSymbol, enumSymbol.ebsNone)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("DropDownSeparator", m_bDropDownSep, False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			.WriteProperty("ColorScheme", m_lXPColor, enumXPThemeColors.ecsBlue)
		End With
		
	End Sub
	
	Private Function Is32BitBMP(ByRef Obj As Object) As Boolean
		Dim uBI As BITMAP
		
		'UPGRADE_ISSUE: Constant vbPicTypeBitmap was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Obj.Type. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Obj.Type = vbPicTypeBitmap Then
			'UPGRADE_WARNING: Couldn't resolve default property of object uBI. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Obj.Handle. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call GetObject_Renamed(Obj.Handle, Len(uBI), uBI)
			Is32BitBMP = uBI.bmBitsPixel = 32
		End If
		
	End Function
	
	'Purpose: Returns True if DLL is present.
	Private Function IsDLLPresent(ByVal sDLL As String) As Boolean
		
		On Error GoTo NotPresent
		Dim hLib As Integer
		hLib = LoadLibraryA(sDLL)
		If hLib <> 0 Then
			FreeLibrary(hLib)
			IsDLLPresent = True
		End If
NotPresent: 
		
	End Function
	
	Private ReadOnly Property IsThemed() As Boolean
		Get
			
			On Error Resume Next
			Static m_bInit As Object
			If HasUxTheme Then
				If Not (m_bInit) Then
					m_bIsThemed = IsAppThemed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_bInit. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_bInit = True
				End If
			End If
			IsThemed = m_bIsThemed
			
		End Get
	End Property
	
	Private ReadOnly Property HasUxTheme() As Boolean
		Get
			
			Static m_bInit As Object
			If Not (m_bInit) Then
				m_bHasUxTheme = IsDLLPresent("uxtheme.dll")
				'UPGRADE_WARNING: Couldn't resolve default property of object m_bInit. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				m_bInit = True
			End If
			HasUxTheme = m_bHasUxTheme
			
		End Get
	End Property
	
	
	Public Overrides Property BackColor() As System.Drawing.Color
		Get
			
			BackColor = System.Drawing.ColorTranslator.FromOle(m_bColors.tBackColor)
			
		End Get
		Set(ByVal Value As System.Drawing.Color)
			
			m_bColors.tBackColor = System.Drawing.ColorTranslator.ToOle(Value)
			If m_ButtonStyle <> enumButtonStlyes.eOfficeXP Then
				m_lXPColor = enumXPThemeColors.ecsCustom
			End If
			RedrawButton()
			RaiseEvent BackColorChange()
			
		End Set
	End Property
	
	
	Public Property ButtonStyle() As enumButtonStlyes
		Get
			
			ButtonStyle = m_ButtonStyle
			
		End Get
		Set(ByVal Value As enumButtonStlyes)
			
			m_ButtonStyle = Value
			InitThemeColors()
			SetThemeColors() 'Set colors
			CreateRegion() 'Create Region Again
			RedrawButton() 'Obviously, force redraw!!!
			RaiseEvent ButtonStyleChange()
			
		End Set
	End Property
	
	
	Public Property Caption() As String
		Get
			
			Caption = m_Caption
			
		End Get
		Set(ByVal Value As String)
			
			m_Caption = Value
			SetAccessKey()
			RedrawButton()
			RaiseEvent CaptionChange()
			
		End Set
	End Property
	
	
	Public Property CaptionAlign() As enumCaptionAlign
		Get
			
			CaptionAlign = m_CaptionAlign
			
		End Get
		Set(ByVal Value As enumCaptionAlign)
			
			m_CaptionAlign = Value
			RedrawButton()
			RaiseEvent CaptionAlignChange()
			
		End Set
	End Property
	
	
	Public Property DropDownSymbol() As enumSymbol
		Get
			
			DropDownSymbol = m_DropDownSymbol
			
		End Get
		Set(ByVal Value As enumSymbol)
			
			m_DropDownSymbol = Value
			RedrawButton()
			RaiseEvent DropDownSymbolChange()
			
		End Set
	End Property
	
	
	Public Property DropDownSeparator() As Boolean
		Get
			
			DropDownSeparator = m_bDropDownSep
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bDropDownSep = Value
			RedrawButton()
			RaiseEvent DropDownSeparatorChange()
			
		End Set
	End Property
	
	
	Public Property DisabledPictureMode() As enumDisabledPicMode
		Get
			
			DisabledPictureMode = m_PicDisabledMode
			
		End Get
		Set(ByVal Value As enumDisabledPicMode)
			
			m_PicDisabledMode = Value
			RedrawButton()
			RaiseEvent DisabledPictureModeChange()
			
		End Set
	End Property
	
	
	Public Shadows Property Enabled() As Boolean
		Get
			
			Enabled = m_bEnabled
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bEnabled = Value
			MyBase.Enabled = m_bEnabled
			RedrawButton()
			RaiseEvent EnabledChange()
			
		End Set
	End Property
	
	
	Public Overrides Property Font() As System.Drawing.Font
		Get
			
			Font = mFont
			
		End Get
		Set(ByVal Value As System.Drawing.Font)
			
			mFont = Value
			Refresh()
			RedrawButton()
			RaiseEvent FontChange()
			'UPGRADE_ISSUE: Font event mFont.FontChanged was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
			mFont_FontChanged(vbNullString)
			
		End Set
	End Property
	
	
	Public Overrides Property ForeColor() As System.Drawing.Color
		Get
			
			ForeColor = System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColor)
			
		End Get
		Set(ByVal Value As System.Drawing.Color)
			
			m_bColors.tForeColor = System.Drawing.ColorTranslator.ToOle(Value)
			MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColor)
			RedrawButton()
			RaiseEvent ForeColorChange()
			
		End Set
	End Property
	
	
	Public Property ForeColorHover() As System.Drawing.Color
		Get
			
			ForeColorHover = System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColorOver)
			
		End Get
		Set(ByVal Value As System.Drawing.Color)
			
			m_bColors.tForeColorOver = System.Drawing.ColorTranslator.ToOle(Value)
			MyBase.ForeColor = System.Drawing.ColorTranslator.FromOle(m_bColors.tForeColorOver)
			RedrawButton()
			RaiseEvent ForeColorHoverChange()
			
		End Set
	End Property
	
	
	Public Property HandPointer() As Boolean
		Get
			
			HandPointer = m_bHandPointer
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bHandPointer = Value
			If m_bHandPointer Then
				MyBase.Cursor = System.Windows.Forms.Cursors.Default
			End If
			RedrawButton()
			RaiseEvent HandPointerChange()
			
		End Set
	End Property
	
	Public ReadOnly Property hwnd() As Integer
		Get
			
			' --Handle that uniquely identifies the control
			hwnd = MyBase.Handle.ToInt32
			
		End Get
	End Property
	
	
	Public Property MaskColor() As System.Drawing.Color
		Get
			
			MaskColor = System.Drawing.ColorTranslator.FromOle(m_lMaskColor)
			
		End Get
		Set(ByVal Value As System.Drawing.Color)
			
			m_lMaskColor = System.Drawing.ColorTranslator.ToOle(Value)
			RedrawButton()
			RaiseEvent MaskColorChange()
			
		End Set
	End Property
	
	
	Public Property Mode() As enumButtonModes
		Get
			
			Mode = m_ButtonMode
			
		End Get
		Set(ByVal Value As enumButtonModes)
			
			m_ButtonMode = Value
			If m_ButtonMode = enumButtonModes.ebmCommandButton Then
				m_Buttonstate = enumButtonStates.eStateNormal 'Force Normal State for command buttons
			End If
			RedrawButton()
			RaiseEvent ValueChange()
			RaiseEvent ModeChange()
			
		End Set
	End Property
	
	
	Public Property MouseIcon() As System.Drawing.Image
		Get
			
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			MouseIcon = MyBase.MouseIcon
			
		End Get
		Set(ByVal Value As System.Drawing.Image)
			
			On Error Resume Next
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			MyBase.MouseIcon = Value
			If (Value Is Nothing) Then
				MyBase.Cursor = System.Windows.Forms.Cursors.Default ' vbDefault
			Else
				m_bHandPointer = False
				RaiseEvent HandPointerChange()
				'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
				MyBase.Cursor = vbCustom ' vbCustom
			End If
			RaiseEvent MouseIconChange()
			
		End Set
	End Property
	
	
	Public Property MousePointer() As System.Windows.Forms.Cursor
		Get
			
			MousePointer = MyBase.Cursor
			
		End Get
		Set(ByVal Value As System.Windows.Forms.Cursor)
			
			'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
			MyBase.Cursor = Value
			RaiseEvent MousePointerChange()
			
		End Set
	End Property
	
	
	Public Property PictureNormal() As System.Drawing.Image
		Get
			
			PictureNormal = m_Picture
			
		End Get
		Set(ByVal Value As System.Drawing.Image)
			
			m_Picture = Value
			If Not Value Is Nothing Then
				RedrawButton()
				RaiseEvent PictureNormalChange()
			Else
				jcbutton_Resize(Me, New System.EventArgs())
				'UPGRADE_NOTE: Object m_PictureHot may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				m_PictureHot = Nothing
				'UPGRADE_NOTE: Object m_PictureDown may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				m_PictureDown = Nothing
				RaiseEvent PictureHotChange()
				RaiseEvent PictureDownChange()
			End If
			
		End Set
	End Property
	
	
	Public Property PictureHot() As System.Drawing.Image
		Get
			
			PictureHot = m_PictureHot
			
		End Get
		Set(ByVal Value As System.Drawing.Image)
			
			If m_Picture Is Nothing Then
				m_Picture = Value
				RaiseEvent PictureNormalChange()
				Exit Property
			End If
			
			m_PictureHot = Value
			RaiseEvent PictureHotChange()
			RedrawButton()
			
		End Set
	End Property
	
	
	Public Property PictureDown() As System.Drawing.Image
		Get
			
			PictureDown = m_PictureDown
			
		End Get
		Set(ByVal Value As System.Drawing.Image)
			
			If m_Picture Is Nothing Then
				m_Picture = Value
				RaiseEvent PictureNormalChange()
				Exit Property
			End If
			
			m_PictureDown = Value
			RaiseEvent PictureDownChange()
			RedrawButton()
			
		End Set
	End Property
	
	
	Public Property PictureAlign() As enumPictureAlign
		Get
			
			PictureAlign = m_PictureAlign
			
		End Get
		Set(ByVal Value As enumPictureAlign)
			
			m_PictureAlign = Value
			If Not m_Picture Is Nothing Then
				RedrawButton()
			End If
			RaiseEvent PictureAlignChange()
			
		End Set
	End Property
	
	
	Public Property PictureShadow() As Boolean
		Get
			
			PictureShadow = m_PictureShadow
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_PictureShadow = Value
			RedrawButton()
			RaiseEvent PictureShadowChange()
			
		End Set
	End Property
	
	
	Public Property PictureEffectOnOver() As enumPicEffect
		Get
			
			PictureEffectOnOver = m_PicEffectonOver
			
		End Get
		Set(ByVal Value As enumPicEffect)
			
			m_PicEffectonOver = Value
			RedrawButton()
			RaiseEvent PictureEffectOnOverChange()
			
		End Set
	End Property
	
	
	Public Property PictureEffectOnDown() As enumPicEffect
		Get
			
			PictureEffectOnDown = m_PicEffectonDown
			
		End Get
		Set(ByVal Value As enumPicEffect)
			
			m_PicEffectonDown = Value
			RedrawButton()
			RaiseEvent PictureEffectOnDownChange()
			
		End Set
	End Property
	
	
	Public Property PicturePushOnHover() As Boolean
		Get
			
			PicturePushOnHover = m_bPicPushOnHover
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bPicPushOnHover = Value
			RedrawButton()
			RaiseEvent PicturePushOnHoverChange()
			
		End Set
	End Property
	
	
	Public Property PictureOpacity() As Byte
		Get
			
			PictureOpacity = m_PictureOpacity
			
		End Get
		Set(ByVal Value As Byte)
			
			m_PictureOpacity = Value
			RedrawButton()
			RaiseEvent PictureOpacityChange()
			
		End Set
	End Property
	
	
	Public Property PictureOpacityOnOver() As Byte
		Get
			
			PictureOpacityOnOver = m_PicOpacityOnOver
			
		End Get
		Set(ByVal Value As Byte)
			
			m_PicOpacityOnOver = Value
			RedrawButton()
			RaiseEvent PictureOpacityOnOverChange()
			
		End Set
	End Property
	
	
	Public Shadows Property RightToLeft() As Boolean
		Get
			
			RightToLeft = m_bRTL
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bttRTL = Value
			m_bRTL = Value
			RedrawButton()
			RaiseEvent RightToLeftChange()
			
		End Set
	End Property
	
	
	Public Property CaptionEffects() As enumCaptionEffects
		Get
			
			CaptionEffects = m_CaptionEffects
			
		End Get
		Set(ByVal Value As enumCaptionEffects)
			
			m_CaptionEffects = Value
			RedrawButton()
			RaiseEvent CaptionEffectsChange()
			
		End Set
	End Property
	
	
	Public Property ShowFocusRect() As Boolean
		Get
			
			ShowFocusRect = m_bShowFocus
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bShowFocus = Value
			RaiseEvent ShowFocusRectChange()
			
		End Set
	End Property
	
	
	Public Property UseMaskColor() As Boolean
		Get
			
			UseMaskColor = m_bUseMaskColor
			
		End Get
		Set(ByVal Value As Boolean)
			
			m_bUseMaskColor = Value
			If Not m_Picture Is Nothing Then
				RedrawButton()
			End If
			RaiseEvent UseMaskColorChange()
			
		End Set
	End Property
	
	
	Public Property value() As Boolean
		Get
			
			value = m_bValue
			
		End Get
		Set(ByVal Value As Boolean)
			
			If m_ButtonMode <> enumButtonModes.ebmCommandButton Then
				m_bValue = Value
				If Not m_bValue Then
					m_Buttonstate = enumButtonStates.eStateNormal
				End If
				RedrawButton()
				RaiseEvent ValueChange()
			Else
				m_Buttonstate = enumButtonStates.eStateNormal
				RedrawButton()
			End If
			
		End Set
	End Property
	
	
	Public Property TooltipTitle() As String
		Get
			
			TooltipTitle = m_sTooltiptitle
			
		End Get
		Set(ByVal Value As String)
			
			m_sTooltiptitle = Value
			RedrawButton()
			RaiseEvent TooltipTitleChange()
			
		End Set
	End Property
	
	
	Public Property ToolTip() As String
		Get
			
			ToolTip = m_sTooltipText
			
		End Get
		Set(ByVal Value As String)
			
			m_sTooltipText = Value
			RedrawButton()
			RaiseEvent ToolTipChange()
			
		End Set
	End Property
	
	
	Public Property TooltipBackColor() As System.Drawing.Color
		Get
			
			TooltipBackColor = System.Drawing.ColorTranslator.FromOle(m_lttBackColor)
			
		End Get
		Set(ByVal Value As System.Drawing.Color)
			
			m_lttBackColor = System.Drawing.ColorTranslator.ToOle(Value)
			RedrawButton()
			RaiseEvent TooltipBackcolorChange()
			
		End Set
	End Property
	
	
	Public Property ToolTipIcon() As enumIconType
		Get
			
			ToolTipIcon = m_lToolTipIcon
			
		End Get
		Set(ByVal Value As enumIconType)
			
			m_lToolTipIcon = Value
			RedrawButton()
			RaiseEvent TooltipIconChange()
			
		End Set
	End Property
	
	
	Public Property ToolTipType() As enumTooltipStyle
		Get
			
			ToolTipType = m_lTooltipType
			
		End Get
		Set(ByVal Value As enumTooltipStyle)
			
			m_lTooltipType = Value
			RedrawButton()
			RaiseEvent ToolTipTypeChange()
			
		End Set
	End Property
	
	
	Public Property ColorScheme() As enumXPThemeColors
		Get
			
			ColorScheme = m_lXPColor
			
		End Get
		Set(ByVal Value As enumXPThemeColors)
			
			m_lXPColor = Value
			SetThemeColors()
			RedrawButton()
			RaiseEvent ColorSchemeChange()
			
		End Set
	End Property
	
	'Determine if the passed function is supported
	Private Function IsFunctionSupported(ByVal sFunction As String, ByVal sModule As String) As Boolean
		
		Dim lngModule As Integer
		
		lngModule = GetModuleHandle(sModule)
		
		If lngModule = 0 Then lngModule = LoadLibraryA(sModule)
		
		If lngModule Then
			IsFunctionSupported = GetProcAddress(lngModule, sFunction)
			FreeLibrary(lngModule)
		End If
		
	End Function
	
	
	'Track the mouse leaving the indicated window
	Private Sub TrackMouseLeave(ByVal lng_hWnd As Integer)
		
		Dim tme As TRACKMOUSEEVENT_STRUCT
		
		If TrackUser32 Then
			With tme
				.cbSize = Len(tme)
				.dwFlags = TRACKMOUSEEVENT_FLAGS.TME_LEAVE
				.hwndTrack = lng_hWnd
			End With
			
			If TrackUser32 Then
				TrackMouseEvent(tme)
			Else
				TrackMouseEventComCtl(tme)
			End If
		End If
		
	End Sub
	
	'=========================================================================
	'PUBLIC ROUTINES including subclassing & public button properties
	
	' CREDITS: Paul Caton
	'======================================================================================================
	'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
	
	Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Integer, ByRef lhWnd As Integer, ByRef uMsg As Integer, ByRef wParam As Integer, ByRef lParam As Integer)
		
		'Parameters:
		'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
		'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
		'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
		'hWnd     - The window handle
		'uMsg     - The message number
		'wParam   - Message related data
		'lParam   - Message related data
		'Notes:
		'If you really know what you're doing, it's possible to change the values of the
		'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
		'values get passed to the default handler.. and optionaly, the 'after' callback
		
		Static bMoving As Boolean
		
		Select Case uMsg
			
			Case WM_MOUSELEAVE
				
				m_bMouseInCtl = False
				If m_bPopupEnabled Then
					If m_bPopupInit Then
						m_bPopupInit = False
						m_bPopupShown = True
						Exit Sub
					Else
						m_bPopupShown = False
					End If
				End If
				
				If m_bIsSpaceBarDown Then Exit Sub
				If m_Buttonstate <> enumButtonStates.eStateNormal Then
					m_Buttonstate = enumButtonStates.eStateNormal
					RedrawButton()
				End If
				RaiseEvent MouseLeave(Me, Nothing)
				
			Case WM_NCACTIVATE, WM_ACTIVATE
				If wParam Then
					m_bParentActive = True
					If m_Buttonstate <> enumButtonStates.eStateNormal Then m_Buttonstate = enumButtonStates.eStateNormal
					If m_bDefault Then
						RedrawButton()
					End If
					RedrawButton()
				Else
					m_bIsDown = False
					m_bIsSpaceBarDown = False
					m_bHasFocus = False
					m_bParentActive = False
					If m_Buttonstate <> enumButtonStates.eStateNormal Then m_Buttonstate = enumButtonStates.eStateNormal
					RedrawButton()
				End If
				
			Case WM_THEMECHANGED
				RedrawButton()
				
			Case WM_SYSCOLORCHANGE
				RedrawButton()
		End Select
		
	End Sub
	
	Public Sub SetPopupMenu(ByRef Menu As Object, Optional ByRef Align As enumMenuAlign = 0, Optional ByRef flags As Object = 0, Optional ByRef DefaultMenu As Object = Nothing)
		
		If Not (Menu Is Nothing) Then
			If (TypeOf Menu Is System.Windows.Forms.ToolStripMenuItem) Then
				
				DropDownMenu = Menu
				MenuAlign = Align
				'UPGRADE_WARNING: Couldn't resolve default property of object flags. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				MenuFlags = flags
				DefaultMenu = DefaultMenu
				m_bPopupEnabled = True
			End If
		End If
		
	End Sub
	
	Public Sub UnsetPopupMenu()
		
		' --Free the popup menu
		'UPGRADE_NOTE: Object DropDownMenu may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		DropDownMenu = Nothing
		'UPGRADE_NOTE: Object DefaultMenu may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		DefaultMenu = Nothing
		m_bPopupEnabled = False
		m_bPopupShown = False
		
	End Sub
	
	Public Sub OpenWebsite(ByVal sAddress As String)
		
		ShellExecute(hwnd, "open", sAddress, vbNullString, vbNullString, 1)
		
	End Sub
	
	Public Sub About()
		
		MsgBox("JCButton v 1.02" & vbNewLine & "Author: Juned S. Chhipa" & vbNewLine & "Contact: juned.chhipa@yahoo.com" & vbNewLine & vbNewLine & "Copyright © 2008-2009 Juned Chhipa. All rights reserved.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, "About")
		
	End Sub
	
	'UPGRADE_ISSUE: Font event mFont.FontChanged was not upgraded. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub mFont_FontChanged(ByVal PropertyName As String)
		
		MyBase.Font = mFont
		Refresh()
		RedrawButton()
		RaiseEvent FontChange()
		
	End Sub
	
	'======================================================================================================
	'Subclass code - The programmer may call any of the following Subclass_??? routines
	
	'Stop subclassing the passed window handle
	
	Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Integer
		
		Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
		System.Diagnostics.Debug.Assert(Subclass_AddrFunc, "")
		
	End Function
	
	Private Function Subclass_Index(ByVal lhWnd As Integer, Optional ByVal bAdd As Boolean = False) As Integer
		
		For Subclass_Index = UBound(SubclassData) To 0 Step -1
			If SubclassData(Subclass_Index).hwnd = lhWnd Then
				If Not bAdd Then Exit Function
				
			ElseIf SubclassData(Subclass_Index).hwnd = 0 Then 
				If bAdd Then Exit Function
			End If
		Next  'Subclass_Index
		
		If Not bAdd Then System.Diagnostics.Debug.Assert(False, "")
		
	End Function
	
	Private Function Subclass_InIDE() As Boolean
		
		System.Diagnostics.Debug.Assert(Subclass_SetTrue(Subclass_InIDE), "")
		
	End Function
	
	Private Function Subclass_Initialize(ByVal lhWnd As Integer) As Integer
		
		Const CODE_LEN As Integer = 200
		Const GMEM_FIXED As Integer = 0
		Const PATCH_01 As Integer = 18
		Const PATCH_02 As Integer = 68
		Const PATCH_03 As Integer = 78
		Const PATCH_06 As Integer = 116
		Const PATCH_07 As Integer = 121
		Const PATCH_0A As Integer = 186
		Const FUNC_CWP As String = "CallWindowProcA"
		Const FUNC_EBM As String = "EbMode"
		Const FUNC_SWL As String = "SetWindowLongA"
		Const MOD_USER As String = "User32"
		Const MOD_VBA5 As String = "vba5"
		Const MOD_VBA6 As String = "vba6"
		
		'UPGRADE_WARNING: Lower bound of array bytBuffer was changed from 1 to 0. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Static bytBuffer(CODE_LEN) As Byte
		Static lngCWP As Integer
		Static lngEbMode As Integer
		Static lngSWL As Integer
		
		Dim lngCount As Integer
		Dim lngIndex As Integer
		Dim strHex As String
		
		If bytBuffer(1) Then
			lngIndex = Subclass_Index(lhWnd, True)
			
			If lngIndex = -1 Then
				lngIndex = UBound(SubclassData) + 1
				
				ReDim Preserve SubclassData(lngIndex)
			End If
			
			Subclass_Initialize = lngIndex
			
		Else
			strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
			
			For lngCount = 1 To CODE_LEN
				bytBuffer(lngCount) = Val("&H" & VB.Left(strHex, 2))
				strHex = Mid(strHex, 3)
			Next  'lngCount
			
			If Subclass_InIDE Then
				bytBuffer(16) = &H90s
				bytBuffer(17) = &H90s
				lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
				
				If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
			End If
			
			lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
			lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
			
			ReDim SubclassData(0)
		End If
		
		With SubclassData(lngIndex)
			.hwnd = lhWnd
			.nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
			.nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSclass)
			
			Call CopyMemory(.nAddrSclass, bytBuffer(1), CODE_LEN)
			Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
			Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
			Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
			Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
			Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
			'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
		End With
		
	End Function
	
	Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean
		
		Subclass_SetTrue = True
		bValue = True
		
	End Function
	
	'UPGRADE_NOTE: When was upgraded to When_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Subclass_AddMsg(ByVal lhWnd As Integer, ByVal uMsg As Integer, Optional ByVal When_Renamed As MsgWhen = MsgWhen.MSG_AFTER)
		
		With SubclassData(Subclass_Index(lhWnd))
			If When_Renamed And MsgWhen.MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MsgWhen.MSG_BEFORE, .nAddrSclass)
			If When_Renamed And MsgWhen.MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MsgWhen.MSG_AFTER, .nAddrSclass)
		End With
		
	End Sub
	
	'UPGRADE_NOTE: When was upgraded to When_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Subclass_DelMsg(ByVal lhWnd As Integer, ByVal uMsg As Integer, Optional ByVal When_Renamed As MsgWhen = MsgWhen.MSG_AFTER)
		
		With SubclassData(Subclass_Index(lhWnd))
			If When_Renamed And MsgWhen.MSG_BEFORE Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelB, .nMsgCountB, MsgWhen.MSG_BEFORE, .nAddrSclass)
			If When_Renamed And MsgWhen.MSG_AFTER Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelA, .nMsgCountA, MsgWhen.MSG_AFTER, .nAddrSclass)
		End With
		
	End Sub
	
	'UPGRADE_NOTE: When was upgraded to When_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Subclass_DoAddMsg(ByVal uMsg As Integer, ByRef aMsgTabel() As Integer, ByRef nMsgCount As Integer, ByVal When_Renamed As MsgWhen, ByVal nAddr As Integer)
		
		Const PATCH_04 As Integer = 88
		Const PATCH_08 As Integer = 132
		
		Dim lngEntry As Integer
		
		Dim lngOffset(1) As Integer
		
		If uMsg = ALL_MESSAGES Then
			nMsgCount = ALL_MESSAGES
			
		Else
			For lngEntry = 1 To nMsgCount - 1
				If aMsgTabel(lngEntry) = 0 Then
					aMsgTabel(lngEntry) = uMsg
					
					GoTo ExitSub
					
				ElseIf aMsgTabel(lngEntry) = uMsg Then 
					GoTo ExitSub
				End If
			Next  'lngEntry
			
			nMsgCount = nMsgCount + 1
			
			'UPGRADE_WARNING: Lower bound of array aMsgTabel was changed from 1 to 0. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim Preserve aMsgTabel(nMsgCount)
			
			aMsgTabel(nMsgCount) = uMsg
		End If
		
		If When_Renamed = MsgWhen.MSG_BEFORE Then
			lngOffset(0) = PATCH_04
			lngOffset(1) = PATCH_05
			
		Else
			lngOffset(0) = PATCH_08
			lngOffset(1) = PATCH_09
		End If
		
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
		
		Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
		
ExitSub: 
		'UPGRADE_NOTE: Erase was upgraded to System.Array.Clear. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		System.Array.Clear(lngOffset, 0, lngOffset.Length)
		
	End Sub
	
	'UPGRADE_NOTE: When was upgraded to When_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Subclass_DoDelMsg(ByVal uMsg As Integer, ByRef aMsgTabel() As Integer, ByRef nMsgCount As Integer, ByVal When_Renamed As MsgWhen, ByVal nAddr As Integer)
		
		Dim lngEntry As Integer
		
		If uMsg = ALL_MESSAGES Then
			nMsgCount = 0
			
			If When_Renamed = MsgWhen.MSG_BEFORE Then
				lngEntry = PATCH_05
				
			Else
				lngEntry = PATCH_09
			End If
			
			Call Subclass_PatchVal(nAddr, lngEntry, 0)
			
		Else
			For lngEntry = 1 To nMsgCount - 1
				If aMsgTabel(lngEntry) = uMsg Then
					aMsgTabel(lngEntry) = 0
					Exit For
				End If
			Next  'lngEntry
		End If
		
	End Sub
	
	Private Sub Subclass_PatchRel(ByVal nAddr As Integer, ByVal nOffset As Integer, ByVal nTargetAddr As Integer)
		
		Call CopyMemory(nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
		
	End Sub
	
	Private Sub Subclass_PatchVal(ByVal nAddr As Integer, ByVal nOffset As Integer, ByVal nValue As Integer)
		
		Call CopyMemory(nAddr + nOffset, nValue, 4)
		
	End Sub
	
	Private Sub Subclass_Stop(ByVal lhWnd As Integer)
		
		With SubclassData(Subclass_Index(lhWnd))
			SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
			
			Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
			Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
			
			GlobalFree(.nAddrSclass)
			.hwnd = 0
			.nMsgCountA = 0
			.nMsgCountB = 0
			Erase .aMsgTabelA
			Erase .aMsgTabelB
		End With
		
	End Sub
	
	Private Sub Subclass_Terminate()
		
		Dim lngCount As Integer
		
		For lngCount = UBound(SubclassData) To 0 Step -1
			If SubclassData(lngCount).hwnd Then Call Subclass_Stop(SubclassData(lngCount).hwnd)
		Next  'lngCount
		
	End Sub
	
	'---------------x---------------x--------------x--------------x-----------x---
	' Oops! Control resulted Longer than expected!
	' Lots of hours and lots of tedious work!   This is my first submission on PSC
	' So if you want to vote for this, just do it ;)
	' Comments are greatly appreciated...
	' Enjoy!
	'---------------x---------------x--------------x--------------x-----------x---
End Class