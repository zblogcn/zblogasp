Option Strict Off
Option Explicit On
Friend Class clsAero
	
	Public hWnd As Integer
	Public hDc As Integer
	Public m_transparencyKey As Integer
	
	
	Private LWA_COLORKEY As Integer
	Private GWL_EXSTYLE As Integer
	Private WS_EX_LAYERED As Integer
	
	
	'UPGRADE_WARNING: 结构 MARGINS 可能要求封送处理属性作为此 Declare 语句中的参数传递。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"”
	Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Integer, ByRef margin As MARGINS) As Integer
	Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef enabledptr As Integer) As Integer
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: 结构 RECT 可能要求封送处理属性作为此 Declare 语句中的参数传递。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"”
	Private Declare Function FillRect Lib "user32" (ByVal hDc As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	'UPGRADE_WARNING: 结构 RECT 可能要求封送处理属性作为此 Declare 语句中的参数传递。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"”
	Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Function SetLayeredWindowAttributesByColor Lib "user32"  Alias "SetLayeredWindowAttributes"(ByVal hWnd As Integer, ByVal crey As Integer, ByVal bAlpha As Byte, ByVal dwFlags As Integer) As Integer
	Private Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	Private Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
	
	Private Structure MARGINS
		Dim m_Left As Integer
		Dim m_Right As Integer
		Dim m_Top As Integer
		Dim m_Button As Integer
	End Structure
	
	Private Structure RECT
		'UPGRADE_NOTE: Left 已升级到 Left_Renamed。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"”
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right 已升级到 Right_Renamed。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"”
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	
	
	'UPGRADE_NOTE: Class_Initialize 已升级到 Class_Initialize_Renamed。 单击以获得更多信息:“ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"”
	Sub Class_Initialize_Renamed()
		LWA_COLORKEY = &H1s
		GWL_EXSTYLE = -20
		WS_EX_LAYERED = &H80000
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub Init()
		m_transparencyKey = RGB(255, 255, 1)
		SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
		SetLayeredWindowAttributesByColor(hWnd, m_transparencyKey, 0, LWA_COLORKEY)
		Dim mg As MARGINS
		Dim en As Integer
		mg.m_Left = -1
		mg.m_Button = -1
		mg.m_Right = -1
		mg.m_Top = -1
		DwmIsCompositionEnabled(en)
		If en Then
			DwmExtendFrameIntoClientArea(hWnd, mg)
		End If
		Exit Sub
	End Sub
	
	Public Sub Paint()
		Dim hBrush, hBrushOld As Integer
		Dim m_Rect As RECT
		hBrush = CreateSolidBrush(m_transparencyKey)
		hBrushOld = SelectObject(hDc, hBrush)
		GetClientRect(hWnd, m_Rect)
		FillRect(hDc, m_Rect, hBrush)
		SelectObject(hDc, hBrushOld)
		DeleteObject(hBrush)
	End Sub
End Class