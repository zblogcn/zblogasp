Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices

Friend Class clsAero

    Public Form As Form
    <StructLayout(LayoutKind.Sequential)> Public Structure MARGINS
        Public cxLeftWidth As Integer
        Public cxRightWidth As Integer
        Public cyTopHeight As Integer
        Public cyButtomheight As Integer
    End Structure


    <DllImport("dwmapi.dll")> Public Shared Function DwmExtendFrameIntoClientArea(ByVal hWnd As IntPtr, ByRef pMarinset As MARGINS) As Integer
    End Function


    Public Sub Go()
        Form.TransparencyKey = Color.FromArgb(255, 255, 1)
        Form.BackColor = Form.TransparencyKey

        Dim margins As MARGINS = New MARGINS
        margins.cxLeftWidth = -1
        margins.cxRightWidth = -1
        margins.cyTopHeight = -1
        margins.cyButtomheight = -1


        DwmExtendFrameIntoClientArea(Form.Handle, margins)
    End Sub
End Class