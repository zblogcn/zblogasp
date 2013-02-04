Option Strict Off
Option Explicit On
Module mdlBase
	Public objFSO As Scripting.FileSystemObject
	Public objRegExp As VBScript_RegExp_55.RegExp
	Public objADO As Object
	Public objXML As New MSXML2.DOMDocument
	Public bolAero As Boolean
	
	Public Structure OSVERSIONINFO
		Dim dwOSVersionInfoSize As Integer
		Dim dwMajorVersion As Integer
		Dim dwMinorVersion As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformID As Integer
		'UPGRADE_WARNING: �̶������ַ����Ĵ�С�����ʺϻ������� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"��
		<VBFixedString(128),System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray,SizeConst:=128)> Public szCSDVersion() As Char
	End Structure
	
	
	Public Structure BROWSEINFO
		Dim hOwner As Integer
		Dim pidlRoot As Integer
		Dim pszDisplayName As String
		Dim lpszTitle As String
		Dim ulFlags As Integer
		Dim lpfn As Integer
		Dim lParam As Integer
		Dim iImage As Integer
	End Structure
	
	
	'����ļ���·��
	Public Declare Function SHGetPathFromIDList Lib "shell32.dll"  Alias "SHGetPathFromIDListA"(ByVal pidl As Integer, ByVal pszPath As String) As Integer
	'��ʾ�ļ����б��
	'UPGRADE_WARNING: �ṹ BROWSEINFO ����Ҫ����ʹ���������Ϊ�� Declare ����еĲ������ݡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"��
	Public Declare Function SHBrowseForFolder Lib "shell32.dll"  Alias "SHBrowseForFolderA"(ByRef lpBrowseInfo As BROWSEINFO) As Integer
	
	'�õ�����ϵͳ�汾
	'UPGRADE_WARNING: �ṹ OSVERSIONINFO ����Ҫ����ʹ���������Ϊ�� Declare ����еĲ������ݡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"��
	Public Declare Function GetVersionEx Lib "kernel32"  Alias "GetVersionExA"(ByRef lpVersionInformation As OSVERSIONINFO) As Integer
	
	Const BIF_RETURNONLYFSDIRS As Short = &H1s
	Const adOpenForwardOnly As Short = 0
	Const adOpenKeyset As Short = 1
	Const adOpenDynamic As Short = 2
	Const adOpenStatic As Short = 3
	
	Const adLockReadOnly As Short = 1
	Const adLockPessimistic As Short = 2
	Const adLockOptimistic As Short = 3
	Const adLockBatchOptimistic As Short = 4
	
	Const ForReading As Short = 1
	Const ForWriting As Short = 2
	Const ForAppending As Short = 8
	
	Const adTypeBinary As Short = 1
	Const adTypeText As Short = 2
	
	Const adModeRead As Short = 1
	Const adModeReadWrite As Short = 3
	
	Const adSaveCreateNotExist As Short = 1
	Const adSaveCreateOverWrite As Short = 2
	
	
	
	
	'Usage:��һ���ļ���ѡȡ����
	'Param:strMsg--���⣬hWnd--Form���
	Function GetFolderPath(ByVal strMsg As String, ByRef hWnd As Integer) As Object
		
		Dim broInfo As BROWSEINFO
		Dim lngGet As Integer
		Dim lngPID As Integer
		Dim strPath As String
		broInfo.hOwner = hWnd
		broInfo.pidlRoot = 0
		broInfo.lpszTitle = strMsg
		broInfo.ulFlags = &H1s 'BIF_RETURNONLYFSDIRS
		lngPID = SHBrowseForFolder(broInfo)
		strPath = Space(512)
		lngGet = SHGetPathFromIDList(lngPID, strPath)
		If lngGet Then
			'API��ȡ������Space���ȽϿӵ�
			GetFolderPath = Left(strPath, InStr(strPath, Chr(0)) - 1)
		Else
			'UPGRADE_WARNING: δ�ܽ������� GetFolderPath ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			GetFolderPath = False
		End If
		
	End Function
	
	'Usage:�õ�ϵͳ�汾
	Function GetSystemVersion() As Object
		Dim objOS As OSVERSIONINFO
		objOS.dwOSVersionInfoSize = 148
		objOS.szCSDVersion = Space(128)
		Call GetVersionEx(objOS)
		'UPGRADE_WARNING: δ�ܽ������� GetSystemVersion ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		GetSystemVersion = CDbl(objOS.dwMajorVersion & "." & objOS.dwMinorVersion)
		'UPGRADE_WARNING: δ�ܽ������� GetSystemVersion ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
		If GetSystemVersion >= 6 Then
			bolAero = True
		End If
	End Function
	
	Function LoadFromFile(ByVal strPath As String, Optional ByRef strCharset As String = "UTF-8") As Object
		
		With objADO
			'UPGRADE_WARNING: δ�ܽ������� objADO.Type ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Type = adTypeText
			'UPGRADE_WARNING: δ�ܽ������� objADO.Mode ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Mode = adModeReadWrite
			'UPGRADE_WARNING: δ�ܽ������� objADO.Open ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Open()
			'UPGRADE_WARNING: δ�ܽ������� objADO.Charset ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Charset = strCharset
			'UPGRADE_WARNING: δ�ܽ������� objADO.Position ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			'UPGRADE_WARNING: δ�ܽ������� objADO.Size ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Position = .Size
			'UPGRADE_WARNING: δ�ܽ������� objADO.LoadFromFile ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.LoadFromFile(strPath)
			'UPGRADE_WARNING: δ�ܽ������� objADO.ReadText ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			'UPGRADE_WARNING: δ�ܽ������� LoadFromFile ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			LoadFromFile = .ReadText
			'UPGRADE_WARNING: δ�ܽ������� objADO.Close ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Close()
		End With
		
		Err.Clear()
		
	End Function
	
	Function SaveToFile(ByRef strFullName As String, ByRef strContent As String, Optional ByRef strCharset As String = "UTF-8") As Object
		
		
		With objADO
			'UPGRADE_WARNING: δ�ܽ������� objADO.Type ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Type = adTypeText
			'UPGRADE_WARNING: δ�ܽ������� objADO.Mode ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Mode = adModeReadWrite
			'UPGRADE_WARNING: δ�ܽ������� objADO.Open ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Open()
			'UPGRADE_WARNING: δ�ܽ������� objADO.Charset ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Charset = strCharset
			'UPGRADE_WARNING: δ�ܽ������� objADO.Position ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			'UPGRADE_WARNING: δ�ܽ������� objADO.Size ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Position = .Size
			'UPGRADE_WARNING: δ�ܽ������� objADO.WriteText ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.WriteText = strContent
			'UPGRADE_WARNING: δ�ܽ������� objADO.SaveToFile ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.SaveToFile(strFullName, adSaveCreateOverWrite)
			'UPGRADE_WARNING: δ�ܽ������� objADO.Close ��Ĭ�����ԡ� �����Ի�ø�����Ϣ:��ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"��
			.Close()
		End With
		
		
	End Function
End Module