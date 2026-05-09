VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RDWorks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#If VBA7 Then
Private Declare PtrSafe Function ShellExecuteA Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare PtrSafe Function GetTempPathA Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As LongPtr) As Long
Private Declare PtrSafe Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Dim Key As LongPtr
#Else
Private Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Dim Key As Long
#End If
Const HKEY_LOCAL_MACHINE = &H80000002
Const KEY_QUERY_VALUE = 1
Const SW_SHOWMAXIMIZED = 3
Const KEY_WOW64_64KEY = 256
Sub Start()
  Dim ProgPath As String
  Dim fileName As String
  Dim ExePath As String * 255
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrBottomLeft
  If ActiveDocument.Selection.Shapes.Count > 0 Then
    Dim expopt As New StructExportOptions
    fileName = Environ("TEMP") + "\RD.ai"
    expopt.UseColorProfile = False
    With ActiveDocument.ExportEx(fileName, cdrAI, cdrSelection, expopt)
      .Version = 2
      .TextAsCurves = True
      .ConvertSpotColors = False
      .UseColorProfile = False
      .SimulateOutlines = False
      .SimulateFills = False
      .IncludePlacedImages = True
      .IncludePreview = False
      .Finish
    End With
    RegOpenKeyExA HKEY_LOCAL_MACHINE, "SOFTWARE\RDSOFT", 0, KEY_QUERY_VALUE + KEY_WOW64_64KEY, Key
    ExePathLen = 255
    RegQueryValueExA Key, "ExePath", 0, 0, ExePath, ExePathLen
    RegCloseKey Key
    ShellExecuteA 0, "open", ExePath, fileName, 0, SW_SHOWMAXIMIZED
  End If
End Sub
