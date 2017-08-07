Attribute VB_Name = "FileDialogue"
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public SelInd As Integer
Public CurrentFile As String

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Function GetDialog(ByVal sMethod As String, ByVal sTitle As String, ByVal sFileName As String, sFilter As String) As String
    On Error GoTo myError
    Dim rtn As Long, pos As Integer
    Dim File As OPENFILENAME
    File.lStructSize = Len(File)
    File.hInstance = App.hInstance
    File.lpstrFile = sFileName & String$(255 - Len(sFileName), 0)
    File.nMaxFile = 255
    File.lpstrFileTitle = String$(255, 0)
    File.nMaxFileTitle = 255
    File.lpstrInitialDir = ""
    File.lpstrFilter = sFilter
    File.lpstrTitle = sTitle
    If UCase(sMethod) = "OPEN" Then
        File.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST
        rtn = GetOpenFileName(File)
    Else
        File.lpstrDefExt = "exe"
        File.flags = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_OVERWRITEPROMPT
        rtn = GetSaveFileName(File)
    End If
    If rtn > 0 Then
        pos = InStr(File.lpstrFile, Chr$(0))
        If pos > 0 Then
            GetDialog = Left$(File.lpstrFile, pos - 1)
        End If
    End If
    Exit Function
myError:
    MsgBox "未知原因导致操作失败！", vbCritical + vbOKOnly, APP_NAME
End Function

