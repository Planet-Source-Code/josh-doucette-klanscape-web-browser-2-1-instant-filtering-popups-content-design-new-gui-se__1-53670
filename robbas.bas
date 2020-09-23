Attribute VB_Name = "robbas"
Private Const BUFFER_LEN = 256
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
Private Declare Function SetWindowPos& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal cX As Long, _
                ByVal cY As Long, _
ByVal wFlags As Long)
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1
Public Function StringBetween(TextString As String, String1 As String, String2 As String) As String
    Dim MyError As Boolean
    Dim Start As Integer
    Dim Length As Integer
    Dim PosString1, PosString2 As Integer
    MyError = MyError Or Len(TextString) = 0
    MyError = MyError Or Len(String1) = 0
    MyError = MyError Or Len(String2) = 0
    If Not MyError Then
        PosString1 = InStr(1, TextString, String1, vbTextCompare)
        If PosString1 > 0 Then
            PosString2 = InStr(PosString1, TextString, String2, vbTextCompare)
        Else: PosString2 = 0
        End If
        MyError = MyError Or PosString1 = 0 Or PosString2 = 0
    End If
    If Not MyError And PosString1 < PosString2 Then
        Start = PosString1 + Len(String1)
        Length = PosString2 - Start
        StringBetween = Mid(TextString, Start, Length)
    End If
    If MyError Then
        StringBetween = 0
    End If
End Function
Public Sub StayOnTop(Form As Form)
    SetWindowPos Form.hWnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub
Public Sub StayOnBottom(Form As Form)
    SetWindowPos Form.hWnd, -2, 0, 0, 0, 0, 1 Or 2
End Sub
Public Sub MakeNewKlanBrowser(Optional TargetURL As String = Empty, Optional bSilent As Boolean = False)
    If TargetURL = Empty Then TargetURL = "http://www.google.com"
        Dim frmB As New Form1
    Load frmB
    With frmB
        .WebBrowser1.Silent = bSilent
        .WebBrowser1.Navigate TargetURL
        .Show
    End With
    Set frmB = Nothing
End Sub
Public Function GetUrlSource(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData
End Function
Public Function Exists_In_String(text1 As String, Arg As String) As String
    Dim x As String
    Dim i As Integer
    x = LOS(Arg)
    For i = 1 To 255


        If LCase(Mid(text1, i, x)) = LCase(Arg) Then
            Exists_In_String = "True"
            Exit Function
        Else
            Exists_In_String = "False"
        End If
    Next i
End Function
Public Function LOS(String1 As String) As String
Dim i
    Dim x As Integer
    Dim y As String
    For i = 1 To 255
        y = Mid(String1, i, 1)
        If y = "" Then GoTo Fin
        x = x + 1
    Next i
Fin:
    LOS = x
End Function
