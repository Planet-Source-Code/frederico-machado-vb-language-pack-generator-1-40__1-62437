Attribute VB_Name = "mUrlSource"
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Const IF_FROM_CACHE         As Long = &H1000000
Private Const IF_MAKE_PERSISTENT    As Long = &H2000000
Private Const IF_NO_CACHE_WRITE     As Long = &H4000000
       
Private Const BUFFER_LEN            As Long = 25

Public Function DownloadUrlSource(ByVal URLString As String) As String
On Local Error GoTo errHandler

Dim lSession    As Long
Dim lInternet   As Long
    ' Get the handle of the currently connected internet connection.
    lSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    ' Get the handle of the selected URL.
    If lSession Then lInternet = InternetOpenUrl(lSession, URLString, vbNullString, 0, IF_NO_CACHE_WRITE, 0)

Dim sBuffer     As String * BUFFER_LEN
Dim sData       As String
Dim iResult     As Integer
Dim lReturn     As Long
    ' If we have the handle, then start reading the web page.
    If lInternet Then
        ' Get the first chunk of data.
        iResult = InternetReadFile(lInternet, sBuffer, BUFFER_LEN, lReturn)
        ' Store the Buffer to the Returning Data String.
        sData = sBuffer
        ' If there's more data then keep reading into the buffer until finished.
        Do While lReturn <> 0
            ' Get the Next chunk of data.
            iResult = InternetReadFile(lInternet, sBuffer, BUFFER_LEN, lReturn)
            ' Add the Next set of buffered data to the Returning Data String.
            sData = sData & Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    ' Close the URL.
    iResult = InternetCloseHandle(lInternet)
    ' Return the Resulting Data.
    DownloadUrlSource = sData
Exit Function
errHandler:
    ' Return the Resulting Data.
    DownloadUrlSource = vbNullString
End Function
