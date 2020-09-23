Attribute VB_Name = "modFileAccess"
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const OFS_MAXPATHNAME = 128

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Function GetFileQuick(strFilePath As String, Optional bolAsString = True)

  Dim arrFileMain() As Byte
  Dim lngSize As Long, lngRet As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT

    'Open the two files
    lngFileHandle = OpenFile(strFilePath, ofData, OF_READ)

    'Get the file size
    lngSize = GetFileSize(lngFileHandle, 0)

    'Create an array of bytes
    ReDim arrFileMain(lngSize) As Byte

    'Read from the file
    ReadFile lngFileHandle, arrFileMain(0), UBound(arrFileMain), lngRet, ByVal 0&

    'Close the file
    CloseHandle lngFileHandle

    ReDim Preserve arrFileMain(UBound(arrFileMain) - 1)

    If bolAsString Then 'Return as string
        GetFileQuick = StrConv(arrFileMain(), vbUnicode)
      Else 'Return as Byte Array
        GetFileQuick = arrFileMain()
    End If

End Function

'Rounds a Byte amount and returns KB with 2 decimal places
Public Function GetRoundedKB(lngNumber As Long)

    GetRoundedKB = Int((lngNumber / 1024) * 100 + 0.5) / 100

End Function

'Rounds a Byte amount and returns MB with 2 decimal places
Public Function GetRoundedMB(lngNumber As Long)

    GetRoundedMB = Int((lngNumber / 1048576) * 100 + 0.5) / 100

End Function

':) Ulli's VB Code Formatter V2.10.8 (15.08.2002 14:36:54) 26 + 160 = 186 Lines
