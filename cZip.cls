VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MyID As String * 3 = "mlh"

'/* Property Variables:
Private m_CompressedSize As Long
Private m_OriginalSize As Long
Private m_CompressLevel As Integer

'/* Declares
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'Private Declare Function Compress Lib "zlib.dll" Alias "compress" (Dest As Any, DestLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function Compress Lib "zlib.dll" Alias "compress2" (Dest As Any, DestLen As Long, Source As Any, ByVal SourceLen As Long, ByVal Level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (Dest As Any, DestLen As Any, src As Any, ByVal srcLen As Long) As Long

'/* Needed to decompress file
Private Type FileHeaderType
  ID As String * 3
  OrigSize As Long
  FileExt As String * 3
End Type


'/* Required for getting Disk Information
Private Type DISK_FREE_TOTAL
    dbFree As Double
    dbTotal As Double
End Type
Dim gudtDiskInfo As DISK_FREE_TOTAL

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEX Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, largIntFreeBytesAvailableToCaller As LARGE_INTEGER, largIntTotalNumberOfBytes As LARGE_INTEGER, largIntTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Function DriveType(ByVal DriveLetter As String) As Integer
  Dim retval As Long
    '/* Return Codes
    '/* 2=Removable
    '/* 3=Fixed
    '/* 4=Network
    '/* 5=CDROM
    '/* 6=RamDisk
    
    DriveLetter = left$(DriveLetter, 1) & ":"
    retval = GetDriveType(DriveLetter)
    If retval = 1 Then
        DriveType = False
    Else
        DriveType = retval
    End If
    
End Function

Public Function SpaceTotal(ByVal DrivePath As String) As Double
    If mGetDiskInfo(DrivePath) Then
        SpaceTotal = gudtDiskInfo.dbTotal
    Else
        SpaceTotal = -1
    End If
End Function

Public Function SpaceFree(ByVal DrivePath As String) As Double
    If mGetDiskInfo(DrivePath) Then
        SpaceFree = gudtDiskInfo.dbFree
    Else
        SpaceFree = -1
    End If
End Function

Private Function mGetDiskInfo(strDrv As String) As Boolean
 Dim largIntFreeBytesAvailableToCaller As LARGE_INTEGER
 Dim largIntTotalNumberOfBytes As LARGE_INTEGER
 Dim largIntTotalNumberOfFreeBytes As LARGE_INTEGER
 Dim lngSectorsPerCluster&, lngBytesPerSector&
 Dim lngNumberOfFreeClusters&, lngTotalNumberOfClusters&
 Dim lngRet&, dbFree#, dbTotal#
 
    Select Case Len(strDrv)
    Case Is < 1
        mGetDiskInfo = False
        Exit Function
    Case 1
        strDrv = strDrv & ":\"
    Case Is > 1
        If right$(strDrv, 1) <> "\" Then strDrv = strDrv & "\"
    End Select
    
    '/* Assume NT or OSR2
    On Error Resume Next
    lngRet = GetDiskFreeSpaceEX(strDrv, _
            largIntFreeBytesAvailableToCaller, _
            largIntTotalNumberOfBytes, _
            largIntTotalNumberOfFreeBytes)
    
    If Err.Number = 0 Then
        dbFree = mUnsigned2Dbl(largIntFreeBytesAvailableToCaller.highpart) * 2# ^ 32
        dbFree = dbFree + mUnsigned2Dbl(largIntFreeBytesAvailableToCaller.lowpart)
        
        dbTotal = mUnsigned2Dbl(largIntTotalNumberOfBytes.highpart) * 2# ^ 32
        dbTotal = dbTotal + mUnsigned2Dbl(largIntTotalNumberOfBytes.lowpart)
    Else
        Dim dbTemp#
        lngRet = GetDiskFreeSpace(strDrv, _
                lngSectorsPerCluster, _
                lngBytesPerSector, _
                lngNumberOfFreeClusters, _
                lngTotalNumberOfClusters)
        dbTemp = mUnsigned2Dbl(lngBytesPerSector) * mUnsigned2Dbl(lngSectorsPerCluster)
        dbFree = mUnsigned2Dbl(lngNumberOfFreeClusters) * dbTemp
        dbTotal = mUnsigned2Dbl(lngTotalNumberOfClusters) * dbTemp
    End If
    
    gudtDiskInfo.dbFree = dbFree
    gudtDiskInfo.dbTotal = dbTotal
    mGetDiskInfo = True

End Function
Private Function mUnsigned2Dbl(ruNumber As Long) As Double
  Const MAXINT As Long = 2147483647
  
    On Error Resume Next
    If ruNumber < 0 Then
        mUnsigned2Dbl = CDbl(ruNumber And MAXINT) + 2# ^ 31
    Else
        mUnsigned2Dbl = CDbl(ruNumber)
    End If
End Function

Private Function CompressData(TheData() As Byte) As Long
  Dim Result As Long
  Dim BufferSize As Long
  Dim TempBuffer() As Byte

    m_OriginalSize = UBound(TheData) + 1
    
    '/* Allocate memory for byte array
    BufferSize = UBound(TheData) + 1
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    
    '/* Compress byte array (data)
    Result = Compress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1, m_CompressLevel)
    
    '/* Truncate to compressed size
    ReDim Preserve TheData(BufferSize - 1)
    CopyMemory TheData(0), TempBuffer(0), BufferSize
    
    '/* Cleanup
    Erase TempBuffer
    
    '/* Set properties if no error occurred
    If Result = 0 Then m_CompressedSize = UBound(TheData) + 1
    
    '/* Return error code (if any)
    CompressData = Result

End Function

Private Function DecompressData(TheData() As Byte, OrigSize As Long) As Long
  Dim Result As Long
  Dim BufferSize As Long
  Dim TempBuffer() As Byte
    
    '/* Allocate memory for buffers
    BufferSize = OrigSize
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    
    '/* Decompress data
    Result = uncompress(TempBuffer(0), BufferSize, TheData(0), UBound(TheData) + 1)
    
    '/* Truncate buffer to compressed size
    ReDim Preserve TheData(BufferSize - 1)
    CopyMemory TheData(0), TempBuffer(0), BufferSize
    
    '/* Reset properties
    If Result = 0 Then
        m_CompressedSize = 0
        m_OriginalSize = 0
    End If
    
    '/* Return error code (if any)
    DecompressData = Result

End Function

Public Property Get CompressedSize() As Long
    CompressedSize = m_CompressedSize
End Property

Public Property Get OriginalSize() As Long
    OriginalSize = m_OriginalSize
End Property

Private Sub Class_Initialize()
    m_CompressedSize = 0
    m_OriginalSize = 0
    '/* 1=more compression, 9=less compression, 0=no compression
    m_CompressLevel = 6
End Sub



Public Sub CompressFile(ByVal SourceFile As String, ByVal DestinationFile As String)
  Dim Buffer() As Byte, Result As Long
  Dim FF As Integer, FH As FileHeaderType
    
    '/* Fix Buffer size
    ReDim Buffer(FileLen(SourceFile) - 1) As Byte
    
    '/* Read byte array from file
    FF = FreeFile
    Open SourceFile For Binary Access Read As #FF
        Get #FF, , Buffer()
    Close #FF
    
    '/* Compress file
    Result = CompressData(Buffer())
    If Result <> 0 Then
        Err.Raise Result, "clsZip.CompressFile()", "Error compressing byte array"
    End If
    
    '/* Save original size in bytes (needed for decompress)
    FH.OrigSize = m_OriginalSize
    FH.FileExt = right(SourceFile, 3)
    FH.ID = MyID
    
    '/* Delete old backup file if necessary
    If Dir(DestinationFile) > vbNullString Then Kill DestinationFile
    
    If CheckForRoom(DestinationFile, m_CompressedSize) Then
        '/* Write compressed file
        Open DestinationFile For Binary Access Write As #FF
            Put #FF, , FH
            Put #FF, , Buffer()
        Close #FF
    End If
    
    '/* Clean up
    Erase Buffer

End Sub
Public Sub DecompressFile(ByVal SourceFile As String, ByVal DestinationFile As String)
  Dim Buffer() As Byte, Result As Long, m_FileExt As String, i As Integer
  Dim FF As Integer, FH As FileHeaderType
    
    FF = FreeFile
    
    '/* Fix Buffer size
    ReDim Buffer(FileLen(SourceFile) - 1) As Byte
    
    '/* Read byte array from file
    Open SourceFile For Binary Access Read As #FF
        Get #FF, , FH
        Get #FF, , Buffer()
    Close #FF
    
    If FH.ID <> MyID Then
        MsgBox "This file is not recognized as a compressed file", vbInformation
        Erase Buffer
        Exit Sub
    End If
    
    '/* Get original file size before compression
    m_OriginalSize = FH.OrigSize
    
    '/* Get original file extention if none is supplied
    If Mid(DestinationFile, Len(DestinationFile) - 3, 1) <> "." Then
        DestinationFile = DestinationFile & "." & FH.FileExt
    End If
    
    Result = DecompressData(Buffer(), m_OriginalSize)
    If Result <> 0 Then
        Err.Raise Result, "clsZip.DecompressFile()", "Error decompressing byte array"
    End If
    
    '/* Delete old file if necessary
    If Dir(DestinationFile) > vbNullString Then Kill DestinationFile
    
    If CheckForRoom(DestinationFile, m_OriginalSize) Then
        '/* Write uncompressed file
        Open DestinationFile For Binary Access Write As #FF
            Put #FF, , Buffer()
        Close #FF
    End If
    
    '/* Clean up
    Erase Buffer

End Sub


Public Property Get CompressLevel() As Variant
    CompressLevel = m_CompressLevel
End Property

Public Property Let CompressLevel(ByVal vNewValue As Variant)
    m_CompressLevel = Val(vNewValue)
    If m_CompressLevel < 1 Then m_CompressLevel = 1
    If m_CompressLevel > 9 Then m_CompressLevel = 9
End Property

Private Function CheckForRoom(ByVal DestinationFile As String, ByVal Size As Long) As Boolean
  Dim tString As String
  Dim FreeSpace As Double, SpaceNeeded As Double
  Dim Result As Boolean

TryAgain:

    Result = True
    FreeSpace = SpaceFree(left(DestinationFile, 3))
    SpaceNeeded = CDbl(m_CompressedSize + 1)
    
    If SpaceNeeded > FreeSpace Then
        tString = "Drive " & left(DestinationFile, 3) & " does not have enough free space." & vbCrLf & _
            Format((SpaceNeeded - FreeSpace), "0") & " additional bytes needed."
        If MsgBox(tString, vbInformation + vbRetryCancel, "Compress File") = vbRetry Then
            GoTo TryAgain
        Else
            Result = False
        End If
    End If
    
    CheckForRoom = Result
  
End Function
