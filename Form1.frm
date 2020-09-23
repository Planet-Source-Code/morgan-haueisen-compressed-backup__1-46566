VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo"
   ClientHeight    =   4320
   ClientLeft      =   2265
   ClientTop       =   1710
   ClientWidth     =   5505
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   5505
   Begin VB.CommandButton Command2 
      Caption         =   "Decompress"
      Height          =   345
      Left            =   3555
      TabIndex        =   4
      Top             =   3720
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compress"
      Height          =   405
      Left            =   3570
      TabIndex        =   3
      Top             =   3195
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3180
      TabIndex        =   2
      Top             =   90
      Width           =   2280
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   135
      TabIndex        =   1
      Top             =   420
      Width           =   2970
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   2985
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cZip As clsZip

Private Sub Command1_Click()
  Dim ZipFile As String
  Dim FilePath As String, FileName As String
    
    FileName = File1.FileName
    FilePath = File1.Path
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    If Dir(FilePath & FileName) > vbNullString And FileName > vbNullString Then
        '/* Change file extention for compressed file
        ZipFile = Left(FileName, Len(FileName) - 3) & "ssz"

        Set cZip = New clsZip
        cZip.CompressFile FilePath & FileName, FilePath & ZipFile
        Set cZip = Nothing
        File1.Refresh
    Else
        MsgBox "Select a file"
    End If
    
End Sub

Private Sub Command2_Click()
  Dim UnzipFile As String
  Dim FilePath As String, FileName As String
  
    FileName = File1.FileName
    FilePath = File1.Path
    If Right(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
    
    If Dir(FilePath & FileName) > vbNullString And FileName > vbNullString Then
        '/* change the name of the uncompressed file so that
        '/* the original is not overwritten
        UnzipFile = Left(FileName, Len(FileName) - 4) & "_unzipped"

        Set cZip = New clsZip
        cZip.DecompressFile FilePath & FileName, FilePath & UnzipFile
        Set cZip = Nothing
        File1.Refresh
    Else
        MsgBox "Select a file"
    End If
    
End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1.Refresh
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    Dir1.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
End Sub


