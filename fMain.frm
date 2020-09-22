VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "AVI File  to BMP"
   ClientHeight    =   1860
   ClientLeft      =   4155
   ClientTop       =   2070
   ClientWidth     =   4410
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   4410
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   270
      TabIndex        =   1
      Top             =   1125
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpenAVIFile 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select AVI File"
      Height          =   420
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   345
      Width           =   1575
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call AVIFileInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call AVIFileExit
End Sub

Private Sub cmdOpenAVIFile_Click()
    Dim res As Long
    Dim OpenFileIN As FDIBFileDlg
    Dim szFile As String
    Dim pAVIFile As Long
    Dim pAVIStream As Long
    Dim numFrames As Long
    Dim firstFrame As Long
    Dim fileInfo As AVI_FILE_INFO
    Dim streamInfo As AVI_STREAM_INFO
    Dim dib As FDIBPointer
    Dim pGetFrameObj As Long
    Dim pDIB As Long
    Dim bih As BITMAPINFOHEADER
    Dim i As Long

    Set OpenFileIN = New FDIBFileDlg
    With OpenFileIN
        .OwnerHwnd = Me.hWnd
        .Filter = "AVI Files|*.avi"
        .DlgTitle = "Open AVI File"
    End With
    
    res = OpenFileIN.VBGetOpenFileNamePreview(szFile)
    If res = False Then GoTo ErrorOut

    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut

    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut

    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut

    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut

    res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut

    res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut

    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.bottom - streamInfo.rcFrame.top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.right - streamInfo.rcFrame.left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight
    End With

    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih)
    If pGetFrameObj = 0 Then
        MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.title
        GoTo ErrorOut
    End If

    Set dib = New FDIBPointer
    For i = firstFrame To (numFrames - 1) + firstFrame
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)
        If dib.CreateFromPackedDIBPointer(pDIB) Then
            Call dib.WriteToFile(App.Path & "\" & i & ".bmp")
            txtStatus = "Bitmap " & i + 1 & " of " & numFrames & " written to app folder"
            txtStatus.Refresh
        Else

        End If
    Next

    Set dib = Nothing

ErrorOut:
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj)
    End If
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream)
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile)
    End If
    If (res <> AVIERR_OK) Then
        MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
    End If
End Sub

