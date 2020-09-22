Attribute VB_Name = "AVIFileIN"
Option Explicit

Public Declare Sub AVIFileInit Lib "avifil32.dll" ()
Public Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT
Public Declare Function AVIFileInfo Lib "avifil32.dll" (ByVal pfile As Long, pfi As AVI_FILE_INFO, ByVal lSize As Long) As Long 'HRESULT
Public Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, _
                                                                ByRef bih As Any) As Long 'returns pointer to GETFRAME object on success (or NULL on error)
Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, _
                                                                ByVal lPos As Long) As Long 'returns pointer to packed DIB on success (or NULL on error)
Public Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (ByVal pGetFrameObj As Long) As Long ' returns zero on success (error number) after calling this function the GETFRAME object pointer is invalid
Public Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
Public Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long
Public Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long
Public Declare Sub AVIFileExit Lib "avifil32.dll" ()

Public Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type


Public Type AVI_RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type AVI_STREAM_INFO
    fccType As Long
    fccHandler As Long
    dwFlags As Long
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As AVI_RECT
    dwEditCount  As Long
    dwFormatChangeCount As Long
    szName As String * 64
End Type

Public Type AVI_FILE_INFO
    dwMaxBytesPerSecond As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

Public Type AVI_COMPRESS_OPTIONS
    fccType As Long
    fccHandler As Long
    dwKeyFrameEvery As Long
    dwQuality As Long
    dwBytesPerSecond As Long
    dwFlags As Long
    lpFormat As Long
    cbFormat As Long
    lpParms As Long
    cbParms As Long
    dwInterleaveEvery As Long
End Type

Global Const AVIERR_OK             As Long = 0&
Private Const SEVERITY_ERROR       As Long = &H80000000
Private Const FACILITY_ITF         As Long = &H40000
Private Const AVIERR_BASE          As Long = &H4000
Global Const OF_SHARE_DENY_WRITE   As Long = &H20
Global Const BI_RGB                As Long = 0
Global Const streamtypeVIDEO       As Long = 1935960438






