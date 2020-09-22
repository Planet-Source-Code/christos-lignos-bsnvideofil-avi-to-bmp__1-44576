VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDIBPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal dwLen As Long)

Private Const BMP_MAGIC_COOKIE As Integer = 19778
Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
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

Private m_memBits() As Byte
Private m_memBitmapInfo() As Byte
Private m_bih As BITMAPINFOHEADER
Private m_bfh As BITMAPFILEHEADER
Public Function CreateFromPackedDIBPointer(ByRef pDIB As Long) As Boolean
Debug.Assert pDIB <> 0

Call CopyMemory(ByVal VarPtr(m_bih.biSize), ByVal pDIB, Len(m_bih))
If m_bih.biBitCount < 16 Then
    Debug.Print "Error! DIB was less than 16 colors."
    Exit Function
End If

If m_bih.biSizeImage < 1 Then Exit Function
ReDim m_memBits(0 To m_bih.biSizeImage - 1)
Call CopyMemory(m_memBits(0), ByVal pDIB + 40, m_bih.biSizeImage)

ReDim m_memBitmapInfo(0 To 39)
Call CopyMemory(m_memBitmapInfo(0), m_bih, Len(m_bih))

With m_bfh
    .bfType = BMP_MAGIC_COOKIE
    .bfSize = 55 + m_bih.biSizeImage
    .bfReserved1 = 0&
    .bfReserved2 = 0&
    .bfOffBits = 54
End With

CreateFromPackedDIBPointer = True

End Function

Public Function WriteToFile(ByVal filename As String) As Boolean

Dim hFile As Integer
On Error Resume Next
hFile = FreeFile()

Open filename For Binary As hFile
    Put hFile, 1, m_bfh
    Put hFile, Len(m_bfh) + 1, m_memBitmapInfo
    Put hFile, , m_memBits
Close hFile

WriteToFile = True

End Function



















