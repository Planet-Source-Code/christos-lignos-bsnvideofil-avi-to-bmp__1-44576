VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDIBFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
        (ByVal hWndOwner As Long, ByVal nFolder As SPECIAL_FOLDERS, _
         pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
   ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function GetOpenFileNamePreview Lib "MSVFW32" _
    Alias "GetOpenFileNamePreviewA" (filestruct As OPENFILENAME) As Long

Private Const MAX_PATH = 1024
Private Const MAX_FILE = 512

Public Enum SPECIAL_FOLDERS
    vbCSIDL_DESKTOP = &H0&
    vbCSIDL_PROGRAMS = &H2&
    vbCSIDL_CONTROLS = &H3&
    vbCSIDL_PRINTERS = &H4&
    vbCSIDL_PERSONAL = &H5&
    vbCSIDL_FAVORITES = &H6&
    vbCSIDL_STARTUP = &H7&
    vbCSIDL_RECENT = &H8&
    vbCSIDL_SENDTO = &H9&
    vbCSIDL_BITBUCKET = &HA&
    vbCSIDL_STARTMENU = &HB&
    vbCSIDL_DESKTOPDIRECTORY = &H10&
    vbCSIDL_DRIVES = &H11&
    vbCSIDL_NETWORK = &H12&
    vbCSIDL_NETHOOD = &H13&
    vbCSIDL_FONTS = &H14&
    vbCSIDL_TEMPLATES = &H15&
End Enum

Private Const NOERROR As Long = &H0

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    Flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Public Enum EOpenFile
    OFN_READONLY = &H1&
    OFN_OVERWRITEPROMPT = &H2&
    OFN_HIDEREADONLY = &H4&
    OFN_NOCHANGEDIR = &H8&
    OFN_SHOWHELP = &H10&
    OFN_ENABLEHOOK = &H20&
    OFN_ENABLETEMPLATE = &H40&
    OFN_ENABLETEMPLATEHANDLE = &H80&
    OFN_NOVALIDATE = &H100&
    OFN_ALLOWMULTISELECT = &H200&
    OFN_EXTENSIONDIFFERENT = &H400&
    OFN_PATHMUSTEXIST = &H800&
    OFN_FILEMUSTEXIST = &H1000&
    OFN_CREATEPROMPT = &H2000&
    OFN_SHAREAWARE = &H4000&
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private m_FileMustExist As Boolean
Private m_MultiSelect As Boolean
Private m_ReadOnly As Boolean 'read only
Private m_HideReadOnly As Boolean
Private m_Filter As String
Private m_FilterIndex As Long
Private m_InitDir As String
Private m_DlgTitle As String
Private m_DefaultExt As String
Private m_Flags As Long
Private m_OverwritePrompt As Boolean
Private m_hWnd As Long

Private Sub Class_Initialize()

m_hWnd = -1&
m_FileMustExist = True
m_MultiSelect = False
m_HideReadOnly = True
m_DlgTitle = App.title
m_OverwritePrompt = True
m_InitDir = GetSpecialFolderLocation(vbCSIDL_PERSONAL) 'default to My Documents folder

End Sub

Property Get OwnerHwnd() As Long
    OwnerHwnd = m_hWnd
End Property
Property Let OwnerHwnd(ByVal vHwnd As Long)
    m_hWnd = vHwnd
End Property





Property Get Filter() As String
    Filter = m_Filter
End Property
Property Let Filter(ByVal vFilterString As String)
    m_Filter = vFilterString
End Property

Property Get FilterIndex() As Long
    FilterIndex = m_FilterIndex
End Property
Property Let FilterIndex(ByVal vIndex As Long)
    m_FilterIndex = vIndex
End Property



Property Get DlgTitle() As String
    DlgTitle = m_DlgTitle
End Property
Property Let DlgTitle(ByVal title As String)
    m_DlgTitle = title
End Property


Property Get Flags() As EOpenFile
    Flags = m_Flags
End Property
Property Let Flags(ByVal vFlags As EOpenFile)
    m_Flags = vFlags
End Property










Public Function VBGetOpenFileNamePreview(filename As String, _
                           Optional FileTitle As String) As Boolean
    Dim opfile As OPENFILENAME
    Dim s As String
    Dim ch As String
    Dim i As Integer

With opfile
    .lStructSize = Len(opfile)
    .Flags = (-m_FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-m_MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-m_ReadOnly * OFN_READONLY) Or _
             (-m_HideReadOnly * OFN_HIDEREADONLY) Or _
             (m_Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    If m_hWnd <> -1 Then .hWndOwner = m_hWnd
    .lpstrInitialDir = m_InitDir
    .lpstrDefExt = m_DefaultExt
    .lpstrTitle = m_DlgTitle
    
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = m_FilterIndex
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE

    If GetOpenFileNamePreview(opfile) Then
        VBGetOpenFileNamePreview = True
        filename = left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        FileTitle = left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
        m_Flags = .Flags
        m_FilterIndex = .nFilterIndex
        m_Filter = FilterLookup(.lpstrFilter, m_FilterIndex)
        If (.Flags And OFN_READONLY) Then m_ReadOnly = True
        m_InitDir = .lpstrFile
    Else
        VBGetOpenFileNamePreview = False
        filename = vbNullChar
        FileTitle = vbNullChar
        Flags = 0
        FilterIndex = -1
        Filter = vbNullChar
    End If
End With


End Function







Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long
    Dim iEnd As Long
    Dim s As String

    iStart = 1
    If sFilters = vbNullChar Then Exit Function
    Do
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function



Private Function GetSpecialFolderLocation(CSIDL As SPECIAL_FOLDERS) As String

   Dim sPath As String
   Dim pidl As Long

   If SHGetSpecialFolderLocation(m_hWnd, CSIDL, pidl) = NOERROR Then
      sPath = Space$(MAX_PATH)
      If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then
         Call CoTaskMemFree(ByVal VarPtr(pidl))
         GetSpecialFolderLocation = left(sPath, InStr(sPath, Chr$(0)) - 1)
      End If
    End If

End Function
