Imports System.Runtime.InteropServices
Imports System.Drawing

''' <summary>extract an icon based on file extension</summary>
''' <see cref="https://msdn.microsoft.com/en-us/library/windows/desktop/bb762179(v=vs.85).aspx"/>
''' <seealso cref="http://www.pinvoke.net/default.aspx/shell32.SHGetFileInfo"/> 
''' <seealso cref="https://www.google.com/#q=SHGetFileInfo"/> 
Public Class clsExtractIcon
    ' DLL Import
    <DllImport("shell32")>
    Private Shared Function SHGetFileInfo(pszPath As String, dwFileAttributes As UInteger, ByRef psfi As SHFILEINFO, cbFileInfo As UInteger, uFlags As UInteger) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function DestroyIcon(hIcon As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function
    ' Constants/Enums

    Private Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
    ''' <summary>Maximal Length of unmanaged Windows-Path-strings</summary>
    Private Const MAX_PATH As Integer = 260
    ''' <summary>Maximal Length of unmanaged Typename</summary>
    Private Const MAX_TYPE As Integer = 80
    <Flags>
    Private Enum SHGetFileInfoConstants As Integer
        SHGFI_ICON = &H100
        ' get icon
        SHGFI_DISPLAYNAME = &H200
        ' get display name
        SHGFI_TYPENAME = &H400
        ' get type name
        SHGFI_ATTRIBUTES = &H800
        ' get attributes
        SHGFI_ICONLOCATION = &H1000
        ' get icon location
        SHGFI_EXETYPE = &H2000
        ' return exe type
        SHGFI_SYSICONINDEX = &H4000
        ' get system icon index
        SHGFI_LINKOVERLAY = &H8000
        ' put a link overlay on icon
        SHGFI_SELECTED = &H10000
        ' show icon in selected state
        SHGFI_ATTR_SPECIFIED = &H20000
        ' get only specified attributes
        SHGFI_LARGEICON = &H0
        ' get large icon
        SHGFI_SMALLICON = &H1
        ' get small icon
        SHGFI_OPENICON = &H2
        ' get open icon
        SHGFI_SHELLICONSIZE = &H4
        ' get shell size icon
        SHGFI_PIDL = &H8
        ' pszPath is a pidl
        SHGFI_USEFILEATTRIBUTES = &H10
        ' use passed dwFileAttribute
        SHGFI_ADDOVERLAYS = &H20
        ' apply the appropriate overlays
        SHGFI_OVERLAYINDEX = &H40
        ' Get the index of the overlay
    End Enum

    <Flags>
    Public Enum SHGFI_ICON_SIZE As Integer
        ''' <summary>get jumbo icon</summary>
        JumboIcon = &H4 '256x256
        ''' <summary>get extra large icon</summary>
        ExtraLargeIcon = &H3 '48x48
        ''' <summary>get large icon</summary>
        LargeIcon = &H0 '32x32
        ''' <summary>get small icon</summary>
        SmallIcon = &H1 '16x16
    End Enum


    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Private Structure SHFILEINFO
        Public Sub New(b As Boolean)
            hIcon = IntPtr.Zero
            iIcon = 0
            dwAttributes = 0
            szDisplayName = ""
            szTypeName = ""
        End Sub
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_TYPE)>
        Public szTypeName As String
    End Structure



    '''<example>PictureBox1.Image=ExtractIcon.GetIconForExtension(".txt",SHGFI_ICON_SIZE.SmallIcon).ToBitmap</example>
    '''<example>ImageList1.Images.Add(".txt",ExtractIcon.GetIconForExtension(".txt",ExtractIcon.SHGFI_ICON_SIZE.SmallIcon).ToBitmap) </example>
    Public Shared Function GetIconForExtension(extension As String, iconsize As SHGFI_ICON_SIZE) As Icon

        Dim shinfo As New SHFILEINFO()
        Dim options As SHGetFileInfoConstants
        options = SHGetFileInfoConstants.SHGFI_ICON Or iconsize Or SHGetFileInfoConstants.SHGFI_USEFILEATTRIBUTES Or SHGetFileInfoConstants.SHGFI_TYPENAME
        Dim ptr As IntPtr = SHGetFileInfo(extension, FILE_ATTRIBUTE_NORMAL, shinfo, CUInt(Marshal.SizeOf(shinfo)), CInt(options))
        Dim icon__1 As Icon = TryCast(Icon.FromHandle(shinfo.hIcon).Clone(), Icon)
        DestroyIcon(shinfo.hIcon)
        Return icon__1
    End Function


End Class
