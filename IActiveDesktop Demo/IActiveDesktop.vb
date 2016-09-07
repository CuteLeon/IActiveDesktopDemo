Imports System.Runtime.InteropServices
Imports System.Text

Public Enum WPSTYLE
    CENTER = 0
    TILE = 1
    STRETCH = 2
    FIT = 3
    FILL = 4
    SPAN = 5
    MAX = 6
End Enum

Public Structure WALLPAPEROPT
        Public dwSize As Integer
        Public dwStyle As WPSTYLE
    End Structure

    Public Structure COMPONENTSOPT
        Public dwSize As Integer
        <MarshalAs(UnmanagedType.Bool)>
        Public fActiveDesktop As Boolean
        <MarshalAs(UnmanagedType.Bool)>
        Public fEnableComponents As Boolean
    End Structure

    Public Structure COMPPOS
        Public Const COMPONENT_TOP As Integer = &H3FFFFFFF
        Public Const COMPONENT_DEFAULT_LEFT As Integer = &HFFFF
        Public Const COMPONENT_DEFAULT_TOP As Integer = &HFFFF
        Public dwSize As Integer
        Public iLeft As Integer
        Public iTop As Integer
        Public dwWidth As Integer
        Public dwHeight As Integer
        Public izIndex As Integer
        <MarshalAs(UnmanagedType.Bool)>
        Public fCanResize As Boolean
        <MarshalAs(UnmanagedType.Bool)>
        Public fCanResizeX As Boolean
        <MarshalAs(UnmanagedType.Bool)>
        Public fCanResizeY As Boolean
        Public iPreferredLeftPercent As Integer
        Public iPreferredTopPercent As Integer
    End Structure

    <Flags>
    Public Enum ITEMSTATE
        NORMAL = &H1
        FULLSCREEN = 2
        SPLIT = &H4
        VALIDSIZESTATEBITS = NORMAL Or SPLIT Or FULLSCREEN
        VALIDSTATEBITS = NORMAL Or SPLIT Or FULLSCREEN Or &H80000000 Or &H40000000
    End Enum

    Public Structure COMPSTATEINFO
        Public dwHeight As Integer
        Public dwItemState As Integer
        Public dwSize As Integer
        Public dwWidth As Integer
        Public iLeft As Integer
        Public iTop As Integer
    End Structure

    Public Enum COMP_TYPE
        HTMLDOC = 0
        PICTURE = 1
        WEBSITE = 2
        CONTROL = 3
        CFHTML = 4
        MAX = 4
    End Enum

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure COMPONENT
    Private Const INTERNET_MAX_URL_LENGTH As Integer = 2084
    Public dwSize As Integer
    Public dwID As Integer
    Public iComponentType As COMP_TYPE
    <MarshalAs(UnmanagedType.Bool)>
    Public fChecked As Boolean
    <MarshalAs(UnmanagedType.Bool)>
    Public fDirty As Boolean
    <MarshalAs(UnmanagedType.Bool)>
    Public fNoScroll As Boolean
    Public cpPos As COMPPOS
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
    Public wszFriendlyName As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=INTERNET_MAX_URL_LENGTH)>
    Public wszSource As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=INTERNET_MAX_URL_LENGTH)>
    Public wszSubscribedURL As String
    '——————————————
    Public dwCurItemState As Integer
    Public csiOriginal As COMPSTATEINFO
    Public csiRestored As COMPSTATEINFO
End Structure

Public Enum DTI_ADTIWUI
        DTI_ADDUI_DEFAULT = &H0
        DTI_ADDUI_DISPSUBWIZARD = &H1
        DTI_ADDUI_POSITIONITEM = &H2
    End Enum

    <Flags>
    Public Enum AD_APPLY
        SAVE = &H1
        HTMLGEN = &H2
        REFRESH = &H4
        ALL = SAVE Or HTMLGEN Or REFRESH
        FORCE = &H8
        BUFFERED_REFRESH = &H10
        DYNAMICREFRESH = &H20
    End Enum

<Flags>
Public Enum COMP_ELEM
    TYPE = &H1
    CHECKED = &H2
    DIRTY = &H4
    NOSCROLL = &H8
    POS_LEFT = &H10
    POS_TOP = &H20
    SIZE_WIDTH = &H40
    SIZE_HEIGHT = &H80
    POS_ZINDEX = &H100
    SOURCE = &H200
    FRIENDLYNAME = &H400
    SUBSCRIBEDURL = &H800
    ORIGINAL_CSI = &H1000
    RESTORED_CSI = &H2000
    CURITEMSTATE = &H4000
    ALL = TYPE Or CHECKED Or DIRTY Or NOSCROLL Or POS_LEFT Or
        SIZE_WIDTH Or SIZE_HEIGHT Or POS_ZINDEX Or SOURCE Or
        FRIENDLYNAME Or POS_TOP Or SUBSCRIBEDURL Or ORIGINAL_CSI Or
        RESTORED_CSI Or CURITEMSTATE
End Enum

<Flags>
    Public Enum ADDURL
        SILENT = &H1
    End Enum

<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("F490EB00-1240-11D1-9888-006097DEACF9")>
Public Interface IActiveDesktop
    <PreserveSig()>
    Function ApplyChanges(ByVal dwFlags As AD_APPLY) As Integer
    <PreserveSig()>
    Function GetWallpaper(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwszWallpaper As StringBuilder, ByVal cchWallpaper As Integer, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function SetWallpaper(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwszWallpaper As String, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GetWallpaperOptions(ByRef pwpo As WALLPAPEROPT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function SetWallpaperOptions(ByRef pwpo As WALLPAPEROPT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GetPattern(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwszPattern As StringBuilder, ByVal cchPattern As Integer, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function SetPattern(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwszPattern As String, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GetDesktopItemOptions(ByRef pco As COMPONENTSOPT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function SetDesktopItemOptions(ByRef pco As COMPONENTSOPT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function AddDesktopItem(ByRef pcomp As COMPONENT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function AddDesktopItemWithUI(ByVal hwnd As IntPtr, ByRef pcomp As COMPONENT, ByVal dwFlags As DTI_ADTIWUI) As Integer
    <PreserveSig()>
    Function ModifyDesktopItem(ByRef pcomp As COMPONENT, ByVal dwFlags As COMP_ELEM) As Integer
    <PreserveSig()>
    Function RemoveDesktopItem(ByRef pcomp As COMPONENT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GetDesktopItemCount(<Out()> ByRef lpiCount As Integer, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GetDesktopItem(ByVal nComponent As Integer, ByRef pcomp As COMPONENT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GetDesktopItemByID(ByVal dwID As IntPtr, ByRef pcomp As COMPONENT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function GenerateDesktopItemHtml(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwszFileName As String, ByRef pcomp As COMPONENT, ByVal dwReserved As Integer) As Integer
    <PreserveSig()>
    Function AddUrl(ByVal hwnd As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal pszSource As String, ByRef pcomp As COMPONENT, ByVal dwFlags As ADDURL) As Integer
    <PreserveSig()>
    Function GetDesktopItemBySource(<MarshalAs(UnmanagedType.LPWStr)> ByVal pwszSource As String, ByRef pcomp As COMPONENT, ByVal dwReserved As Integer) As Integer
End Interface