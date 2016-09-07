Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text

Public Class ActiveDesktopForm
    '在使用之前要用CoCreateObject来进行创建。 
    <DllImport("OLE32.DLL")>
    Public Shared Function CoCreateInstance(
         ByRef ClassGuid As Guid,
         ByVal pUnkOuter As IntPtr,
         ByVal dwClsContext As Integer,
         ByRef InterfaceGuid As Guid,
         ByRef Result As IActiveDesktop) As IntPtr
    End Function

    Private Const CLSCTX_INPROC_SERVER As Integer = 1
    Private CLSID_ActiveDesktop As New Guid("75048700-EF1F-11D0-9888-006097DEACF9")
    Private IID_IActiveDesktop As New Guid("F490EB00-1240-11D1-9888-006097DEACF9")
    Private ActiveDesktop As IActiveDesktop
    Private COMPonent As COMPONENT = New COMPONENT

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '创建IActiveDesktop实例
        CoCreateInstance(CLSID_ActiveDesktop, IntPtr.Zero, CLSCTX_INPROC_SERVER, IID_IActiveDesktop, ActiveDesktop)
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        '释放ActiveDesktop对象
        Marshal.ReleaseComObject(ActiveDesktop)
    End Sub

    ''' <summary>
    ''' 获取桌面壁纸文件路径
    ''' </summary>
    Private Function GetWallpaperFilePath() As String
        Dim WallpaperFilePath As New StringBuilder(260)
        ActiveDesktop.GetWallpaper(WallpaperFilePath, WallpaperFilePath.Capacity, 0)
        Return WallpaperFilePath.ToString()
    End Function

    ''' <summary>
    ''' 设置桌面壁纸
    ''' </summary>
    ''' <param name="WallpaperFilePath">壁纸文件的路径</param>
    ''' <returns>接口函数返回值</returns>
    Private Function SetWallpaperFile(ByVal WallpaperFilePath As String) As Integer
        Dim SetWallpaperResult As Integer = ActiveDesktop.SetWallpaper(WallpaperFilePath, 0)
        ActiveDesktop.ApplyChanges(AD_APPLY.FORCE Or AD_APPLY.SAVE Or AD_APPLY.REFRESH)
        Return SetWallpaperResult
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Templates) & "\" & My.Computer.Clock.TickCount & ".jpg"
        Button1.BackgroundImage.Save(FilePath, Imaging.ImageFormat.Jpeg)
        SetWallpaperFile(FilePath)
        Kill(FilePath)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim FilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Templates) & "\" & My.Computer.Clock.TickCount & ".jpg"
        Button2.BackgroundImage.Save(FilePath, Imaging.ImageFormat.Jpeg)
        SetWallpaperFile(FilePath)
        Kill(FilePath)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim FilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Templates) & "\" & My.Computer.Clock.TickCount & ".jpg"
        Button3.BackgroundImage.Save(FilePath, Imaging.ImageFormat.Jpeg)
        SetWallpaperFile(FilePath)
        Kill(FilePath)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim FilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Templates) & "\" & My.Computer.Clock.TickCount & ".jpg"
        Button4.BackgroundImage.Save(FilePath, Imaging.ImageFormat.Jpeg)
        SetWallpaperFile(FilePath)
        Kill(FilePath)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '&H80070057 //参数错误
        '&H80004001 //未实现

        Dim PonentPosition As COMPPOS
        COMPonent.dwSize = Marshal.SizeOf(COMPonent) '！MSDN: 必须有！

        PonentPosition.dwHeight = Me.Height
        PonentPosition.dwWidth = Me.Width
        PonentPosition.fCanResize = True
        PonentPosition.iLeft = 0
        PonentPosition.iPreferredLeftPercent = 50
        PonentPosition.iTop = 0
        PonentPosition.iPreferredTopPercent = 50
        PonentPosition.izIndex = 0
        COMPonent.cpPos = PonentPosition
        COMPonent.dwID = 1001
        COMPonent.fChecked = True
        COMPonent.fDirty = True
        COMPonent.fNoScroll = True
        COMPonent.iComponentType = COMP_TYPE.PICTURE
        COMPonent.wszFriendlyName = "LeonTest"
        COMPonent.wszSource = "LeonTest"
        COMPonent.wszSubscribedURL = "http://www.baidu.com"
        Dim Result As Integer = ActiveDesktop.AddDesktopItem(COMPonent, 0) 'MSDN: Reserved. Must be set to zero.
        If Result <> 0 Then MsgBox("失败...      &H" & Hex(Result))

        ActiveDesktop.ApplyChanges(AD_APPLY.ALL)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim intCount As Integer
        Dim intIndex As Integer = 0
        ActiveDesktop.GetDesktopItemCount(intCount, 0)
        MsgBox("共有 " & intCount & " 个桌面元素")

        COMPonent.dwSize = Marshal.SizeOf(COMPonent)
        For intIndex = 0 To intCount - 1

        Next

        While (intIndex <= (intCount - 1))
            ActiveDesktop.GetDesktopItem(intIndex, COMPonent, 0)
        End While

        MsgBox(COMPonent.cpPos.iLeft.ToString & vbCrLf &
            COMPonent.iComponentType.ToString & vbCrLf &
             COMPonent.wszFriendlyName & vbCrLf &
             COMPonent.wszSource & vbCrLf &
             COMPonent.wszSubscribedURL & vbCrLf &
             COMPonent.csiOriginal.dwItemState.ToString
            )
    End Sub
End Class