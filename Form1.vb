Imports System.Runtime.InteropServices
Imports System.Diagnostics

Public Class Form1
    ' Import the SHEmptyRecycleBin function from shell32.dll
    <DllImport("shell32.dll", SetLastError:=True)>
    Private Shared Function SHEmptyRecycleBin(hwnd As IntPtr, pszRootPath As String, dwFlags As UInteger) As Integer
    End Function

    Private Const SHERB_NOCONFIRMATION As UInteger = &H1
    Private Const SHERB_NOPROGRESSUI As UInteger = &H2
    Private Const SHERB_NOSOUND As UInteger = &H4

    Private notifyIcon As NotifyIcon
    Private contextMenu As ContextMenuStrip

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Create context menu
        contextMenu = New ContextMenuStrip()
        contextMenu.Items.Add("Empty Recycle Bin", Nothing, AddressOf EmptyRecycleBin_Click)
        contextMenu.Items.Add("Open Recycle Bin", Nothing, AddressOf OpenRecycleBin_Click)

        ' Create notify icon
        notifyIcon = New NotifyIcon()
        notifyIcon.Icon = SystemIcons.Information ' You can use a custom icon here
        notifyIcon.Visible = True
        notifyIcon.ContextMenuStrip = contextMenu
        notifyIcon.Text = "Recycle Bin Manager"
    End Sub

    Private Sub EmptyRecycleBin_Click(sender As Object, e As EventArgs)
        SHEmptyRecycleBin(Me.Handle, Nothing, SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND)
    End Sub

    Private Sub OpenRecycleBin_Click(sender As Object, e As EventArgs)
        Process.Start("explorer.exe", "shell:RecycleBinFolder")
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If notifyIcon IsNot Nothing Then
            notifyIcon.Visible = False
            notifyIcon.Dispose()
        End If
        MyBase.OnFormClosing(e)
    End Sub
End Class
