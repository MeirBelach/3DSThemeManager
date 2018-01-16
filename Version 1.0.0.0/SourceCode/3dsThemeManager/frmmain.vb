Imports System.IO
Imports System.IO.Compression

Public Class frmmain
    Dim msgtitle = "3DS Theme Manager"

    Private Sub themeloader()
        For Each Dir As String In Directory.GetDirectories(My.Settings.ThemesFolder)
            ListBox1.Items.Add(Path.GetFileName(Dir))
        Next
    End Sub

    Private Sub compression(ByVal folderin As String, ByVal strname As String)
        Dim startPath As String = folderin
        Dim zipPath As String = My.Settings.ZipFolder & "\" & strname & ".zip"
        Dim tAf = False
        If File.Exists(zipPath) = False Then
            tAf = False
        Else
            If MsgBox("Overwrite: " & zipPath & "?", vbYesNo + vbCritical, msgtitle) = vbYes Then
                File.Delete(zipPath)
            Else
                tAf = True
            End If
        End If

        If tAf = False Then
            ZipFile.CreateFromDirectory(startPath, zipPath)
        End If
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsToolStripMenuItem.Click
        frmoptions.Show()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PictureBox1.Image = Me.Icon.ToBitmap

        If Directory.Exists(My.Settings.ThemesFolder) = False Then
            My.Settings.ThemesFolder = ""
            My.Settings.Save()
        End If
        If Directory.Exists(My.Settings.ZipFolder) = False Then
            My.Settings.ZipFolder = ""
            My.Settings.Save()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ListBox1.SelectedItem <> "" Then
            PictureBox1.ImageLocation = "G:\3dsThemes\Custom\" & ListBox1.SelectedItem & "\preview.png"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListBox1.SelectedItem <> "" Then
            If ListBox2.Items.Contains(ListBox1.SelectedItem) = False Then
                ListBox2.Items.Add(ListBox1.SelectedItem)
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If ListBox2.SelectedItem <> "" Then
            ListBox2.Items.Remove(ListBox2.SelectedItem)
        End If
    End Sub

    Private Sub CompressThemesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompressThemesToolStripMenuItem.Click
        If ListBox2.Items.Count = 0 Then
            MsgBox("No Theme(s) to Zip.", vbOKOnly + vbExclamation, msgtitle)
        Else
            If My.Settings.ZipFolder <> "" Then

                For Each item In ListBox2.Items
                    compression(My.Settings.ThemesFolder & "\" & item.ToString, item.ToString)
                Next

                ListBox2.Items.Clear()
                MsgBox("Done Zipping.", vbOKOnly + vbExclamation, msgtitle)
                If MsgBox("Open Zip Folder?", vbYesNo + vbQuestion, msgtitle) = vbYes Then
                    Shell("explorer " & My.Settings.ZipFolder, AppWinStyle.NormalFocus)
                End If
            End If
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If My.Settings.ThemesFolder <> "" Then
            themeloader()
            Timer2.Stop()
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        frmabout.Show()
    End Sub
End Class
