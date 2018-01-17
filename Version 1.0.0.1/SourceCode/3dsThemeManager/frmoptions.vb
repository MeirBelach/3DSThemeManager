Public Class frmoptions

    Private Function openfolder() As String
        Dim nam = ""
        FolderBrowserDialog1.SelectedPath = ""
        FolderBrowserDialog1.ShowDialog()
        If FolderBrowserDialog1.SelectedPath <> "" Then
            nam = FolderBrowserDialog1.SelectedPath
        End If

        Return nam
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = openfolder()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text = openfolder()
    End Sub

    Private Sub SaveSettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveSettingsToolStripMenuItem.Click
        My.Settings.ThemesFolder = TextBox1.Text
        My.Settings.ZipFolder = TextBox2.Text
        If CheckBox1.Checked = True Then
            My.Settings.Playmusic = "True"
        Else
            My.Settings.Playmusic = "False"
        End If
        My.Settings.Save()
    End Sub

    Private Sub options_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = My.Settings.ThemesFolder
        TextBox2.Text = My.Settings.ZipFolder
        If My.Settings.Playmusic = "True" Then
            CheckBox1.Checked = True
        End If
    End Sub
End Class