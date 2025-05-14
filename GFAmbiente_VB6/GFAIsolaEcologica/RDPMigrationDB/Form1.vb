Imports System.Text

Public Class Form1

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.Filter = "RDP DataBase (*.mdb)|*.mdb|All files (*.*)|*.*"
        OpenFileDialog1.DefaultExt = "mdb"
        OpenFileDialog1.FilterIndex = 0
        OpenFileDialog1.Title = "Selezione Database"

        ProgressBar1.Style = ProgressBarStyle.Blocks
        ProgressBar1.Visible = False

        lblDataBaseDestinazione.Text = ""
        lblDataBasePartenza.Text = ""
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        OpenFileDialog1.FileName = "IsolaEcologica_old.mdb"
        lblDataBasePartenza.Text = ""
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            lblDataBasePartenza.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OpenFileDialog1.FileName = "IsolaEcologica_new.mdb"
        lblDataBaseDestinazione.Text = ""
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            lblDataBaseDestinazione.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim m = New Migrationtool.Database

        m.DestinazioneName = lblDataBaseDestinazione.Text
        m.PartenzaName = lblDataBasePartenza.Text

        Try
            m.ConnectionToStart()
            m.ConnectionToDestination()
            '
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            '
            m.StartMigration()
        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace.ToString, "Migration Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ProgressBar1.Style = ProgressBarStyle.Blocks
        End Try

    End Sub

    Public Sub UpdateLblInfo(ByVal txt As String)
        lblInfo.Text = txt
        Application.DoEvents()
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim m = New Migrationtool.Database

        m.DestinazioneName = lblDataBaseDestinazione.Text
        m.PartenzaName = lblDataBasePartenza.Text

        Try
            m.ConnectionToStart()
            m.ConnectionToDestination()
            '
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            '
            m.GenerateScript()


        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace.ToString, "Migration Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ProgressBar1.Style = ProgressBarStyle.Blocks
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Dim m = New Migrationtool.Database

        m.DestinazioneName = lblDataBaseDestinazione.Text
        m.PartenzaName = lblDataBasePartenza.Text

        Try
            m.ConnectionToStart()
            m.ConnectionToDestination()
            '
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Marquee
            '
            m.AnalisyTables()

            RptViewTableNews.Show()

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + ex.StackTrace.ToString, "Migration Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ProgressBar1.Style = ProgressBarStyle.Blocks
        End Try

    End Sub
End Class
