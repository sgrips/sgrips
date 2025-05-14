Public Class RptViewTableNews

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Report.TabelleNuove' table. You can move, or remove it, as needed.
        Me.DataTable1TableAdapter.Connection.ConnectionString = My.Settings.ReportConnectionString
        Me.DataTable1TableAdapter.Connection.Open()


        Me.DataTable1TableAdapter.Fill(Me.ReportDataSet.DataTable1)
        
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class