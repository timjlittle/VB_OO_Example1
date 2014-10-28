Public Class frmProduct

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim prod As Product

        prod = New Product(1)

        txtDescription.Text = prod.Description
        txtPrice.Text = prod.PriceExVat.ToString

    End Sub
End Class