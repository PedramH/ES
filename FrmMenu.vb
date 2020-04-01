Public Class FrmMenu
    Private Sub FrmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        M_user_info.Text = loggedInUserName
    End Sub

    Private Sub M_user_logout_Click(sender As Object, e As EventArgs) Handles M_user_logout.Click
        My.Settings.loggedin = ""
        My.Settings.loginDate = New System.DateTime(1900, 1, 1, 12, 0, 0)
        My.Settings.validation = ""
        My.Settings.usersName = ""
        My.Settings.userGroup = ""
        LoginForm.Show()
        Me.Close()
    End Sub

    Private Sub M_emkansanji_orders_Click(sender As Object, e As EventArgs) Handles M_emkansanji_orders.Click
        emkanSanjiForm.Show()
    End Sub

    Private Sub M_emkansanji_new_Click(sender As Object, e As EventArgs) Handles M_emkansanji_new.Click
        FrmNewEmkansanji.Show()
    End Sub

    Private Sub M_production_orders_Click(sender As Object, e As EventArgs) Handles M_production_orders.Click
        emkanSanjiForm.Show()
        '' TODO
    End Sub

    Private Sub M_prodcution_wireInventory_Click(sender As Object, e As EventArgs) Handles M_prodcution_wireInventory.Click
        wires.Show()
    End Sub

    Private Sub M_customers_list_Click(sender As Object, e As EventArgs) Handles M_customers_list.Click
        '' TODO
    End Sub

    Private Sub M_customer_new_Click(sender As Object, e As EventArgs) Handles M_customer_new.Click
        customerFormState = "new"
        customerForm.Show()
    End Sub

    Private Sub M_products_new_Click(sender As Object, e As EventArgs) Handles M_products_new.Click
        productFormState = "new"
        productForm.Show()
    End Sub
    Public wireForm As New FrmNewEmkansanji


End Class