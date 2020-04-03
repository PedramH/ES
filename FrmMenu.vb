Public Class FrmMenu
    Private Sub FrmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        M_user_info.Text = loggedInUserName

        '' Disable Features that are not yes implemented
        M_production_plan.Enabled = False
        M_production_data.Enabled = False
        M_production_shippment.Enabled = False
        '' ---------------------------------------------

        '' Add emkansanji state parirs to the dictionary    



        HandleUserPermissions()
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
        Dim f As New emkanSanjiForm
        Select Case loggedInUserGroup
            Case "Tolid1"
                f.CBOrderState.Text = "امکان سنجی اولیه تولید"
            Case "Tolid2"
                f.CBOrderState.Text = "امکان سنجی نهایی تولید"
            Case "QA"
                f.CBOrderState.Text = "امکان سنجی نهایی کیفی"
        End Select
        f.Show()
    End Sub

    Private Sub M_emkansanji_new_Click(sender As Object, e As EventArgs) Handles M_emkansanji_new.Click
        FrmNewEmkansanji.Show()
    End Sub

    Private Sub M_production_orders_Click(sender As Object, e As EventArgs) Handles M_production_orders.Click
        '' This shows only the orders that are validated to produce
        Dim f As New emkanSanjiForm
        f.CBOrderState.Text = "تایید شده"
        f.Show()
        '' TODO
    End Sub

    Private Sub M_prodcution_wireInventory_Click(sender As Object, e As EventArgs) Handles M_prodcution_wireInventory.Click
        Dim f As New wires
        f.Show()
    End Sub

    Private Sub M_customers_list_Click(sender As Object, e As EventArgs) Handles M_customers_list.Click
        Dim f As New FrmNewEmkansanji
        f.formState = "customerSearch"
        f.Show()
        '' TODO
    End Sub

    Private Sub M_customer_new_Click(sender As Object, e As EventArgs) Handles M_customer_new.Click
        Dim f As New customerForm
        customerFormState = "new"
        f.Show()
    End Sub

    Private Sub M_products_new_Click(sender As Object, e As EventArgs) Handles M_products_new.Click
        Dim f As New productForm
        productFormState = "new"
        f.Show()
    End Sub




    '' ----------------------------- Permission handling -----------------------------------------
    Private Sub HandleUserPermissions()
        If loggedInUserGroup <> "Admin" And loggedInUserGroup <> "QC" Then
            M_emkansanji_new.Enabled = False
            M_products_new.Enabled = False
        End If
        If loggedInUserGroup <> "Admin" And loggedInUserGroup <> "Tolid1" And loggedInUserGroup <> "Tolid2" Then
            ' M_warehouse_updateInventory.Enabled = False

        End If
        If loggedInUserGroup <> "Admin" And loggedInUserGroup <> "Anbar" And loggedInUserGroup <> "Tolid1" And loggedInUserGroup <> "Tolid2" Then
            M_warehouse_productCode.Enabled = False
        End If
        If loggedInUserGroup = "Anbar" Then
            Menu_emkansanji.Visible = False
            Menu_production.Visible = False
            Menu_customers.Visible = False
            Menu_products.Visible = False

        End If

    End Sub

    Private Sub M_products_list_Click(sender As Object, e As EventArgs) Handles M_products_list.Click
        Dim frmProductList As New FrmNewEmkansanji
        frmProductList.formState = "productSearch"
        frmProductList.Show()
    End Sub

    Private Sub M_warehouse_mandrels_Click(sender As Object, e As EventArgs) Handles M_warehouse_mandrels.Click
        Dim f As New mandrels
        f.Show()
    End Sub

    Private Sub M_warehouse_wires_Click(sender As Object, e As EventArgs) Handles M_warehouse_wires.Click
        Dim f As New wires
        f.Show()
    End Sub

    Private Sub M_warehouse_productCode_Click(sender As Object, e As EventArgs) Handles M_warehouse_productCode.Click
        Dim f As New FrmNewEmkansanji
        f.formState = "newProductCode"
        f.Show()
    End Sub
End Class