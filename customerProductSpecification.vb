Imports System.ComponentModel

Public Class customerProductSpecification
    Public colA As New Collection
    Public thisFormsCaller As FrmNewEmkansanji
    Public thisforms_newcaller As emkanSanjiForm
    Private Sub customerProductSpecification_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        colA.Add(CBMaterial)
        colA.Add(TBWireDiameter)
        colA.Add(TBdTol)
        colA.Add(TBOD)
        colA.Add(TBODTol)
        colA.Add(TBDi)
        colA.Add(TBDiTol)
        colA.Add(TBNt)
        colA.Add(TBNtTol)
        colA.Add(TBNActive)
        colA.Add(TBL0)
        colA.Add(TBL0Tol)
        colA.Add(CBCoilingDirection)
        colA.Add(TBSpringRate)
        colA.Add(TBRateTol)
        colA.Add(CBScoilType)
        colA.Add(CBEcoilType)
        colA.Add(TBF1)
        colA.Add(TBL1)
        colA.Add(TBF1Tol)
        colA.Add(TBF2)
        colA.Add(TBL2)
        colA.Add(TBF2Tol)
        colA.Add(TBF3)
        colA.Add(TBL3)
        colA.Add(TBF3Tol)
        colA.Add(TBUnit)
        colA.Add(TBTooli)
        colA.Add(TBghotri)
        colA.Add(TBMinHardness)
        colA.Add(TBMaxHardness)
        colA.Add(TBhardnessUnit)
        For i = 1 To colA.Count()
            colA(i).Tabindex = i
        Next

        ParseCustomerProductSpec(colA, TBCustomerProductSpec.Text)

    End Sub

    Public Function ParseCustomerProductSpec(c As Collection, productSpec As String)
        '' Gets a Code and Fill the Boxes
        If Len(productSpec) = 0 Then
            Exit Function
        End If
        Dim productSpecArray As String() = productSpec.Split(New Char() {"|"c})
        On Error Resume Next
        Dim i As Integer
        For i = 1 To c.Count()
            c(i).text = productSpecArray(i - 1)
        Next
    End Function

    Public Function GenerateCutsomerProductSpec(c As Collection)
        '' Gets condition of checkboxes and generates a code
        On Error Resume Next
        Dim productSpec As String = ""
        Dim i As Integer
        productSpec = c(1).text
        For i = 2 To c.Count()
            productSpec = productSpec & "|" & c(i).text
        Next
        Return productSpec
    End Function

    Private Sub BTSave_Click(sender As Object, e As EventArgs) Handles BTSave.Click
        TBCustomerProductSpec.Text = GenerateCutsomerProductSpec(colA)
        On Error Resume Next
        thisFormsCaller.TBCustomerProductSpec.Text = TBCustomerProductSpec.Text
        thisforms_newcaller.TBCustomerProductSpec.Text = TBCustomerProductSpec.Text
        Me.Close()
    End Sub
End Class