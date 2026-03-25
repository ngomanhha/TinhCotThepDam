Public Class frmMain

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnTinhToan_Click(sender As Object, e As EventArgs) Handles btnTinhToan.Click
        Dim BeRong, ChieuCao, a, aPhay, M As Double
        Dim Betong, cotthep As String

        If IsNumeric(txtBeRong.Text) = False Then
            MsgBox("Vui lòng nhập số vào ô chiều rộng")
            Exit Sub
        End If
        BeRong = Double.Parse(txtBeRong.Text)
        ChieuCao = txtChieuCao.Text
        'a = txtA.Text
        'aPhay = txtAPhay.Text
        'M = txtM.Text * 10 ^ 6

        'Betong = txtBeTong.Text
        'cotthep = txtCotThep.Text

        Dim DiemTichTietDien As Double = BeRong * ChieuCao
        MsgBox(DiemTichTietDien)

    End Sub
End Class
