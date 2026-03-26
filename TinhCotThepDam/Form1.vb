Imports System.Globalization

Public Class frmMain

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnTinhToan_Click(sender As Object, e As EventArgs) Handles btnTinhToan.Click
        Dim BeRong, ChieuCao, a, aPhay, M As Double
        Dim Betong, cotthep As String

        If IsNumeric(txtBeRong.Text) = True Then
            BeRong = Double.Parse(txtBeRong.Text)
        Else
            MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
            Exit Sub
        End If

        If IsNumeric(txtChieuCao.Text) = True Then
            ChieuCao = Double.Parse(txtChieuCao.Text)
        Else
            MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
            Exit Sub
        End If

        If IsNumeric(txtA.Text) = True Then
            a = Double.Parse(txtA.Text)
        Else
            MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
            Exit Sub
        End If

        If IsNumeric(txtAPhay.Text) = True Then
            aPhay = Double.Parse(txtAPhay.Text)
        Else
            MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
            Exit Sub
        End If

        If IsNumeric(txtM.Text) = True Then
            M = Double.Parse(txtM.Text * 10 ^ 6)
        Else
            MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
            Exit Sub
        End If

        Betong = txtBeTong.Text
        cotthep = txtCotThep.Text
        Dim ho As Double = ChieuCao - a
        Dim Rb, Rs, Es, Rsc, pxiR, alphaR As Double

        Select Case Betong
            Case "B10"
                Rb = 11.5
            Case "B15"
                Rb = 14.5
            Case "B20"
                Rb = 11.5
            Case "B25"
                Rb = 14.5
            Case "B30"
                Rb = 17.5
            Case "B35"
                Rb = 20.0
            Case "B40"
                Rb = 22.0
            Case Else
                MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
                Exit Sub
        End Select

        'If Betong = "B20" Then
        '    Rb = 11.5
        'End If
        'If cotthep = "CB400-V" Then
        '    Rs = 350
        '    Es = 200000
        '    Rsc = 350

        'End If
        Select Case cotthep
            Case "CB400-V"
                Rs = 350
                Es = 200000
                Rsc = 350
            Case "CB400-B"
                Rs = 360
                Es = 200000
                Rsc = 360
            Case "CB500-V"
                Rs = 490
                Es = 200000
                Rsc = 490
            Case "CB500-B"
                Rs = 510
                Es = 200000
                Rsc = 510
            Case Else
                MsgBox("Kiểm tra lại dữ liệu đầu vào !!!")
                Exit Sub
        End Select
        pxiR = 0.8 / (1 + Rs / (0.0035 * Es))
        alphaR = pxiR * (1 - 0.5 * pxiR)

        Dim alphaM As Double = M / (Rb * BeRong * ho ^ 2)
        Dim As1, AsPhay As Double
        If alphaM <= alphaR Then
            Dim pxi As Double = 1 - Math.Sqrt(1 - 2 * alphaM)
            As1 = pxi * Rb * ho * BeRong / Rs
            AsPhay = 0
        Else
            AsPhay = (M - alphaR * Rb * BeRong * ho ^ 2) / (Rsc * (ho - aPhay))
            As1 = (pxiR * Rb * BeRong * ho + Rsc * aPhay) / Rs
        End If
        txtAs.Text = Math.Round(As1 / 100, 2)
        txtAsPhay.Text = Math.Round(AsPhay / 100, 2)




    End Sub

    Private Sub txtBeRong_TextChanged(sender As Object, e As EventArgs) Handles txtBeRong.TextChanged
        txtAs.Text = ""
        txtAsPhay.Text = ""
    End Sub
End Class
