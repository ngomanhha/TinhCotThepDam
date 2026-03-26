Imports System.Globalization
Imports System.Security.Cryptography

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

        Betong = cbxBeTong.SelectedItem
        cotthep = cbxCotThep.SelectedItem

        Dim As1, AsPhay As Double
        TinhThep(BeRong, ChieuCao, Betong, cotthep, a, aPhay, M, As1, AsPhay)
        txtAs.Text = Math.Round(As1 / 100, 2)
        txtAsPhay.Text = Math.Round(AsPhay / 100, 2)


    End Sub

    Sub TinhThep(b As Double, h As Double, BeTong As String, CotThep As String, a As Double, aPhay As Double, M As Double,
                 ByRef As1 As Double, ByRef AsPhay As Double)
        Dim ho As Double = h - a
        Dim Rb, Rs, Es, Rsc, pxiR, alphaR As Double

        Rb = ReturnRb(BeTong)
        Rs = ReturnRs(CotThep)
        Es = ReturnEs(CotThep)
        Rsc = ReturnRsc(CotThep)

        pxiR = 0.8 / (1 + Rs / (0.0035 * Es))
        alphaR = pxiR * (1 - 0.5 * pxiR)

        Dim alphaM As Double = M / (Rb * b * ho ^ 2)

        If alphaM <= alphaR Then
            Dim pxi As Double = 1 - Math.Sqrt(1 - 2 * alphaM)
            As1 = pxi * Rb * ho * b / Rs
            AsPhay = 0
        Else
            AsPhay = (M - alphaR * Rb * b * ho ^ 2) / (Rsc * (ho - aPhay))
            As1 = (pxiR * Rb * b * ho + Rsc * aPhay) / Rs
        End If
        txtAs.Text = Math.Round(As1 / 100, 2)
        txtAsPhay.Text = Math.Round(AsPhay / 100, 2)

    End Sub

    Function ReturnRb(BeTong As String) As Double
        Dim Rb As Double
        Select Case BeTong
            Case "B20"
                Rb = "11.5"
            Case "B25"
                Rb = "14.5"
            Case "B30"
                Rb = "17.5"
            Case Else
                Rb = 0
        End Select
        Return Rb
    End Function

    Function ReturnRs(CotThep As String) As Double
        Dim Rs As Double
        Select Case CotThep
            Case "CB200-T"
                Rs = 350
            Case "CB300-V"
                Rs = 360
            Case "CB400-V"
                Rs = 490
            Case Else
                Rs = 0
        End Select
        Return Rs
    End Function

    Function ReturnEs(CotThep As String) As Double
        Dim Es As Double
        Select Case CotThep
            Case "CB200-T"
                Es = 200000
            Case "CB300-V"
                Es = 200000
            Case "CB400-V"
                Es = 200000
            Case Else
                Es = 0
        End Select
        Return Es
    End Function

    Function ReturnRsc(CotThep As String) As Double
        Dim Rsc As Double

        Select Case CotThep
            Case "CB200-T"
                Rsc = 350
            Case "CB300-V"
                Rsc = 360
            Case "CB400-V"
                Rsc = 490
            Case Else
                Rsc = 0
        End Select
        ReturnRsc = Rsc
    End Function



    Private Sub txtBeRong_TextChanged(sender As Object, e As EventArgs) Handles txtBeRong.TextChanged
        txtAs.Text = ""
        txtAsPhay.Text = ""
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim DsBeTong As List(Of String) = New List(Of String)({"B20", "B25", "B30"})
        Dim DsCotThep As List(Of String) = New List(Of String)({"CB240-T", "CB300-V", "CB400-V"})


        For Each item As String In DsBeTong
            cbxBeTong.Items.Add(item)
        Next
        For Each item As String In DsCotThep
            cbxCotThep.Items.Add(item)
        Next

        cbxBeTong.SelectedIndex = 0
        cbxCotThep.SelectedIndex = 0



    End Sub
End Class
