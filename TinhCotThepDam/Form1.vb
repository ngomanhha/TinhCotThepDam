Imports System.Globalization
Imports System.Net.Http.Headers
Imports System.Security.Cryptography

Public Class frmMain

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

        SYS_List_BeTong = New List(Of BeTong)
        SYS_List_CotThep = New List(Of CotThep)
        SYS_List_MatCat = New List(Of MatCatDam)

        Dim bt1 As New BeTong("B20", 11.5, 27000)
        Dim bt2 As New BeTong("B25", 14.5, 30000)

        SYS_List_BeTong.Add(bt1)
        SYS_List_BeTong.Add(bt2)

        Dim ct1 As New CotThep("CB300-V", 260, 260, 200000)
        Dim ct2 As New CotThep("CB400-V", 350, 350, 200000)

        SYS_List_CotThep.Add(ct1)
        SYS_List_CotThep.Add(ct2)



        For Each betong In SYS_List_BeTong
            cbxBeTong.Items.Add(betong.Name)
        Next
        For Each cothep In SYS_List_CotThep
            cbxCotThep.Items.Add(cothep.Name)
        Next

        cbxBeTong.SelectedIndex = 0
        cbxCotThep.SelectedIndex = 0
    End Sub

    Dim MatCat As MatCatDam

    Private Sub tsRun_Click(sender As Object, e As EventArgs) Handles tsRun.Click
        Dim BeRong, ChieuCao, a, aPhay, M As Double
        Dim Betong As BeTong
        Dim Cotthep As CotThep

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

        For Each bt As BeTong In SYS_List_BeTong
            If bt.Name = cbxBeTong.SelectedItem.ToString Then
                Betong = bt
                Exit For
            End If
        Next

        For Each ct As CotThep In SYS_List_CotThep
            If ct.Name = cbxCotThep.SelectedItem.ToString Then
                Cotthep = ct
                Exit For
            End If
        Next

        Dim tenMC As String = txtName.Text
        MatCat = New MatCatDam(tenMC, Betong, Cotthep, BeRong, ChieuCao, a, aPhay, M)
        MatCat.TinhThep()

        txtAs.Text = Math.Round(MatCat.As1 / 100, 2)
        txtAsPhay.Text = Math.Round(MatCat.AsPhay / 100, 2)


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

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If MatCat Is Nothing Then
            MsgBox("Vui lòng tính toán trước khi thêm vào danh sách !!!")
            Exit Sub
        End If
        SYS_List_MatCat.Add(MatCat)
        HienThiDanhSachMatCat()

    End Sub

    Sub HienThiDanhSachMatCat()
        lstBxMatCat.Items.Clear()
        For Each mc As MatCatDam In SYS_List_MatCat
            lstBxMatCat.Items.Add(mc.Name)
        Next
    End Sub
End Class
