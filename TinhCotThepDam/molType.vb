Module molType
    '1. mặt cắt dầm
    ' - Tên
    ' - Bê tông ...

    Public Class MatCatDam
        Public Name As String
        Public BeTong As BeTong
        Public Thep As CotThep
        Public B As Double
        Public H As Double
        Public A As Double
        Public APhay As Double
        Public M As Double

        Public As1 As Double
        Public AsPhay As Double

        Sub New(_name As String, _bt As BeTong, _cotthep As CotThep, _b As Double, _a As Double, _aphay As Double, _h As Double, _m As Double)
            Name = _name
            BeTong = _bt
            Thep = _cotthep
            B = _b
            A = _a
            APhay = _aphay
            H = _h
            M = _m

        End Sub

        Sub TinhThep()
            Dim ho As Double = H - A
            Dim Rb, Rs, Es, Rsc, pxiR, alphaR As Double

            Rb = BeTong.Rb
            Rs = Thep.Rs
            Es = Thep.Es
            Rsc = Thep.Rsc

            pxiR = 0.8 / (1 + Rs / (0.0035 * Es))
            alphaR = pxiR * (1 - 0.5 * pxiR)

            Dim alphaM As Double = M / (Rb * B * ho ^ 2)

            If alphaM <= alphaR Then
                Dim pxi As Double = 1 - Math.Sqrt(1 - 2 * alphaM)
                As1 = pxi * Rb * ho * B / Rs
                AsPhay = 0
            Else
                AsPhay = (M - alphaR * Rb * B * ho ^ 2) / (Rsc * (ho - APhay))
                As1 = (pxiR * Rb * B * ho + Rsc * APhay) / Rs
            End If

        End Sub
    End Class



    '2. Bê tông
    'Rb...
    Public Class BeTong
        Public Name As String
        Public Rb As Double
        Public Eb As Double

        Sub New(_name As String, _rb As Double, _eb As Double)
            Name = _name
            Rb = _rb
            Eb = _eb
        End Sub

        Sub KhaiBaoTenVaCuongDo()
            MsgBox("tên tôi là:" & Name & ", cường độ là:" & Rb)
        End Sub

    End Class

    '3. Cốt thép
    'Rs, Es, Rsc...
    Public Class CotThep
        Public Name As String
        Public Rs As Double
        Public Es As Double
        Public Rsc As Double

        Sub New(_name As String, _rs As Double, _es As Double, _rsc As Double)
            Name = _name
            Rs = _rs
            Es = _es
            Rsc = _rsc
        End Sub

    End Class
End Module
