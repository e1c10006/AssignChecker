Imports System.Drawing
Public Class SettingForm
    Private TaikbCtls As Label()
    Private SizekbCtls As Label()
    Private EsckbCtls As Label()
    Private IconkbCtls As Label()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Settings("MaxView") = MaxView.Text
        My.Settings("TaiKb") = SUBKB.Text
        My.Settings("SizeKb") = SIZEKB.Text
        If SIZEKB.Text = "1" Then
            My.Settings("UserFormSizeX") = H_WIDTH.Text
            My.Settings("UserFormSizeY") = H_HEIGHT.Text
        Else
            My.Settings("DefaultFormSizeX") = H_WIDTH.Text
            My.Settings("DefaultFormSizeY") = H_HEIGHT.Text
        End If
        My.Settings("DefaultStyle") = テーマ.Text
        My.Settings("EscClose") = ESCKB.Text

        If ICONKB.Text = "1" Then
            My.Settings("IconVisible") = True
        Else
            My.Settings("IconVisible") = False
        End If
        Me.Close()
    End Sub

    Private Sub SettingForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TaikbCtls = {H_SUBKBA, H_SUBKBB, H_SUBKBC, H_SUBKBD, H_SUBKBE}
        SizekbCtls = {H_SIZEKBA, H_SIZEKBB}
        EsckbCtls = {H_ESCKBA, H_ESCKBB}
        IconkbCtls = {H_ICONKBA, H_ICONKBB}

        Select Case My.Settings("DefaultStyle")
            Case "Silver"
                Me.Style = MetroFramework.MetroColorStyle.Silver
            Case "Black"
                Me.Style = MetroFramework.MetroColorStyle.Black
            Case "Orange"
                Me.Style = MetroFramework.MetroColorStyle.Orange
            Case "Purple"
                Me.Style = MetroFramework.MetroColorStyle.Purple
            Case "Lime"
                Me.Style = MetroFramework.MetroColorStyle.Lime
            Case "Green"
                Me.Style = MetroFramework.MetroColorStyle.Green
            Case "Blue", "Default"
                Me.Style = MetroFramework.MetroColorStyle.Default
        End Select

        MaxView.Text = My.Settings("MaxView")
        SUBKB.Text = My.Settings("TaiKb")
        SIZEKB.Text = My.Settings("SizeKb")
        If SIZEKB.Text = "0" Then
            H_WIDTH.Text = My.Settings("DefaultFormSizeX")
            H_HEIGHT.Text = My.Settings("DefaultFormSizeY")
        Else
            H_WIDTH.Text = My.Settings("UserFormSizeX")
            H_HEIGHT.Text = My.Settings("UserFormSizeY")
        End If
        ESCKB.Text = My.Settings("EscClose")
        テーマ.Text = My.Settings("DefaultStyle")
        If DirectCast(My.Settings("IconVisible"), Boolean) Then
            ICONKB.Text = "1"
        Else
            ICONKB.Text = "0"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub SettingForm_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                If My.Settings("EscClose").ToString.TrimEnd() = "1" Then
                    Me.Close()
                End If
        End Select
    End Sub

#Region "サブフォーム表示区分"""
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) _
        Handles SUBKB.TextChanged

        Select Case SUBKB.Text.ToString
            Case "0"
                ColorChange(H_SUBKBA, TaikbCtls)
            Case "1"
                ColorChange(H_SUBKBB, TaikbCtls)
            Case "2"
                ColorChange(H_SUBKBC, TaikbCtls)
            Case "3"
                ColorChange(H_SUBKBD, TaikbCtls)
            Case "4"
                ColorChange(H_SUBKBE, TaikbCtls)
        End Select

    End Sub

    Private Sub SUBKBA_Click(sender As Object, e As EventArgs) Handles H_SUBKBA.Click
        SUBKB.Text = "0"
        ColorChange(H_SUBKBA, TaikbCtls)
    End Sub

    Private Sub SUBKBB_Click(sender As Object, e As EventArgs) Handles H_SUBKBB.Click
        SUBKB.Text = "1"
        ColorChange(H_SUBKBB, TaikbCtls)
    End Sub

    Private Sub SUBKBC_Click(sender As Object, e As EventArgs) Handles H_SUBKBC.Click
        SUBKB.Text = "2"
        ColorChange(H_SUBKBC, TaikbCtls)
    End Sub

    Private Sub SUBKBD_Click(sender As Object, e As EventArgs) Handles H_SUBKBD.Click
        SUBKB.Text = "3"
        ColorChange(H_SUBKBD, TaikbCtls)
    End Sub

    Private Sub SUBKBE_Click(sender As Object, e As EventArgs) Handles H_SUBKBE.Click
        SUBKB.Text = "4"
        ColorChange(H_SUBKBE, TaikbCtls)
    End Sub
#End Region

#Region "画面サイズ"
    Private Sub SIZEKB_TextChanged(sender As Object, e As EventArgs) Handles SIZEKB.TextChanged
        If SIZEKB.Text.ToString.TrimEnd = "0" Then
            ColorChange(H_SIZEKBA, SizekbCtls)
        Else
            ColorChange(H_SIZEKBB, SizekbCtls)
        End If
    End Sub
    Private Sub LBL3_Click(sender As Object, e As EventArgs) Handles H_SIZEKBA.Click
        SIZEKB.Text = "0"
        H_WIDTH.Text = My.Settings("DefaultFormSizeX")
        H_HEIGHT.Text = My.Settings("DefaultFormSizeY")
        ColorChange(H_SIZEKBA, SizekbCtls)
    End Sub
    Private Sub LBL4_Click(sender As Object, e As EventArgs) Handles H_SIZEKBB.Click
        SIZEKB.Text = "1"
        H_WIDTH.Text = My.Settings("UserFormSizeX")
        H_HEIGHT.Text = My.Settings("UserFormSizeY")
        ColorChange(H_SIZEKBB, SizekbCtls)
    End Sub
#End Region

#Region "Escapeキーで終了"
    Private Sub ESCKB_TextChanged(sender As Object, e As EventArgs) Handles ESCKB.TextChanged, ESCKB.TextChanged
        If ESCKB.Text.ToString.TrimEnd = "0" Then
            ColorChange(H_ESCKBA, EsckbCtls)
        Else
            ColorChange(H_ESCKBB, EsckbCtls)
        End If
    End Sub
    Private Sub ESCKBA_Click(sender As Object, e As EventArgs) Handles H_ESCKBA.Click
        ESCKB.Text = "0"
        ColorChange(H_ESCKBA, EsckbCtls)
    End Sub
    Private Sub ESCKBB_Click(sender As Object, e As EventArgs) Handles H_ESCKBB.Click
        ESCKB.Text = "1"
        ColorChange(H_ESCKBB, EsckbCtls)
    End Sub
#End Region

#Region "Icon表示"
    Private Sub ICONKB_TextChanged(sender As Object, e As EventArgs) Handles ICONKB.TextChanged
        If ICONKB.Text.ToString.TrimEnd = "0" Then
            ColorChange(H_ICONKBA, IconkbCtls)
        Else
            ColorChange(H_ICONKBB, IconkbCtls)
        End If
    End Sub
    Private Sub ICONKBA_Click(sender As Object, e As EventArgs) Handles H_ICONKBA.Click
        ICONKB.Text = "0"
        ColorChange(H_ICONKBA, ICONKBCtls)
    End Sub
    Private Sub ICONKBB_Click(sender As Object, e As EventArgs) Handles H_ICONKBB.Click
        ICONKB.Text = "1"
        ColorChange(H_ICONKBB, ICONKBCtls)
    End Sub
#End Region


    Private Sub ColorChange(ByRef Control As Label, ByVal Ctls As Label())
        For Each Ctl As Label In Ctls
            If Ctl IsNot Nothing Then
                Ctl.ForeColor = Color.Black
                Ctl.BackColor = Color.FromArgb(240, 240, 240)
            End If
        Next
        Control.ForeColor = Color.FromArgb(240, 240, 240)
        Control.BackColor = Color.DarkBlue
    End Sub


End Class