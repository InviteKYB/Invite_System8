Imports Frame8
Imports Base8
Imports Base8.Shared
Imports System8.ENV

Public Class GWM100_VIEW

    Private Sub Me_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        g10.ShowRowHeaders = False
        g10.ShowColumnHeaders = False
        g10.AllowAddRows = True
        g40.ShowRowHeaders = False
        g40.ShowColumnHeaders = False
        g40.AllowAddRows = True
        g50.ShowRowHeaders = False
        g50.ShowColumnHeaders = False
        g50.AllowAddRows = True
        g60.ShowRowHeaders = False
        g60.ShowColumnHeaders = False
        g60.AllowAddRows = True
        g70.ShowRowHeaders = False
        g70.ShowColumnHeaders = False
        g70.AllowAddRows = True

        If Comp_Type() = "HK" Then 'HK인경우 폼로드에도 숨김항목 처리해야한다. 21.03.17 테스트 변경
            btn_dmb100.Visible = False '파트리스트버튼
            btn_pms300_jump.Visible = False '유효성버튼
            btnClose.Visible = False

            '구매의뢰출력관련
            Label2.Visible = False
            btn_prt2.Visible = False
            g31.Visible = False

        End If

    End Sub

    Public Overrides Sub MenuButton_Click(ByVal mty As MenuType)
        Select Case mty
            Case MenuType.Open

            Case MenuType.Save

            Case MenuType.Delete

            Case MenuType.New

            Case Else
                MyBase.MenuButton_Click(mty)
        End Select

    End Sub

    Private Sub PrintPreviewControl1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      
    End Sub

    Private Sub btnAppr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAppr.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub btnDeputy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeputy.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub btnAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAll.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub EPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles EPanel1.Paint

    End Sub

    Private Sub g20_AfterMoveRow(ByVal sender As System.Object, ByVal PrevRowIndex As System.Int32, ByVal RowIndex As System.Int32) Handles g20.AfterMoveRow

    End Sub

    Private Sub TAB1_Selected(ByVal sender As System.Object, ByVal e As DevExpress.XtraTab.TabPageEventArgs) Handles TAB1.Selected

    End Sub

    Private Sub btn_dmb100_Click(sender As System.Object, e As System.EventArgs) Handles btn_dmb100.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub
    Private Sub btn_pms300_jump_Click(sender As System.Object, e As System.EventArgs) Handles btn_pms300_jump.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    '이전 결재자 대결 (2020-09-14)
    Private Sub DeputyMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeputyMenuItem1.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    '다음 결재자 대결 (2020-09-14)
    Private Sub DeputyMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeputyMenuItem2.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub

    Private Sub TAB1_Click(sender As System.Object, e As System.EventArgs) Handles TAB1.Click

    End Sub

    '출력2 버튼 (HL 전용) (2021-02-15) 
    Private Sub btn_prt2_Click(sender As System.Object, e As System.EventArgs) Handles btn_prt2.Click
        'PopUp_VIEW에 코딩되어있음
    End Sub
End Class
