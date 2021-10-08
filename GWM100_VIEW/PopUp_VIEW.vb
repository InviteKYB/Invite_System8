Imports Frame8
Imports Base8
Imports Base8.Shared
Imports System8.ENV
Imports Base8.Parameter
Imports System.Data
Imports System8

Imports System.Net.Mail
Imports System.IO
'Imports CrystalDecisions.CrystalReports.Engine
'Imports CrystalDecisions.Shared
'Imports CrystalDecisions.ReportSource
Imports System.Runtime.CompilerServices
Imports System.Net
Imports System.Net.Configuration
Imports System.Windows
Imports System.Reflection '2020-05-13. YANG 추가.
Imports System.Text.RegularExpressions

Public Class PopUp_VIEW

    Private popup As GWM100_VIEW

    Public Appr_No As String
    Public Appr_Tab As Long
    Public Appr_Ref As String
    Public Appr_All As Boolean

    Public btn_ReadOnly As Boolean = False

    Private Form_Cd As String
    Private Ftp_Form As String
    Private Ref_No1 As String
    Private Ref_No2 As String
    Private Ref_No3 As String
    Private Ref_No4 As String
    Private Ftp_ID As Long
    Private Appr_Sw As String
    Private Appr_Bc As String
    Public Appr_Sort As String

    Private Appr_Self As Boolean = False    '자가결재여부

    Dim re_emp As String
    Dim re_chk As String = "0"
    Dim msql As String = ""
    Dim dSet As Data.DataSet = Nothing
    Dim _chk_dept As String = ""
    Dim _DeputyBC As String = ""            '대결 구분

    Public Dev_CD As String
    Public Dev_dSet As DataSet
    Public Dev_dSet_Sub As DataSet = Nothing

    Private Call_Form_Cd As Object

    Private Sub PopUp_VIEW_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim p As New OpenParameters

        p.Add("@appr_no", Appr_No)

        Dim dSet As DataSet = popup.OpenDataSet("gwm100_view_info1", p)

        If IsEmpty(dSet) Then
            MessageInfo("[결재] 마스터정보 오류!!!")
            Me.Close()
        End If

        ' 결재마스터정보(전역변수)
        '================================================
        Form_Cd = ToStr(dSet.Tables(0).Rows(0).Item("form_cd"))
        Ref_No1 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No1"))
        Ref_No2 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No2"))
        Ref_No3 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No3"))
        Ref_No4 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No4"))
        Ftp_ID = ToStr(dSet.Tables(0).Rows(0).Item("Ftp_ID"))
        '================================================

        ' 결재정보
        '===============================================================
        popup.appr_no.Text = ToStr(dSet.Tables(0).Rows(0).Item("appr_no"))
        popup.cdt.Text = ToStr(dSet.Tables(0).Rows(0).Item("cdt"))
        popup.cidNm.Text = ToStr(dSet.Tables(0).Rows(0).Item("cidNm"))
        popup.dept_nm.Text = ToStr(dSet.Tables(0).Rows(0).Item("dept_nm"))
        popup.title.Text = ToStr(dSet.Tables(0).Rows(0).Item("title"))

        ' 결재버튼잠금설정
        '===============================================================
        If btn_ReadOnly = True Then
            popup.btnAppr.Enabled = False
            popup.btnReturn.Enabled = False
            popup.btnDeputy.Enabled = False
            popup.btnAll.Enabled = False
        Else
            popup.btnAppr.Enabled = True
            popup.btnReturn.Enabled = True
            popup.btnDeputy.Enabled = True
            popup.btnAll.Enabled = True
        End If

        Me.Text = ToStr(dSet.Tables(0).Rows(0).Item("title"))

        Dim agree_Cnt As Long
        Dim ref_Cnt As Long

        agree_Cnt = 0
        ref_Cnt = 0

        For Each dRow In dSet.Tables(0).Rows
            Select Case ToStr(dRow("appr_bc"))
                Case "GW100300"
                    If agree_Cnt = 0 Then
                        popup.txt_agree.Text = ToStr(dRow("nm"))
                    End If

                    agree_Cnt = agree_Cnt + 1

                Case "GW100500"
                    If ref_Cnt = 0 Then
                        popup.txt_ref.Text = ToStr(dRow("nm"))
                    End If

                    ref_Cnt = ref_Cnt + 1
            End Select
        Next

        If agree_Cnt > 1 Then
            popup.txt_agree.Text = popup.txt_agree.Text & " 외 " & CStr(agree_Cnt - 1)
        End If

        If ref_Cnt > 1 Then
            popup.txt_ref.Text = popup.txt_ref.Text & " 외 " & CStr(ref_Cnt - 1)
        End If
        '===============================================================

        Dim p1 As New OpenParameters

        If Appr_All = True Then

            popup.btnAppr.Text = "확인"
            popup.btnAppr.Enabled = False
            popup.btnReturn.Enabled = False
            popup.btnDeputy.Enabled = False
            popup.btnAll.Enabled = False

            popup.btn_dmb100.Visible = False

            popup.btn_pms300_jump.Visible = False ' 2020-05-13. yang 추가

        Else
            p1.Add("@appr_no", Appr_No)
            p1.Add("@appr_chrg", Parameter.Login.RegId)

            dSet = popup.OpenDataSet("gwm100_view_info2", p1)

            If IsEmpty(dSet) Then
                MessageInfo("[결재] 상세정보 오류!!!")
                Me.Close()
            End If

            ' 결재상세정보(전역변수)
            '================================================
            Appr_Sw = ToStr(dSet.Tables(0).Rows(0).Item("appr_sw"))
            Appr_Bc = ToStr(dSet.Tables(0).Rows(0).Item("appr_bc"))
            Appr_Sort = ToStr(dSet.Tables(0).Rows(0).Item("appr_sort"))
            '================================================

            If dSet.Tables(0).Rows.Count >= 2 Then
                Appr_Self = True
            End If

            ' 버튼상태
            '================================================
            Select Case Appr_Sw
                Case "BC210100" '등록상태
                    Select Case Appr_Bc
                        Case "GW100500" '참조
                            ' 참조일경우 확인 버튼만..
                            popup.btnAppr.Text = "확인"
                            popup.btnReturn.Enabled = False
                            popup.btnDeputy.Enabled = False
                            popup.btnAll.Enabled = False
                    End Select

                Case "BC210300"         '승인
                    Select Case Appr_Bc
                        Case "GW100500" '참조
                            ' 참조일경우 확인 버튼만..
                            popup.btnAppr.Text = "확인"
                            popup.btnAppr.Enabled = False
                            popup.btnReturn.Enabled = False
                            popup.btnDeputy.Enabled = False
                            popup.btnAll.Enabled = False

                        Case Else
                            If Comp_Type() = "SSP" Then
                                '받은결재함에서 열었을경우 대결,전결로 미리 승인처리가 되어 있을 때
                                '결재자, 실제결재자가 다르면 확인버튼만 활성화 (2019-02-13 이재일)
                                If Appr_Tab = 1 Then
                                    Dim sSql As String = ""
                                    Dim appr_chrg As String = ""
                                    Dim real_chrg As String = ""

                                    sSql = "select appr_chrg, real_chrg  from GWM110 " & _
                                          "where appr_no = '" & Appr_No & "' " & _
                                          "and appr_chrg = " & Login.RegId
                                    dSet = Link.ExcuteQuery(sSql)

                                    If IsEmpty(dSet) = False Then
                                        appr_chrg = DataValue(dSet, "appr_chrg")
                                        real_chrg = DataValue(dSet, "real_chrg")
                                    End If

                                    If appr_chrg <> real_chrg Then
                                        popup.btnAppr.Text = "확인"
                                        popup.btnAppr.Enabled = True
                                        popup.btnReturn.Enabled = False
                                        popup.btnDeputy.Enabled = False
                                        popup.btnAll.Enabled = False
                                    End If
                                Else
                                    popup.btnAppr.Enabled = False
                                    popup.btnReturn.Enabled = False
                                    popup.btnDeputy.Enabled = False
                                    popup.btnAll.Enabled = False
                                End If
                            Else
                                popup.btnAppr.Enabled = False
                                popup.btnReturn.Enabled = False
                                popup.btnDeputy.Enabled = False
                                popup.btnAll.Enabled = False
                            End If

                            ' 보낸결재함에서 열었을경우 이전결재자가 결재하지 않았으면 대결 활성화
                            If Appr_Tab = 2 Then

                                Dim p3 As New OpenParameters

                                p3.Add("@appr_no", Appr_No)
                                p3.Add("@appr_sort", Appr_Sort)

                                dSet = popup.OpenDataSet("gwm100_view_next", p3)

                                If IsEmpty(dSet) <> True Then
                                    If dSet.Tables(0).Rows(0).Item("appr_sw") = "BC210100" Then
                                        popup.btnDeputy.Enabled = True
                                    End If
                                End If
                            End If
                    End Select

                Case Else       '반려
                    ' 반려는 닫기 버튼만...
                    popup.btnAppr.Enabled = False
                    popup.btnReturn.Enabled = False
                    popup.btnDeputy.Enabled = False
                    popup.btnAll.Enabled = False
            End Select

            '우리엠텍 -> 자가결재는 반려 버튼 비활성화 (2021-08-11)
            If Comp_Type() = "WR" And Appr_Self = True Then
                popup.btnReturn.Enabled = False
            End If
        End If

        '================================
        ' 결재라인
        '================================
        popup.Open("gwm100_view_g40")
        popup.Open("gwm100_view_g50")
        popup.Open("gwm100_view_g60")
        popup.Open("gwm100_view_g70")
        popup.Open("gwm100_view_g10")

        popup.g40.AddNewRow()
        popup.g50.AddNewRow()
        popup.g60.AddNewRow()
        popup.g70.AddNewRow()
        popup.g10.AddNewRow()
        popup.g10.RowHeight = 54

        Dim p5 As New OpenParameters
        Dim tCol As Long

        p5.Add("@appr_no", Appr_No)

        dSet = popup.OpenDataSet("gwm100_view_line", p5)

        If IsEmpty(dSet) Then
            MessageInfo("[결재] 결재라인 오류!!!")
            Me.Close()
        End If

        Dim timg As eImage = Nothing

        tCol = 1
        For Each dRow In dSet.Tables(0).Rows

            popup.g40.Text("c" & tCol, 0) = ToStr(dRow("title"))
            popup.g50.Text("c" & tCol, 0) = ToStr(dRow("nm"))

            Select Case tCol
                Case 1
                    timg = popup.img1
                Case 2
                    timg = popup.img2
                Case 3
                    timg = popup.img3
                Case 4
                    timg = popup.img4
                Case 5
                    timg = popup.img5
                Case 6
                    timg = popup.img6
                Case 7
                    timg = popup.img7
            End Select

            Select Case ToStr(dRow("appr_sw"))
                Case "BC210300" '승인
                    Select Case ToStr(dRow("deputy_bc"))
                        Case "GW110100" '전결
                            If IsDBNull(dRow("img4")) = False Then
                                Dim b1() As Byte = dRow("img4")
                                Dim st1 As New System.IO.MemoryStream(b1)

                                timg.Image = System.Drawing.Image.FromStream(st1)
                            End If

                        Case ("GW110200") '대결
                            If IsDBNull(dRow("img3")) = False Then
                                Dim b2() As Byte = dRow("img3")
                                Dim st2 As New System.IO.MemoryStream(b2)

                                timg.Image = System.Drawing.Image.FromStream(st2)
                            End If

                        Case Else       '승인
                            If IsDBNull(dRow("img")) = True Then
                                If IsDBNull(dRow("img1")) = False Then
                                    Dim b3() As Byte = dRow("img1")
                                    Dim st3 As New System.IO.MemoryStream(b3)

                                    timg.Image = System.Drawing.Image.FromStream(st3)
                                End If

                            Else
                                Dim b4() As Byte = dRow("img")
                                Dim st4 As New System.IO.MemoryStream(b4)

                                timg.Image = System.Drawing.Image.FromStream(st4)
                            End If
                    End Select

                Case "BC210400" '반려
                    If IsDBNull(dRow("img2")) = False Then
                        Dim b5() As Byte = dRow("img2")
                        Dim st5 As New System.IO.MemoryStream(b5)

                        timg.Image = System.Drawing.Image.FromStream(st5)
                    End If
            End Select

            tCol = tCol + 1
        Next

        dSet = popup.OpenDataSet("gwm100_view_line_2", p5)

        tCol = 1
        For Each dRow In dSet.Tables(0).Rows

            popup.g60.Text("c" & tCol, 0) = ToStr(dRow("title"))
            popup.g70.Text("c" & tCol, 0) = ToStr(dRow("nm"))

            Select Case tCol
                Case 1
                    timg = popup.img8
                Case 2
                    timg = popup.img9
                Case 3
                    timg = popup.img10
                Case 4
                    timg = popup.img11
                Case 5
                    timg = popup.img12
                Case 6
                    timg = popup.img13
                Case 7
                    timg = popup.img14
            End Select

            Select Case ToStr(dRow("appr_sw"))
                Case "BC210300"     '승인
                    Select Case ToStr(dRow("deputy_bc"))
                        Case "GW110100"     '전결
                            If IsDBNull(dRow("img4")) = False Then
                                Dim b1() As Byte = dRow("img4")
                                Dim st1 As New System.IO.MemoryStream(b1)

                                timg.Image = System.Drawing.Image.FromStream(st1)
                            End If

                        Case ("GW110200")   '대결
                            If IsDBNull(dRow("img3")) = False Then
                                Dim b2() As Byte = dRow("img3")
                                Dim st2 As New System.IO.MemoryStream(b2)

                                timg.Image = System.Drawing.Image.FromStream(st2)
                            End If

                        Case Else           '승인
                            If IsDBNull(dRow("img")) = True Then
                                If IsDBNull(dRow("img1")) = False Then
                                    Dim b3() As Byte = dRow("img1")
                                    Dim st3 As New System.IO.MemoryStream(b3)

                                    timg.Image = System.Drawing.Image.FromStream(st3)
                                End If
                            Else
                                Dim b4() As Byte = dRow("img")
                                Dim st4 As New System.IO.MemoryStream(b4)

                                timg.Image = System.Drawing.Image.FromStream(st4)
                            End If
                    End Select

                Case "BC210400"     '반려
                    If IsDBNull(dRow("img2")) = False Then
                        Dim b5() As Byte = dRow("img2")
                        Dim st5 As New System.IO.MemoryStream(b5)

                        timg.Image = System.Drawing.Image.FromStream(st5)
                    End If
            End Select

            tCol = tCol + 1
        Next
        '================================
        ' 의견
        '================================
        Dim p10 As New OpenParameters

        p10.Add("@appr_no", Appr_No)

        popup.Open("gwm100_view_g20", p10)
        '================================

        '서로 다른 폼들이 같은 FTP Form Code로 등록하여 FTP Form Code 형식을 분리해줌
        Select Case Form_Cd
            Case "PRP200_BJ"    '제작요청서 (설비/지그/부품)
                Ftp_Form = "PRP100_BJ"
            Case "PRP300_BJ"    '제작요청서 (자동화 가공요청)
                Ftp_Form = "PRP100_BJ"
            Case Else
                Ftp_Form = Form_Cd
        End Select

        ' 첨부파일
        '================================
        Dim p20 As New OpenParameters

        p20.Add("@form_cd", Ftp_Form)
        p20.Add("@ftp_id", Ftp_ID)

        popup.Open("gwm100_view_g30", p20)

        '한림 구매발주일 때 첨부파일2(구매의뢰) 보이게 추가 (2021-02-15)
        If Comp_Type() = "HL" And Ftp_Form = "MMB100_HL" Then
            popup.Label1.Text = "발주 첨부파일"
            popup.Label2.Text = "구매의뢰 첨부파일"
            popup.Label2.Visible = True
            popup.g31.Visible = True
            popup.btn_prt2.Visible = True
            popup.btn_prt2.Text = "구매의뢰서"

            p20.Clear()
            p20.Add("@po_no", Ref_No1)
            popup.Open("gwm100_view_g31", p20)
        Else
            popup.Label2.Visible = False
            popup.g31.Visible = False
            popup.btn_prt2.Visible = False
        End If
        '================================

        '데브리포트
        '================================
        Call Dev_Report()

        If Appr_All = True Then
            Exit Sub
        End If

        Dim SQL As String

        SQL = "update GWM110 set " & _
                " read_dt = getdate() " & _
                 " where appr_no = '" & Appr_No & "'" & _
                " and appr_chrg = " & Parameter.Login.RegId & _
                " and isnull(read_dt,'') = ''"
        Link.ExcuteQuery(SQL)

        '결재상신자가 들어올경우 전결대결 버튼 비활성화 2019-06-04
        Dim dSet2 As DataSet
        Dim SQL2 As String

        SQL2 = ""
        SQL2 = "select appr_bc, isnull(appr_sort, 0) as appr_sort from GWM110  " & _
               " where appr_no = '" & Appr_No & "'" & _
               " and appr_chrg = " & Parameter.Login.RegId

        dSet2 = Link.ExcuteQuery(SQL2)

        If dSet2.Tables(0).Rows(0).Item("appr_bc") <> "GW100500" And dSet2.Tables(0).Rows(0).Item("appr_sort") <> "0" Then
            popup.btnDeputy.Enabled = True
            popup.btnAll.Enabled = True
        Else
            popup.btnDeputy.Enabled = False
            popup.btnAll.Enabled = False
        End If

        '구매승인 BOM등록 점프 추가 20190504
        If Form_Cd = "DMB105_SP" Then
            popup.btn_dmb100.Visible = True
        Else
            popup.btn_dmb100.Visible = False
        End If

        '유효성검토 점프 버튼 추가. 2020-05-13. YANG
        If Form_Cd = "SDB100_SSP" Then
            Dim _so_type As String = ""
            Dim _grp1_cd As String = ""

            p.Clear()
            p.Add("@appr_No", Appr_No)

            Dim ds As DataSet = Base8.Link.ReadDataSet("SDB100_SPP_get_so_info_appr", p)

            If Not IsEmpty(ds) Then
                _so_type = DataValue(ds, "so_type")
                _grp1_cd = DataValue(ds, "grp1_cd")

                If (_grp1_cd = "A100100" Or _grp1_cd = "A100101") And (_so_type = "SD204300" Or _so_type = "SD204400" Or _so_type = "SD204500") Then '시스템 또 파트면서 as, ias, 개선일격우
                    popup.btn_pms300_jump.Visible = True
                Else
                    popup.btn_pms300_jump.Visible = False
                End If
            End If

        Else
            popup.btn_pms300_jump.Visible = False
        End If
    End Sub

    Public Sub Dev_Report()
        ' 데브리포트
        '================================
        'If MOD_APPR.Dev_Print(Form_Cd, Appr_No, Ref_No1, Ref_No2, Ref_No3, Ref_No4, popup.PrintControl, popup.g20) <> True Then
        '    Me.Close()

        'End If

        '데브리포트 
        '================================
        Dim p As New OpenParameters
        Dim p_r As New OpenParameters
        Dim dSet_r As New DataSet
        Dim re_cd As String = ""
        Dim p_r2 As New OpenParameters
        Dim dSet_r2 As New DataSet
        Dim re_cd2 As String = ""

        Dim Rpt_Afr As DevExpress.XtraReports.UI.XtraReport = Nothing   '최초레포트
        Dim Rpt_Afr2 As DevExpress.XtraReports.UI.XtraReport = Nothing  '후속레포트

        p.Add("@appr_no", Appr_No)

        Dim dSet As DataSet = Base8.Link.ReadDataSet("gwm100_view_info1", p) ' popup.OpenDataSet("gwm100_view_info1", p)

        If IsEmpty(dSet) Then
            MessageInfo("[결재] 마스터정보 오류!!!")
            Me.Close()
        End If

        ' 결재마스터정보(전역변수)
        '================================================
        Form_Cd = ToStr(dSet.Tables(0).Rows(0).Item("form_cd"))
        Ref_No1 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No1"))
        Ref_No2 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No2"))
        Ref_No3 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No3"))
        Ref_No4 = ToStr(dSet.Tables(0).Rows(0).Item("Ref_No4"))
        Ftp_ID = ToStr(dSet.Tables(0).Rows(0).Item("Ftp_ID"))
        '================================================

        Select Case Form_Cd
            Case "GWB100"
                p_r.Add("@type_bc", "")
                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb100_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB100"
                End If

            Case "GWB100_HL", "GWB100_SK", "GWB100_WR", "GWB100_TST"        '공용보고서 (HL, SK, WR)

                p_r.Add("@type_bc", "")
                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB100_SSP"

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb100_ssp_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB100_SSP_R01"
                End If

            Case "GWB100_DS"

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb100_ds_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB100_DS_R"
                End If

            Case "MQA100_BJ"

                p_r.Add("@make_req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mqa100_bj_print", p_r)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "MQA100_BJ_01"
                End If

            Case "MQA100_HH"

                p_r.Add("@make_req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mqa100_hh_print", p_r)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "MQA100_HH_01"
                End If

            Case "MQA120_BJ"

                p.Add("@make_req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mqa120_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MQA120_BJ_01"
                End If

            Case "QMG100_BJ"

                p.Add("@bad_id", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg100_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG100_BJ_01"
                End If

            Case "QMG100_HK"

                p.Add("@bad_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg100_hk_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG100_HK_R"
                End If

            Case "QMG100_SF"

                p.Add("@bad_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("QMG100_SF_Print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG100_Print_R01_SF"
                End If

            Case "QMG200_WTS"

                p.Add("@eco_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg200_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG200_WTS_R"
                End If

            Case "QMG200_HK"

                p.Add("@bad_id", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg200_hk_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG200_HK_R"
                End If

                dSet_r2 = Nothing
                dSet_r2 = Base8.Link.ReadDataSet("qmg200_hk_print2", p)

                Rpt_Afr = New QMG200_HK_R2(dSet_r2)
                Rpt_Afr.CreateDocument()

            Case "QMG600_WTS", "QMG600_WTS_2"

                p.Add("@act_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg600_wts_prt", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG600_WTS_R"
                End If

            Case "QMG800_WTS"

                p.Add("@eval_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg800_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMG800_WTS_R"
                End If

            Case "FMR100_BJ"

                p.Add("@rep_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fmr100_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "FMR100_BJ_01"
                End If

            Case "FMR110_BJ"

                p.Add("@rep_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fmr110_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "FMR110_BJ_01"
                End If

            Case "MMA100_SK", "MMA100_JM" '수경화학 구매요청 전자결재, 진명구매의뢰서 출력물 '21.08.011 추가

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("MMA100_SK_print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R01"
                    're_cd = "MMA100_SK_R01"
                End If


            Case "MMA100_TST" '구매요청 전자결재

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("MMA100_tst_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA100_TST_R"
                End If

            Case "MMA100_SY" '삼양 구매요청 전자결재 ' 20200110 YANG. 추가.

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("MMA100_SY_print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA100_SY_R01"
                End If

            Case "MMA100_WTS", "MMA100_HK", "MMA100_HH"      '위더스, HK, 한일 구매요청

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "MMA200_BJ"

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mma200_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA200_BJ_01"
                End If

            Case "MMA200_TST" '구매요청 전자결재

                p.Add("@res_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("MMA200_tst_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA200_TST_R"
                End If

            Case "PPD500_SK"

                p.Add("@Id", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("ppd500_sk_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PPD500_SK_R"
                End If

            Case "PPT650_BJ"

                p.Add("@res_id", Ref_No1)
                p.Add("@p_r_no", Ref_No2)
                p.Add("@m_pr_seq", Ref_No3)
                p.Add("@div_seq", Ref_No4)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("ppt650_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PPT650_BJ_01"
                End If

            Case "MQA400_BJ"

                p.Add("@res_id", Ref_No1)
                p.Add("@make_req_no", Ref_No2)
                p.Add("@make_req_seq", Ref_No3)
                p.Add("@mold_sq", Ref_No4)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mqa400_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MQA400_BJ"
                End If

            Case "PRP100_BJ"

                p.Add("@p_r_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("prp100_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PRP100_BJ"
                End If

            Case "PRP200_BJ"

                p.Add("@p_r_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("prp200_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PRP200_BJ"
                End If

            Case "PRP300_BJ"

                p.Add("@p_r_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("prp300_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PRP300_BJ"
                End If

            Case "QMT700_BJ"

                p.Add("@ts_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmt700_bj_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMT700_BJ"
                End If

            Case "RAK150_BJ"

                p.Add("@group_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("rak150_bj_report", p)

                If Not IsEmpty(dSet_r) Then
                    're_cd = "RAK150_BJ_R"
                    re_cd = "RAK150_BJ_R2"
                End If

            Case "LER320_HL", "LER300_HL"

                p.Add("@req_no", Ref_No1)
                p.Add("@print_non_show", "0")
                p.Add("@print_fac_cd", "")

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("LER320_HL_print", p)

                If Not IsEmpty(dSet_r) Then
                    Select Case DataValue(dSet_r.Tables(1), "out_bc_cd")
                        
                        Case "LE200120"                 '출고구분 = 임대
                            re_cd = "LER320_HL_R_01"

                        Case "LE200125", "LE200126"     '출고구분 = 판매, 수리보관
                            re_cd = "LER320_HL_R_02"

                        Case Else
                            re_cd = "LER320_HL_R_02"

                    End Select
                End If

            Case "MMB100_SPD"

                p.Add("@po_no", Ref_No1)

                Dim tPo_1_bc As String = Ref_No2
                Dim tPo_kd As String = Ref_No3

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_spd_print", p)

                If Not IsEmpty(dSet_r) Then
                    If tPo_1_bc = "MM102100" Then '트레이이면서
                        If tPo_kd = "MM101200" Or tPo_kd = "MM101150" Then ' 외주발주, 상품일경우도 외주출력물로 
                            re_cd = "MMB100_SPD_R"
                        Else '외주발주가 아니면 자재발주
                            re_cd = "MMB100_SPD_R2"
                        End If
                    ElseIf tPo_1_bc = "MM102200" Then '판촉이라면..
                        re_cd = "MMB100_SPD_R3"
                    Else
                        're_cd = "MMB100_SPD"
                        re_cd = ""
                    End If

                End If

            Case "SDB190_SPD" '개발비관리비용 20180403

                p.Add("@reg_mon", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb190_spd_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "SDB190R_SPD"
                End If

            Case "QMM100_SPD" '수입검사등록(삼원) 20180410

                p.Add("@iqc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmm100_spd_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMM100_SPD_01"
                End If

            Case "QMM102_SPD" '수입검사등록(삼원) 20190125

                p.Add("@iqc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmm102_spd_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMM102_SPD_01"
                End If

            Case "GWB110" '회의록 20180621

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb110_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB110"
                End If

            Case "GWB110_SK", "GWB110_WR", "GWB110_TST" '회의록 20180621

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB110_SSP" '회의록 20180621

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb110_ssp_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB110_SSP"
                End If

            Case "GWB110_DS" '회의록_DS

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb110_ds_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB110_DS_R"
                End If

            Case "GWB120" '기안작성 20180621
                '코딩변경 btn_read_only 변수로 변경 mma100_hl외 다른 메뉴에서도 잠금처리 필요
                'If Appr_Ref = "MMA100_HL" Then '구매의뢰등록 참조보기시
                '    popup.btnAppr.Enabled = False
                '    popup.btnReturn.Enabled = False
                '    popup.btnDeputy.Enabled = False
                '    popup.btnAll.Enabled = False
                'Else
                '    popup.btnAppr.Enabled = True
                '    popup.btnReturn.Enabled = True
                '    popup.btnDeputy.Enabled = True
                '    popup.btnAll.Enabled = True
                'End If
                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb120_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB120"
                End If

            Case "GWB120_SSP" '기안작성_SSP 20190905

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb120_ssp_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB120"
                End If

            Case "GWB120_WTS" '기안작성_WTS 20191112

                p_r.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb120_wts_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB120R_WTS"
                End If

            Case "GWB120_WR", "GWB120_TST"    '기안작성_WR

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB130_WTS" '기안작성_WTS 20191112

                p_r.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb130_wts_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB130R_WTS"
                End If

            Case "GWB120_DS" '기안작성_DS

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb120_ds_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB120_DS_R"
                End If

            Case "GWB110_WTS" '회의록_WTS 20191112

                p_r.Add("@doc_no", Ref_No1)
                p_r.Add("@seq", Ftp_ID)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb110_wts_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB110_WTS_R"
                End If

            Case "GWB130"   '제안서

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb130_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB130"
                End If


            Case "GWB130_TST"    '제안서_TST

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB270", "GWB270_HL" '휴가신청서 20180621 ' 20210109 한림 추가

                p_r.Add("@seq", Ref_No1)
                'p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)
                'dSet_r = Base8.Link.ReadDataSet("gwb270_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd
                    're_cd = "GWB270"
                End If

            Case "GWB270_TST"    '휴가신청서_TST

                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GAT120_DD" '연차휴가 신청서(대동)

                p.Add("@ftp_id", Ref_No1)
                p.Add("@emp_no", Ref_No2)
                p.Add("@atn_cd", "A010")
                p.Add("@lan_no", "1")

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gat120_dd_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "GAT120_DD_R1"
                End If

            Case "GWB200" '사유서/경위서 20180621

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb200_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB200"
                End If

            Case "GWB210" '시말서 20180622

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb200_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB210"
                End If

            Case "GWB220" '휴직원/사직서/퇴직금정산 20180622

                p.Add("@seq", Ref_No1)

                Dim t_type_bc As String = ""

                dSet = Base8.Link.ReadDataSet("gwb220_gettype", p)

                If Not IsEmpty(dSet) Then
                    t_type_bc = ToStr(DataValue(dSet, "type_bc"))

                End If

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_spd_print", p)

                If t_type_bc = "GW200250" Then '휴직원
                    re_cd = "GWB220"

                ElseIf t_type_bc = "GW200300" Then '사직서
                    re_cd = "GWB220_1"

                ElseIf t_type_bc = "GW200400" Then '퇴직금중간정산신청서
                    re_cd = "GWB220_2"

                Else
                    re_cd = ""
                End If

            Case "GWB230" '상조회대출신청서 20180622

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb200_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB230"
                End If

            Case "GWB240" '차량사용신청서 20180622

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb200_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB240"
                End If

            Case "GWB250" '학자금신청서 20180622

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb200_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB250"
                End If

            Case "GWB260" '대출금신청서 20180625

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb200_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB260"
                End If

            Case "GWB300_SK", "GWB300_WR", "GWB300_TST"   '사외 교육보고서

                p_r.Add("@type_bc", "")
                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB300_SSP"

                p_r.Add("@type_bc", "")
                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb300_ssp_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB300_SSP_R"
                End If

            Case "GWB310_TST" '경조금 신청서

                p_r.Add("@type_bc", "")
                p_r.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWO100_SSP"   'ISO 문서 표지_SSP

                p_r.Add("@type_bc", "")
                p_r.Add("@id", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwo100_ssp_print", p_r)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWO100_SSP_R"
                End If

            Case "GAT120_SP" '근태계신청등록

                p.Add("@fr_dt", Ref_No1)
                p.Add("@emp_no", Ref_No2)
                p.Add("@atn_cd", Ref_No3)
                p.Add("@sq_no", Ref_No4)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gat120_sp_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "GAT120_SP"
                End If

            Case "GAT120_PUB" '근태계신청등록

                p.Add("@fr_dt", Ref_No1)
                p.Add("@emp_no", Ref_No2)
                p.Add("@atn_cd", Ref_No3)
                p.Add("@sq_no", Ref_No4)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gat120_pub_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "GAT120_PUB"
                End If

            Case "FAH110_SP" '품의신청서

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah110_sp_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "FAH110_SP"
                End If

            Case "FAH110_HS" '지출품의서_HS

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah110_hs_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "FAH110_HS_R"
                End If

            Case "FAH110_WTS" '지출결의/구매의뢰서

                p.Add("@doc_no", Ref_No1)
                p.Add("@doc_bc", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah110_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    Select Case Ref_No2
                        Case "FA993100" '구매의뢰서
                            re_cd = "FAH110R_WTS_03"
                        Case "FA993200" '지출(개인)
                            re_cd = "FAH110R_WTS_01"
                        Case "FA993300" '지출(팀비)
                            re_cd = "FAH110R_WTS_02"
                        Case Else
                            re_cd = "FAH110R_WTS_01"
                    End Select
                End If

            Case "FAH110_DS" '지출결의/구매의뢰서

                p.Add("@doc_no", Ref_No1)
                p.Add("@doc_bc", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah110_ds_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    Select Case Ref_No2
                        Case "FA993100" '구매의뢰서
                            re_cd = "FAH110R_DS_03"
                        Case "FA993200" '지출(개인)
                            re_cd = "FAH110R_DS_01"
                        Case "FA993300" '여비정산
                            re_cd = "FAH110R_DS_02"
                        Case Else
                            re_cd = "FAH110R_DS_01"
                    End Select
                End If

            Case "FAH110_SY" '지출결의/구매의뢰서

                p.Add("@doc_no", Ref_No1)
                p.Add("@doc_bc", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah110_sy_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    Select Case Ref_No2
                        Case "FA993100" '구매의뢰서
                            re_cd = "FAH110R_SY_03"
                        Case "FA993200" '지출(개인)
                            re_cd = "FAH110R_SY_01"
                        Case "FA993300" '여비정산
                            re_cd = "FAH110R_SY_02"
                        Case Else
                            re_cd = "FAH110R_SY_01"
                    End Select
                End If

            Case "FAH110_SK", "FAH110_HH" '지출결의/구매의뢰서

                p.Add("@doc_no", Ref_No1)
                p.Add("@doc_bc", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    Select Case Ref_No2
                        Case "FA993100" '구매의뢰서
                            re_cd = "FAH110R_SK_03"
                        Case "FA993200" '지출(개인)
                            re_cd = "FAH110R_SK_01"
                        Case "FA993300" '여비정산
                            re_cd = "FAH110R_SK_02"
                        Case Else
                            re_cd = "FAH110R_SK_01"
                    End Select
                End If

            Case "FAH110_WR" '지출결의/구매의뢰서

                p.Add("@doc_no", Ref_No1)
                p.Add("@doc_bc", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah110_wr_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    Select Case Ref_No2
                        Case "FA993100" '구매의뢰서
                            re_cd = "FAH110_WR_R3"
                        Case "FA993200" '지출(개인)
                            re_cd = "FAH110_WR_R"
                        Case "FA993300" '여비정산
                            re_cd = "FAH110_WR_R2"
                        Case Else
                            re_cd = "FAH110_WR_R"
                    End Select
                End If

            Case "FAH120_SP" '정산내역서

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah120_sp_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "FAH120_SP"
                End If

            Case "FAH120_HS" '정산내역서

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah120_hs_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "FAH110_HS_R"
                End If

            Case "FAH120_WTS" '정산내역서

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("fah120_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "FAH120_WTS"
                End If

            Case "SDB100_SSP" 'SSP 생산지시서(수주) 2019-01-17.YANG

                p.Add("@so_no", Ref_No1)

                Dim ds As DataSet = Base8.Link.ReadDataSet("SDB105_SSP_GetPrtCode", p) '여기서 수주번호로 어떤 워크셋과 레포트를 사용할지 결정해서 바꿔준다.

                '1.리포트 데이터 가져오기
                Dim _wSet_cd As String = "SDB105_SSP_R01" '맞는게 없으면 001번으로 띄운다.
                Dim _rpt_cd As String = "SDB105_SSP_print001"

                If Not IsEmpty(ds) Then
                    _rpt_cd = ToStr(DataValue(ds, 0))
                    _wSet_cd = ToStr(DataValue(ds, 1))

                End If

                dSet_r = Base8.Link.ReadDataSet(_wSet_cd, p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = _rpt_cd
                End If

            Case "PMS600_check_SSP" 'SSP 기술검토서. 2019-06-28. 7.YANG

                p.Add("@so_no", Ref_No1)

                Dim ds As DataSet = Base8.Link.ReadDataSet("PMS600_SSP_GetPrtCode", p) '여기서 수주번호로 어떤 워크셋과 레포트를 사용할지 결정해서 바꿔준다.

                '1.리포트 데이터 가져오기
                Dim _wSet_cd As String = "PMS600_SSP_Print_R01" '맞는게 없으면 001번으로 띄운다.
                Dim _rpt_cd As String = "PMS600_SSP_Print_R01"

                If Not IsEmpty(ds) Then
                    _rpt_cd = ToStr(DataValue(ds, 0))
                    _wSet_cd = ToStr(DataValue(ds, 1))

                End If

                dSet_r = Base8.Link.ReadDataSet(_wSet_cd, p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = _rpt_cd
                End If


                '--아래는 기존코드. 2020-04-25. YANG 위와같ㅇ ㅣ수정함. 시스템이면서 본물량인것은 기존출력물 사용하고 나머지는 새로운 출력물을 사용한다.
                'p.Add("@so_no", Ref_No1)

                ''1.리포트 데이터 가져오기
                'dSet_r = Base8.Link.ReadDataSet("PMS600_SSP_Print_R01", p)

                'If Not IsEmpty(dSet_r) Then
                '    '선택된 정보가 없습니다.
                '    re_cd = "PMS600_SSP_R01"
                'End If

            Case "PMS600_out_SSP" 'SSP 출하성적서. 2019-06-28. 7.YANG

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("PMS600_SSP_Print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "PMS600_SSP_R01"
                End If

            Case "SDB130_HL" '공사등록_HL. 2019-06-19. YANG추가.

                p.Add("@const_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("SDB130_HL_print_001", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    If DataValue(dSet_r, "dlv_bc") = "SD490100" Then '지상이라면
                        re_cd = "SDB130_HL_R01" '지상
                    ElseIf DataValue(dSet_r, "dlv_bc") = "SD490200" Then '지하라면
                        re_cd = "SDB130_HL_R02" '지하
                    Else
                        re_cd = ""
                    End If
                End If

            Case "SDB140_HL" '물량리스트_HL. 2019-06-20. YANG추가.

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("SDB140_HL_print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB140_HL_R01"
                End If

            Case "PMS550_HL" '오류분추가의뢰서_HL. 2019-06-24. YANG추가.

                p.Add("@req_doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("PMS550_HL_get_req_bc", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    If DataValue(dSet_r, "req_bc") = "SD579100" Then '오류분 추가의뢰서
                        re_cd = "PMS550_HL_R01" '오류분 추가의뢰서
                        dSet_r = Base8.Link.ReadDataSet("PMS550_HL_print_R01", p)
                    ElseIf DataValue(dSet_r, "req_bc") = "SD579200" Then '추가분 의뢰서
                        re_cd = "PMS550_HL_R02" '추가분 의뢰서
                        dSet_r = Base8.Link.ReadDataSet("PMS550_HL_print_R02", p)
                    Else
                        re_cd = ""
                    End If
                End If

            Case "SDZ700_HC" '4M 신청서 등록

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdz700_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDZ700_HC_R"
                End If

            Case "DMB105_SP" 'BOM 구매승인

                p.Add("@appr_no", Appr_No)
                p.Add("@gbn", "1")

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("dmb105_sp_print", p)

                If Not IsEmpty(dSet_r) Then
                    If ToStr(dSet_r.Tables(0).Rows(0).Item("so_type")) = "SD204100" Or ToStr(dSet_r.Tables(0).Rows(0).Item("so_type")) = "SD204200" Then
                        re_cd = "DMB105_1SP"
                        'ElseIf ToStr(dSet_r.Tables(0).Rows(0).Item("so_type")) = "SD204300" Or ToStr(dSet_r.Tables(0).Rows(0).Item("so_type")) = "SD204400" Then
                    Else
                        p.Add("@appr_no", Appr_No)
                        '1.리포트 데이터 가져오기
                        dSet_r = Base8.Link.ReadDataSet("dmb105_sp_print2", p)

                        re_cd = "DMB105_2SP"
                    End If
                End If

            Case "DMO100_SY" '금형관리_SY

                p.Clear()
                p.Add("@mold_cd", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("dmo100_sy_print", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "DMO100_SY_R"
                End If

            Case "DMP250" '개발의뢰서 2019-03-13.YANG

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Ref_No2, p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Ref_No3
                End If

            Case "DMP400" '승인원 등록관리. 2019-03-13.YANG

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Ref_No2, p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Ref_No3
                End If

            Case "MMC510_HC" '사급대행처리. 2019-03-27.YANG

                p.Add("@agent_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("MMC510_HC_Print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "MMC510_HC_R01"
                End If

            Case "QMG100_SSP"       '부적합 등록_SSP

                p.Add("@bad_id", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg100_ssp_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG100_SSP"
                End If

            Case "QMG140_HC" '수입검사 불합격 통보서_HC  (2019-03-28)

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg140_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG140_HC_R"
                End If

            Case "QMG141_HC" '출하검사 불합격 통보서_HC  (2019-03-28)

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg141_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG141_HC_R"
                End If

            Case "QMG150_HC" '시정조치요구서_HC  (2019-03-28)

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg150_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG150_HC_R"
                End If

            Case "QMG140_SY" '수입검사 부적합보고서_SY  (2020-01-18)

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg140_sy_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG140_SY_R"
                End If

            Case "QMG141_SY" '출하검사 부적합보고서_SY  (2020-01-18)

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg141_sy_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG141_SY_R"
                End If

            Case "QMG150_SY" '시정(예방)조치요구서_SY  (2020-01-18)

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg150_sy_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG150_SY_R"
                End If

            Case "QME300_HC" '시방변경요청서_HC  (2019-04-05)

                p.Add("@pub_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qme300_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QME300_HC_R"
                End If

            Case "QME320_HC" '시방변경서_HC  (2019-04-05)

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qme320_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QME320_HC_R"
                End If

            Case "QME300_SY" '설계변경요청서_HC  (2019-12-13)

                p.Add("@pub_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qme300_sy_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QME300_SY_R"
                End If

            Case "QME320_HC" '설계변경서_sy  (2019-12-13)

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qme320_sy_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QME320_SY_R"
                End If


            Case "PMB210_HC", "PMB210_WR" 'PJT결재보고서  (2019-04-11)

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R01"
                End If

            Case "MMB100_HC" '발주서(힐세리온)

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_hc_print2", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "MMB100_HC_R"
                End If

            Case "MMB100_DB", "MMB100_JM", "MMB100_SWT", "MMB100_HH" '발주서(동방)'2020-08-05. YANG 추가, 발주서(JM) 추가 '2021.06.08 김용범

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "MMB100_HL"

                'p.Add("@po_no", Ref_No1)

                ''1.리포트 데이터 가져오기
                'dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                'If Not IsEmpty(dSet_r) Then
                '    '선택된 정보가 없습니다.
                '    re_cd = Form_Cd & "_R"
                'End If

                p.Clear()
                p.Add("@po_no", Ref_No1)

                dSet_r = Nothing
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

                '구매의뢰서 Dataset 전역변수 세팅
                p.Clear()
                p.Add("@appr_no", Ref_No2)

                Dev_dSet_Sub = Nothing
                Dev_dSet_Sub = Base8.Link.ReadDataSet(Form_Cd & "_print2", p)

            Case "MMB100_DD" '발주서(대동)

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_dd_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "MMB100_DD_R01"
                End If

            Case "MMB100_HK" '발주서(HK)

                p.Clear()
                p.Add("@po_no", Ref_No1)
                p.Add("@gubun", Ref_No2)

                If Ref_No2 = "MM060100" Then    '내자
                    '1.리포트 데이터 가져오기
                    dSet_r = Base8.Link.ReadDataSet("mmb100_hk_print", p)
                    If Not IsEmpty(dSet_r) Then
                        re_cd = "MMB100_HK_R"
                    End If

                Else                            '엔지니어링사업부
                    '1.리포트 데이터 가져오기
                    dSet_r = Base8.Link.ReadDataSet("mmb100_hk_print", p)
                    If Not IsEmpty(dSet_r) Then
                        re_cd = "MMB100_HK_R2"
                    End If
                End If

            Case "SDA800_HC" '주문서(힐세리온)

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sda800_hc_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "sda800_hc_r"
                End If

            Case "SDA700_SK", "SDA700_CH", "SDA700_WR", "SDA700_HH" '견적등록

                p.Add("@estm_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd + "_R"
                End If

            Case "SDA700_JM"

                p.Add("@estm_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sda700_jm_print_gw", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd + "_R" + dSet_r.Tables(1).Rows(0).Item("lanNo")
                End If

            Case "SDB100_SK" '주문서(힐세리온)

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb100_sk_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB100_SK_R"
                End If

            Case "SDB100_HH" '주문서(한일하이테크) 2021-07-30. YANG 추가.

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb100_HH_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB100_HH_R"
                End If


            Case "SDB100_JM" '주문서(진명파워텍)

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb100_jm_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB100_JM_R"
                End If

            Case "SDB100_CH" '주문서등록(씨에치산업)

                p.Clear()
                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb100_ch_print", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "SDB100_CH_R"
                End If

            Case "SDB100_DRUM" '수주서(대세)

                p.Clear()
                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb100_drum_print", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "SDB100_DRUM_R"
                End If


            Case "SDB400_HC" '내부용제품신청_HC. 2019-05.02 추가. YANG.

                p.Add("@prd_req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("SDB400_HC_R01", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB400_HC_R01"
                End If

            Case "QMC105_WTS" '출하검사성적서_WTS

                p.Add("@qc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmc105_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "QMC105_WTS_R"
                End If


            Case "QMC105_HK" '출하특기사항_HK

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If


            Case "QMF100_HK" '출하특기사항_HK

                p.Add("@eqc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "MMB100_WTS" '구매발주서_WTS

                p.Add("@po_no", Ref_No1)
                p.Add("@po_bs", Ref_No2)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMB100_WTS_R"
                End If

            Case "GWB210_HL" '지출결의서_HL 2019.07.05

                p.Add("@doc_id", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb210_hl_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB210_HL_R"
                End If

            Case "MMA100_HL" '구매발주서_HL 20190701

                p.Add("@req_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mma100_hl_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA100_HL_R"
                End If

            Case "MMA100_WR" '구매발주서_HL 20190701

                p.Add("@req_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mma100_wr_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA100R_WR"
                End If

            Case "SDB258_SSP" '체크리스트_SSP 20190711

                p.Add("@chk_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("LK_SDB258_SSP_Print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "SDB258_SSP_R01"
                End If

            Case "PPE200_WTS" '외주발주등록

                p.Add("@po_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("ppe200_wts_bal_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PPE200R_WTS_BAL"
                End If

            Case "GAT120_WTS" '근태계신청등록

                p.Add("@ftp_id", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gat120_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GAT120R_WTS"
                End If

            Case "GAT500_WTS" '특근신청등록

                p.Add("@ftp_id", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gat500_wts_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GAT500R_WTS"
                End If

            Case "LEC100_SK" '자재변경요청_SK 20190711

                p.Clear()
                p.Add("@chg_doc_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("LEC100_SK_print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "LEC100_SK_R01"
                End If

            Case "LEJ300_SK" '자재 재고조정 월별'

                p.Clear()
                p.Add("@std_month", "")
                p.Add("@appr_no", Appr_No)
                p.Add("@appr_no_2", Appr_No)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("LEJ300_SK_Print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "LEJ300_SK_R"
                End If

            Case "LEO150_in_SK" '판매마감_부대비용포함_SK 20190711

                p.Clear()
                p.Add("@appr_no_1", Appr_No)
                p.Add("@appr_no_2", "")
                p.Add("@f_fr_dt", "")
                p.Add("@f_to_dt", "")
                p.Add("@f_itm_cd", "")
                p.Add("@f_itm_nm", "")
                p.Add("@f_cust_nm", "")
                p.Add("@f_chk_sub_amt", "1")
                p.Add("@prt_sw", "2")
                p.Add("@outstore_no", "")
                p.Add("@outstore_sq", "")

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("LEO113_SK_print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "LEO113_SK_R01"
                End If

            Case "LEO150_notin_SK" '판매마감_부대비용불포함_SK 20190711

                p.Clear()
                p.Add("@appr_no_1", "")
                p.Add("@appr_no_2", Appr_No)
                p.Add("@f_fr_dt", "")
                p.Add("@f_to_dt", "")
                p.Add("@f_itm_cd", "")
                p.Add("@f_itm_nm", "")
                p.Add("@f_cust_nm", "")
                p.Add("@f_chk_sub_amt", "")
                p.Add("@prt_sw", "3")
                p.Add("@outstore_no", "")
                p.Add("@outstore_sq", "")

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("LEO113_SK_print_R01", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "LEO113_SK_R01"
                End If

            Case "LEA505_HK"    '매입 품목 현황_HK
                p.Clear()
                p.Add("@appr_no", Appr_No)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB280_HL" '휴일근무신청서

                p.Clear()
                p.Add("@doc_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb280_hl_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB280R_HL"
                End If

            Case "MMA100_DS" '구매요청서_DS
                '2가지 출력물 합쳐져서 나오는 것 확인해야 함.

                p.Clear()
                p.Add("@req_no", Ref_No1)

                dSet_r = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R01", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMA100_DS_R01"
                End If

                'If Ref_No2 = "MM043100" Then   '일반
                '    '1.리포트 데이터 가져오기
                '    dSet_r2 = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R02", p)
                '    If Not IsEmpty(dSet_r2) Then
                '        re_cd2 = "MMA100_DS_R02"
                '    End If
                'ElseIf Ref_No2 = "MM043200" Then '볼트/너트
                '    '1.리포트 데이터 가져오기
                '    dSet_r2 = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R03", p)
                '    If Not IsEmpty(dSet_r2) Then
                '        re_cd2 = "MMA100_DS_R03"
                '    End If
                'ElseIf Ref_No2 = "MM043300" Then '프렌지
                '    '1.리포트 데이터 가져오기
                '    dSet_r2 = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R04", p)
                '    If Not IsEmpty(dSet_r2) Then
                '        re_cd2 = "MMA100_DS_R04"
                '    End If
                'End If

                'If Ref_No2 = "MM043100" Then '일반                    

                '    dSet_r2 = Nothing
                '    dSet_r2 = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R02", p)

                '    Rpt_Afr = New MMA100_DS_R02(dSet_r2)
                '    Rpt_Afr.CreateDocument()
                '    'Rpt_Fir.Pages.AddRange(Rpt_Afr.Pages)

                'ElseIf Ref_No2 = "MM043200" Then '볼트/너트

                '    dSet_r2 = Nothing
                '    dSet_r2 = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R03", p)

                '    Rpt_Afr = New MMA100_DS_R03(dSet_r2)
                '    Rpt_Afr.CreateDocument()
                '    'Rpt_Fir.Pages.AddRange(Rpt_Afr.Pages)

                'ElseIf Ref_No2 = "MM043300" Then '프랜지

                '    dSet_r2 = Nothing
                '    dSet_r2 = Base8.Link.ReadDataSet("MMA100_DS_PRINT_R04", p)

                '    Rpt_Afr = New MMA100_DS_R04(dSet_r2)
                '    Rpt_Afr.CreateDocument()
                '    'Rpt_Fir.Pages.AddRange(Rpt_Afr.Pages)
                'End If

            Case "MMB100_DS" '구매품의서

                p.Clear()
                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_ds_print", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "MMB100_DS_R01"
                End If

            Case "PMA300_SY" '프로젝트(현장) 관리_SY

                p.Clear()
                p.Add("@prj_cd", Ref_No1)

                If Ref_No2 = "PM300100" Then    '영업부
                    '1.리포트 데이터 가져오기
                    dSet_r = Base8.Link.ReadDataSet("pma300_sy_print", p)
                    If Not IsEmpty(dSet_r) Then
                        re_cd = "PMA300_SY_R2"
                    End If

                Else                            '엔지니어링사업부
                    '1.리포트 데이터 가져오기
                    dSet_r = Base8.Link.ReadDataSet("pma300_sy_print", p)
                    If Not IsEmpty(dSet_r) Then
                        re_cd = "PMA300_SY_R"
                    End If
                End If

            Case "PPA300_SY" '구매의뢰등록_SY. 2019-11-24. YANG 추가.

                p.Clear()
                p.Add("@preq_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("PPA300_SY_Print_R01", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "PPA300_SY_R01"
                End If

            Case "PPA300_WR" '구매의뢰등록_WR. 

                p.Clear()
                p.Add("@preq_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("PPA300_WR_Print_R01", p)
                If Not IsEmpty(dSet_r) Then
                    If Ref_No2 = "PP935100" Then 'SET외주
                        re_cd = "PPA300_WR_R02"
                    ElseIf Ref_No2 = "PP935200" Then '설계외주
                        re_cd = "PPA300_WR_R03"
                    ElseIf Ref_No2 = "PP935300" Then '임가공외주
                        re_cd = "PPA300_WR_R04"
                    ElseIf Ref_No2 = "PP935400" Then '시사출외주
                        re_cd = "PPA300_WR_R05"
                    ElseIf Ref_No2 = "PP935500" Then '명판발주
                        re_cd = "PPA300_WR_R01"
                    ElseIf Ref_No2 = "PP935600" Then '포장외주
                        re_cd = "PPA300_WR_R06"
                    End If
                    're_cd = "PPA300_WR_R01"
                End If

            Case "GWB120_HL" '기안작성_HL 20191112

                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb120_hl_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB120_HL"
                End If

            Case "GWB320" '외근(출장신청서) '20191115

                p.Clear()
                p.Add("@doc_no", Ref_No1)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gwb320_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GWB320R"
                End If

            Case "MMB100_SY" '발주서(삼양) '20191121

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_sy_print2", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "MMB100_SY_R"
                End If


            Case "MMB100_SF" '발주서(세이프코리아) '20200807

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mmb100_sf_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "MMB100_SF_R"
                End If

            Case "SDB100_HK" '수주등록(영업 CS) 2020-03-27

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb100_hk_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB100_hk_R"
                End If

            Case "SDB105_HK" '수주등록(장비) 2020-04-18

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdb105_hk_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDB105_HK_R"
                End If


                p.Clear()
                p.Add("@dept_cd", Login.DeptCd)
                dSet = Base8.Link.ReadDataSet("sdb105_hk_deptchk", p)

                If Not IsEmpty(dSet) Then
                    _chk_dept = DataValue(dSet)
                End If

                '재무, 영업, 영업관리만 가능.
                If _chk_dept <> "" And re_cd <> "" Then

                    p.Clear()
                    p.Add("@so_no", Ref_No1)

                    dSet_r2 = Nothing
                    dSet_r2 = Base8.Link.ReadDataSet("sdb105_hk_print4", p)

                    Rpt_Afr2 = New SDB105_HK_R4(dSet_r2)
                    Rpt_Afr2.CreateDocument()
                End If

            Case "QMG140_WR" '부적합보고서(우리엠텍) 2020-05-16 Jang

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg140_wr_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "qmg140_wr_R"
                End If

            Case "SDA600_WR" '견적의뢰서(우리엠텍) 2020-05-19 Seol

                p.Add("@rqs_estm_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sda600_wr_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "sda600_wr_R"
                End If

            Case "MQA120_WR" '시사출/샘플의뢰서(우리엠텍) 2020-05-19 Seol

                p.Add("@make_req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mqa120_wr_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "mqa120_wr_R"
                End If

            Case "LER200_CH" '출고의뢰등록(CH산업) 2020-06-23 Seol

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("ler200_ch_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "LER200_CH_R"
                End If

            Case "LTA130_CH" '폐기처리등록(CH산업) 2020-07-21 Seol

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("lta130_ch_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "LTA130_CH_R"
                End If
            Case "SDA770_HL" '견적체크리스트_HL

                p.Clear()
                p.Add("@doc_seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sda770_hl_print", p)
                If Not IsEmpty(dSet_r) Then
                    re_cd = "SDA770_HL_R"
                End If

            Case "PMA100_SK", "PMA900_WR" '프로젝트등록 2020-08-18 Seol, 프로젝트기안서 2020-10-15 Seol

                p.Add("@prj_doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "SDB100_WR", "SDB105_WR", "SDB105_SWT", "SDB100_SWT" '우리엠텍 수주서. 2020-09-07. 김용범추가, 서우 수주서. 2021-07-20

                p.Add("@so_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "PLM100_WR" '우리엠텍 PLM파일 결재

                p.Clear()
                p.Add("@plm_list", "")
                p.Add("@appr_no", Appr_No)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "PLM100_WR_01"
                End If

            Case "FAD300_HK", "FAD305_HK" '장비시연/샘플/타임스터디 신청 및 진행 등록 Seol

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "CSA170_HK" '서비스처리등록. 2020-12-30. 김용범추가.

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "CSA170R_HK"
                End If

            Case "QMM110_SF" '수입검사 등록 2021-01-14 Seol

                p.Add("@iqc_no", Ref_No1)
                p.Add("@iqc_sq", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print2", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "QMM121_SF", "QMM131_SF" '공정검사 등록, 파이널검사 등록 2021-01-14 Seol

                p.Add("@iqc_no", Ref_No1)
                p.Add("@iqc_sq", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB290_HL" '탄력근무신청서. 2021-01-14. 김용범추가.

                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "GWB290_TST" '교육수강신청서_TST
                p.Add("@type_bc", "")
                p.Add("@seq", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "GAT500_HK" '연장근무신청서. 2021-01-29. 설현수추가.

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "SEF270_HK" 'CI /PL 등록(부품) 2021-03-02 장민식 추가

                p.Add("@invoice_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sef270_hk_1print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SEF270_HK_01"
                End If

            Case "CSS130_HK" '장비 BS보고서, 2021-03-12 설현수 추가
                p.Add("@doc_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "CSS130_HK_R"
                End If

            Case "SDZ700_DRUM" '4M 신청서 등록 seol 2021-04-06

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("sdz700_drum_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDZ700_DRUM_R"
                End If

            Case "QMG140_DRUM" '수입검사 불합격 통보서

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg140_drum_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG140_DRUM_R"
                End If

            Case "QMG150_DRUM" '시정조치요구서

                p.Add("@iss_id", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("qmg150_drum_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "QMG150_DRUM_R"
                End If

            Case "MDA200_JM" '치형구 의뢰서 2021-05-28
                p.Add("@ftp_id", Ftp_ID)
                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("mda200_jm_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "MDA200_JM_R"
                End If

            Case "SDZ400_JM", "SDZ401_JM" '4M 신청서 등록
                p.Add("@occ_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDZ400_JM_R"
                End If

            Case "SDZ700_JM", "SDZ701_JM" '4M 신청서 등록
                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDZ700_JM_R"
                End If

            Case "SDZ800_JM" '검사의뢰 접수등록
                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd + "_R"
                End If

            Case "SDT500_JM", "SDT510_JM" '시정 및 예방조치 요구서 등록, 시정 및 예방조치 대책서 등록
                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "SDT500_JM_R"
                End If

            Case "DMP200_TST" '개발의뢰서_TST
                p.Add("@prj_no", Ref_No1)
                p.Add("@seq", Ref_No2)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = "DMP200_TST_R"
                End If

            Case "MQA900_HH"
                p.Add("@make_req_no", Ref_No1)
                p.Add("@make_req_seq", Ref_No2)
                p.Add("@mas_bc", Ref_No3)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd + "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd + "_R"
                End If

            Case "GAT500_HH"

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet("gat500_hh_print", p)

                If Not IsEmpty(dSet_r) Then
                    re_cd = "GAT500_HH_R"
                End If

                dSet_r2 = Nothing
                dSet_r2 = Base8.Link.ReadDataSet("gat500_hh_print2", p)

                Rpt_Afr = New GAT500_HH_R2(dSet_r2)
                Rpt_Afr.CreateDocument()

            Case "DMB410_JM" '진명부품단가의뢰 2021-08-25. 김용범추가

                p.Add("@req_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "MQA120_HH" '시사출의뢰등록 2021-09-01 설

                p.Add("@make_req_no", Ref_No1)
                p.Add("@mas_bc", Ref_No2)
                p.Add("@num", 5)
                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then
                    '선택된 정보가 없습니다.
                    re_cd = Form_Cd & "_R"
                End If

            Case "DMB470_JM" '진명가격합의서 2021-09-07. 김용범추가

                p.Add("@po_no", Ref_No1)

                '1.리포트 데이터 가져오기
                dSet_r = Base8.Link.ReadDataSet(Form_Cd & "_print", p)

                If Not IsEmpty(dSet_r) Then

                    If Ref_No2 = "DM470100" Then '신규
                        re_cd = Form_Cd & "_R1"
                    ElseIf Ref_No2 = "DM470200" Then '변경
                        re_cd = Form_Cd & "_R2"
                    End If

                End If

        End Select

        If re_cd = "" Then '리포트코드 등록메세지 추가
            PutMessage("GWM_100_VIEW_01", "", Form_Cd)
            Me.Close()
        Else

            'If Form_Cd = "MMA100_DS" Then
            '    popup.PrintViewControl(popup.PrintControl, re_cd, dSet_r) '레포트코드, 워크셋
            '    popup.PrintControl.PrintingSystem.Pages.AddRange(Rpt_Afr.Pages)
            'Else

            If Form_Cd = "SDB105_HK" And _chk_dept <> "" Then
                popup.PrintViewControl(popup.PrintControl, re_cd, dSet_r) '레포트코드, 워크셋
                popup.PrintControl.PrintingSystem.Pages.AddRange(Rpt_Afr2.Pages)

            ElseIf Form_Cd = "QMG200_HK" Then
                popup.PrintViewControl(popup.PrintControl, re_cd, dSet_r) '레포트코드, 워크셋
                popup.PrintControl.PrintingSystem.Pages.AddRange(Rpt_Afr.Pages)

            ElseIf Form_Cd = "GAT500_HH" Then
                popup.PrintViewControl(popup.PrintControl, re_cd, dSet_r) '레포트코드, 워크셋
                popup.PrintControl.PrintingSystem.Pages.AddRange(Rpt_Afr.Pages)

            Else
                popup.PrintViewControl(popup.PrintControl, re_cd, dSet_r)

                Dev_CD = re_cd
                Dev_dSet = dSet_r
            End If
        End If
    End Sub

    Public Function InitPopup() As Boolean

        InitPopup = False

        popup = New GWM100_VIEW      'Form1 안에 불려지는 폼을 여기서 호출한다. 
        ' 참조부분에 추가되어 있어야 한다
        ' 불려지는 폼은 다른 일반 메뉴와 동일하게 구현하면 된다

        If popup IsNot Nothing Then
            popup.Dock = System.Windows.Forms.DockStyle.Fill

            'popup 화면에 버턴이 있다면 popup내에서 제어해도 되고 여기서도 제어 가능하다

            AddHandler popup.btnAppr.Click, AddressOf btnAppr_Click
            AddHandler popup.btnReturn.Click, AddressOf btnReturn_Click
            AddHandler popup.btnDeputy.Click, AddressOf btnDeputy_Click
            AddHandler popup.btnAll.Click, AddressOf btnAll_Click
            AddHandler popup.btnClose.Click, AddressOf btnClose_Click
            AddHandler popup.btnSave.Click, AddressOf btnSave_Click

            AddHandler popup.DeputyMenuItem1.Click, AddressOf DeputyMenuItem1_Click     '이전 결재자 대결
            AddHandler popup.DeputyMenuItem2.Click, AddressOf DeputyMenuItem2_Click     '다음 결재자 대결

            AddHandler popup.btn_prt2.Click, AddressOf btn_prt2_Click                   '출력2 (HL 전용) (2021-02-15)

            AddHandler popup.g30.ButtonPressed, AddressOf g30_ButtonPressed
            AddHandler popup.g31.ButtonPressed, AddressOf g31_ButtonPressed
            AddHandler popup.g20.AfterMoveRow, AddressOf g20_AfterMoveRow
            AddHandler popup.TAB1.Selected, AddressOf TAB1_Selected
            AddHandler popup.btn_dmb100.Click, AddressOf btn_dmb100_Click

            AddHandler popup.btn_pms300_jump.Click, AddressOf btn_pms300_jump_Click

            Me.Controls.Add(popup)

            popup.Show()
        End If

        InitPopup = True
    End Function

    Public Function InitPopup2(ByRef _call_frm As Control, ByRef _call_frm_cd As Object) As Boolean

        InitPopup2 = False

        popup = New GWM100_VIEW      'Form1 안에 불려지는 폼을 여기서 호출한다. 
        ' 참조부분에 추가되어 있어야 한다
        ' 불려지는 폼은 다른 일반 메뉴와 동일하게 구현하면 된다

        If popup IsNot Nothing Then
            popup.Dock = System.Windows.Forms.DockStyle.Fill

            'popup 화면에 버턴이 있다면 popup내에서 제어해도 되고 여기서도 제어 가능하다

            AddHandler popup.btnAppr.Click, AddressOf btnAppr_Click
            AddHandler popup.btnReturn.Click, AddressOf btnReturn_Click
            AddHandler popup.btnDeputy.Click, AddressOf btnDeputy_Click
            AddHandler popup.btnAll.Click, AddressOf btnAll_Click
            AddHandler popup.btnClose.Click, AddressOf btnClose_Click
            AddHandler popup.btnSave.Click, AddressOf btnSave_Click

            AddHandler popup.DeputyMenuItem1.Click, AddressOf DeputyMenuItem1_Click     '이전 결재자 대결
            AddHandler popup.DeputyMenuItem2.Click, AddressOf DeputyMenuItem2_Click     '다음 결재자 대결

            AddHandler popup.g30.ButtonPressed, AddressOf g30_ButtonPressed
            AddHandler popup.g20.AfterMoveRow, AddressOf g20_AfterMoveRow
            AddHandler popup.TAB1.Selected, AddressOf TAB1_Selected
            AddHandler popup.btn_dmb100.Click, AddressOf btn_dmb100_Click

            AddHandler popup.btn_pms300_jump.Click, AddressOf btn_pms300_jump_Click

            'Me.Controls.Add(popup)
            _call_frm.Controls.Add(popup)
            popup.Dock = DockStyle.Fill


            'popup.Show()
        End If
        If Appr_No <> "" Then
            Me.PopUp_VIEW_Load(Nothing, Nothing)
        End If

        Call_Form_Cd = _call_frm_cd

        InitPopup2 = True

    End Function


    Private Function Get_PreChk() As Boolean

        Dim p As New OpenParameters

        p.Add("@appr_no", Appr_No)
        p.Add("@appr_sort", Appr_Sort)
        p.Add("@appr_bc", Appr_Bc)

        Select Case Comp_Type()
            Case "IVT", "WTS", "DS"
                p.Add("@deputy_bc", _DeputyBC)
        End Select

        Dim dSet As DataSet = popup.OpenDataSet("gwm100_view_pre", p)

        If IsEmpty(dSet) Then
            Get_PreChk = False
            MessageInfo("[결재] 상세정보 오류!!!")
            Exit Function
        End If

        'cnt = 결재 하지않은 이전결재자의 수
        If dSet.Tables(0).Rows(0).Item("cnt") = 0 Then
            Get_PreChk = True
        Else
            Get_PreChk = False
        End If
    End Function

    ' Sw => 1 = 승인, 2 = 대결, 3 = 전결
    Private Function Approval(ByVal Sw As Long) As Boolean
        Dim SQL As String
        Dim tAppr_Sw As String
        Dim dSet As DataSet
        Dim dSet2 As DataSet = Nothing
        Dim dSet3 As DataSet = Nothing
        Dim phone_num As String = ""
        Dim tMng_Sw As Integer = 0


        If Appr_No = "" Then Exit Function '결재번호 없을시 이후 쿼리수행안함 21.03.17 

        Approval = False

        Select Case Sw
            Case 1
                ' 결재상세테이블 승인처리
                SQL = "update GWM110 set " & _
                      " appr_sw = 'BC210300', " & _
                      " pms_sw = '1', " & _
                      " pms_dt = getdate(), " & _
                      " real_chrg = " & Parameter.Login.RegId & _
                      " where appr_no = '" & Appr_No & "'" & _
                      " and appr_chrg = " & Parameter.Login.RegId
                Link.ExcuteQuery(SQL)

            Case 2  '대결
                If _DeputyBC = "1" Then     '이전 결재자 대결 결재상세테이블 승인처리 (2020-09-15)
                    SQL = "update GWM110 set " & _
                          " appr_sw = 'BC210300', " & _
                          " pms_sw = '1', " & _
                          " pms_dt = getdate(), " & _
                          " real_chrg = " & Parameter.Login.RegId & ", " & _
                          " deputy_bc = 'GW110200'" & _
                          " where appr_no = '" & Appr_No & "'" & _
                          " and appr_bc <> 'GW100500' " & _
                          " and appr_sw <> 'BC210300' " & _
                          " and appr_sort = " & CInt(Appr_Sort) - 1
                    Link.ExcuteQuery(SQL)
                Else


                    '대결 결재상세테이블 승인처리
                    SQL = "update GWM110 set " & _
                              " appr_sw = 'BC210300', " & _
                              " pms_sw = '1', " & _
                              " pms_dt = getdate(), " & _
                              " real_chrg = " & Parameter.Login.RegId & ", " & _
                              " deputy_bc = 'GW110200'" & _
                              " where appr_no = '" & Appr_No & "'" & _
                              " and appr_bc <> 'GW100500' " & _
                              " and appr_sort = " & CInt(Appr_Sort) + 1
                    Link.ExcuteQuery(SQL)


                    '결재상세테이블 승인처리
                    SQL = "update GWM110 set " & _
                          " appr_sw = 'BC210300', " & _
                          " pms_sw = '1', " & _
                          " pms_dt = getdate(), " & _
                          " real_chrg = " & Parameter.Login.RegId & _
                          " where appr_no = '" & Appr_No & "'" & _
                          " and appr_chrg = " & Parameter.Login.RegId
                    Link.ExcuteQuery(SQL)

                End If



            Case 3  '전결

                If Comp_Type() = "SSP" Then

                    If (Form_Cd = "FAH110_SP" Or Form_Cd = "FAH120_SP" Or Form_Cd = "GAT120_SP") Then


                        SQL = "select mng_sw = isnull(mng_sw,0) " & _
                              " from GWM110 " & _
                              " where appr_no = '" & Appr_No & "'" & _
                              " and appr_chrg = " & Parameter.Login.RegId

                        dSet = Link.ExcuteQuery(SQL)

                        tMng_Sw = ToDec(dSet.Tables(0).Rows(0).Item("mng_sw"))


                        ' 2020.05.08 최경식
                        ' 카카오톡 발송때문에 전결부터 승인처리

                        ' 전결 결재상세테이블 승인처리
                        SQL = "update GWM110 set " & _
                              " appr_sw = 'BC210300', " & _
                              " pms_sw = '1', " & _
                              " pms_dt = getdate(), " & _
                              " real_chrg = " & Parameter.Login.RegId & ", " & _
                              " deputy_bc = 'GW110100'" & _
                              " where appr_no = '" & Appr_No & "'" & _
                              " and appr_bc <> 'GW100500' " & _
                              " and appr_sort > " & CInt(Appr_Sort) & _
                              " and mng_sw = " & tMng_Sw
                        Link.ExcuteQuery(SQL)

                    Else

                        ' 2020.05.08 최경식
                        ' 카카오톡 발송때문에 전결부터 승인처리

                        ' 전결 결재상세테이블 승인처리
                        SQL = "update GWM110 set " & _
                              " appr_sw = 'BC210300', " & _
                              " pms_sw = '1', " & _
                              " pms_dt = getdate(), " & _
                              " real_chrg = " & Parameter.Login.RegId & ", " & _
                              " deputy_bc = 'GW110100'" & _
                              " where appr_no = '" & Appr_No & "'" & _
                              " and appr_bc <> 'GW100500' " & _
                              " and appr_sort > " & CInt(Appr_Sort)
                        Link.ExcuteQuery(SQL)

                    End If


                    ' 결재상세테이블 승인처리
                    SQL = "update GWM110 set " & _
                          " appr_sw = 'BC210300', " & _
                          " pms_sw = '1', " & _
                          " pms_dt = getdate(), " & _
                          " real_chrg = " & Parameter.Login.RegId & _
                          " where appr_no = '" & Appr_No & "'" & _
                          " and appr_chrg = " & Parameter.Login.RegId
                    Link.ExcuteQuery(SQL)

                Else

                    ' 결재상세테이블 승인처리
                    SQL = "update GWM110 set " & _
                          " appr_sw = 'BC210300', " & _
                          " pms_sw = '1', " & _
                          " pms_dt = getdate(), " & _
                          " real_chrg = " & Parameter.Login.RegId & _
                          " where appr_no = '" & Appr_No & "'" & _
                          " and appr_chrg = " & Parameter.Login.RegId
                    Link.ExcuteQuery(SQL)

                    ' 전결 결재상세테이블 승인처리
                    SQL = "update GWM110 set " & _
                          " appr_sw = 'BC210300', " & _
                          " pms_sw = '1', " & _
                          " pms_dt = getdate(), " & _
                          " real_chrg = " & Parameter.Login.RegId & ", " & _
                          " deputy_bc = 'GW110100'" & _
                          " where appr_no = '" & Appr_No & "'" & _
                          " and appr_bc <> 'GW100500' " & _
                          " and appr_sort > " & CInt(Appr_Sort)
                    Link.ExcuteQuery(SQL)

                End If

        End Select

        Dim p As New OpenParameters

        p.Add("@appr_no", Appr_No)

        dSet = popup.OpenDataSet("gwm100_view_last", p)

        If IsEmpty(dSet) Then
            MessageInfo("[결재] 승인 오류!!!")
            Exit Function
        End If

        ' 마지막 결재자 체크
        ' appr_ok = 승인한 결재자 수,  appr_all = 결재자 수
        If ((dSet.Tables(0).Rows(0).Item("appr_ok") = dSet.Tables(0).Rows(0).Item("appr_all")) Or dSet.Tables(0).Rows(0).Item("appr_last") = "BC210300") And dSet.Tables(0).Rows(0).Item("appr_sw") <> "BC210300" Then


            If Comp_Type() = "SSP" Then
                SQL = "update GWM100 set " & _
                      " before_update = '2'" & _
                      " where appr_no = '" & Appr_No & "'"
                Link.ExcuteQuery(SQL)
            End If



            ' 결재마스터테이블 결재완료처리
            SQL = "update GWM100 set " & _
                  " comp_sw = 'BC210300'" & _
                  " where appr_no = '" & Appr_No & "'"
            Link.ExcuteQuery(SQL)



            If Comp_Type() = "SSP" Then
                SQL = "update GWM100 set " & _
                      " after_update = '2'" & _
                      " where appr_no = '" & Appr_No & "'"
                Link.ExcuteQuery(SQL)
            End If


            

            tAppr_Sw = "BC210300"   '최종결재

            '삼양 -> 최종 결재 시 선정대리점 -> 확정대리점 Insert (2019-12-04)
            If Comp_Type() = "SY" Then
                SQL = "if not exists (select * from pma325 where prj_cd = '" & Ref_No1 & "' and cust_bc = 'PM325100') " & _
                      "begin " & _
                      "     insert into pma325 (prj_cd, cust_cd, rmks, cust_bc, cid, cdt, mid, mdt) " & _
                      "     select a.prj_cd, a.cust_cd, a.rmks, 'PM325100', " & Parameter.Login.RegId & ", getdate(), " & Parameter.Login.RegId & ", getdate() " & _
                      "     from pma320 a " & _
                      "     where a.prj_cd = '" & Ref_No1 & "' " & _
                      "end "
                Link.ExcuteQuery(SQL)
            End If

        ElseIf dSet.Tables(0).Rows(0).Item("appr_sw") = "BC210300" Then
            tAppr_Sw = "BC210300"   '최종결재
        Else
            tAppr_Sw = "BC210200"   '중간결재
        End If


        If Comp_Type() = "SSP" Then

            SQL = "update GWM100 set " & _
                  " before_update = '3'" & _
                  " where appr_no = '" & Appr_No & "'"
            Link.ExcuteQuery(SQL)

        End If


        ' 테이블 Update
        SQL = "exec GWAPPR_Mod1 3, " & _
              "                   '" & Form_Cd & "', " & _
              "                   '" & Appr_No & "', " & _
              "                   '" & Ref_No1 & "', " & _
              "                   '" & Ref_No2 & "', " & _
              "                   '" & Ref_No3 & "', " & _
              "                   '" & Ref_No4 & "', " & _
              "                   0, " & _
              "                   '" & tAppr_Sw & "', " & _
              "                    " & Parameter.Login.RegId
        Link.ExcuteQuery(SQL)


        If Comp_Type() = "SSP" Then
            SQL = "update GWM100 set " & _
                  " after_update = '3'" & _
                  " where appr_no = '" & Appr_No & "'"
            Link.ExcuteQuery(SQL)
        End If
       

        Approval = True
    End Function

    '승인 버튼
    Private Sub btnAppr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SQL As String

        If Appr_No = "" Then Exit Sub '결재번호 없을시 이후 쿼리 수행하지 않도록 처리 2021.03.17

        If Comp_Type() = "HK" Then
            If MsgBox(GetMessage("GWM100_VIEW_003", "선택한 결재를 승인하시겠습니까?"), vbQuestion + vbYesNo, GetMessage("COM_WRD_CHK")) = vbNo Then Exit Sub
        End If

        Select Case Appr_Bc
            Case "GW100500" '참조
                ' 결재상세테이블 승인처리
                SQL = "update GWM110 set " & _
                      " appr_sw = 'BC210300', " & _
                      " pms_sw = '1', " & _
                      " pms_dt = getdate(), " & _
                      " real_chrg = " & Parameter.Login.RegId & _
                      " where appr_no = '" & Appr_No & "'" & _
                      " and appr_chrg = " & Parameter.Login.RegId
                Link.ExcuteQuery(SQL)

            Case Else   '결재 또는 협조 등등...
                ' 이전결재자가 승인했는지 체크

                Dim dSet As DataSet = popup.OpenDataSet("gwm100_view_stat")

                If Get_PreChk() <> True Then

                    If Comp_Type() = "SSP" Then
                        If (Form_Cd = "FAH110_SP" Or Form_Cd = "FAH120_SP" Or Form_Cd = "GAT120_SP") And dSet.Tables(0).Rows(0).Item("stat_bc") = "1" Then
                            MessageInfo("이전 결재자가 결재처리를 하지 않았습니다.")
                            Exit Sub
                        Else
                            If MsgBox(GetMessage("GWM100_VIEW_001", "이전 결재자가 결재처리를 하지 않았습니다.<br> 결재를 진행하시겠습니까?"), vbQuestion + vbYesNo, GetMessage("COM_WRD_CHK")) = vbNo Then Exit Sub
                        End If

                    Else
                        MessageInfo("이전 결재자가 결재처리를 하지 않았습니다.")
                        Exit Sub
                    End If
                End If

                If Approval(1) <> True Then Exit Sub

                'E-Mail 전송 추가 (2019-05-22)
                'HK 시정조치요구(8D)_작성팀 -> 최종결재자 승인일 때 결재자 모두 메일 전송 (2020-08-24)
                'If Comp_Type() = "HC" Or Comp_Type() = "WTS" Or Comp_Type() = "SK" Or (Comp_Type() = "HK" And Form_Cd = "QMG200_HK") Then
                '    Me.mail_send("1")
                'End If

                '메일 전송 업체가 증가함에 따라 CASE 문으로 변경 (2020-09-11)
                Select Case Comp_Type()
                    Case "HC", "WTS", "SK", "JM", "DS", "HH"
                        Me.mail_send("1")

                    Case "HK"
                        If Form_Cd = "QMG200_HK" Then
                            Me.mail_send("1")
                        End If
                End Select
        End Select

        ''SMS 전송
        'If Comp_Type() = "SSP" Then
        '    Call SMS_SSP("1")
        'End If

        If Comp_Type() = "HK" Then

            Call_Form_Cd.Call_Refresh()

        End If


        Me.DialogResult = Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    '반려 버튼
    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SQL As String

        If Appr_No = "" Then Exit Sub '결재번호 없을시 이후 쿼리 수행하지 않도록 처리 2021.03.17

        If Comp_Type() = "HK" Then
            If MsgBox(GetMessage("GWM100_VIEW_004", "선택한 결재를 반려하시겠습니까?"), vbQuestion + vbYesNo, GetMessage("COM_WRD_CHK")) = vbNo Then Exit Sub
        End If

        ' 이전결재자가 승인했는지 체크
        '2019-08-16 이전승인자에 상관없이 반려처리 될 수 있도록 수정 SSP만
        If Comp_Type() <> "SSP" Then
            If Get_PreChk() <> True Then
                MessageInfo("이전 결재자가 결재처리를 하지 않았습니다.")
                Exit Sub
            End If
        End If

        '구매승인일 경우 발주 데이터가 생성되어 있다면 결재취소를 할수 없게 처리 2019-06-24 장민식
        If Form_Cd = "DMB105_SP" Then
            msql = "select bomreg_id as cnt " & _
                   " from mmb150 " & _
                   " /*where '" & Ref_No2 & "' like '%' + convert(varchar(20), bomreg_id) + '%' */" & _
                   " where bomreg_id in (select strVALUE from dbo.FN_SPLIT('" & Ref_No2 & "',','))" & _
                   " and bomreg_id <> 0 "
            dSet = Link.ExcuteQuery(msql)

            If IsEmpty(dSet) = False Then
                MsgBox(GetMessage("DMB105_SP_003"), , GetMessage("MSG_INFO")) '발주데이터가 생성된 항목입니다. 결재취소 할 수 없습니다.
                Exit Sub
            End If
        End If

        ' 결재상세테이블 반려처리
        SQL = "update GWM110 set " & _
              " appr_sw = 'BC210400', " & _
              " pms_sw = '1', " & _
              " pms_dt = getdate(), " & _
              " real_chrg = " & Parameter.Login.RegId & _
              " where appr_no = '" & Appr_No & "'" & _
              " and appr_chrg = " & Parameter.Login.RegId
        Link.ExcuteQuery(SQL)

        ' 결재마스터테이블 반려처리
        SQL = "update GWM100 set " & _
              " comp_sw = 'BC210400'," & _
              " conf_rtn = '1'" & _
              " where appr_no = '" & Appr_No & "'"
        Link.ExcuteQuery(SQL)

        ' 최종결재
        SQL = "exec GWAPPR_Mod1 3, " & _
              "                   '" & Form_Cd & "', " & _
              "                   '" & Appr_No & "', " & _
              "                   '" & Ref_No1 & "', " & _
              "                   '" & Ref_No2 & "', " & _
              "                   '" & Ref_No3 & "', " & _
              "                   '" & Ref_No4 & "', " & _
              "                   0, " & _
              "                   'BC210400', " & _
              "                    " & Parameter.Login.RegId & " "
        '"                    " & IIf(chk_re_Appr() = True, 1, 0)

        '타업체 에러로 인한 수정 (2021-07-21)
        If Comp_Type() = "SSP" Then
            SQL = SQL & ", " & IIf(chk_re_Appr() = True, 1, 0)
        End If

        Link.ExcuteQuery(SQL)

        ' 반려 이전결재자 새로운결재 확인여부 초기화
        SQL = "update b set " & _
              " b.new_sw = '1'" & _
              " from GWM100 as a" & _
              " inner join GWM110 as b on b.appr_no = a.appr_no" & _
              " where a.appr_no = '" & Appr_No & "'"
        Link.ExcuteQuery(SQL)

        'E-Mail 전송 추가 (2019-05-22)
        Select Case Comp_Type()
            Case "HC", "WTS", "SK", "JM", "HH"
                Me.mail_send("3")
        End Select

        ''SMS 전송
        'If Comp_Type() = "SSP" Then
        '    If MsgBox("문자발송 하시겠습니까?", vbQuestion + vbYesNo, "질의") = vbNo Then Exit Sub

        '    Call MOD_APPR.gFunSendAppr(2, Appr_No)
        'End If

        If Comp_Type() = "HK" Then

            Call_Form_Cd.Call_Refresh()

        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    '대결 버튼
    Private Sub btnDeputy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Appr_No = "" Then Exit Sub '결재번호 없을시 이후 쿼리 수행하지 않도록 처리 2021.03.17

        If Comp_Type() = "HK" Then
            If MsgBox(GetMessage("GWM100_VIEW_005", "선택한 결재를 대결하시겠습니까?"), vbQuestion + vbYesNo, GetMessage("COM_WRD_CHK")) = vbNo Then Exit Sub
        End If

        Select Case Comp_Type()
            Case "IVT", "WTS"      '이전/다음 결재자 대결 추가 (2020-09-14)
                Dim sMenu As New ContextMenuStrip

                sMenu = popup.DeputyMenuStrip

                '대결 소메뉴 보이기
                sMenu.Show()
                sMenu.Left = MousePosition.X
                sMenu.Top = MousePosition.Y

            Case Else
                Call _Deputy()
        End Select
    End Sub

    '이전 결재자 대결 버튼
    Private Sub DeputyMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        _DeputyBC = "1"
        Call _Deputy()
    End Sub

    '다음 결재자 대결 버튼
    Private Sub DeputyMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        _DeputyBC = ""
        Call _Deputy()
    End Sub

    '대결
    Private Sub _Deputy()
        ' 이전 결재자가 승인했는지 체크
        If Get_PreChk() <> True Then
            MsgBox(GetMessage("GWM100_VIEW_002", "이전 결재자가 결재처리를 하지 않았습니다."), , GetMessage("COM_WRD_CHK"))
            'MessageInfo("이전 결재자가 결재처리를 하지 않았습니다.")
            Exit Sub
        End If

        If Approval(2) <> True Then Exit Sub

        'E-Mail 전송 추가 (2019-05-22)
        'HK 시정조치요구(8D)_작성팀 -> 최종결재자 승인일 때 결재자 모두 메일 전송 (2020-08-24)
        'If Comp_Type() = "HC" Or Comp_Type() = "WTS" Or Comp_Type() = "SK" Or (Comp_Type() = "HK" And Form_Cd = "QMG200_HK") Then
        '    Me.mail_send("2")
        'End If

        '메일 전송 업체가 증가함에 따라 CASE 문으로 변경 (2020-09-11)
        Select Case Comp_Type()
            Case "HC", "WTS", "SK", "JM", "DS", "HH"
                Me.mail_send("2")

            Case "HK"
                If Form_Cd = "QMG200_HK" Then
                    Me.mail_send("2")
                End If
        End Select

        ''SMS 전송
        'If Comp_Type() = "SSP" Then
        '    Call SMS_SSP("2")
        'End If

        If Comp_Type() = "HK" Then

            Call_Form_Cd.Call_Refresh()

        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    '전결 버튼
    Private Sub btnAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If Appr_No = "" Then Exit Sub '결재번호 없을시 이후 쿼리 수행하지 않도록 처리 2021.03.17

        If Comp_Type() = "HK" Then
            If MsgBox(GetMessage("GWM100_VIEW_006", "선택한 결재를 전결하시겠습니까?"), vbQuestion + vbYesNo, GetMessage("COM_WRD_CHK")) = vbNo Then Exit Sub
        End If

        ' 이전결재자가 승인했는지 체크
        If Get_PreChk() <> True Then
            MsgBox(GetMessage("GWM100_VIEW_002", "이전 결재자가 결재처리를 하지 않았습니다."), , GetMessage("COM_WRD_CHK"))
            'MessageInfo("이전 결재자가 결재처리를 하지 않았습니다.")
            Exit Sub
        End If

        If Approval(3) <> True Then Exit Sub

        'E-Mail 전송 추가 (2019-05-22)
        'HK 시정조치요구(8D)_작성팀 -> 최종결재자 승인일 때 결재자 모두 메일 전송 (2020-08-24)
        'If Comp_Type() = "HC" Or Comp_Type() = "WTS" Or Comp_Type() = "SK" Or (Comp_Type() = "HK" And Form_Cd = "QMG200_HK") Then
        '    Me.mail_send("4")
        'End If

        '메일 전송 업체가 증가함에 따라 CASE 문으로 변경 (2020-09-11)
        Select Case Comp_Type()
            Case "HC", "WTS", "SK", "JM", "DS", "HH"
                Me.mail_send("4")

            Case "HK"
                If Form_Cd = "QMG200_HK" Then
                    Me.mail_send("4")
                End If
        End Select

        If Comp_Type() = "HK" Then

            Call_Form_Cd.Call_Refresh()

        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    '닫기 버튼
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    '저장 버튼
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SQL As String

        SQL = "update GWM110 set " & _
            " rmks = '" & Replace(popup.rmks2.Text, "'", "''") & "'" & _
            " where appr_no = '" & Appr_No & "'" & _
            " and appr_chrg = " & Parameter.Login.RegId
        Link.ExcuteQuery(SQL)

        If Comp_Type() = "SK" Or Comp_Type() = "JM" Or Comp_Type() = "HH" Then
            If PutYesNo("GWM100_VIEW_005") = MsgBoxResult.Yes Then
                Me.mail_send_rmks(popup.rmks2)
            End If

        ElseIf Comp_Type() = "DD" Then '대동철강의 경우 결재후에 쿼리를 날려준다. (발주서에 내용을 넣어줘야 함.) 2020-07-28. YANG 추가.
            Dim _p As New OpenParameters

            _p.Add("@appr_no", Appr_No)

            Base8.Link.ReadDataSet("GWM100_btn_save_after_DD", _p)

            Me.PopUp_VIEW_Load(Nothing, Nothing)
        End If

        Dim p1 As New OpenParameters
        p1.Add("@appr_no", Appr_No)
        popup.Open("gwm100_view_g20", p1)
    End Sub

    '출력2 버튼 (HL 전용) (2021-02-15)
    Private Sub btn_prt2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Comp_Type() <> "HL" Then Exit Sub

        '구매의뢰서 미리보기
        popup.PrintForm(Dev_dSet_Sub, "MMA100_HL_R")
    End Sub

    Private Sub g20_AfterMoveRow(ByVal sender As System.Object, ByVal PrevRowIndex As System.Int32, ByVal RowIndex As System.Int32)
        popup.rmks1.Text = popup.g20.Text("rmks", RowIndex)
    End Sub

    Private Sub TAB1_Selected(ByVal sender As System.Object, ByVal e As DevExpress.XtraTab.TabPageEventArgs)
        Dim SQL As String
        Dim dSet As DataSet

        If e.PageIndex = 1 Then

            SQL = "select rmks " & _
                  " from GWM110 " & _
                  " where appr_no = '" & Appr_No & "'" & _
                  " and appr_chrg = " & Parameter.Login.RegId

            dSet = Link.ExcuteQuery(SQL)

            If Not IsEmpty(dSet) Then
                popup.rmks2.Text = dSet.Tables(0).Rows(0).Item("rmks")
            End If


        End If
    End Sub

    '유효성검토 (생산일정표의 생산지시/이력사항 탭으로 점프)
    Private Sub btn_pms300_jump_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctr_prs As New Windows.Forms.Form
        Dim _p As OpenParameters = New OpenParameters

        _p.Clear()
        _p.Add("@appr_No", Appr_No)

        Dim ds As DataSet = Base8.Link.ReadDataSet("SDB100_SPP_get_so_info_appr", _p)

        If IsEmpty(ds) Then Return

        Dim _so_no As String = DataValue(ds, "so_no")

        Dim ctr As Object = New PMS300_SSP
        ctr_prs.Dispose()

        ctr_prs = New Windows.Forms.Form
        ctr_prs.Controls.Add(ctr)

        Dim size = Me.Size
        size.Width = size.Width ' - (size.Width / 4)
        ctr_prs.Size = SystemInformation.PrimaryMonitorMaximizedWindowSize


        ctr.Dock = DockStyle.Fill
        ctr_prs.Show()
        ctr_prs.Location = New System.Drawing.Point(0, 0)

        ctr.btn_popup_save.visible = True
        ctr.Fr_Jump_PMS312_SSP(_so_no)
    End Sub

    '구매승인 BOm등록 점프 추가 20190504
    Private Sub btn_dmb100_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SQL As String
        Dim dSet As DataSet
        Dim f_prd_id As String
        Dim f_prd_sq As String
        Dim f_prd_cd As String
        Dim f_prd_nm As String
        Dim f_plan_bc As String
        Dim f_so_no As String
        Dim f_Appr_No As String
        Dim ctr_prs As New Windows.Forms.Form

        SQL = "select top 1 isnull(b.prd_id, '') as prd_id, isnull(b.prd_sq, '') as prd_sq, isnull(c.itm_cd, '') as itm_cd, isnull(c.itm_nm, '') as itm_nm, " & _
              " isnull(dbo.fnbasenm(b.so_type, 1) + '/' + convert(varchar(10), b.so_dt, 23), '') as plan_bc, isnull(b.so_no, '') as so_no " & _
              " from dmb100 a " & _
              " inner join sdb100 b on a.prd_id = b.prd_id and a.prd_sq = b.prd_sq " & _
              " left join dma100 c on b.prd_id = c.itm_id " & _
              " where a.appr_no = '" & Appr_No & "'"

        dSet = Link.ExcuteQuery(SQL)

        f_prd_id = dSet.Tables(0).Rows(0).Item("prd_id")
        f_prd_sq = dSet.Tables(0).Rows(0).Item("prd_sq")
        f_prd_cd = dSet.Tables(0).Rows(0).Item("itm_cd")
        f_prd_nm = dSet.Tables(0).Rows(0).Item("itm_nm")
        f_plan_bc = dSet.Tables(0).Rows(0).Item("plan_bc")
        f_so_no = dSet.Tables(0).Rows(0).Item("so_no")
        f_Appr_No = Appr_No

        Dim ctr As Object = New DMB105_SP
        ctr_prs.Dispose()

        ctr_prs = New Windows.Forms.Form
        ctr_prs.Controls.Add(ctr)

        Dim size = Me.Size
        size.Width = size.Width ' - (size.Width / 4)
        ctr_prs.Size = size

        ctr.Dock = DockStyle.Fill
        ctr_prs.Show()
        ctr.Fr_Jump_GWM100_View(f_prd_id, f_prd_sq, f_prd_cd, f_prd_nm, f_plan_bc, f_so_no, f_Appr_No)


        'Dim ctr As Object = Parameter.MainFrame.Frame.CallMenuForm("DMB100_SP")
        'ctr.Fr_Jump_GWM100_View(f_prd_id, f_prd_sq, f_prd_cd, f_prd_nm, f_plan_bc, f_so_no)
    End Sub

    '2021.01.29 최경식
    'SSP(반려관련)
    Private Function chk_re_Appr() As Boolean
        chk_re_Appr = False

        Dim dSet As DataSet
        Dim is_reAppr As String = ""

        mSQL = " Select m10 " & _
               " From bca200v " & _
               " Where base_cd = 'BC210400'"
        dSet = Link.ExcuteQuery(mSQL)

        If IsEmpty(dSet) = False Then
            is_reAppr = ToStr(DataValue(dSet, "m10"))
        End If

        If is_reAppr = "1" Then
            chk_re_Appr = True
        End If

    End Function

    'E-Mail 전송 추가 (2019-05-22)
#Region "E-Mail"

    Private Sub mail_send(ByVal _mail_chk As String)

        Dim cnt As Integer = 0
        Dim totcnt As Integer = 0
        Dim emp As String = ""

        Dim p3 As OpenParameters = New OpenParameters()
        Dim p1 As OpenParameters = New OpenParameters()
        Dim p2 As OpenParameters = New OpenParameters()
        Dim dSet2 As Data.DataSet = Nothing
        Dim dSet1 As Data.DataSet = Nothing
        Dim dSet3 As Data.DataSet = Nothing
        Dim dSet As Data.DataSet = Nothing

        Dim log_appr As String = Appr_No
        Dim log_sub_ject As String = ""
        Dim log_to_mail As String = ""

        Dim tMsg As String = ""
        Dim mSQL As String = ""
        Dim mSQL1 As String = ""
        Dim mSQL2 As String = ""
        Dim mSQL3 As String = ""
        Dim tRow As Integer = 0

        Dim mHost_Nm As String = ""
        Dim mHost_Port As String = ""
        Dim mServer_ID As String = ""
        Dim mServer_PW As String = ""
        Dim mBase_Mail As String = ""
        Dim mSSL As String = ""
        Dim mCust_Mail As String = ""
        Dim email_adr As String = ""
        Dim excel_nm As String = ""
        Dim seq_no As String = ""
        Dim flienm As String = ""
        Dim _chk_mail As String = ""

        'DS 메일 전송 여부 체크 추가 (2020-09-21)
        If Comp_Type() = "DS" And GetMessage("SYS_APPR_MAIL") <> "1" Then
            Exit Sub
        End If

        mSQL = " Select SMTP_ServerNm, SMTP_Port, SMTP_MailNm, SMTP_MailPw, SMTP_BaseNm, SMTP_SSL " & _
               " From bcc200 " & _
               " Where bs_cd = '" & Login.BsCd & "'"
        dSet = Link.ExcuteQuery(mSQL)

        If IsEmpty(dSet) = False Then
            mHost_Nm = ToStr(DataValue(dSet, "SMTP_ServerNm"))
            mHost_Port = ToStr(DataValue(dSet, "SMTP_Port"))
            mServer_ID = ToStr(DataValue(dSet, "SMTP_MailNm"))
            mServer_PW = ToStr(DataValue(dSet, "SMTP_MailPw"))
            mBase_Mail = ToStr(DataValue(dSet, "SMTP_BaseNm"))
            mSSL = ToStr(DataValue(dSet, "SMTP_SSL"))
        Else
            'MsgBox("SMTP정보가 설정되어 있는지 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP_NOT"), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        End If

        If mHost_Nm = "" Then
            'MsgBox("SMTP정보의 SMTPServer 정보를 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP", "SMTPServer"), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        ElseIf mServer_ID = "" Then
            'MsgBox("SMTP정보의 메일계정 정보를 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP", GetMessage("COM_MSG_SMTP_MailNm")), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        ElseIf mServer_PW = "" Then
            'MsgBox("SMTP정보의 계정암호 정보를 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP", GetMessage("COM_MSG_SMTP_MailPw")), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        End If

        totcnt = totcnt + 1

        Try
            Dim Message As MailMessage = New MailMessage()
            'Dim smtp As New SmtpClient()
            '컴퓨터 이름이 한글인 컴퓨터는 이름을 MAIL로 지정하여 보낸다.
            Dim smtp As New Frame8_SmtpClient.SmtpClientEx("MAIL")
            Dim smtpUser As New System.Net.NetworkCredential()

            If mServer_ID = "" Or mServer_PW = "" Then
                MessageInfo("[보내는E-Mail / 패스워드]는 필수입니다. 확인해주세요")
                Exit Sub
            End If

            'BCC200정보
            smtpUser.UserName = mServer_ID
            smtpUser.Password = mServer_PW

            '보내는사람
            Message.From = New MailAddress(mBase_Mail)

            '받는사람 = 메일가져오기
            '받는 사람
            Dim SQL_R As String = ""
            Dim dSet_R As Data.DataSet = Nothing

            If _mail_chk = "1" Then '승인
                '받는 사람
                If Comp_Type() = "HK" Then
                    'HK 시정조치요구(8D)_작성팀 -> 최종결재자 승인일 때 결재자 모두 메일 전송 (2020-08-24)
                    SQL_R = "if exists (select * from gwm110  " & _
                            "           where appr_no = '" & Appr_No & "' " & _
                            "           and isnull(appr_sort, 0) = (select isnull(max(appr_sort), 0) " & _
                            "                                       from gwm110 " & _
                            "                                       where appr_no = '" & Appr_No & "' " & _
                            "                                       and appr_bc <> 'GW100500') " & _
                            "           and isnull(appr_bc, '') <> 'GW100500' " & _
                            "           and isnull(appr_sw, '') = 'BC210300') " & _
                            " begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            " end "

                ElseIf Comp_Type() = "DS" Then
                    'DS 최종 결재 시 결재자, 참조자 모두 메일 발송, 최종 결재 아닐시 결재자 모두 메일 발송 (2020-09-10)
                    SQL_R = "if exists (select * from gwm110  " & _
                            "           where appr_no = '" & Appr_No & "' " & _
                            "           and isnull(appr_sort, 0) = (select isnull(max(appr_sort), 0) " & _
                            "                                       from gwm110 " & _
                            "                                       where appr_no = '" & Appr_No & "' " & _
                            "                                       and appr_bc <> 'GW100500') " & _
                            "           and isnull(appr_bc, '') <> 'GW100500' " & _
                            "           and isnull(appr_sw, '') = 'BC210300') " & _
                            "begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            "end " & _
                            "else " & _
                            "begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            "   and isnull(a.appr_bc, '') = 'GW100100' " & _
                            "end "

                Else
                    SQL_R = "if (select count(*) from gwm110 " & _
                            "    where appr_no = '" & Appr_No & "' " & _
                            "    and isnull(appr_sort, 0) = " & Appr_Sort + 1 & _
                            "    and appr_bc <> 'GW100500' ) = 1 " & _
                            " begin " & _
                            "       select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg) , '1' as re_chk " & _
                            "       from gwm110 a " & _
                            "       where a.appr_no = '" & Appr_No & "' " & _
                            "       and isnull(a.appr_sort, 0) = " & Appr_Sort + 1 & _
                            "       and a.appr_bc <> 'GW100500' " & _
                            " end " & _
                            " else " & _
                            " begin " & _
                            "       select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg), '2' as re_chk  " & _
                            "       from gwm110 a " & _
                            "       where a.appr_no = '" & Appr_No & "' " & _
                            "       and ((a.appr_bc <> 'GW100500' and isnull(a.appr_sort, 0) = 0 ) or a.appr_bc = 'GW100500')  " & _
                            " end "
                End If

                dSet_R = Link.ExcuteQuery(SQL_R)

                If IsEmpty(dSet_R) = False Then
                    For Each dRow In dSet_R.Tables(0).Rows
                        If chk_email_match(ToStr(dRow("email"))) Then       '메일 형식 체크
                            re_chk = ToStr(dRow("re_chk"))
                            re_emp = ToStr(dRow("email"))
                            Message.To.Add(New MailAddress(re_emp))
                        End If
                    Next
                Else
                    re_chk = ""
                    re_emp = ""
                End If

            ElseIf _mail_chk = "2" Then '대결
                '받는 사람
                If Comp_Type() = "HK" Then
                    'HK 시정조치요구(8D)_작성팀 -> 최종결재자 승인일 때 결재자 모두 메일 전송 (2020-08-24)
                    SQL_R = "if exists (select * from gwm110  " & _
                            "           where appr_no = '" & Appr_No & "' " & _
                            "           and isnull(appr_sort, 0) = (select isnull(max(appr_sort), 0) " & _
                            "                                       from gwm110 " & _
                            "                                       where appr_no = '" & Appr_No & "' " & _
                            "                                       and appr_bc <> 'GW100500') " & _
                            "           and isnull(appr_bc, '') <> 'GW100500' " & _
                            "           and isnull(appr_sw, '') = 'BC210300') " & _
                            " begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            " end "

                ElseIf Comp_Type() = "DS" Then
                    'DS 최종 결재 시 결재자, 참조자 모두 메일 발송, 최종 결재 아닐시 결재자 모두 메일 발송 (2020-09-10)
                    SQL_R = "if exists (select * from gwm110  " & _
                            "           where appr_no = '" & Appr_No & "' " & _
                            "           and isnull(appr_sort, 0) = (select isnull(max(appr_sort), 0) " & _
                            "                                       from gwm110 " & _
                            "                                       where appr_no = '" & Appr_No & "' " & _
                            "                                       and appr_bc <> 'GW100500') " & _
                            "           and isnull(appr_bc, '') <> 'GW100500' " & _
                            "           and isnull(appr_sw, '') = 'BC210300') " & _
                            "begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            "end " & _
                            "else " & _
                            "begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            "   and isnull(a.appr_bc, '') = 'GW100100' " & _
                            "end "

                Else
                    SQL_R = "if (select count(*) from gwm110 " & _
                           "         where appr_no = '" & Appr_No & "' " & _
                           "         and isnull(appr_sort, 0) = " & Appr_Sort + 2 & _
                           "         and appr_bc <> 'GW100500' ) = 1 " & _
                           " begin " & _
                           "        select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg), '1' as re_chk  " & _
                           "        from gwm110 a " & _
                           "        where a.appr_no = '" & Appr_No & "' " & _
                           "        and isnull(a.appr_sort, 0) = " & Appr_Sort + 2 & _
                           "        and a.appr_bc <> 'GW100500' " & _
                           " end " & _
                           " else " & _
                           " begin " & _
                           "        select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg), '2' as re_chk  " & _
                           "        from gwm110 a " & _
                           "        where a.appr_no = '" & Appr_No & "' " & _
                           "        and isnull(a.appr_sort, 0) = 0 " & _
                           "        and a.appr_bc <> 'GW100500' " & _
                           " end "
                End If

                dSet_R = Link.ExcuteQuery(SQL_R)

                If IsEmpty(dSet_R) = False Then
                    For Each dRow In dSet_R.Tables(0).Rows
                        If chk_email_match(ToStr(dRow("email"))) Then       '메일 형식 체크
                            re_chk = ToStr(dRow("re_chk"))
                            re_emp = ToStr(dRow("email"))
                            Message.To.Add(New MailAddress(re_emp))
                        End If
                    Next
                Else
                    re_chk = ""
                    re_emp = ""
                End If

            ElseIf _mail_chk = "3" Then '반송
                '받는 사람
                SQL_R = "select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg), '3' as re_chk " & _
                        "from gwm110 a " & _
                        "where a.appr_no = '" & Appr_No & "'" & _
                        "and isnull(a.appr_sort, 0) = 0 " & _
                        "and a.appr_bc <> 'GW100500' "
                dSet_R = Link.ExcuteQuery(SQL_R)

                If IsEmpty(dSet_R) = False Then
                    If chk_email_match(DataValue(dSet_R, "email")) Then       '메일 형식 체크
                        re_chk = DataValue(dSet_R, "re_chk")
                        re_emp = DataValue(dSet_R, "email")
                        Message.To.Add(New MailAddress(re_emp))
                    End If
                Else
                    re_chk = ""
                    re_emp = ""
                End If

            ElseIf _mail_chk = "4" Then '전결
                '받는 사람
                If Comp_Type() = "HK" Then      'HK 시정조치요구(8D)_작성팀 -> 최종결재자 승인일 때 결재자 모두 메일 전송 (2020-08-24)
                    SQL_R = "if exists (select * from gwm110  " & _
                            "           where appr_no = '" & Appr_No & "' " & _
                            "           and isnull(appr_sort, 0) = (select isnull(max(appr_sort), 0) " & _
                            "                                       from gwm110 " & _
                            "                                       where appr_no = '" & Appr_No & "' " & _
                            "                                       and appr_bc <> 'GW100500') " & _
                            "           and isnull(appr_bc, '') <> 'GW100500' " & _
                            "           and isnull(appr_sw, '') = 'BC210300') " & _
                            " begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            " end "

                ElseIf Comp_Type() = "DS" Then
                    'DS 최종 결재 시 결재자, 참조자 모두 메일 발송, 최종 결재 아닐시 결재자 모두 메일 발송 (2020-09-10)
                    SQL_R = "if exists (select * from gwm110  " & _
                            "           where appr_no = '" & Appr_No & "' " & _
                            "           and isnull(appr_sort, 0) = (select isnull(max(appr_sort), 0) " & _
                            "                                       from gwm110 " & _
                            "                                       where appr_no = '" & Appr_No & "' " & _
                            "                                       and appr_bc <> 'GW100500') " & _
                            "           and isnull(appr_bc, '') <> 'GW100500' " & _
                            "           and isnull(appr_sw, '') = 'BC210300') " & _
                            "begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            "end " & _
                            "else " & _
                            "begin " & _
                            "   select c.email, '2' as re_chk " & _
                            "   from gwm110 a " & _
                            "       left join scu100 b on a.appr_chrg = b.reg_id " & _
                            "       left join hra150 c on b.emp_no = c.emp_no " & _
                            "   where a.appr_no = '" & Appr_No & "' " & _
                            "   and isnull(c.email, '') <> '' " & _
                            "   and isnull(a.appr_bc, '') = 'GW100100' " & _
                            "end "

                Else
                    SQL_R = " select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg), '2' as re_chk " & _
                            " from gwm110 a " & _
                            " where a.appr_no = '" & Appr_No & "' " & _
                            " and isnull(a.appr_sort, 0) = 0 " & _
                            " and a.appr_bc <> 'GW100500' "
                End If

                dSet_R = Link.ExcuteQuery(SQL_R)

                If IsEmpty(dSet_R) = False Then
                    For Each dRow In dSet_R.Tables(0).Rows
                        If chk_email_match(ToStr(dRow("email"))) Then       '메일 형식 체크
                            re_chk = ToStr(dRow("re_chk"))
                            re_emp = ToStr(dRow("email"))
                            Message.To.Add(New MailAddress(re_emp))
                        End If
                    Next
                Else
                    re_chk = ""
                    re_emp = ""
                End If
            End If

            '바디 내용 가져오기
            mSQL3 = ""
            mSQL3 = " exec GWAPPR_MAIL '0','" & Login.ID & "','" & Appr_No & "'"
            dSet3 = Link.ExcuteQuery(mSQL3)

            Message.IsBodyHtml = True
            Message.Priority = MailPriority.High

            Select Case Comp_Type()
                Case "HC"
                    If re_chk = "1" Then  '결재중
                        Message.Subject = " [힐세리온 결재 : 결재처리]  " + Appr_No
                    ElseIf re_chk = "2" Then '승인
                        Message.Subject = " [힐세리온 결재 : 결재완료]  " + Appr_No
                    ElseIf re_chk = "3" Then '반려
                        Message.Subject = " [힐세리온 결재 : 반송건]  " + Appr_No
                    End If

                Case "SK"
                    If re_chk = "1" Then  '결재중
                        Message.Subject = " [수경화학 결재 : 결재처리]  " + Appr_No
                    ElseIf re_chk = "2" Then '승인
                        Message.Subject = " [수경화학 결재 : 결재완료]  " + Appr_No
                    ElseIf re_chk = "3" Then '반려
                        Message.Subject = " [수경화학 결재 : 반송건]  " + Appr_No
                    End If

                Case "JM"
                    If re_chk = "1" Then  '결재중
                        Message.Subject = " [진명파워텍 결재 : 결재처리]  " + Appr_No
                    ElseIf re_chk = "2" Then '승인
                        Message.Subject = " [진명파워텍 결재 : 결재완료]  " + Appr_No
                    ElseIf re_chk = "3" Then '반려
                        Message.Subject = " [진명파워텍 결재 : 반송건]  " + Appr_No
                    End If

                Case "WTS"
                    If re_chk = "1" Then  '결재중
                        Message.Subject = " [위더스 결재 : 결재처리]  " + Appr_No
                    ElseIf re_chk = "2" Then '승인
                        Message.Subject = " [위더스 결재 : 결재완료]  " + Appr_No
                    ElseIf re_chk = "3" Then '반려
                        Message.Subject = " [위더스 결재 : 반송건]  " + Appr_No
                    End If

                Case "HK"
                    Dim title As String = ""

                    '시정조치요구서(8D) 제목 가져오기
                    SQL_R = " select isnull(title, '') as title " & _
                            " from qmg200 " & _
                            " where appr_no = '" & Appr_No & "' "
                    dSet_R = Link.ExcuteQuery(SQL_R)

                    If IsEmpty(dSet_R) = False Then
                        title = DataValue(dSet_R)
                    Else
                        title = ""
                    End If

                    If re_chk = "2" Then  '결재중
                        Message.Subject = " [시정조치요구서] " + Appr_No + " / " + title
                    End If

                Case Else
                    If re_chk = "1" Then  '결재중
                        Message.Subject = " [전자결재 : 결재처리]  " + Appr_No
                    ElseIf re_chk = "2" Then '승인
                        Message.Subject = " [전자결재 : 결재완료]  " + Appr_No
                    ElseIf re_chk = "3" Then '반려
                        Message.Subject = " [전자결재 : 반송건]  " + Appr_No
                    End If
            End Select

            log_to_mail = Message.To.ToString
            log_sub_ject = Message.Subject

            Message.SubjectEncoding = System.Text.Encoding.UTF8
            Message.Body = DataValue(dSet3, "title")

            Try

            Catch ex As Exception
                MessageError(ex)
                Me.log_update(log_appr, log_sub_ject, log_to_mail, ex.ToString, "F")
            End Try

            Message.BodyEncoding = System.Text.Encoding.UTF8

            smtp.Host = mHost_Nm
            smtp.Port = mHost_Port
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.UseDefaultCredentials = True
            smtp.Credentials = smtpUser

            '받는 사람이 있으면 메일 전송 (2020-09-14)
            If Message.To.Count > 0 Then
                smtp.Send(Message)
                Me.log_update(log_appr, log_sub_ject, log_to_mail, "", "S")
            End If

            Message.Attachments.Dispose()

        Catch ex As Exception
            'MessageInfo(ex.ToString())
            Me.log_update(log_appr, log_sub_ject, log_to_mail, ex.ToString, "F")
            cnt = cnt + 1
        End Try

        If Comp_Type() <> "HK" Then
            If cnt > 0 Then
                Dim err As String = "메일(" & cnt & ")건 발송이 실패하였습니다."
                err = err + emp
                MessageInfo(err)
            End If
        End If
    End Sub

    'E-Mail 로그
    Private Sub log_update(ByVal log_appr As String, ByVal log_sub_ject As String, ByVal log_to_mail As String, ByVal log_msg As String, ByVal log_send As String)
        Dim log_mSQL As String = ""

        log_mSQL = " insert into GWM100_mail_LOG (log_appr_no, log_sub_ject, log_to_mail, log_msg, log_send, cid, cdt) " & _
                   " values ('" & log_appr & "', '" & Microsoft.VisualBasic.Left(log_sub_ject, 250) & "', '" & Microsoft.VisualBasic.Left(log_to_mail, 250) & "', '" & _
                             Microsoft.VisualBasic.Left(log_msg, 1000) & "', '" & log_send & "', " & Parameter.Login.RegId & ", getdate()) "
        Link.ExcuteQuery(log_mSQL)
    End Sub

    'E-Mail 형식 체크 (2020-09-14)
    Function chk_email_match(ByVal email As String) As Boolean
        Dim _email As New Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")

        If _email.IsMatch(email) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "E-Mail(의견작성) 수경"

    Private Sub mail_send_rmks(ByVal send_rmks As eMemo)

        Dim cnt As Integer = 0
        Dim totcnt As Integer = 0
        Dim emp As String = ""

        Dim p3 As OpenParameters = New OpenParameters()
        Dim p1 As OpenParameters = New OpenParameters()
        Dim p2 As OpenParameters = New OpenParameters()
        Dim dSet2 As Data.DataSet = Nothing
        Dim dSet1 As Data.DataSet = Nothing
        Dim dSet3 As Data.DataSet = Nothing
        Dim dSet As Data.DataSet = Nothing

        Dim log_appr As String = Appr_No
        Dim log_sub_ject As String = ""
        Dim log_to_mail As String = ""

        Dim tMsg As String = ""
        Dim mSQL As String = ""
        Dim mSQL1 As String = ""
        Dim mSQL2 As String = ""
        Dim mSQL3 As String = ""
        Dim tRow As Integer = 0

        Dim mHost_Nm As String = ""
        Dim mHost_Port As String = ""
        Dim mServer_ID As String = ""
        Dim mServer_PW As String = ""
        Dim mBase_Mail As String = ""
        Dim mSSL As String = ""
        Dim mCust_Mail As String = ""
        Dim email_adr As String = ""
        Dim excel_nm As String = ""
        Dim seq_no As String = ""
        Dim flienm As String = ""
        Dim _chk_mail As String = ""

        mSQL = " Select SMTP_ServerNm, SMTP_Port, SMTP_MailNm, SMTP_MailPw, SMTP_BaseNm, SMTP_SSL " & _
               " From bcc200 " & _
               " Where bs_cd = '" & Login.BsCd & "'"
        dSet = Link.ExcuteQuery(mSQL)

        If IsEmpty(dSet) = False Then
            mHost_Nm = ToStr(DataValue(dSet, "SMTP_ServerNm"))
            mHost_Port = ToStr(DataValue(dSet, "SMTP_Port"))
            mServer_ID = ToStr(DataValue(dSet, "SMTP_MailNm"))
            mServer_PW = ToStr(DataValue(dSet, "SMTP_MailPw"))
            mBase_Mail = ToStr(DataValue(dSet, "SMTP_BaseNm"))
            mSSL = ToStr(DataValue(dSet, "SMTP_SSL"))
        Else
            'MsgBox("SMTP정보가 설정되어 있는지 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP_NOT"), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        End If

        If mHost_Nm = "" Then
            'MsgBox("SMTP정보의 SMTPServer 정보를 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP", "SMTPServer"), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        ElseIf mServer_ID = "" Then
            'MsgBox("SMTP정보의 메일계정 정보를 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP", GetMessage("COM_MSG_SMTP_MailNm")), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        ElseIf mServer_PW = "" Then
            'MsgBox("SMTP정보의 계정암호 정보를 확인바랍니다.")
            MsgBox(GetMessage("COM_MSG_SMTP", GetMessage("COM_MSG_SMTP_MailPw")), vbInformation + vbOKOnly, GetMessage("SYS_MSGBOX_TITLE"))
            Exit Sub
        End If

        totcnt = totcnt + 1

        Try
            Dim Message As MailMessage = New MailMessage()
            Dim smtp As New SmtpClient()
            Dim smtpUser As New System.Net.NetworkCredential()

            If mServer_ID = "" Or mServer_PW = "" Then
                MessageInfo("[보내는E-Mail / 패스워드]는 필수입니다. 확인해주세요")
                Exit Sub
            End If

            'BCC200정보
            smtpUser.UserName = mServer_ID
            smtpUser.Password = mServer_PW

            '보내는사람
            Message.From = New MailAddress(mBase_Mail)

            '받는사람 = 메일가져오기
            '받는 사람
            Dim SQL_R As String = ""
            Dim dSet_R As Data.DataSet = Nothing


            '받는(사람) 기안자에게 알림 메일 전송
            SQL_R = "       select email =  (select h.email from scu100 s  left join hra150 h on s.emp_no = h.emp_no  where s.reg_id = a.appr_chrg) , '1' as re_chk " & _
                    "       from gwm110 a " & _
                    "       where a.appr_no = '" & Appr_No & "' " & _
                    "       and a.appr_bc <> 'GW100500' "

            dSet_R = Link.ExcuteQuery(SQL_R)

            If IsEmpty(dSet_R) = False Then
                For Each dRow In dSet_R.Tables(0).Rows
                    re_chk = DataValue(dSet_R, "re_chk")
                    re_emp = DataValue(dSet_R, "email")
                    Message.To.Add(New MailAddress(re_emp))
                Next
            Else
                re_chk = ""
                re_emp = ""
            End If


            Message.IsBodyHtml = True
            Message.Priority = MailPriority.High

            Select Case Comp_Type()

                Case "SK", "JM", "HH"
                    Message.Subject = " 결재의견(" + Login.EmpNm + ") " + Appr_No
            End Select

            log_to_mail = Message.To.ToString
            log_sub_ject = Message.Subject

            Message.SubjectEncoding = System.Text.Encoding.UTF8
            Message.Body = Replace(send_rmks.Text, Environment.NewLine, "<br>")

            Try

            Catch ex As Exception
                MessageError(ex)
                Me.log_update(log_appr, log_sub_ject, log_to_mail, ex.ToString, "F")
            End Try

            Message.BodyEncoding = System.Text.Encoding.UTF8

            smtp.Host = mHost_Nm
            smtp.Port = mHost_Port
            smtp.DeliveryMethod = SmtpDeliveryMethod.Network
            smtp.UseDefaultCredentials = True
            smtp.Credentials = smtpUser
            smtp.Send(Message)
            Me.log_update(log_appr, log_sub_ject, log_to_mail, "", "S")

            Message.Attachments.Dispose()

        Catch ex As Exception
            'MessageInfo(ex.ToString())
            Me.log_update(log_appr, log_sub_ject, log_to_mail, ex.ToString, "F")
            cnt = cnt + 1
        End Try

        If cnt > 0 Then
            Dim err As String = "메일(" & cnt & ")건 발송이 실패하였습니다."
            err = err + emp
            MessageInfo(err)
        End If
    End Sub

#End Region

#Region "FTP관리"

    Private Sub g30_ButtonPressed(ByVal sender As Object, ByVal columnName As String)
        Dim FileID As String = popup.g30.Text("file_id")
        Dim FileNm As String = popup.g30.Text("file_nm")

        If FileNm = "" Then
            Exit Sub
        End If

        Select Case columnName
            Case "show"
                'FTPShared.FileDownLoad(FileID, FileNm, , True, Form_Cd)
                FTPShared.FileDownLoad(FileID, FileNm, , True, Form_Cd, "c:\Temp_Invite")
            Case "down"
                FTPShared.FileDownLoad(FileID, FileNm, , False, Form_Cd)
        End Select
    End Sub

    Private Sub g31_ButtonPressed(ByVal sender As Object, ByVal columnName As String)
        Dim FileID As String = popup.g31.Text("file_id")
        Dim FileNm As String = popup.g31.Text("file_nm")

        If FileNm = "" Then
            Exit Sub
        End If

        Select Case columnName
            Case "show"
                'FTPShared.FileDownLoad(FileID, FileNm, , True, Form_Cd)
                FTPShared.FileDownLoad(FileID, FileNm, , True, "MMA100_HL", "c:\Temp_Invite")
            Case "down"
                FTPShared.FileDownLoad(FileID, FileNm, , False, "MMA100_HL")
        End Select
    End Sub

#End Region

#Region "SMS_SSP"
    Private Sub SMS_SSP(ByVal gubun As String)

        'Dim sSql As String = ""
        'Dim sSql2 As String = ""
        'Dim dSet As DataSet = Nothing
        'Dim cnt As Long = 0
        'Dim cnt2 As Long = 0
        'Dim CompSW As String = ""
        'Dim FinalSort As String = ""
        'Dim NextID As String = ""

        'If Comp_Type() <> "SSP" Then Exit Sub

        ''최종결재자 순번 체크
        'sSql = "select max(appr_sort) as FinalSort  from gwm110 " & _
        '       "where appr_no = '" & Appr_No & "' " & _
        '       "and appr_bc <> 'GW100500'"
        'dSet = Link.ExcuteQuery(sSql)

        'If Not IsEmpty(dSet) Then
        '    FinalSort = DataValue(dSet, "FinalSort")
        'Else
        '    Exit Sub
        'End If

        ''승인된 결재자 Count
        'sSql = "select isnull(count(*), 0) as cnt  from gwm110 " & _
        '       "where appr_no = '" & Appr_No & "' " & _
        '       "and appr_sw = 'BC210300'"

        ''모든 결재자 Count
        'sSql2 = "select isnull(count(*), 0) as cnt, (" & sSql & ") as cnt2, max(b.comp_sw) as comp_sw " & _
        '        "from gwm110 a " & _
        '        "   inner join gwm100 b on a.appr_no = b.appr_no " & _
        '        "where a.appr_no = '" & Appr_No & "' "
        'dSet = Link.ExcuteQuery(sSql2)

        'If Not IsEmpty(dSet) Then
        '    cnt = DataValue(dSet, "cnt")
        '    cnt2 = DataValue(dSet, "cnt2")
        '    CompSW = DataValue(dSet, "comp_sw")
        'End If

        'If gubun = "1" Then
        '    '모든 결재자 Count = 승인된 결재자 Count or 결재자 순번 = 최종결재자 순번이면 '결재완료' SMS
        '    If (cnt = cnt2 Or Appr_Sort = FinalSort) And CompSW <> "BC100300" Then
        '        Call MOD_APPR.gFunSendAppr(3, Appr_No)

        '    ElseIf CompSW <> "BC100300" Then
        '        Call MOD_APPR.gFunSendAppr(1, Appr_No, Login.RegId)
        '    End If

        'ElseIf gubun = "2" Then
        '    '다음결재자 ID 체크
        '    sSql = "select top 1 appr_chrg  from gwm110 " & _
        '           "where appr_no = '" & Appr_No & "' " & _
        '           "and appr_sort = " & Appr_Sort & " + 1" & _
        '           "order by appr_sort "
        '    dSet = Link.ExcuteQuery(sSql)

        '    If Not IsEmpty(dSet) Then
        '        NextID = DataValue(dSet, "appr_chrg")
        '    End If

        '    '모든 결재자 Count = 승인된 결재자 Count or 결재자 순번 = 최종결재자 순번이면 '결재완료' SMS
        '    If (cnt = cnt2 Or Appr_Sort = FinalSort) Then

        '    Else
        '        Call MOD_APPR.gFunSendAppr(1, Appr_No, NextID)
        '    End If
        'End If
    End Sub
#End Region

End Class