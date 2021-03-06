<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase
'*  2. Function Name        :
'*  3. Program ID           : D1211MA1
'*  4. Program Name         : 전자계산서 발행(구매) - 도큐빌 
'*  5. Program Desc         : 전자계산서에 대하여 발행 또는 발행취소하는 기능 
'*  6. Component List       :
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 2009/10/31
'*  9. Modifier (First)     : Lee MIn Hyung
'* 10. Modifier (Last)      : Chon, Jaehyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBSCRIPT">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim iDBSYSDate


iDBSYSDate = "<%=GetSvrDate%>"

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>
End Sub

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "D1213MB1.asp"
Const BIZ_PGM_ID2 = "D1213MB2.asp"
Const BIZ_PGM_ID3 = "D1213MB3.asp"
Const BIZ_PGM_ID4 = "D1213MB4.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'=                       4.2 Constant variables
'========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"

'==========================================================================================================

'add header datatable column
Dim C1_send_check
Dim C1_proc_flag_nm
Dim C1_is_send_nm
Dim C1_where_flag
Dim C1_tax_bill_no
Dim C1_tax_doc_no
Dim C1_inv_no
Dim C1_inv_amend_type
Dim C1_inv_amend_type_nm
Dim C1_remark
Dim C1_remark2
Dim C1_remark3
Dim C1_tax_bill_type_nm
Dim C1_bp_cd
Dim C1_bp_nm
Dim C1_issued_dt
Dim C1_vat_calc_type_nm
Dim C1_vat_inc_flag_nm
Dim C1_vat_type
Dim C1_vat_type_nm
Dim C1_vat_rate
Dim C1_cur
Dim C1_total_amt
Dim C1_fi_total_amt
Dim C1_net_amt
Dim C1_fi_net_amt
Dim C1_vat_amt
Dim C1_fi_vat_amt
Dim C1_total_loc_amt
Dim C1_fi_total_loc_amt
Dim C1_net_loc_amt
Dim C1_fi_net_loc_amt
Dim C1_vat_loc_amt
Dim C1_fi_vat_loc_amt
Dim C1_report_biz_area
Dim C1_tax_biz_area_nm
Dim C1_sales_grp
Dim C1_sales_grp_nm
Dim C1_remarks
Dim C1_disuse_reason
Dim C1_legacy_pk
Dim C1_sale_no

Dim C1_proc_flag
Dim C1_is_send
Dim C1_issue_dt_fg

Dim C1_name
Dim C1_name_B
Dim C1_email
Dim C1_bp_seq
Dim C1_name2
Dim C1_name2_B
Dim C1_email2
Dim C1_bp_seq2

'add detail datatable column
Dim	C2_item_cd
Dim	C2_item_nm
Dim	C2_spec
Dim	C2_bill_qty
Dim	C2_bill_unit
Dim	C2_bill_price
Dim	C2_total_amt
Dim	C2_bill_amt
Dim	C2_vat_amt
Dim	C2_total_amt_loc
Dim	C2_bill_amt_loc
Dim	C2_vat_amt_loc
Dim	C2_bill_no
Dim	C2_bill_seq
Dim	C2_tax_bill_no


Dim docStatusMake   '작성 
Dim docStatusIssue 	'발행 
Dim docStatusReject  	' 반려 
Dim docStatusRequestDisuse  ' 폐기요청 
Dim docStatusCancelDisuse 	' 폐기취소 
Dim docStatusDisuse  	' 폐기 
Dim docStatusDelete  	' 삭제 

Dim sendStatusFail


docStatusMake = "T0" '작성 
docStatusIssue = "10"	'발행 
docStatusReject = "60"	' 반려 
docStatusRequestDisuse = "81"	' 폐기요청 
docStatusCancelDisuse = "82"	' 폐기취소 
docStatusDisuse = "90"	' 폐기 
docStatusDelete = "D0"	' 삭제 

sendStatusFail = "2" '실패 

Dim lgStrPrevKeyTempGlNo
Dim lgStrPrevKeyTempGlDt
Dim lgQueryFlag					' 신규조회 및 추가조회 구분 Flag
Dim lgGridPoupMenu              ' Grid Popup Menu Setting
Dim lgAllSelect

Dim lgIsOpenPop
Dim IsOpenPop
Dim lgPageNo_B
Dim lgSortKey_B
Dim lgOldRow1
Dim lgOldRow, lgRow
Dim lgSortKey1, lgSortKey2
Dim IntRetCD

Const C_MaxKey = 3
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
'=========================================================================================================
'=========================================================================================================
Sub Form_Load()
   Call LoadInfTB19029

   With frm1
      Call FormatDATEField(.txtIssuedFromDt)
      Call LockObjectField(.txtIssuedFromDt,"R")
      Call FormatDATEField(.txtIssuedToDt)
      Call LockObjectField(.txtIssuedToDt, "R")
      Call InitSpreadSheet()
      Call InitSpreadSheet2

      Call SetDefaultVal
      Call InitVariables
      Call InitComboBox()
      Call InitSpreadComboBox

      Call SetToolbar("110000000000111")										'⊙: 버튼 툴바 제어 

      .txtSupplierCd.focus
      .btnResend.disabled	= false
      .btnPublish.disabled	= false
      .btnApprovalDisuse.disabled	= false
      .btnCancelDisuse.disabled	= false
   End With
End Sub

'=========================================================================================================
Sub InitComboBox()
   Dim iCodeArr
   Dim iNameArr
   Dim iDx

   '자료유형(Data Type)
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTaxDocumentType, lgF0, lgF1, Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTransmitStatus, lgF0, lgF1, Chr(11))



End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
    C1_send_check        = 1
    C1_proc_flag_nm      = 2
    C1_is_send_nm        = 3
    C1_where_flag        = 4
    C1_tax_bill_no       = 5
    C1_tax_doc_no        = 6
    C1_inv_no            = 7
    C1_inv_amend_type    = 8
    C1_inv_amend_type_nm = 9
    C1_remark            = 10
    C1_remark2           = 11
    C1_remark3           = 12
    C1_tax_bill_type_nm  = 13
    C1_bp_cd             = 14
    C1_bp_nm             = 15
    C1_issued_dt         = 16
    C1_vat_calc_type_nm  = 17
    C1_vat_inc_flag_nm   = 18
    C1_vat_type          = 19
    C1_vat_type_nm       = 20
    C1_vat_rate          = 21
    C1_cur               = 22
    C1_total_amt         = 23
    C1_fi_total_amt      = 24
    C1_net_amt           = 25
    C1_fi_net_amt        = 26
    C1_vat_amt           = 27
    C1_fi_vat_amt        = 28
    C1_total_loc_amt     = 29
    C1_fi_total_loc_amt  = 30
    C1_net_loc_amt       = 31
    C1_fi_net_loc_amt    = 32
    C1_vat_loc_amt       = 33
    C1_fi_vat_loc_amt    = 34
    C1_report_biz_area   = 35
    C1_tax_biz_area_nm   = 36
    C1_sales_grp         = 37
    C1_sales_grp_nm      = 38
    C1_remarks           = 39
    C1_disuse_reason     = 40
    C1_legacy_pk         = 41
    C1_sale_no           = 42

    C1_proc_flag         = 43
    C1_is_send           = 44
    C1_issue_dt_fg       = 45

    C1_name              = 46
    C1_name_B            = 47
    C1_email             = 48
    C1_bp_seq            = 49
    C1_name2             = 50
    C1_name2_B           = 51
    C1_email2            = 52
    C1_bp_seq2           = 53
End Sub

Sub InitSpreadPosVariables2()
	'add tab1 detail datatable column
	C2_item_cd = 1
	C2_item_nm = 2
	C2_spec = 3
	C2_bill_qty = 4
	C2_bill_unit = 5
	C2_bill_price = 6

	C2_total_amt = 7
	C2_bill_amt = 8
	C2_vat_amt = 9
	C2_total_amt_loc = 10
	C2_bill_amt_loc = 11
	C2_vat_amt_loc = 12
	C2_bill_no = 13
	C2_bill_seq = 14
	C2_tax_bill_no = 15

End Sub

'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size

    lgStrPrevKeyTempGlNo = ""							'initializes Previous Key
    lgLngCurRows = 0									   'initializes Deleted Rows Count

    lgPageNo_B		= ""                          'initializes Previous Key for spreadsheet #2
    lgSortKey_B	= "1"

    lgOldRow = 0
    lgRow = 0
End Sub

'=========================================================================================================
Sub SetDefaultVal()
	'승인의 일자는 당일의 일자만 조회한다.
    Dim EndDate
	EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

	'승인의 일자는 당일 ~ 당일 이다.
	frm1.txtIssuedFromDt.text  = EndDate
	frm1.txtIssuedToDt.text    = EndDate
	frm1.txtSupplierCd.focus
	'lgGridPoupMenu          = GRID_POPUP_MENU_PRT
End Sub

'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData
		.MaxCols = C1_bp_seq2 + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols												'☆: 사용자 별 Hidden Column
		.ColHidden = True
		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20090708",, parent.gAllowDragDropSpread

		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetCheck	C1_send_check,			"선택",     			4,  -10, "", True, -1
		ggoSpread.SSSetEdit  	C1_proc_flag_nm, 		"계산서상태", 			10
		ggoSpread.SSSetEdit  	C1_is_send_nm, 			"송신상태", 				8
		ggoSpread.SSSetEdit  	C1_where_flag, 			"업무구분", 				8
		ggoSpread.SSSetEdit  	C1_tax_bill_no,			"세금계산서관리번호",		15

		ggoSpread.SSSetEdit  	C1_tax_doc_no,			"세금계산서번호",			15
        ggoSpread.SSSetEdit     C1_inv_no           ,   "계산서번호",              15
        ggoSpread.SSSetCombo    C1_inv_amend_type   ,   "수정사유",                15
        ggoSpread.SSSetCombo    C1_inv_amend_type_nm,   "수정사유명",              15
        ggoSpread.SSSetEdit     C1_remark           ,   "비고1",                   15, ,,150

        ggoSpread.SSSetEdit     C1_remark2          ,   "비고2",                   15, ,,150
        ggoSpread.SSSetEdit     C1_remark3          ,   "비고3",                   15, ,,150
		ggoSpread.SSSetEdit  	C1_tax_bill_type_nm,	"계산서형태명",			10
		ggoSpread.SSSetEdit  	C1_bp_cd,				"발행처",   				10
		ggoSpread.SSSetEdit  	C1_bp_nm,				"발행처명",   				15

        ggoSpread.SSSetDate  	C1_issued_dt,       	"발행일",     				10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit  	C1_vat_calc_type_nm,	"VAT적용기준",				10
		ggoSpread.SSSetEdit  	C1_vat_inc_flag_nm,		"VAT포함구분",     			10
		ggoSpread.SSSetEdit		C1_vat_type,			"VAT유형",					8

        ggoSpread.SSSetEdit		C1_vat_type_nm,			"VAT유형명",				12
		ggoSpread.SSSetFloat	C1_vat_rate,			"VAT율",	     			10, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  	C1_cur,					"통화",						8
		ggoSpread.SSSetFloat	C1_total_amt,			"합계금액",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_total_amt,		"(회계)합계금액",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

        ggoSpread.SSSetFloat	C1_net_amt,       		"공급가액",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_net_amt,			"(회계)공급가액",     		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_vat_amt,       		"VAT금액",     				18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_vat_amt,			"(회계)VAT금액",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_total_loc_amt,		"합계금액(자국)",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

        ggoSpread.SSSetFloat	C1_fi_total_loc_amt,	"(회계)합계금액(자국)",		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_net_loc_amt,			"공급가액(자국)",     		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_net_loc_amt,		"(회계)공급가액(자국)",		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_vat_loc_amt,			"VAT금액(자국)",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_vat_loc_amt,		"(회계)VAT금액(자국)",		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetEdit		C1_report_biz_area,		"세금신고사업장",			12
		ggoSpread.SSSetEdit		C1_tax_biz_area_nm,		"세금신고사업장명",			20
		ggoSpread.SSSetEdit		C1_sales_grp,			"영업그룹",					12
		ggoSpread.SSSetEdit		C1_sales_grp_nm,		"영업그룹명",				20
		ggoSpread.SSSetEdit		C1_remarks,				"비고",						20

		'=============
		ggoSpread.SSSetEdit  	C1_disuse_reason,		"반려/폐기사유",			15
		ggoSpread.SSSetEdit  	C1_legacy_pk,			"legacy PK",				15
		ggoSpread.SSSetEdit  	C1_sale_no,				"명세서번호",				15

		ggoSpread.SSSetEdit  	C1_proc_flag, 			"계산서상태", 				10
		ggoSpread.SSSetEdit  	C1_is_send, 			"송신상태", 				10
		ggoSpread.SSSetEdit  	C1_issue_dt_fg, 		"발행여부", 				10

        ggoSpread.SSSetEdit     C1_name   ,         "담당자",               10, ,,30
		ggoSpread.SSSetButton   C1_name_B
        ggoSpread.SSSetEdit     C1_email    ,       "E-Mail",               20
        ggoSpread.SSSetEdit     C1_bp_seq   ,       "순번",                 10
        ggoSpread.SSSetEdit     C1_name2    ,       "담당자2",              10, ,,30
		ggoSpread.SSSetButton   C1_name2_B
        ggoSpread.SSSetEdit     C1_email2   ,       "E-Mail2",              20
        ggoSpread.SSSetEdit     C1_bp_seq2,         "순번2",                10

		'Call ggoSpread.MakePairsColumn(C1_change_reason_cd, C1_change_reason, "1")
        Call ggoSpread.SSSetColHidden(C1_proc_flag, C1_issue_dt_fg, True)
        Call ggoSpread.SSSetColHidden(C1_inv_amend_type, C1_inv_amend_type, True)

        Call ggoSpread.SSSetColHidden(C1_bp_seq, C1_bp_seq, True)
        Call ggoSpread.SSSetColHidden(C1_bp_seq2, C1_bp_seq2, True)

		Call ggoSpread.MakePairsColumn(C1_name, C1_name_B)
		Call ggoSpread.MakePairsColumn(C1_name2, C1_name2_B)

		.ReDraw = True
	End With

	Call SetSpreadLock()
End Sub

Sub InitSpreadSheet2()
	Call initSpreadPosVariables2()
	With frm1.vspdData2
		.MaxCols = C2_tax_bill_no + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols												'☆: 사용자 별 Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		.ReDraw = False
		ggoSpread.Spreadinit "V20090708",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos2("A")
		ggoSpread.SSSetEdit  	C2_item_cd, 			"품목", 			18
		ggoSpread.SSSetEdit  	C2_item_nm, 			"품목명", 			20
		ggoSpread.SSSetEdit  	C2_spec, 				"규격", 			20

		ggoSpread.SSSetFloat	C2_bill_qty,			"수량",				15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit  	C2_bill_unit, 			"단위", 			8
		ggoSpread.SSSetFloat  	C2_bill_price, 			"단가", 			15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetFloat  	C2_total_amt, 			"합계금액", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_bill_amt, 			"공급가액", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_vat_amt, 			"VAT금액", 			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetFloat  	C2_total_amt_loc,		"합계금액(자국)", 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_bill_amt_loc, 		"공급가액(자국)", 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_vat_amt_loc, 		"VAT금액(자국)", 	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetEdit  	C2_bill_no, 			"Bill No.", 		15
		ggoSpread.SSSetEdit  	C2_bill_seq, 			"Bill Seq.", 		10
		ggoSpread.SSSetEdit  	C2_tax_bill_no, 		"Tax Bill No", 		15
		.ReDraw = True
	End With
	Call SetSpreadLock_B()
End Sub

'========================== 2.2.6 InitSpreadComboBox()  =====================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitSpreadComboBox()

	Dim strCboData    ''lgF0
	Dim strCboData2    ''lgF1

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("DT006", "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	strCboData = Replace(lgF0,chr(11),vbTab)
    strCboData2 = Replace(lgF1,chr(11),vbTab)
    strCboData = Left(strCboData,Len(strCboData) - 1)
    strCboData2 = Left(strCboData2,Len(strCboData2) - 1)
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboData,  C1_inv_amend_type
	ggoSpread.SetCombo strCboData2, C1_inv_amend_type_nm

End Sub

'========================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

		frm1.vspddata.col = C1_send_check
		frm1.vspddata.row = 0
		'frm1.vspddata.ColHeadersShow = True

		ggoSpread.SpreadLock 	C1_proc_flag_nm	,  -1, 	C1_proc_flag_nm
		ggoSpread.SpreadLock 	C1_is_send_nm 	,  -1, 	C1_is_send_nm
		ggoSpread.SpreadLock 	C1_where_flag	,  -1, 	C1_where_flag
		ggoSpread.SpreadLock 	C1_tax_bill_no 	,  -1, 	C1_tax_bill_no
		ggoSpread.SpreadLock 	C1_tax_doc_no	,  -1, 	C1_tax_doc_no
		ggoSpread.SpreadLock 	C1_inv_no	,  -1, 	C1_inv_no
		ggoSpread.SpreadLock 	C1_tax_bill_type_nm	,  -1, 	C1_tax_bill_type_nm
		ggoSpread.SpreadLock 	C1_bp_cd	,  -1, 	C1_bp_cd
		ggoSpread.SpreadLock 	C1_bp_nm	,  -1, 	C1_bp_nm
		ggoSpread.SpreadLock 	C1_issued_dt	,  -1, 	C1_issued_dt
		ggoSpread.SpreadLock 	C1_vat_calc_type_nm	,  -1, 	C1_vat_calc_type_nm
		ggoSpread.SpreadLock 	C1_vat_inc_flag_nm	,  -1, 	C1_vat_inc_flag_nm
		ggoSpread.SpreadLock 	C1_vat_type	,  -1, 	C1_vat_type
		ggoSpread.SpreadLock 	C1_vat_type_nm	,  -1, 	C1_vat_type_nm
		ggoSpread.SpreadLock 	C1_vat_rate	,  -1, 	C1_vat_rate
		ggoSpread.SpreadLock 	C1_cur	,  -1, 	C1_cur
		ggoSpread.SpreadLock 	C1_total_amt	,  -1, 	C1_total_amt
		ggoSpread.SpreadLock 	C1_fi_total_amt	,  -1, 	C1_fi_total_amt
		ggoSpread.SpreadLock 	C1_net_amt	,  -1, 	C1_net_amt
		ggoSpread.SpreadLock 	C1_fi_net_amt	,  -1, 	C1_fi_net_amt
		ggoSpread.SpreadLock 	C1_vat_amt	,  -1, 	C1_vat_amt
		ggoSpread.SpreadLock 	C1_fi_vat_amt	,  -1, 	C1_fi_vat_amt
		ggoSpread.SpreadLock 	C1_total_loc_amt	,  -1, 	C1_total_loc_amt
		ggoSpread.SpreadLock 	C1_fi_total_loc_amt	,  -1, 	C1_fi_total_loc_amt
		ggoSpread.SpreadLock 	C1_net_loc_amt	,  -1, 	C1_net_loc_amt
		ggoSpread.SpreadLock 	C1_fi_net_loc_amt	,  -1, 	C1_fi_net_loc_amt
		ggoSpread.SpreadLock 	C1_vat_loc_amt	,  -1, 	C1_vat_loc_amt
		ggoSpread.SpreadLock 	C1_fi_vat_loc_amt	,  -1, 	C1_fi_vat_loc_amt
		ggoSpread.SpreadLock 	C1_report_biz_area	,  -1, 	C1_report_biz_area
		ggoSpread.SpreadLock 	C1_tax_biz_area_nm	,  -1, 	C1_tax_biz_area_nm
		ggoSpread.SpreadLock 	C1_sales_grp	,  -1, 	C1_sales_grp
		ggoSpread.SpreadLock 	C1_sales_grp_nm	,  -1, 	C1_sales_grp_nm
		ggoSpread.SpreadLock 	C1_remarks	,  -1, 	C1_remarks
		ggoSpread.SpreadLock 	C1_disuse_reason	,  -1, 	C1_disuse_reason
		ggoSpread.SpreadLock 	C1_legacy_pk	,  -1, 	C1_legacy_pk
		ggoSpread.SpreadLock 	C1_sale_no	,  -1, 	C1_sale_no
		ggoSpread.SpreadLock 	C1_proc_flag	,  -1, 	C1_proc_flag
		ggoSpread.SpreadLock 	C1_is_send	,  -1, 	C1_is_send
		ggoSpread.SpreadLock 	C1_issue_dt_fg	,  -1, 	C1_issue_dt_fg

        ggoSpread.SpreadUNLock  C1_inv_amend_type   , -1, C1_inv_no
        ggoSpread.SpreadUNLock  C1_inv_amend_type_nm, -1, C1_inv_no
        ggoSpread.SpreadUNLock  C1_remark           , -1, C1_remark
        ggoSpread.SpreadUNLock  C1_remark2          , -1, C1_remark2
        ggoSpread.SpreadUNLock  C1_remark3          , -1, C1_remark3

		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1	,-1

        ggoSpread.SSSetRequired C1_inv_amend_type,	-1, -1
		ggoSpread.SSSetRequired C1_inv_amend_type_nm,	-1, -1
		ggoSpread.SSSetRequired C1_remark,			-1, -1

        ggoSpread.SpreadUNLock    C1_name   ,  -1,  C1_name   
        ggoSpread.SpreadUNLock    C1_name_B  ,  -1,  C1_name_B 
        ggoSpread.SpreadLock    C1_email  ,  -1,  C1_email  
        ggoSpread.SpreadLock    C1_bp_seq ,  -1,  C1_bp_seq 
        ggoSpread.SpreadUNLock    C1_name2  ,  -1,  C1_name2 
        ggoSpread.SpreadUNLock    C1_name2_B  ,  -1,  C1_name2_B  
        ggoSpread.SpreadLock    C1_email2 ,  -1,  C1_email2 
        ggoSpread.SpreadLock    C1_bp_seq2,  -1,  C1_bp_seq2

		.vspdData.ReDraw = True
	End With
End Sub

Sub SetSpreadLock_B()
	With frm1
		.vspdData2.ReDraw = False
		ggoSpread.Source = .vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData2.ReDraw = True
	End With
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C1_send_check        = iCurColumnPos(1 )
            C1_proc_flag_nm      = iCurColumnPos(2 )
            C1_is_send_nm        = iCurColumnPos(3 )
            C1_where_flag        = iCurColumnPos(4 )
            C1_tax_bill_no       = iCurColumnPos(5 )
            C1_tax_doc_no        = iCurColumnPos(6 )
            C1_inv_no            = iCurColumnPos(7 )
            C1_inv_amend_type    = iCurColumnPos(8 )
            C1_inv_amend_type_nm = iCurColumnPos(9 )
            C1_remark            = iCurColumnPos(10)
            C1_remark2           = iCurColumnPos(11)
            C1_remark3           = iCurColumnPos(12)
            C1_tax_bill_type_nm  = iCurColumnPos(13)
            C1_bp_cd             = iCurColumnPos(14)
            C1_bp_nm             = iCurColumnPos(15)
            C1_issued_dt         = iCurColumnPos(16)
            C1_vat_calc_type_nm  = iCurColumnPos(17)
            C1_vat_inc_flag_nm   = iCurColumnPos(18)
            C1_vat_type          = iCurColumnPos(19)
            C1_vat_type_nm       = iCurColumnPos(20)
            C1_vat_rate          = iCurColumnPos(21)
            C1_cur               = iCurColumnPos(22)
            C1_total_amt         = iCurColumnPos(23)
            C1_fi_total_amt      = iCurColumnPos(24)
            C1_net_amt           = iCurColumnPos(25)
            C1_fi_net_amt        = iCurColumnPos(26)
            C1_vat_amt           = iCurColumnPos(27)
            C1_fi_vat_amt        = iCurColumnPos(28)
            C1_total_loc_amt     = iCurColumnPos(29)
            C1_fi_total_loc_amt  = iCurColumnPos(30)
            C1_net_loc_amt       = iCurColumnPos(31)
            C1_fi_net_loc_amt    = iCurColumnPos(32)
            C1_vat_loc_amt       = iCurColumnPos(33)
            C1_fi_vat_loc_amt    = iCurColumnPos(34)
            C1_report_biz_area   = iCurColumnPos(35)
            C1_tax_biz_area_nm   = iCurColumnPos(36)
            C1_sales_grp         = iCurColumnPos(37)
            C1_sales_grp_nm      = iCurColumnPos(38)
            C1_remarks           = iCurColumnPos(39)
            C1_disuse_reason     = iCurColumnPos(40)
            C1_legacy_pk         = iCurColumnPos(41)
            C1_sale_no           = iCurColumnPos(42)

            C1_proc_flag         = iCurColumnPos(43)
            C1_is_send           = iCurColumnPos(44)
            C1_issue_dt_fg       = iCurColumnPos(45)

            C1_name              = iCurColumnPos(46)
            C1_name_B            = iCurColumnPos(47)
            C1_email             = iCurColumnPos(48)
            C1_bp_seq            = iCurColumnPos(49)
            C1_name2             = iCurColumnPos(50)
            C1_name2_B           = iCurColumnPos(51)
            C1_email2            = iCurColumnPos(52)
            C1_bp_seq2           = iCurColumnPos(53)

	End Select
End Sub

'========================================================================================
Sub GetSpreadColumnPos2(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C2_item_cd = iCurColumnPos(1)
			C2_item_nm = iCurColumnPos(2)
			C2_spec = iCurColumnPos(3)
			C2_bill_qty = iCurColumnPos(4)
			C2_bill_unit = iCurColumnPos(5)
			C2_bill_price = iCurColumnPos(6)
			C2_total_amt = iCurColumnPos(7)
			C2_bill_amt = iCurColumnPos(8)
			C2_vat_amt = iCurColumnPos(9)
			C2_total_amt_loc = iCurColumnPos(10)
			C2_bill_amt_loc = iCurColumnPos(11)
			C2_vat_amt_loc = iCurColumnPos(12)
			C2_bill_no = iCurColumnPos(13)
			C2_bill_seq = iCurColumnPos(14)
			C2_tax_bill_no = iCurColumnPos(15)

	End Select
End Sub

'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발행처"
	arrParam(1) = "b_biz_partner"

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE In ('C', 'CS')"
	arrParam(5) = "발행처"

	arrField(0) = "bp_cd"
	arrField(1) = "bp_nm"
	arrField(2) = "bp_rgst_no"

	arrHeader(0) = "발행처"
	arrHeader(1) = "발행처명"
	arrHeader(2) = "사업자등록번호"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
											  Array(arrParam, arrField, arrHeader), _
											  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtSupplierCd.focus
	End If

	Set gActiveElement = document.activeElement
End Function

'=========================================================================================================================
Function OpenSalesGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "영업그룹"
	arrParam(1) = "b_sales_grp"

	arrParam(2) = Trim(frm1.txtSalesGrpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "usage_flag = 'Y'"
	arrParam(5) = "영업그룹"

	arrField(0) = "sales_grp"
	arrField(1) = "sales_grp_nm"

	arrHeader(0) = "영업그룹"
	arrHeader(1) = "영업그룹명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
											  Array(arrParam, arrField, arrHeader), _
											  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSalesGrpCd.Value = arrRet(0)
		frm1.txtSalesGrpNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtSalesGrpCd.focus
	End If

	Set gActiveElement = document.activeElement
End Function

'=========================================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "세금신고사업장"
	arrParam(1) = "b_tax_biz_area"

	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = ""
	arrParam(5) = "세금신고사업장"

	arrField(0) = "tax_biz_area_cd"
	arrField(1) = "tax_biz_area_nm"

	arrHeader(0) = "세금신고사업장"
	arrHeader(1) = "세금신고사업장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
											  Array(arrParam, arrField, arrHeader), _
											  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd.Value    = arrRet(0)
		frm1.txtBizAreaNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtBizAreaCd.focus
	End If

	Set gActiveElement = document.activeElement
End Function


'------------------------------------------  OpenHistoryRef()  -------------------------------------------------
'	Name : OpenHistoryRef()
'	Description : Altered Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenHistoryRef()
	Dim arrRet
	Dim arrParam(0)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	iCalledAspName = AskPRAspName("D1211PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "D1211PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData.Row =frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C1_inv_no

	If Trim(frm1.vspdData.Text) = "" Then
		IntRetCD = DisplayMsgBox("205903", parent.VB_INFORMATION, "X", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.vspdData.Text)			'☜: 조회 조건 데이타 

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

End Function


'------------------------------------------  OpenBillRef()  -------------------------------------------------
'	Name : OpenBillRef()
'	Description : Altered Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBillRef()
	Dim arrRet
	Dim arrParam(0)
	Dim iCalledAspName


	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	iCalledAspName = AskPRAspName("D1212PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "D1212PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData.Row =frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C1_sale_no
	If Trim(frm1.vspdData.Text) = "" Then
		IntRetCD = DisplayMsgBox("205928", parent.VB_INFORMATION, "X", "X")
		IsOpenPop = False
		Exit Function
	End If

	frm1.vspdData.Col = C1_inv_no

	arrParam(0) = Trim(frm1.vspdData.Text)			'☜: 조회 조건 데이타 

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0)), _
		"dialogWidth=760px; dialogHeight=640px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function



'========================================================================================
Function fnResend()
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDtlVal
	Dim strCheckSend, iSelectCnt
	Dim strCheckProc

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, is_send_flag, proc_flag

	Dim objDTI
	Dim arrayM, userDN, userInfo, userInfoSet

	Dim RetFlag

	Dim StrSaveFlag , StrMessageNo

	'⊙: Processing is NG
	Call LayerShowHide(1)

	StrSaveFlag = "SD"

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDtlVal = ""
		iSelectCnt = 0
		lgAllSelect = False


		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1" Then

				.vspdData.Col = C1_is_send
				is_send_flag = Trim(.vspdData.text)
				.vspdData.Col = C1_proc_flag
				proc_flag = Trim(.vspdData.text)

				If Not (is_send_flag = sendStatusFail _
							AND (proc_flag = docStatusIssue or proc_flag = docStatusCancelDisuse or proc_flag = docStatusDisuse) ) Then

					.vspdData.Col = C1_tax_bill_no

					Call DisplayMsgBox("205902","X", .vspdData.text, "X")   '☜ 바뀐부분 
					Call LayerShowHide(0)
					Exit Function

				End If

				.vspdData.Col = C1_inv_no :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_proc_flag :	strDtlVal = strDtlVal & Trim(.vspdData.Text) & parent.gColSep

			End If
		Next

		If strVal = "" Then
			Call DisplayMsgBox("181216","X", .vspdData.text, "X")   '☜ 바뀐부분 
			Call LayerShowHide(0)
			Exit Function
		End If

		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strVal
		.txtDtlSpread.value = strDtlVal

		.txtuserDN.value = userDN
		.txtuserInfo.value = userInfo
		.txtbtnFlag.value = "Resend"

		Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ☜: 비지니스 ASP 를 가동 
	End With
End Function


Function fnPublish()
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	Dim strCheckSend, iSelectCnt
	Dim strCheckProc
    Dim strAmendType, strRemark, strCheckMsg

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, proc_flag

	Dim objDTI
	Dim arrayM, userDN, userInfo, userInfoSet

	Dim RetFlag

	Dim StrSaveFlag , StrMessageNo

	'⊙: Processing is NG
	Call LayerShowHide(1)

	StrSaveFlag = "SD"

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		iSelectCnt = 0
		lgAllSelect = False

        Dim loopCount
        loopCount = 0
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = C1_send_check

            If .vspdData.text = "1" Then
                loopCount = loopCount + 1
            End If
        Next

        If loopCount > 1 Then
            Call DisplayMsgBox("205926","X", "X", "X")   '☜ 바뀐부분 
            Call LayerShowHide(0)
            Exit Function            
        End If

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1" Then

                .vspdData.Col = C1_name
                If Trim(.vspdData.Text) = "" Then 
                    Call LayerShowHide(0)
                    Call DisplayMsgBox("970021","X", "담당자", "X")   '☜ 바뀐부분 
                    Exit Function
                End If 

                .vspdData.Col = C1_inv_amend_type
                strAmendType = .vspdData.Text

                .vspdData.Col = C1_remark
                strRemark = .vspdData.Text

                If strAmendType = "" Or strRemark = "" Then
                    If strRemark = "" Then
                        strCheckMsg = "비고"
                    End If 

                    If strAmendType = "" Then
                        strCheckMsg = "수정사유"
                    End If 

                    Call DisplayMsgBox("970021","X", strCheckMsg, "X")   '☜ 바뀐부분 
                    Call LayerShowHide(0)
                    Exit Function            
                End If 

				.vspdData.Col = C1_net_loc_amt
				net_loc_amt = CDbl(.vspdData.text)

				.vspdData.Col = C1_fi_net_loc_amt
				fi_net_loc_amt = CDbl(.vspdData.text)

				.vspdData.Col = C1_vat_loc_amt
				vat_loc_amt = CDbl(.vspdData.text)

				.vspdData.Col = C1_fi_vat_loc_amt
				fi_vat_loc_amt = CDbl(.vspdData.text)

				If net_loc_amt <> fi_net_loc_amt Or vat_loc_amt <> fi_vat_loc_amt Then

					.vspdData.Col = C1_tax_bill_no
            		RetFlag = DisplayMsgBox("205911", parent.VB_YES_NO, .vspdData.text, "X")   '☜ 바뀐부분 

					If RetFlag = VBNO Then
						Call LayerShowHide(0)
						Exit Function
					End If
				End If

				.vspdData.Col = C1_proc_flag
				proc_flag = Trim(.vspdData.text)


				If Not (proc_flag = docStatusMake or proc_flag = "")  Then

					.vspdData.Col = C1_tax_bill_no

					If (proc_flag = docStatusDelete Or proc_flag = docStatusDisuse or proc_flag = docStatusReject) Then

            			RetFlag = DisplayMsgBox("205905", parent.VB_YES_NO, .vspdData.text, "X")   '☜ 바뀐부분 

						If RetFlag = VBNO Then
							Call LayerShowHide(0)
							Exit Function
						End If
					Else
						Call DisplayMsgBox("205901","X", .vspdData.text, "X")   '☜ 바뀐부분 
						Call LayerShowHide(0)
						Exit Function
					End If
				End If

				.vspdData.Col = C1_where_flag :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_tax_bill_no :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				strVal = strVal & "" & parent.gColSep						'BCDT03_IG1_vat_no
				strVal = strVal & "" & parent.gColSep						'BCDT03_IG1_iv_no
				.vspdData.Col = C1_inv_amend_type :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_remark  :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_remark2 :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col = C1_remark3 :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C1_bp_seq :     strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                .vspdData.Col = C1_bp_seq2 :    strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				strVal = strVal & lRow & parent.gRowSep

				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1
			End If
		Next

		If strVal = "" Then
			Call DisplayMsgBox("181216","X", .vspdData.text, "X")   '☜ 바뀐부분 
			Call LayerShowHide(0)
			Exit Function
		End If

		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strVal
		.txtDtlSpread.value = ""

		.txtuserDN.value = userDN
		.txtuserInfo.value = userInfo
		.txtbtnFlag.value = "Publish"

		Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ☜: 비지니스 ASP 를 가동 
	End With
End Function

Function fnApprovalDisuse()
	With frm1
		Call ChangeDocStatus(docStatusDisuse, .btnApprovalDisuse.title)
	End With
End Function

Function fnCancelDisuse()

	With frm1
		Call ChangeDocStatus(docStatusCancelDisuse, .btnCancelDisuse.title)
	End With

End Function

Sub ChangeDocStatus(pvChangeCode, pvMessageText)
	Dim lRow
	Dim lGrpCnt
	Dim strVal, strDtlVal
	Dim strCheckSend, iSelectCnt, StrSaveFlag
	Dim strCheckProc

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, is_send_flag, proc_flag

	Dim objDTI
	Dim arrayM, userDN, userInfo, userInfoSet
	StrSaveFlag = "SD"

	Call LayerShowHide(1)

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDtlVal = ""
		iSelectCnt = 0
		lgAllSelect = False

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1" Then

				.vspdData.Col = C1_proc_flag
				proc_flag = Trim(.vspdData.text)


				If Not (proc_flag = docStatusRequestDisuse )  Then

					.vspdData.Col = C1_tax_bill_no

					Call DisplayMsgBox("205906","X", pvMessageText & " : " & .vspdData.text, "X")   '☜ 바뀐부분 
					Call LayerShowHide(0)
					Exit Sub

				End If

				.vspdData.Col = C1_inv_no :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

			End If
		Next

		If strVal = "" Then
			Call DisplayMsgBox("181216","X", .vspdData.text, "X")   '☜ 바뀐부분 
			Call LayerShowHide(0)
			Exit Sub
		End If

		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = strVal
		.txtDtlSpread.value = ""
		.txtChangeStatus.value = pvChangeCode

		.txtuserDN.value = userDN
		.txtuserInfo.value = userInfo
		.txtbtnFlag.value = "ChangeDocStatus"

		Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ☜: 비지니스 ASP 를 가동 
	End With
End Sub

'=========================================================================================================
Sub txtIssuedFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then
		frm1.txtIssuedToDt.focus
		Call FncQuery
	End If
End Sub

'=========================================================================================================
Sub txtIssuedToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then
		frm1.txtIssuedFromDt.focus
		Call FncQuery
	End If
End Sub

'=========================================================================================================
Sub txtBizAreaCd_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		exit sub
	ElseIf KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

'=========================================================================================================
Sub txtBizAreaCd1_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		exit sub
	ElseIf KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim	strFlag
	
	With frm1.vspdData
	
		.Row = Row
		Select Case Col	
		
			Case C1_inv_amend_type, C1_inv_amend_type_nm
				.Col = Col
				intIndex = .Value
				If Col = C1_inv_amend_type Then
					.Col = C1_inv_amend_type_nm
				Else
					.Col = C1_inv_amend_type
				End If
				.Value = intIndex
		
		End Select		
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )


	Call SetPopupMenuItemInf("0000111111")

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData

 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If

 		frm1.vspddata.Row = frm1.vspdData.ActiveRow
		frm1.vspddata.Col = C1_tax_bill_no

		frm1.vspddata2.MaxRows = 0

		If DbQuery2 = False Then
			Call RestoreToolBar()
			Exit Sub
		End If

		lgOldRow = frm1.vspddata.Row
	Else
		If lgOldRow <> Row Then
 			'------ Developer Coding part (Start)
			frm1.vspddata.Row = Row
			frm1.vspddata.Col = C1_tax_bill_no
			frm1.vspddata2.MaxRows = 0

			lgOldRow = Row

			If DbQuery2 = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
	 		'------ Developer Coding part (End)
	 	End If
 	End If
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	If frm1.vspdData.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If

	Call vspdData_Click(NewCol, NewRow)
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C1_send_check Or NewCol <= C1_send_check Then
        Cancel = True
        Exit Sub
    End If

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
Sub txtFromReqDt_DblClick(Button)
'    If Button = 1 Then
'        frm1.txtFromReqDt.Action = 7
'        Call SetFocusToDocument("M")
'    End If
End Sub

Sub txtFromReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub txttoReqDt_DblClick(Button)
'    If Button = 1 Then
'        frm1.txttoReqDt.Action = 7
'        Call SetFocusToDocument("M")
'        frm1.txttoReqDt.focus
'    End If
End Sub

'=========================================================================================================
Sub txttoReqDt_Change(Button)
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================


'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True													'⊙: Indicates that value changed
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow
    Dim LngMaxRow

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKeyTempGlNo <> "" Then
			'Call DbQuery("1",frm1.vspddata.row)
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           'Call DbQuery("2",frm1.vspddata.ActiveRow)
		End If
   End if
End Sub

'#########################################################################################################
'												4. Common Function부 
'=========================================================================================================
Function FncQuery()

    Dim IntRetCD

    FncQuery = False																		'⊙: Processing is NG

    Err.Clear																				'☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    With frm1
	    ggoSpread.Source = .vspdData
	    If  ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'데이타가 변경되었습니다. 조회하시겠습니까?
	    	If IntRetCD = vbNo Then
		      	Exit Function
	    	End If
	    End If

		'-----------------------
	    'Check condition area
	    '-----------------------
		If Not chkFieldByCell(.txtIssuedFromDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.txtIssuedToDt, "A", "1") Then Exit Function

	   If CompareDateByFormat( .txtIssuedFromDt.text, _
										.txtIssuedToDt.text, _
										.txtIssuedFromDt.Alt, _
										.txtIssuedToDt.Alt, _
										"970025", _
										.txtIssuedFromDt.UserDefinedFormat, _
										parent.gComDateType, _
										True) = False Then
			Exit Function
		End If

		If frm1.txtSupplierCd.value = "" Then
			frm1.txtSupplierNm.value = ""
		End If

		'-----------------------
		'Erase contents area
		'-----------------------
		'	    Call ggoOper.ClearField(Document, "2")												'⊙: Clear Contents  Field
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

		Call InitVariables 																	'⊙: Initializes local global variables


	End With

	If DBquery = False Then
		Call RestoreToolBar()
		Exit Function
	End If

	FncQuery = True
End Function

'========================================================================================
Function FncNew()
	Dim IntRetCD

	FncNew = False																	'⊙: Processing is NG

	Err.Clear																			'☜: Protect system from crashing
	'On Error Resume Next															'☜: Protect system from crashing

	'-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")												'⊙: Clear Condition Field
	Call ggoOper.ClearField(Document, "2")												'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	Call LockObjectField(.txtFromReqDt,"R")
	Call LockObjectField(.txtToReqDt,"R")

	'Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field
	Call SetDefaultVal
	Call InitVariables																	'⊙: Initializes local global variables

	FncNew = True																		'⊙: Processing is OK
End Function


'========================================================================================
Function FncSave()
	Dim IntRetCD
	FncSave = False																		'⊙: Processing is NG



	FncSave = True																		'⊙: Processing is OK
End Function

'========================================================================================
Function FncCancel()
	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo																	'☜: Protect system from crashing
End Function

'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()																'☜: Protect system from crashing
End Function

'=======================================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function


'=======================================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)											'☜:화면 유형, Tab 유무 
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'=======================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")								'데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitSpreadComboBox()
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
Function DbQuery()
	Dim strVal
	Dim txtStatusflag

	DbQuery = False

	With frm1

		.btnResend.disabled	= false
		.btnPublish.disabled	= false
		.btnApprovalDisuse.disabled	= false
		.btnCancelDisuse.disabled	= false

		If .rdoStatusflag1.checked = True Then
			txtStatusflag = ""
		ElseIf .rdoStatusflag2.checked = True Then
			txtStatusflag = .rdoStatusflag2.value
		ElseIf .rdoStatusflag3.checked = True Then
			txtStatusflag = .rdoStatusflag3.value
		End If


		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
		                      "&txtSupplierCd=" & Trim(.txtSupplierCd.value) & _
		                      "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value) & _
		                      "&txtBizAreaCd=" & Trim(.txtBizAreaCd.value) & _
		                      "&cboTaxDocumentType=" & Trim(.cboTaxDocumentType.value) & _
		                      "&cboTransmitStatus=" & Trim(.cboTransmitStatus.value) & _
		                      "&rdoStatusflag=" & Trim(txtStatusflag) & _
		                      "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text) & _
		                      "&txtIssuedToDt=" & Trim(.txtIssuedToDt.text)

	End With

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)																'☜: 비지니스 ASP 를 가동 

	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 And Spread 3 Data 조회 
'========================================================================================
Function DbQuery2()
	DbQuery2 = False

	Dim strVal                                                        			'⊙: Processing is NG
	Dim iTaxBillNo
	Dim strWhereFlag

	Call LayerShowHide(1)

	ggoSpread.Source = frm1.vspdData

	frm1.vspddata.Row = lgOldRow
	frm1.vspddata.Col = C1_tax_bill_no
	iTaxBillNo = frm1.vspddata.Text

    frm1.vspddata.Col = C1_where_flag
    strWhereFlag = frm1.vspddata.Text

	strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001 & _
								  "&txtTaxBillNo=" & Trim(iTaxBillNo) & _
                                  "&strWhereFlag=" & Trim(strWhereFlag)

	Call RunMyBizASP(MyBizASP, strVal)

	DbQuery2 = True
End Function

'========================================================================================
Function DbQueryOk()																		'☆: 조회 성공후 실행로직 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE																'⊙: Indicates that current mode is Update mode

	lgOldRow = 1
	frm1.vspdData.Col = 1
	frm1.vspdData.Row = 1

	With frm1
														'⊙: This function lock the suitable field
		'Call LayerShowHide(0)
		Call SetToolbar("110000000001111")																'⊙: 버튼 툴바 제어 


		If .vspdData.MaxRows > 0 Then
			.btnResend.disabled	= false
			.btnPublish.disabled	= false
			.btnApprovalDisuse.disabled	= false
			.btnCancelDisuse.disabled	= false

            ggoSpread.SSSetRequired    C1_name, -1, -1

			If Dbquery2 = False Then
				Call RestoreToolbar()
				Exit Function
			End If

			Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
			Set gActiveElement = document.activeElement

            Dim loopCount
            For loopCount = 1 To frm1.vspdData.MaxRows
                frm1.vspdData.Row = loopCount
                frm1.vspdData.Col = C1_proc_flag
                
                If Not (frm1.vspdData.text = "" Or frm1.vspdData.text = docStatusReject Or frm1.vspdData.text = docStatusDisuse) Then
                    ggoSpread.SSSetProtected    C1_inv_amend_type   , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_inv_amend_type_nm, loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_remark           , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_remark2          , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_remark3          , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_name             , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_name_B           , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_name2            , loopCount, loopCount
                    ggoSpread.SSSetProtected    C1_name2_B          , loopCount, loopCount
                End If
            Next
		End If
	End With
End Function

'======================================================================================================
Function SetGridFocus()
	with frm1
		.vspdData.Row = 1
		.vspdData.Col = 1
		.vspdData.Action = 1
	end with
End Function

'========================================================================================
Function DbSave()
End Function

'========================================================================================
Function SaveResult()
	Call ExecMyBizASP(frm1, BIZ_PGM_ID4)			' ☜: 비지니스 ASP 를 가동 
End Function

'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

    lgLngCurRows = 0                            'initializes Deleted Rows Count

	ggoSpread.source = frm1.vspddata2
    frm1.vspdData2.MaxRows = 0

	Call MainQuery

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus2(lRow, lCol)
	Dim pvCnt, pvInputCnt

	pvInputCnt = 0

	For pvCnt = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = pvCnt
		frm1.vspdData.Col = C1_send_check
		If frm1.vspdData.text = "1" Then
			If lRow = pvInputCnt Then
				Exit For
			End If
			pvInputCnt = pvInputCnt + 1
		End if

	Next
	frm1.vspdData.focus
	frm1.vspdData.Row = pvCnt
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function


'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssuedFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedFromDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtIssuedFromDt.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssuedToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtIssuedToDt.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And (Col = C1_name_B Or Col = C1_name2_B) Then
		    .Col = Col
		    .Row = Row

            Call OpenEmail(Col, Row)
			Call SetActiveCell(frm1.vspdData,Col - 1,frm1.vspdData.ActiveRow ,"M","X","X")
		End If
    End With
End Sub

'=========================================================================================================
'    Name : OpenEmail()
'    Description : Email PopUp
'=========================================================================================================
Function OpenEmail(Byval Col, Byval Row)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim strBpCd

    frm1.vspdData.Row = frm1.vspdData.Row
    frm1.vspdData.Col = C1_bp_cd
    strBpCd = frm1.vspdData.Text

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "담당자 팝업"                                     ' 팝업 명칭 
    arrParam(1) = "dt_biz_partner_person"                                          ' TABLE 명칭 
    arrParam(2) = ""                                                   ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = " bp_cd = " & FilterVar(strBpCd, "''", "S")               ' Where Condition
    arrParam(5) = "담당자"            
    
    arrField(0) = "name"                                                  ' Field명(0)
    arrField(1) = "email"                                                  ' Field명(1)
    arrField(2) = "dept_nm"                                                  ' Field명(1)
    arrField(3) = "ph_no"                                                  ' Field명(1)
    arrField(4) = "remarks"                                                  ' Field명(1)
    arrField(5) = "bp_seq"                                                  ' Field명(1)
    
    arrHeader(0) = "담당자"                                                ' Header명(0)
    arrHeader(1) = "E-Mail"                                           ' Header명(1)
    arrHeader(2) = "부서명"                                           ' Header명(1)
    arrHeader(3) = "전화번호"                                           ' Header명(1)
    arrHeader(4) = "비고"                                           ' Header명(1)
    arrHeader(5) = "순번"                                           ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetEmail(arrRet, Col, Row)        'return value setting
    End If    

	Set gActiveElement = document.activeElement

End Function

'=========================================================================================================
'    Name : SetEmail()
'    Description : Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetEmail(Byval arrRet, ByVal Col, ByVal Row)
    With frm1
        frm1.vspdData.Row = frm1.vspdData.Row
    
        If Col = C1_name_B Then 
            .vspdData.Col = C1_name
            .vspdData.Text = arrRet(0)
            .vspdData.Col = C1_email
            .vspdData.Text = arrRet(1)
            .vspdData.Col = C1_bp_seq
            .vspdData.Text = arrRet(5)
        Else
            .vspdData.Col = C1_name2
            .vspdData.Text = arrRet(0)
            .vspdData.Col = C1_email2
            .vspdData.Text = arrRet(1)
            .vspdData.Col = C1_bp_seq2
            .vspdData.Text = arrRet(5)            
        End If
    End With
End Function

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
    Dim strBpCd, strName
    Dim nameArr, emailArr, seqArr

    frm1.vspdData.Row = frm1.vspdData.Row

    frm1.vspdData.Col = C1_bp_cd
    strBpCd = frm1.vspdData.Text

    frm1.vspdData.Col = Col
    strName = frm1.vspdData.Text

    If Col = C1_name Or Col = C1_name2 Then 
        Call CommonQueryRs(" name, email, bp_seq "," dt_biz_partner_person "," bp_cd = " & FilterVar(strBpCd, "''", "S") & " and name = " & FilterVar(strName, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        nameArr = Split(lgF0, Chr(11))
        emailArr = Split(lgF1, Chr(11))
        seqArr = Split(lgF2, Chr(11))

        If Col = C1_name Then
            If trim(lgF0) = "" Then
                frm1.vspdData.Col = C1_name
                frm1.vspdData.Text = ""
                frm1.vspdData.Col = C1_email
                frm1.vspdData.Text = ""
                frm1.vspdData.Col = C1_bp_seq
                frm1.vspdData.Text = ""
            Else 
                frm1.vspdData.Col = C1_name
                frm1.vspdData.Text = nameArr(0)
                frm1.vspdData.Col = C1_email
                frm1.vspdData.Text = emailArr(0)
                frm1.vspdData.Col = C1_bp_seq
                frm1.vspdData.Text = seqArr(0)
            End If
        ElseIf Col = C1_name2 Then 
            If trim(lgF0) = "" Then
                frm1.vspdData.Col = C1_name2
                frm1.vspdData.Text = ""
                frm1.vspdData.Col = C1_email2
                frm1.vspdData.Text = ""
                frm1.vspdData.Col = C1_bp_seq2
                frm1.vspdData.Text = ""
            Else
                frm1.vspdData.Col = C1_name2
                frm1.vspdData.Text = nameArr(0)
                frm1.vspdData.Col = C1_email2
                frm1.vspdData.Text = emailArr(0)
                frm1.vspdData.Col = C1_bp_seq2
                frm1.vspdData.Text = seqArr(0)
            End If
        End If
    End If

'	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<% '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### %>
<BODY TABINDEX="-1" SCROLL="no">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE <%=LR_SPACE_TYPE_00%>>
			<TR>
				<TD <%=HEIGHT_TYPE_00%>></TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_10%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출수정정발행등록</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
									 </TR>
								</TABLE>
							</TD>
							<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:OpenBillRef()">거래명세서</A> | <A href="vbscript:OpenHistoryRef()">이력조회</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE <%=LR_SPACE_TYPE_40%>>
										<TR>
											<TD CLASS="TD5"NOWRAP>발행일</TD>
											<TD CLASS="TD6"NOWRAP>
												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssuedFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작일자" class=required></OBJECT>');</SCRIPT> ~
 												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssuedToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료일자" class=required></OBJECT>');</SCRIPT>
 											</TD>
 											<TD CLASS="TD5" NOWRAP>발행처</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
												<INPUT TYPE=TEXT AlT="발행처" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
										</TR>
										<TR>
 											<TD CLASS="TD5" NOWRAP>영업그룹</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSalesGrpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSalesGrp()">
												<INPUT TYPE=TEXT AlT="영업그룹" ID="txtSalesGrpNm" NAME="txtSalesGrpNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
 											<TD CLASS="TD5" NOWRAP>세금신고사업장</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
												<INPUT TYPE=TEXT AlT="세금신고사업장" ID="txtBizAreaNm" NAME="txtBizAreaNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
										</TR>
										<TR>
											<TD CLASS="TD5"NOWRAP>계산서상태</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboTaxDocumentType" ALT="계산서상태" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
											<TD CLASS="TD5"NOWRAP>송신상태</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboTransmitStatus" ALT="송신상태" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>계산서형태</TD>
											<TD CLASS=TD6 NOWRAP>
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag1" value="%" tag = "11X" checked><label for="rdoCfmAll">전체</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag2" value="R" tag = "11X"><label for="rdoCfmreceipt">영수</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag3" value="D" tag = "11X"><label for="rdoCfmdemand">청구</label>
											</TD>
											<TD CLASS=TD5 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>

									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=* valign=top>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR HEIGHT="60%">
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData ID = "A" width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
									</TR>
									<TR HEIGHT="40%">
										<TD WIDTH="100%" colspan="4"><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vspdData2><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_01%>></TD>
			</TR>
			<TR HEIGHT="20">
				<TD WIDTH="100%" >
  					<TABLE <%=LR_SPACE_TYPE_30%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD><BUTTON NAME="btnReSend" CLASS="CLSSBTN" OnClick="VBScript:Call fnResend()">재전송</BUTTON>&nbsp;
							    <BUTTON NAME="btnPublish" CLASS="CLSSBTN" OnClick="VBScript:Call fnPublish()">발행</BUTTON>&nbsp;
							    <BUTTON NAME="btnApprovalDisuse" CLASS="CLSSBTN" OnClick="VBScript:Call fnApprovalDisuse()">폐기승인</BUTTON>&nbsp;
								<BUTTON NAME="btnCancelDisuse" CLASS="CLSSBTN" OnClick="VBScript:Call fnCancelDisuse()">폐기취소</BUTTON>&nbsp;</TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
  					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
		<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
		<TEXTAREA CLASS="hidden" NAME="txtDtlSpread" tag="24" TABINDEX="-1"></TEXTAREA>
		<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserDN" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserInfo" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtbtnFlag" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtChangeStatus" tag="24" TABINDEX="-1">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=280 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
	<FORM NAME=EBAction TARGET="MyBizASP"   METHOD="POST">
		<INPUT TYPE="HIDDEN" NAME="uname"       TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="dbname"      TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="filename"    TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="condvar"     TABINDEX="-1">
		<INPUT TYPE="HIDDEN" NAME="date">
	</Form>
</BODY>
</HTML>
