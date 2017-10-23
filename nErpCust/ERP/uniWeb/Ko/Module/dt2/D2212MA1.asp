<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : m5134ma1
'*  4. Program Name         : 전자계산서 발행(구매) 
'*  5. Program Desc         : 전자계산서에 대하여 발행 또는 발행취소하는 기능 
'*  6. Component List       : PAGG015.dll
'*  7. Modified date(First) : 2000/10/14
'*  8. Modified date(Last)  : 2003/10/31
'*  9. Modifier (First)     : Lee MIn Hyung
'* 10. Modifier (Last)      : Lee Min HYung
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<!--SCRIPT LANGUAGE="VBScript"	SRC="./D2211ma1.vbs"></SCRIPT-->
<SCRIPT LANGUAGE="VBSCRIPT">

Option Explicit  

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim lgOldRow, lgRow
Dim lgSortKey1
Dim lgSortKey2

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
Const BIZ_PGM_ID  = "D2212MB1.asp"
Const BIZ_PGM_ID2 = "D2212MB2.asp"
Const BIZ_PGM_ID3 = "D2212MB3.asp"
Const BIZ_PGM_ID4 = "D2212MB4.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'=                       4.2 Constant variables 
'========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"

'==========================================================================================================

'add header datatable column
Dim 	C1_send_check
Dim 	C1_where_flag
Dim 	C1_issue_dt_flag
Dim 	C1_success_flag
Dim 	C1_process_date
Dim 	C1_dt_inv_no
Dim 	C1_tax_bill_no
Dim 	C1_tax_doc_no
Dim 	C1_change_reason_cd
Dim 	C1_change_reason
Dim 	C1_change_remark
Dim 	C1_change_remark2
Dim 	C1_change_remark3
Dim	C1_tax_bill_type_nm
Dim 	C1_bp_cd
Dim 	C1_bp_nm
	
Dim   C1_issued_dt
Dim 	C1_vat_calc_type_nm
Dim 	C1_vat_inc_flag_nm
Dim 	C1_vat_type
Dim 	C1_vat_type_nm
Dim 	C1_vat_rate
Dim 	C1_cur
Dim 	C1_total_amt
Dim 	C1_fi_total_amt
Dim 	C1_net_amt

Dim 	C1_fi_net_amt
Dim 	C1_vat_amt
Dim 	C1_fi_vat_amt
Dim 	C1_total_loc_amt
Dim 	C1_fi_total_loc_amt
Dim 	C1_net_loc_amt
Dim 	C1_fi_net_loc_amt
Dim 	C1_vat_loc_amt
Dim 	C1_fi_vat_loc_amt

Dim 	C1_report_biz_area
Dim 	C1_tax_biz_area_nm
Dim 	C1_sales_grp
Dim 	C1_sales_grp_nm
Dim	C1_remarks
Dim 	C1_error_desc

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
Dim	C2_where_flag

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
 
      Call SetToolbar("110000000000111")										'⊙: 버튼 툴바 제어    	
 
      .txtSupplierCd.focus
   End With		
End Sub

'========================================================================================================= 
Sub InitComboBox()
   Dim iCodeArr 
   Dim iNameArr
   Dim iDx
	
   '자료유형(Data Type)
   Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("DT006", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    

   iCodeArr = vbTab & lgF0
   iNameArr = vbTab & lgF1

   ggoSpread.SetCombo Replace(iCodeArr, Chr(11), vbTab), C1_change_reason_cd			'COLM_DATA_TYPE
   ggoSpread.SetCombo Replace(iNameArr, Chr(11), vbTab), C1_change_reason
End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
	C1_send_check = 1
	C1_where_flag = 2
	C1_issue_dt_flag = 3
	C1_success_flag = 4
	C1_process_date = 5
	C1_dt_inv_no = 6
	C1_tax_bill_no = 7
	C1_tax_doc_no = 8
	C1_change_reason_cd = 9
	C1_change_reason = 10
	C1_change_remark = 11
	C1_change_remark2 = 12
	C1_change_remark3 = 13
	C1_tax_bill_type_nm = 14
	C1_bp_cd = 15
	C1_bp_nm = 16

	C1_issued_dt = 17
	C1_vat_calc_type_nm = 18
	C1_vat_inc_flag_nm = 19
	C1_vat_type = 20
	C1_vat_type_nm = 21
	C1_vat_rate = 22
	C1_cur = 23

	C1_total_amt = 24
	C1_fi_total_amt = 25
	C1_net_amt = 26
	C1_fi_net_amt = 27
	C1_vat_amt = 28
	C1_fi_vat_amt = 29

	C1_total_loc_amt = 30
	C1_fi_total_loc_amt = 31
	C1_net_loc_amt = 32
	C1_fi_net_loc_amt = 33
	C1_vat_loc_amt = 34
	C1_fi_vat_loc_amt = 35

	C1_report_biz_area = 36
	C1_tax_biz_area_nm = 37
	C1_sales_grp = 38
	C1_sales_grp_nm = 39
	C1_remarks = 40
	C1_error_desc = 41
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
	C2_where_flag = 16
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
		.MaxCols = C1_error_desc + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols												'☆: 사용자 별 Hidden Column
		.ColHidden = True
		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20090707",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetButton	C1_send_check
		ggoSpread.SSSetEdit  	C1_where_flag, 		"업무명", 					 8, ,,3
		ggoSpread.SSSetEdit  	C1_issue_dt_flag, 	"계산서발행여부", 			12, 2,,10
		ggoSpread.SSSetEdit  	C1_success_flag, 		"전송여부", 					10, 2,,10
		ggoSpread.SSSetDate  	C1_process_date,		"전송일",     				13, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C1_dt_inv_no,			"전송관리번호",				30, ,,50
		ggoSpread.SSSetEdit  	C1_tax_bill_no,      "세금계산서관리번호",		15, ,,30
		ggoSpread.SSSetEdit  	C1_tax_doc_no,       "세금계산서번호",			15, ,,30

		ggoSpread.SSSetCombo  	C1_change_reason_cd, "수정사유",					15
		ggoSpread.SSSetCombo  	C1_change_reason,    "수정사유",					15
		ggoSpread.SSSetEdit  	C1_change_remark,		"비고1",						15, ,,30
		ggoSpread.SSSetEdit  	C1_change_remark2,	"비고2",						15, ,,30
		ggoSpread.SSSetEdit  	C1_change_remark3,	"비고3",						15, ,,30
		
		ggoSpread.SSSetEdit  	C1_tax_bill_type_nm, "계산서형태명", 			15, ,,18
		ggoSpread.SSSetEdit  	C1_bp_cd,				"발행처",   					15, ,,18
		ggoSpread.SSSetEdit  	C1_bp_nm,				"발행처명",   				30, ,,50
		ggoSpread.SSSetDate  	C1_issued_dt,       	"발행일",     				13, 2, parent.gDateFormat

		ggoSpread.SSSetEdit  	C1_vat_calc_type_nm, "VAT적용기준",				15, 2,,15 
		ggoSpread.SSSetEdit  	C1_vat_inc_flag_nm,	"VAT포함구분",     		15, 2,,15
		ggoSpread.SSSetEdit		C1_vat_type,			"VAT유형",					10, 2,,10
		ggoSpread.SSSetEdit		C1_vat_type_nm,		"VAT유형명",					20, ,,20
		ggoSpread.SSSetFloat		C1_vat_rate,			"VAT율",	     				18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  	C1_cur,					"통화",						10, 2,,10

		ggoSpread.SSSetFloat		C1_total_amt,			"합계금액",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_fi_total_amt,     "(회계)합계금액",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_net_amt,       	"공급가액",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_fi_net_amt,			"(회계)공급가액",     	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_vat_amt,       	"VAT금액",     				18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_fi_vat_amt,       "(회계)VAT금액",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetFloat		C1_total_loc_amt,		"합계금액(자국)",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_fi_total_loc_amt,	"(회계)합계금액(자국)",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_net_loc_amt,      "공급가액(자국)",     	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_fi_net_loc_amt,	"(회계)공급가액(자국)",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_vat_loc_amt,		"VAT금액(자국)",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat		C1_fi_vat_loc_amt,	"(회계)VAT금액(자국)",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetEdit		C1_report_biz_area,	"세금신고사업장",			10, 2,,10
		ggoSpread.SSSetEdit		C1_tax_biz_area_nm,	"세금신고사업장명",			15, ,,10
		ggoSpread.SSSetEdit		C1_sales_grp,			"영업그룹",					10, 2,,20
		ggoSpread.SSSetEdit		C1_sales_grp_nm,		"영업그룹명",				15, ,,20
		
		ggoSpread.SSSetEdit		C1_remarks,				"비고",						30, ,,50
		ggoSpread.SSSetEdit		C1_error_desc,			"전송에러내용",				30, ,,50
		
		ggoSpread.SSSetEdit		C1_remarks,				"비고",						20, ,,150
		ggoSpread.SSSetEdit		C1_error_desc,			"전송에러내용",				30, ,,200

		Call ggoSpread.MakePairsColumn(C1_change_reason_cd, C1_change_reason, "1")
      Call ggoSpread.SSSetColHidden(C1_change_reason_cd, C1_change_reason_cd, True)

		.ReDraw = True
	End With

	Call InitComboBox()
	Call SetSpreadLock()
End Sub

Sub InitSpreadSheet2()
	Call initSpreadPosVariables2()
	With frm1.vspdData2	
		.MaxCols = C2_where_flag + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols												'☆: 사용자 별 Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		.ReDraw = False 
		ggoSpread.Spreadinit "V20090707",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos2("A")
		ggoSpread.SSSetEdit  	C2_item_cd, 			"품목", 				15, ,,18
		ggoSpread.SSSetEdit  	C2_item_nm, 			"품목명", 			30, ,,30
		ggoSpread.SSSetEdit  	C2_spec, 				"규격", 				15, ,,18

		ggoSpread.SSSetFloat		C2_bill_qty,			"수량",				15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit  	C2_bill_unit, 			"단위", 				15, ,,18
		ggoSpread.SSSetFloat  	C2_bill_price, 		"단가", 				18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec	

		ggoSpread.SSSetFloat  	C2_total_amt, 			"합계금액", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_bill_amt, 			"공급가액", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_vat_amt, 			"VAT금액", 			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetFloat  	C2_total_amt_loc,		"합계금액(자국)", 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_bill_amt_loc, 		"공급가액(자국)", 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_vat_amt_loc, 		"VAT금액(자국)", 	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetEdit  	C2_bill_no, 			"Bill No.", 		15, ,,18
		ggoSpread.SSSetEdit  	C2_bill_seq, 			"Bill Seq.", 		15, ,,18
		ggoSpread.SSSetEdit  	C2_tax_bill_no, 		"Tax Bill No", 	15, ,,18
		ggoSpread.SSSetEdit  	C2_where_flag, 		"Where Flag", 		15, ,,18
		
		Call ggoSpread.SSSetColHidden(C2_bill_no, C2_where_flag, True)

		.ReDraw = True
	End With	
	Call SetSpreadLock_B()
End Sub

'========================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

		frm1.vspddata.col = C1_send_check
		frm1.vspddata.row = 0
		frm1.vspddata.ColHeadersShow = True

		'ggoSpread.SpreadLock    C_send_check,			 -1, C_send_check
		ggoSpread.SpreadLock    C1_where_flag, 		-1, C1_where_flag
		ggoSpread.SpreadLock    C1_issue_dt_flag, 	-1, C1_issue_dt_flag
		ggoSpread.SpreadLock    C1_success_flag, 		-1, C1_success_flag
		ggoSpread.SpreadLock		C1_process_date,		-1, C1_process_date
		ggoSpread.SpreadLock		C1_dt_inv_no,			-1, C1_dt_inv_no
		ggoSpread.SpreadLock    C1_tax_bill_no,      -1, C1_tax_bill_no
		ggoSpread.SpreadLock    C1_tax_doc_no,       -1, C1_tax_doc_no
'		ggoSpread.SpreadLock	   C1_change_reason_cd,	-1, C1_change_reason_cd
'		ggoSpread.SpreadLock	   C1_change_reason,		-1, C1_change_reason
'		ggoSpread.SpreadLock	   C1_change_remark,		-1, C1_change_remark
'		ggoSpread.SpreadLock	   C1_change_remark2,	-1, C1_change_remark2
'		ggoSpread.SpreadLock	   C1_change_remark3,	-1, C1_change_remark3
		ggoSpread.SpreadLock	   C1_tax_bill_type_nm, -1, C1_tax_bill_type_nm
		ggoSpread.SpreadLock		C1_bp_cd,				-1, C1_bp_cd
		ggoSpread.SpreadLock		C1_bp_nm,				-1, C1_bp_nm
																		  
		ggoSpread.SpreadLock		C1_issued_dt,       	-1, C1_issued_dt
		ggoSpread.SpreadLock		C1_vat_calc_type_nm, -1, C1_vat_calc_type_nm
		ggoSpread.SpreadLock		C1_vat_inc_flag_nm,	-1, C1_vat_inc_flag_nm
		ggoSpread.SpreadLock		C1_vat_type,			-1, C1_vat_type
		ggoSpread.SpreadLock		C1_vat_type_nm,		-1, C1_vat_type_nm
		ggoSpread.SpreadLock		C1_vat_rate,			-1, C1_vat_rate
		ggoSpread.SpreadLock		C1_cur,					-1, C1_cur
																		  
		ggoSpread.SpreadLock		C1_total_amt,			-1, C1_total_amt
		ggoSpread.SpreadLock		C1_fi_total_amt,     -1, C1_fi_total_amt
		ggoSpread.SpreadLock		C1_net_amt,       	-1, C1_net_amt
		ggoSpread.SpreadLock		C1_fi_net_amt,			-1, C1_fi_net_amt
		ggoSpread.SpreadLock		C1_vat_amt,       	-1, C1_vat_amt
		ggoSpread.SpreadLock		C1_fi_vat_amt,       -1, C1_fi_vat_amt
																		  
		ggoSpread.SpreadLock		C1_total_loc_amt,		-1, C1_total_loc_amt	
		ggoSpread.SpreadLock		C1_fi_total_loc_amt,	-1, C1_fi_total_loc_amt
		ggoSpread.SpreadLock		C1_net_loc_amt,      -1, C1_net_loc_amt
		ggoSpread.SpreadLock		C1_fi_net_loc_amt,	-1, C1_fi_net_loc_amt
		ggoSpread.SpreadLock		C1_vat_loc_amt,		-1, C1_vat_loc_amt
		ggoSpread.SpreadLock		C1_fi_vat_loc_amt,	-1, C1_fi_vat_loc_amt
																		  
		ggoSpread.SpreadLock		C1_report_biz_area,	-1, C1_report_biz_area
		ggoSpread.SpreadLock		C1_tax_biz_area_nm,	-1, C1_tax_biz_area_nm
		ggoSpread.SpreadLock		C1_sales_grp,			-1, C1_sales_grp
		ggoSpread.SpreadLock		C1_sales_grp_nm,		-1, C1_sales_grp_nm
		ggoSpread.SpreadLock		C1_remarks,				-1, C1_remarks
		ggoSpread.SpreadLock		C1_error_desc,			-1, C1_error_desc

		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1	,-1
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
			C1_send_check = iCurColumnPos(1)
			C1_where_flag = iCurColumnPos(2)
			C1_issue_dt_flag = iCurColumnPos(3)
			C1_success_flag = iCurColumnPos(4)
			C1_process_date = iCurColumnPos(5)
			C1_dt_inv_no = iCurColumnPos(6)
			C1_tax_bill_no = iCurColumnPos(7)
			C1_tax_doc_no = iCurColumnPos(8)
			C1_change_reason_cd = iCurColumnPos(9)
			C1_change_reason = iCurColumnPos(10)
			C1_change_remark = iCurColumnPos(11)
			C1_change_remark2 = iCurColumnPos(12)
			C1_change_remark3 = iCurColumnPos(13)
			C1_tax_bill_type_nm = iCurColumnPos(14)
			C1_bp_cd = iCurColumnPos(15)
			C1_bp_nm = iCurColumnPos(16)

			C1_issued_dt = iCurColumnPos(17)
			C1_vat_calc_type_nm = iCurColumnPos(18)
			C1_vat_inc_flag_nm = iCurColumnPos(19)
			C1_vat_type = iCurColumnPos(20)
			C1_vat_type_nm = iCurColumnPos(21)
			C1_vat_rate = iCurColumnPos(22)
			C1_cur = iCurColumnPos(23)
			C1_total_amt = iCurColumnPos(24)
			C1_fi_total_amt = iCurColumnPos(25)
			C1_net_amt = iCurColumnPos(26)

			C1_fi_net_amt = iCurColumnPos(27)
			C1_vat_amt = iCurColumnPos(28)
			C1_fi_vat_amt = iCurColumnPos(29)
			C1_total_loc_amt = iCurColumnPos(30)
			C1_fi_total_loc_amt = iCurColumnPos(31)
			C1_net_loc_amt = iCurColumnPos(32)
			C1_fi_net_loc_amt = iCurColumnPos(33)
			C1_vat_loc_amt = iCurColumnPos(34)
			C1_fi_vat_loc_amt = iCurColumnPos(35)

			C1_report_biz_area = iCurColumnPos(36)
			C1_tax_biz_area_nm = iCurColumnPos(37)
			C1_sales_grp = iCurColumnPos(38)
			C1_sales_grp_nm = iCurColumnPos(39)
			C1_remarks = iCurColumnPos(40)
			C1_error_desc = iCurColumnPos(41)
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
			C2_where_flag = iCurColumnPos(16)
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
											  "dialogWidth=760px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
											  "dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
											  "dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

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
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	If Col = C1_send_check Then
 		Exit Sub
	End If
	
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
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	'----------  Coding part  -------------------------------------------------------------   
	' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1.vspdData

		.Row = Row

		Select Case Col
			Case  C1_change_reason
				.Col = Col
				intIndex = .Value
				.Col = C1_change_reason_cd
				.Value = intIndex
		End Select
	End With
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
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If frm1.vspdData.MaxRows <= 0 Or NewCol < 0 Or NewRow <= 0 Then
		Exit Sub
	End If
	 
	Call vspdData_Click(NewCol, NewRow)
End Sub

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
		If frm1.txtSalesGrpCd.value = "" Then
			frm1.txtSalesGrpNm.value = ""
		End If
		If frm1.txtBizAreaCd.value = "" Then
			frm1.txtBizAreaNm.value = ""
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

		FncQuery = True	
	End With
	
	Call DBquery()
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
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtBillStatus
	Dim txtTransferStatus
	Dim txtStatusflag
	Dim txtKeyNo

	DbQuery = False
	
	With frm1
      If .rdoBillStatus1.checked = True Then
			txtBillStatus = .rdoBillStatus1.value
		ElseIf .rdoBillStatus2.checked = True Then
			txtBillStatus = .rdoBillStatus2.value
		ElseIf .rdoBillStatus3.checked = True Then
			txtBillStatus = .rdoBillStatus3.value
		End If
		
      If .rdoTransferStatus1.checked = True Then
			txtTransferStatus = .rdoTransferStatus1.value
		ElseIf .rdoTransferStatus2.checked = True Then
			txtTransferStatus = .rdoTransferStatus2.value
		ElseIf .rdoTransferStatus3.checked = True Then
			txtTransferStatus = .rdoTransferStatus3.value
		End If

      If .rdoStatusflag1.checked = True Then
			txtStatusflag = .rdoStatusflag1.value
		ElseIf .rdoStatusflag2.checked = True Then
			txtStatusflag = .rdoStatusflag2.value
		ElseIf .rdoStatusflag3.checked = True Then
			txtStatusflag = .rdoStatusflag3.value
		End If
		
		If Trim(.txtKeyNo.value) = "" Then
			txtKeyNo = "%"
		Else
			txtKeyNo = .txtKeyNo.value
		End If

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
		                      "&txtSupplierCd=" & Trim(.txtSupplierCd.value) & _
		                      "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value) & _
		                      "&txtBizAreaCd=" & Trim(.txtBizAreaCd.value) & _
		                      "&rdoBillStatus=" & Trim(txtBillStatus) & _
		                      "&rdoTransferStatus=" & Trim(txtTransferStatus) & _
		                      "&rdoStatusflag=" & Trim(txtStatusflag) & _
		                      "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text) & _
		                      "&txtKeyNo=" & Trim(txtKeyNo) & _
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

	ggoSpread.Source = frm1.vspdData 
	frm1.vspddata.Row = lgOldRow
	frm1.vspddata.Col = C1_tax_bill_no : iTaxBillNo = frm1.vspddata.Text
	frm1.vspddata.Col = C1_where_flag: strWhereFlag = frm1.vspddata.Text

	strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001 & _
								  "&txtTaxBillNo=" & Trim(iTaxBillNo) & _
								  "&txtWhereFlag=" & Trim(strWhereFlag)
	
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
		If .vspdData.MaxRows > 0 Then
			If Dbquery2 = False Then
				Call RestoreToolbar()
				Exit Function
			End If	
			
			Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
			Set gActiveElement = document.activeElement
		End If

		Call LayerShowHide(0)
		Call SetToolbar("110000000001111")																'⊙: 버튼 툴바 제어 
	End With
End Function

'========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim lRow
	Dim lGrpCnt     
	Dim strVal
	Dim strCheckSend, iSelectCnt
	Dim strCheckProc	

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, issue_dt_flag, success_flag
	
	Dim objDTI 
	Dim arrayM, userDN, userInfo, userInfoSet
	
	Dim RetFlag

	With frm1
		ggoSpread.Source = .vspdData

		If Row > 0 And Col = C1_send_check Then
			lGrpCnt = 1
			strVal = ""
			iSelectCnt = 0
			lgAllSelect = False
			.vspdData.Row = Row
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
					Exit Sub
				End If
			End If

			.vspdData.Col = C1_issue_dt_flag
			issue_dt_flag = .vspdData.text

			.vspdData.Col = C1_success_flag
			success_flag = .vspdData.text

			If issue_dt_flag = "Y" And success_flag = "Y" Then
				.vspdData.Col = C1_tax_bill_no
				RetFlag = DisplayMsgBox("205905", parent.VB_YES_NO, .vspdData.text, "X")   '☜ 바뀐부분 

				If RetFlag = VBNO Then
					Call LayerShowHide(0)
					Exit Sub
				End If
			End If

			.vspdData.Col = C1_where_flag :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			.vspdData.Col = C1_tax_bill_no :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			.vspdData.Col = C1_change_reason_cd :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			.vspdData.Col = C1_change_remark :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			.vspdData.Col = C1_change_remark2 :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
			.vspdData.Col = C1_change_remark3 :	strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

			lGrpCnt = lGrpCnt + 1
			iSelectCnt = iSelectCnt + 1

			.txtMaxRows.value = lGrpCnt - 1
			.txtSpread.value = strVal

			On Error Resume Next									'☜: Protect system from crashing
			Set objDTI = CreateObject("FSSmartBillDTI.DTIInterface")
			On Error Goto 0

			If Err.Number <> 0 Then
				MsgBox "스마트빌 인증키 모듈 생성 에러 입니다."
				Call LayerShowHide(0)
				Exit Sub
			End If

			userInfoSet = objDTI.getUserInfoForBulk()		' 인증서 확인 창을 띄웁니다.(인증서 비밀번호)
			arrayM = split(userInfoSet,";#;")				' 인증 값을 배열로 받음
			userDN = arrayM(1)									' 배열에서 인증정보 DN값을 추출
			userInfo = arrayM(2)									' 배열에서 인증정보 INFO값을 추출

			' 인증서 정보 추출여부 확인
			If Len(userDN) = 0 Then								' userDN이 0이면 실패
				Call LayerShowHide(0)
				Exit Sub
			End If

			.txtuserDN.value = userDN
			.txtuserInfo.value = userInfo

			Call LayerShowHide(1)
			Call ExecMyBizASP(frm1, BIZ_PGM_ID3)			' ☜: 비지니스 ASP 를 가동 
		End If
    End With
End Sub

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
							<TD CLASS="CLSSTABP">
								<TABLE ID="MyTab1" CELLSPACING=0 CELLPADDING=0>
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>(재)매출정발행등록</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
									 </TR>
								</TABLE>
							</TD>
        					<TD WIDTH=*>&nbsp;</TD>
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
												<input type=radio CLASS="RADIO" name="rdoBillStatus" id="rdoBillStatus1" value="%" tag = "11X" checked><label for="Radio1">전체</label>&nbsp;&nbsp;
										      <input type=radio CLASS="RADIO" name="rdoBillStatus" id="rdoBillStatus2" value="N" tag = "11X"><label for="Radio2">미전송</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoBillStatus" id="rdoBillStatus3" value="Y" tag = "11X"><label for="Radio3">전송</label>
                                 </TD>
											<TD CLASS="TD5"NOWRAP>송신상태</TD>
											<TD CLASS="TD6"NOWRAP>
												<input type=radio CLASS="RADIO" name="rdoTransferStatus" id="rdoTransferStatus1" value="%" tag = "11X" checked><label for="Radio4">전체</label>&nbsp;&nbsp;
										      <input type=radio CLASS="RADIO" name="rdoTransferStatus" id="rdoTransferStatus2" value="Y" tag = "11X"><label for="Radio5">성공</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoTransferStatus" id="rdoTransferStatus3" value="N" tag = "11X"><label for="Radio6">실패</label>
										  </TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>계산서형태</TD>
											<TD CLASS=TD6 NOWRAP>
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag1" value="%" tag = "11X" checked><label for="rdoCfmAll">전체</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag2" value="R" tag = "11X"><label for="rdoCfmreceipt">영수</label>&nbsp;&nbsp;
												<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoStatusflag3" value="D" tag = "11X"><label for="rdoCfmdemand">청구</label>
											</TD>
											<TD CLASS=TD5 NOWRAP>전송관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtKeyNo" SIZE=40 MAXLENGTH=40 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="전송관리번호"></TD>
										</TR>
										<TR>
											<TD CLASS=TD6 NOWRAP colspan="4" ><font color="red">본 화면은 세금계산서 발행 화면입니다. 발행 이후의 승인/거절/취소 등의 업무는 스마트빌 웹사이트에서만 가능합니다. 국세청전송 또한 스마트빌 웹사이트를 이용하세요.</font></TD>
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
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>										                            	ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
									</TR>
									<TR HEIGHT="40%">
										<TD WIDTH="100%" colspan="4"><SCRIPT LANGUAGE=JavaScript>										                             	ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData2><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
			<TR>
				<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
		<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
		<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserDN" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserInfo" tag="24" TABINDEX="-1">
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
