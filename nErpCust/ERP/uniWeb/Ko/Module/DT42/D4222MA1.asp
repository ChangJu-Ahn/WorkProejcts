<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 전자세금계산서(스마트빌(양방향))
'*  2. Function Name        : 
'*  3. Program ID           : D4222MA1
'*  4. Program Name         : 매출세금계산서(역발행)
'*  5. Program Desc         :    
'*  6. Component List       : 
'*  7. Modified date(First) : 2011/05/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
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
Const BIZ_PGM_ID  = "D4222MB1.asp"  'Main조회
Const BIZ_PGM_ID2 = "D4222MB2.asp"  'Dtl조회
Const BIZ_PGM_ID3 = "D4222MB3.asp"  '계산서보기

'==========================================  1.2.1 Global 상수 선언  ======================================
'=                       4.2 Constant variables 
'========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"

'==========================================================================================================

'add header datatable column
Dim 	C1_send_check           '선택
Dim     C1_conversation_id      '계산서번호
Dim 	C1_dti_type             '계산서종류
Dim 	C1_dti_type_nm          '계산서종류
Dim 	C1_seq_id               '계산서번호
Dim     C1_dti_wdate            '계산서일자
Dim 	C1_total_amount         '합계금액
Dim 	C1_sup_amount           '공급가액
Dim 	C1_tax_amount           '세액합계
Dim     C1_byr_com_regno        '공급받는자 사업자등록번호
Dim     C1_byr_com_name         '공급받는자 상호
Dim     C1_byr_email            '공급받는자 담당자 E-Mail
Dim     C1_ref_invoice_num      '참조송장번호 
Dim     C1_dti_status           '최종상태
Dim     C1_dti_status_nm        '최종상태
Dim     C1_disuse_reason        '취소/거부사유
Dim     C1_amend_code           '수정코드
Dim     C1_amend_code_nm        '수정사유
Dim     C1_remark               '비고
Dim     C1_remark2              '비고2
Dim     C1_remark3              '비고3
Dim     C1_return_code          '오류코드
Dim     C1_return_description   '오류설명
'Dim     C1_sup_com_regno        'Supply Com Regno
'Dim     C1_sup_com_email        'Supply Com Email


'add detail datatable column
Dim     C2_conversation_id      '계산서번호
Dim	    C2_item_cd              '품목코드
Dim	    C2_item_nm              '품목명
Dim	    C2_item_size            '규격  
Dim	    C2_unit_price           '단가
Dim	    C2_item_qty             '수량
Dim	    C2_sup_amount           '공급가액
Dim     C2_tax_amount           '세액합계
Dim     C2_currency_code        '통화코드
Dim     C2_remark               '비고
Dim     C2_item_md              '구입일자
Dim     C2_dti_line_num         '항목번호


Dim lgStrPrevKey1    
Dim lgQueryFlag					' 신규조회 및 추가조회 구분 Flag
Dim lgGridPoupMenu              ' Grid Popup Menu Setting
Dim lgAllSelect

Dim lgIsOpenPop
Dim IsOpenPop       
Dim lgPageNo_B
Dim lgSortKey_B
Dim lgOldRow1

'Const C_MaxKey = 3
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
'========================================================================================================= 
'========================================================================================================= 
Sub Form_Load()

   Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

   With frm1
      Call InitComboBox()
      Call InitSpreadSheet()
      Call InitSpreadSheet2

      Call SetDefaultVal
      Call InitVariables
 
      Call SetToolbar("110000000000111")										'⊙: 버튼 툴바 제어    	
 
      .txtSupplierCd.focus
            
      .btnEmail.disabled	= true
      .btnPrint.disabled	= true
      
      
   End With		
End Sub

'========================================================================================================= 
Sub InitComboBox()
   Dim iCodeArr 
   Dim iNameArr
   Dim iDx
	
	'계산서의 발행상태
    'Call CommonQueryRs(" B.MINOR_CD , B.MINOR_NM "," B_CONFIGURATION A INNER JOIN B_MINOR B ON (A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B. MINOR_CD) ", _
    '                     " A.MAJOR_CD='DT409' and A.SEQ_NO = 1 and B.MINOR_CD not in ('X', 'S') ORDER BY A.REFERENCE ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    Call CommonQueryRs(" MINOR_CD , MINOR_NM "," B_MINOR (NOLOCK)  ", " MAJOR_CD='DT409' and MINOR_CD in ('I','O','T','V', 'C','R','W','N','M','O')  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
    Call SetCombo2(frm1.cboBillStatus ,lgF0  ,lgF1  ,Chr(11))
        
    
     Call CommonQueryRs(" MINOR_CD , MINOR_NM "," B_MINOR (NOLOCK)  ", " MAJOR_CD='DT408'  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
    Call SetCombo2(frm1.cboAmendCode ,lgF0  ,lgF1  ,Chr(11))
			

End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
	
	 	C1_send_check           =  1     '선택	 	
        C1_conversation_id      =  2     '계산서번호
     	C1_dti_type             =  3     '계산서종류
     	C1_dti_type_nm          =  4     '계산서종류
     	C1_seq_id               =  5     '계산서번호
        C1_dti_wdate            =  6     '계산서일자
     	C1_total_amount         =  7     '합계금액
     	C1_sup_amount           =  8     '공급가액
     	C1_tax_amount           =  9     '세액합계
        C1_byr_com_regno        =  10     '공급받는자 사업자등록번호
        C1_byr_com_name         =  11     '공급받는자 상호
        C1_byr_email            =  12     '공급받는자 담당자 E-Mail
        C1_ref_invoice_num      =  13     '참조송장번호 
        C1_dti_status           =  14     '최종상태
        C1_dti_status_nm        =  15     '최종상태
        C1_disuse_reason        =  16     '취소/거부사유
        C1_amend_code           =  17     '수정코드
        C1_amend_code_nm        =  18     '수정사유
        C1_remark               =  19     '비고
        C1_remark2              =  20     '비고2
        C1_remark3              =  21     '비고3
        C1_return_code          =  22     '오류코드
        C1_return_description   =  23     '오류설명
        
        
        'C1_sup_com_regno        =  22     'Supply Com Regno
        'C1_sup_com_email        =  23     'Supply Com Email
	 		 	    	    	
End Sub

Sub InitSpreadPosVariables2()
	'add tab1 detail datatable column
	
		C2_conversation_id      =  1      '계산서번호
        C2_item_cd              =  2      '품목코드
        C2_item_nm              =  3      '품목명
        C2_item_size            =  4      '규격  
        C2_unit_price           =  5      '단가
        C2_item_qty             =  6      '수량
        C2_sup_amount           =  7      '공급가액
        C2_tax_amount           =  8      '세액합계
        C2_currency_code        =  9      '통화코드
        C2_remark               =  10     '비고
        C2_item_md              =  11     '구입일자
        C2_dti_line_num         =  12     '항목번호


End Sub

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
   
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKey1      = ""                                      '⊙: initializes Previous Key
    lgLngCurRows = 0									   'initializes Deleted Rows Count
    
    lgSortKey       = 1                                       '⊙: initializes sort direction
    lgSortKey1	    = "1"
    lgPageNo_B		= ""                          'initializes Previous Key for spreadsheet #2    
    
    lgOldRow = 0
    lgRow = 0
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	'승인의 일자는 전월 말일부터 당일
    Dim strYear, strMonth, strDay
    Dim StartDate
    Dim PreStartDate
    Dim EndDate
	EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

	Call	ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	
	PreStartDate = UNIDateAdd("D", -1, StartDate,Parent.gServerDateFormat)
			
	frm1.txtIssuedFromDt.text  = PreStartDate
	frm1.txtIssuedToDt.text    = EndDate
	
	
	frm1.txtSupplierCd.focus
	'lgGridPoupMenu          = GRID_POPUP_MENU_PRT
End Sub

'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()

	With frm1.vspdData	
	
	     ggoSpread.Source = frm1.vspdData
		 ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	    .ReDraw = false
        .MaxCols = C1_return_description + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols														'공통콘트롤 사용 Hidden Column
        .ColHidden = True
        .MaxRows = 0


		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetCheck	C1_send_check,		"선택",     			 	 4,  -10, "", True, -1
		ggoSpread.SSSetEdit		C1_conversation_id,	"계산서번호",				30, ,,50
		ggoSpread.SSSetEdit  	C1_dti_type,		"계산서종류",				15, ,,18		
		ggoSpread.SSSetEdit  	C1_dti_type_nm,		"계산서종류",				15, ,,18		
		ggoSpread.SSSetEdit  	C1_seq_id,          "계산서번호",       		15, ,,30
		ggoSpread.SSSetDate  	C1_dti_wdate,       "계산서일자",     			13, 2, parent.gDateFormat
		ggoSpread.SSSetFloat	C1_total_amount,	"합계금액",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetFloat	C1_sup_amount ,    	"공급가액",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_tax_amount ,    	"세액합계",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  	C1_byr_com_regno,   "공급받는자 사업자등록번호",	25, ,,30
		ggoSpread.SSSetEdit  	C1_byr_com_name ,   "공급받는자 상호",	        25, ,,50
		ggoSpread.SSSetEdit		C1_byr_email,		"공급받는자 담당자 E-Mail",	20, ,,40
		ggoSpread.SSSetEdit  	C1_ref_invoice_num, "참조송장번호",		        15, ,,30		
		ggoSpread.SSSetEdit  	C1_dti_status,		 "최종상태",				15, ,,18
		ggoSpread.SSSetEdit  	C1_dti_status_nm,	 "최종상태",			    20, ,,40		
		ggoSpread.SSSetEdit  	C1_disuse_reason,	 "취소/거부사유",			15, ,,50
		ggoSpread.SSSetEdit  	C1_amend_code,		 "수정코드",				15, ,,18		
		ggoSpread.SSSetEdit  	C1_amend_code_nm,	 "수정사유",			    15, ,,40						
		ggoSpread.SSSetEdit		C1_remark,			"비고",				    	20, ,,50
		ggoSpread.SSSetEdit		C1_remark2,			"비고2",		    		20, ,,50
		ggoSpread.SSSetEdit		C1_remark3,			"비고3",	    			20, ,,50		
		ggoSpread.SSSetEdit  	C1_return_code,         "오류코드",		15, ,,30
		ggoSpread.SSSetEdit  	C1_return_description,  "오류설명",		15, ,,40
				
								                        		
		'Call ggoSpread.MakePairsColumn(C1_change_reason_cd, C1_change_reason, "1")
      Call ggoSpread.SSSetColHidden(C1_dti_type, C1_dti_type, True)
      Call ggoSpread.SSSetColHidden(C1_dti_status, C1_dti_status, True)


		.ReDraw = True
	End With
	
	Call SetSpreadLock()
End Sub

Sub InitSpreadSheet2()

	Call initSpreadPosVariables2()
	With frm1.vspdData2	
		.MaxCols = C2_dti_line_num + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols												'☆: 사용자 별 Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		.ReDraw = False 
		ggoSpread.Spreadinit "V20090707",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos("B")
		
		ggoSpread.SSSetEdit		C2_conversation_id,	    "계산서번호",		30, ,,50		
		ggoSpread.SSSetEdit  	C2_item_cd, 			"품목", 			15, ,,18
		ggoSpread.SSSetEdit  	C2_item_nm, 			"품목명", 			30, ,,30		
		ggoSpread.SSSetEdit  	C2_item_size,			"규격", 			15, ,,40
        ggoSpread.SSSetFloat  	C2_unit_price, 		    "단가", 			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec	
		ggoSpread.SSSetFloat	C2_item_qty ,			"수량",				15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat  	C2_sup_amount, 			"공급가액", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_tax_amount, 			"세액합계", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetEdit  	C2_currency_code, 		"통화코드", 		15, ,,18		
		ggoSpread.SSSetEdit  	C2_remark, 		        "비고", 		    20, ,,50		
		ggoSpread.SSSetDate  	C2_item_md,             "구입일자",     	13, 2, parent.gDateFormat
        ggoSpread.SSSetEdit  	C2_dti_line_num,		"항목번호", 		15, ,,18						
		
		'Call ggoSpread.SSSetColHidden(C2_bill_no, C2_where_flag, True)

		.ReDraw = True
	End With	
	Call SetSpreadLock_B()
End Sub

'========================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

		'frm1.vspddata.col = C1_send_check
		'frm1.vspddata.row = 0
		'frm1.vspddata.ColHeadersShow = True

		'ggoSpread.SpreadLock    C_send_check,		-1, C_send_check
		ggoSpread.SpreadLock    C1_conversation_id, -1, C1_conversation_id
		ggoSpread.SpreadLock    C1_dti_type, 		-1, C1_dti_type		
		ggoSpread.SpreadLock    C1_dti_type_nm, 	-1, C1_dti_type_nm				
		ggoSpread.SpreadLock    C1_seq_id, 		    -1, C1_seq_id		
		ggoSpread.SpreadLock    C1_dti_wdate, 		-1, C1_dti_wdate				
		ggoSpread.SpreadLock    C1_total_amount , 	-1, C1_total_amount 		
		ggoSpread.SpreadLock    C1_sup_amount , 	-1, C1_sup_amount 		
		ggoSpread.SpreadLock    C1_tax_amount , 	-1, C1_tax_amount 		
		ggoSpread.SpreadLock    C1_byr_com_regno,   -1, C1_byr_com_regno
		ggoSpread.SpreadLock    C1_byr_com_name,    -1, C1_byr_com_name
		ggoSpread.SpreadLock	C1_byr_email ,		-1, C1_byr_email 
		ggoSpread.SpreadLock	C1_ref_invoice_num,	-1, C1_ref_invoice_num
		ggoSpread.SpreadLock    C1_dti_status, 		-1, C1_dti_status		
		ggoSpread.SpreadLock    C1_dti_status_nm, 	-1, C1_dti_status_nm		
		ggoSpread.SpreadLock    C1_disuse_reason,   -1, C1_disuse_reason		
		ggoSpread.SpreadLock    C1_amend_code, 		-1, C1_amend_code
		ggoSpread.SpreadLock    C1_amend_code_nm, 	-1, C1_amend_code_nm
		ggoSpread.SpreadLock	C1_remark,		    -1, C1_remark
		ggoSpread.SpreadLock	C1_remark2,		    -1, C1_remark2
		ggoSpread.SpreadLock	C1_remark3,		    -1, C1_remark3
		
		ggoSpread.SpreadLock	C1_return_code,	-1, C1_return_code
		ggoSpread.SpreadLock	C1_return_description,	    -1, C1_return_description
		
																		
		'ggoSpread.SSSetRequired	  C1_byr_emp_name,	-1, -1
																																
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
			
			C1_send_check           = iCurColumnPos(1)      '선택									
            C1_conversation_id      = iCurColumnPos(2)      '계산서번호
     	    C1_dti_type             = iCurColumnPos(3)      '계산서종류
     	    C1_dti_type_nm          = iCurColumnPos(4)      '계산서종류
     	    C1_seq_id               = iCurColumnPos(5)      '계산서번호
            C1_dti_wdate            = iCurColumnPos(6)      '계산서일자
     	    C1_total_amount         = iCurColumnPos(7)      '합계금액
     	    C1_sup_amount           = iCurColumnPos(8)      '공급가액
     	    C1_tax_amount           = iCurColumnPos(9)      '세액합계
            C1_byr_com_regno        = iCurColumnPos(10)     '공급받는자 사업자등록번호
            C1_byr_com_name         = iCurColumnPos(11)     '공급받는자 상호
            C1_byr_email            = iCurColumnPos(12)     '공급받는자 담당자 E-Mail
            C1_ref_invoice_num      = iCurColumnPos(13)     '참조송장번호 
            C1_dti_status           = iCurColumnPos(14)     '최종상태
            C1_dti_status_nm        = iCurColumnPos(15)     '최종상태
            C1_disuse_reason        = iCurColumnPos(16)     '취소/거부사유
            C1_amend_code           = iCurColumnPos(17)     '수정코드
            C1_amend_code_nm        = iCurColumnPos(18)     '수정사유
            C1_remark               = iCurColumnPos(19)     '비고
            C1_remark2              = iCurColumnPos(20)     '비고2
            C1_remark3              = iCurColumnPos(21)     '비고3            
            C1_return_code          = iCurColumnPos(22)     '오류코드
            C1_return_description   = iCurColumnPos(23)     '오류설명
                                                
        Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C2_conversation_id      = iCurColumnPos(1)      '계산서번호
            C2_item_cd              = iCurColumnPos(2)      '품목코드
            C2_item_nm              = iCurColumnPos(3)      '품목명
            C2_item_size            = iCurColumnPos(4)      '규격  
            C2_unit_price           = iCurColumnPos(5)      '단가
            C2_item_qty             = iCurColumnPos(6)      '수량
            C2_sup_amount           = iCurColumnPos(7)      '공급가액
            C2_tax_amount           = iCurColumnPos(8)      '세액합계
            C2_currency_code        = iCurColumnPos(9)      '통화코드
            C2_remark               = iCurColumnPos(10)     '비고
            C2_item_md              = iCurColumnPos(11)     '구입일자
            C2_dti_line_num         = iCurColumnPos(12)     '항목번호
													    
	End Select    
End Sub


Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처"					
	arrParam(1) = "b_biz_partner"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE In ('C', 'CS')"	
	arrParam(5) = "발행처"						

	arrField(0) = "bp_cd"					
	arrField(1) = "BP_FULL_NM"	
	arrField(2) = "bp_rgst_no"				

	arrHeader(0) = "거래처코드"				
	arrHeader(1) = "거래처처명"	
	arrHeader(2) = "사업자등록번호"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
											  Array(arrParam, arrField, arrHeader), _
											  "dialogWidth=760px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(1)
		frm1.txtSupplierNm.Value = arrRet(2)
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


Function OpenPopup(Byval strcode, Byval iWhere)
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	Dim strAmendCd
	Dim strBpCd

   If IsOpenPop = True Then Exit Function

   IsOpenPop = True


	Select Case iWhere
	    Case 1 
                                                          
         arrParam(0) = "수정코드팝업"
         arrParam(1) = "B_MINOR (nolock) " ' TABLE 명칭 
         arrParam(2) = strcode      ' Code Condition
         arrParam(3) = ""       ' Name Cindition
         arrParam(4) = " MAJOR_CD = 'DT408' "       ' Where Condition
         arrParam(5) = "수정코드"    ' 조건필드의 라벨 명칭 
         

         arrField(0) = "MINOR_CD"     ' Field명(0)
         arrField(1) = "MINOR_NM"     ' Field명(1)

         arrHeader(0) = "코드"    ' Header명(0)
         arrHeader(1) = "코드명"     ' Header명(1)
	    		
		Case 2
        
        
         frm1.vspddata.Col = C1_bp_cd
         strBpCd = Trim(frm1.vspddata.value) 
         
         arrParam(0) = "거래처담당자"
         arrParam(1) = "XXSB_DTI_BP_USER (nolock) " ' TABLE 명칭 
         arrParam(2) = strcode      ' Code Condition
         arrParam(3) = ""       ' Name Cindition
         arrParam(4) = " FND_BP_CD = '" & strBpCd & "'"    ' Where Condition
         arrParam(5) = "거래처담당자"    ' 조건필드의 라벨 명칭 
         

         arrField(0) = "FND_USER_NAME"          ' Field명(0)
         arrField(1) = "FND_BP_CD"              ' Field명(1)
         arrField(2) = "FND_USER_DEPT_NAME"     ' Field명(2)
         arrField(3) = "FND_USER_TEL_NUM"       ' Field명(3)
         arrField(4) = "FND_USER_EMAIL"         ' Field명(4)
         

         arrHeader(0) = "거래처담당자명"        ' Header명(0)
         arrHeader(1) = "거래처"                ' Header명(1)
         arrHeader(2) = "거래처부서명"          ' Header명(2)
         arrHeader(3) = "거래처전화번호"        ' Header명(3)
         arrHeader(4) = "거래처담당자E-Mail"    ' Header명(4)
         
            
		Case Else
		
	     Exit Function
   End Select
	 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
            "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: Yes; status: No;")

   IsOpenPop = False

   If arrRet(0) = "" Then
      Exit Function
   Else
		Call SetPopup(arrRet, iWhere)
   End If 
 
End Function


Function SetPopup(Byval arrRet, Byval iWhere)
   With frm1
      Select Case iWhere
	     Case 1   ' 
           .vspdData.Col = C1_amend_code
           .vspdData.Text = arrRet(0)
           .vspdData.Col = C1_remark
           .vspdData.Text = arrRet(1)
'           Call vspdData_Change(C_PuNo, .vspdData.Row)	     	     
	     Case 2   
            '.vspdData.Col = C1_byr_emp_name
           '.vspdData.Text = arrRet(0)
           
            '.vspdData.Col = C1_byr_dept_name
           '.vspdData.Text = arrRet(1)
           
           '.vspdData.Col = C1_byr_tel_num
           '.vspdData.Text = arrRet(2)
           
           '.vspdData.Col = C1_byr_email
           '.vspdData.Text = arrRet(3)
           
'           Call vspdData_Change(C_PuNo, .vspdData.Row)	     	     
      End Select
   End With
End Function



'이메일 재발송
Function fnEmail()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lGrpCnt, lRow
	Dim strWhere
	Dim IntRetCD
	Dim strVal
	Dim Status
	Dim messageNo
	Dim messageAmend
	Dim strConverid
	Dim strDtiStatus
	Dim RetFlag
	Dim iSelectCnt
	Dim EmailPublishFlag
    Dim DtiStatus
    Dim sbdescription
    Dim strByrEmail
    Dim strConvid
    
    Dim strSmartid, strSmartPW
    Dim strExpDate, NowDate, strYear, strMonth,  strDay
    Dim strUrlinfo, strIssueASP,  strURL
    Dim arrRet

    fnEmail = False

            
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 

    
    sbdescription  = ""
     
    
     '여러건의 세금계산서를 선택할수 없음(체크)
    if Check1() = false then
      Call LayerShowHide(0)
      exit function
    end if  
                   	      	
	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		iSelectCnt = 0

		'-----------------------
		'Data manipulate area
		'-----------------------		
						
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
						
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1"  Then
						                          
                  '발행대상체크
			      .vspdData.Col = C1_dti_status
			      strDtiStatus = Trim(.vspdData.text)
				
			     if	(strDtiStatus = "I")  or  (strDtiStatus = "C") or (strDtiStatus = "V") or (strDtiStatus = "N") then
					
					.vspdData.Col = C1_conversation_id
                     strConvid = Trim(.vspdData.Text)

                     
                     .vspdData.Col = C1_byr_email
                     strByrEmail = Trim(.vspdData.Text)
					
					
							         
                 else
                    .vspdData.Col = C1_conversation_id
			         messageNo = Trim(.vspdData.text)
											
			         Call DisplayMsgBox("DT4207","X", messageNo,"X")	            		
			         'DT4207:  %1은 E-Mail 전송 대상이 아닙니다.
			         Call LayerShowHide(0)
			         Exit Function
                                                            																			
			     end if	  
                 
                                                                                                        
        '------------웹페이지 호출 부분------------------------------------------  
                 If CommonQueryRs(" TOP 1 A.SMART_ID, A.SMART_PASSWORD "," XXSB_DTI_SM_USER A (nolock) "," A.FND_USER = '" & parent.gUsrID & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
                    strSmartid = Trim(Replace(lgF0,Chr(11),""))
                    strSmartPW = Trim(Replace(lgF1,Chr(11),""))
                 else
                    Call DisplayMsgBox("DT4118","X", "X","X")	               		
                    'DT4118:전자세금계산서의 사용자를 확인하세요
                     Call LayerShowHide(0)
                    Exit Function 	  
                 End if
                
      
                   If CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = 'DT400' AND MINOR_CD = '01' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
                        strUrlinfo =  Trim(Replace(lgF0,Chr(11),""))
                        
                        if strUrlinfo = "" then
                             Call DisplayMsgBox("DT4108","X", "X","X")	               		
                           'DT4108: URL 정보가 없습니다.
                            Call LayerShowHide(0)
                            Exit Function 	  
                        end if            
                   else
                       Call DisplayMsgBox("DT4108","X", "X","X")	               		
                       'DT4108: URL 정보가 없습니다.
                        Call LayerShowHide(0)
                       Exit Function 	  
                   End if
      
      
                strIssueASP = "XXSB_DTI_SEND_EMAIL.asp"

          
                strURL =  strUrlinfo & strIssueASP & "?conversation_id=" + strConvid + "&EMAIL=" + strByrEmail + "&ID=" + strSmartid + "&PASS=" + strSmartPW + "&STATUS=" + strDtiStatus + ""

            
           arrRet =  window.showModalDialog(strUrl ,, "dialogWidth=810px; dialogHeight=480px; center: Yes; help: No; resizable: No; status: no; scroll:Yes;")       
                                                                                            																																				     								
		   iSelectCnt = iSelectCnt + 1
       End If
     Next

		if iSelectCnt = 0 then
			Call DisplayMsgBox("DT4205","X", "X","X")	
			'DT4205: 선택된 행이 없습니다.            		
			Call LayerShowHide(0)
			Exit Function
		end if

	End With
        
         Call LayerShowHide(0)
         'Call MainQuery
	
	fnEmail = True	

End Function



'계산서 보기
Function fnPrint()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lGrpCnt, lRow
	Dim strWhere
	Dim IntRetCD
	Dim strVal
	Dim Status
	Dim messageNo
	Dim messageAmend
	Dim strConverid
	Dim strDtiStatus
	Dim RetFlag
	Dim iSelectCnt
	Dim AcceptFlag
    Dim DtiStatus
    Dim sbdescription
    Dim PrintFlag

    fnPrint = False
    
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
    DtiStatus = ""
    sbdescription  = ""
      	      	
	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		iSelectCnt = 0

		'-----------------------
		'Data manipulate area
		'-----------------------		
						
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
						
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1"  Then
						                                                                                																																				     
				'--------------------------------------------------------------------------		
				 strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_conversation_id     :	strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep		'Conversation ID    '2
  			     		     
																																														  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1
           End If
        Next

		if iSelectCnt = 0 then
			Call DisplayMsgBox("DT4205","X", "X","X")	
			'DT4205: 선택된 행이 없습니다.            		
			Call LayerShowHide(0)
			Exit Function
		end if

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID3)
	
	fnPrint = True	

End Function


Function printWebCall(batchid)
	DIm IntRetCD
	DIm strURL
    Dim StrRegNo
    Dim strUrlinfo
    Dim strIssueASP
    Dim arrRet
    Dim strBatchid
    Dim strConvid
    
    
	DIm lRow
	
   
    strBatchid = ""
    strConvid = ""
    
                                               
       If CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = 'DT400' AND MINOR_CD = '01' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
            strUrlinfo =  Trim(Replace(lgF0,Chr(11),""))
            
            if strUrlinfo = "" then
                 Call DisplayMsgBox("DT4108","X", "X","X")	               		
           'DT4108: URL 정보가 없습니다.
           Exit Function 	  
            end if            
       else
           Call DisplayMsgBox("DT4108","X", "X","X")	               		
           'DT4108: URL 정보가 없습니다.
           Exit Function 	  
       End if
                            
       strURL =  strUrlinfo & "XXSB_DTI_PRINT.asp?batch_id=" + batchid +  "&SORTFIELD=A&SORTORDER=1 "
                                                                                        
       arrRet =  window.showModalDialog(strUrl ,, "dialogWidth=810px; dialogHeight=480px; center: Yes; help: No; resizable: No; status: no; scroll:Yes;")       
           
        'DbSaveOk
                                                 	
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


Sub cboBillStatus_OnChange()

'	If frm1.txtAcctCd.value <> "" Then
'		Call DisplayMsgBox("800489","x",frm1.txtAcctCd.alt,"x")
'		frm1.txtAcctCd.focus
'		Exit Sub
'	End If  

'  	Call fncquery()
    frm1.vspdData.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
    call InitVariables
    
    call dbquery

    'if frm1.cboBudgetType.value = "2" then
	'	Call ggoOper.SetReqAttr(frm1.txtVercode, "N")
	'else
	'	Call ggoOper.SetReqAttr(frm1.txtVercode, "D")
    'end if 
End Sub	



Sub vspdData_Change(ByVal Col , ByVal Row )

   Dim strAmend
   Dim strEmpName
   Dim strBpCd


    Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        
                                                                                                                   
    End Select

    'ggoSpread.Source = frm1.vspdData
    'ggoSpread.UpdateRow Row
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

 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 		
 		frm1.vspddata.Row = frm1.vspdData.ActiveRow
		frm1.vspddata.Col = C1_conversation_id
    
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
            frm1.vspddata.Col = C1_conversation_id
            frm1.vspddata2.MaxRows = 0

            lgOldRow = Row
            
            lgStrPrevKey1 = ""  
                
            If DbQuery2 = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
            '------ Developer Coding part (End)
        End If
 	End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SP2C" 
    Set gActiveSpdSheet = frm1.vspdData2
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort ,lgSortKey1
            lgSortKey1 = 1
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
	'With frm1.vspdData

	'	.Row = Row

	'	Select Case Col
	'		Case  C1_change_reason
	'			.Col = Col
	'			intIndex = .Value
	'			.Col = C1_change_reason_cd
	'			.Value = intIndex
	'	End Select
	'End With
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


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
		    Select Case Col
	'		    Case C1_amend_pop
	'			   .Row = Row
	'	           .Col = C1_amend_code
	'	            Call OpenPopup(.Text, 1)
	'		    Case C1_byr_emp_pop
	'		        frm1.vspddata.Col = C1_byr_emp_name			    
	'			    Call OpenPopup(.text, 2)        
		    End Select    
	    End If								
    End With                     
End Sub


'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    if OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    if OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery2 = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

'#########################################################################################################
'												4. Common Function부 
'=========================================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False																		'⊙: Processing is NG


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
	    
	       If Not chkField(Document, "1") Then									         '☜: This function check required field
               Exit Function
            End If
	    
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
		
	End With
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

		Call InitVariables 																	'⊙: Initializes local global variables
        
        Call DisableToolBar(parent.TBC_QUERY)
	    If DbQuery = False Then
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
    
    Select Case gActiveSpdSheet.id
		Case "vspdData"
			Call InitSpreadSheet()      
			Call ggoSpread.ReOrderingSpreadData()
		Case "vspdData2"
			Call InitSpreadSheet2		
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ReOrderingSpreadData()	
	End Select     
End Sub

'========================================================================================
Function DbQuery() 
    Dim strVal
    
	DbQuery = False
	    
    If LayerShowHide(1) = False then
		Exit Function 
   	End if
		
	With Frm1
	    If lgIntFlgMode = parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜: 
		    strVal = strVal & "&txtIssuedFromDt=" & Trim(.htxtIssuedFromDt.value)		
		    strVal = strVal & "&txtIssuedToDt=" & Trim(.htxtIssuedToDt.value)				    
		    strVal = strVal & "&cboBillStatus=" & Trim(.hcboBillStatus.value)						
		    strVal = strVal & "&txtSupplierNm=" & Trim(.htxtSupplierNm.value)
		    strVal = strVal & "&cboAmendCode=" & Trim(.hcboAmendCode.value)						
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        Else
		    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
		    strVal = strVal     &  "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text)
		    strVal = strVal     &  "&txtIssuedToDt=" & Trim(.txtIssuedToDt.text)		    
		    strVal = strVal     &  "&cboBillStatus=" & Trim(.cboBillStatus.Value)						    
		    strVal = strVal     &  "&txtSupplierNm=" & Replace(Trim(.txtSupplierNm.Value),"-","")					
		    strVal = strVal     &  "&cboAmendCode=" & Trim(.cboAmendCode.Value)						    
            strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
            strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
        End if
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
	Dim iConvid
    
    if LayerShowHide(1) = False then
	   Exit Function
	end if                                                         '☜: Show Processing Message
    
	ggoSpread.Source = frm1.vspdData 
	frm1.vspddata.Row = lgOldRow
	frm1.vspddata.Col = C1_conversation_id : iConvid = Trim(frm1.vspddata.Text)
									  
    With Frm1
	    strVal = BIZ_PGM_ID2 & "?txtMode="            & parent.UID_M0001						         
	    strVal = strVal     &  "&txtConvid=" & Trim(iConvid) 
        strVal = strVal     &  "&txtMaxRows="         & .vspdData2.MaxRows
        strVal = strVal     &  "&lgStrPrevKey1=" & lgStrPrevKey1                 '☜: Next key tag
   End With    								  
								  	
	Call RunMyBizASP(MyBizASP, strVal)

	DbQuery2 = True                                                     
End Function

'========================================================================================
Function DbQueryOk()																		'☆: 조회 성공후 실행로직 

    Dim iRow
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE																'⊙: Indicates that current mode is Update mode
     Call ggoOper.LockField(Document, "Q")	               
     
	lgOldRow = 1
	frm1.vspdData.Col = 1
	frm1.vspdData.Row = 1
                
	With frm1
	
	    Select Case Trim(frm1.cboBillStatus.value)
                 case "I":
                     .btnPrint.disabled = true
                     .btnemail.disabled = false
                     
                 case "O":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "T":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "V":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "C":
                     .btnPrint.disabled = false
                     .btnemail.disabled = false
                     
                 case "N":
                     .btnPrint.disabled = true
                     .btnemail.disabled = false
                     
                 case "S":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "R":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "W":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                 case else
                     .btnPrint.disabled = true
                     .btnemail.disabled = true       
                                  
        End Select
			
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

Function DbQueryNotOk()

    With frm1
        Select Case Trim(.cboBillStatus.value)
              case "I":
                     .btnPrint.disabled = true
                     .btnemail.disabled = false
                     
                 case "O":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "T":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "V":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "C":
                     .btnPrint.disabled = false
                     .btnemail.disabled = false
                     
                 case "N":
                     .btnPrint.disabled = true
                     .btnemail.disabled = false
                     
                 case "S":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "R":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case "W":
                     .btnPrint.disabled = true
                     .btnemail.disabled = true
                     
                 case else
                     .btnPrint.disabled = true
                     .btnemail.disabled = true   
                     
        End Select
     End With   

End Function

'========================================================================================================
' Function Name : DbQueryOk2
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk2()
    DIm Row
													     
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	frm1.vspdData.focus	
	
End Function


'========================================================================================================
' Function Name : DbQueryOk2
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryNotOk2()
    DIm Row
													     
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	frm1.vspdData.focus	
	
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
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
    'Call InitVariables															'⊙: Initializes local global variables
    
    frm1.txtIsusse.value = ""
    
	Call MainQuery()
End Function


'========================================================================================
Function DbSaveOk2()													'☆: 저장 성공후 실행 로직 
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
	
	frm1.txtIsusse.value = ""
	
    'Call InitVariables															'⊙: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================
Function DbSaveNotOk2()													'☆: 저장 성공후 실행 로직 

    
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


'계산서 보기(웹페이지 호출은 안한상태)
Function  ExeNumOk()

'    Call DisableToolBar(parent.TBC_QUERY)
            
    Call printWebCall(Trim(frm1.hbatchid.value))
    
End Function



Function  ExeNumNot()

    Call DisplayMsgBox("800407","X","X","X")	'작업실행중 에러입니다.
		
End Function

Function Check1()
    Dim IntRetCD,imRow
    Dim iRow
    Dim iSelectCnt    
        
    iSelectCnt = 0
     
    Check1 = False
          
    With frm1.vspdData	
		For iRow = 1 To .MaxRows 
			.Row = iRow
			
			 .Row = iRow
			 .Col = C1_send_check

			If .text = "1" Then			
			   
			   iSelectCnt = iSelectCnt + 1				    															
			end if	 					 		 					 					 		    
		Next
	End With
	
	if iSelectCnt > 1 then
		IntRetCD = DisplayMsgBox("DT4201","X","X","X")        
		'DT4201: 여러건의 세금계산서를 동시에 처리할수 없습니다.
		Exit Function
	end if
    
    Check1 = True
    	
End Function



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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
 											<TD CLASS="TD5" NOWRAP>거래처</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSupplierCd" SIZE=15 MAXLENGTH=60 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
												<INPUT TYPE=TEXT AlT="거래처" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD> 											
										</TR>
										<TR> 											
											<TD CLASS="TD5"NOWRAP>계산서상태</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboBillStatus" ALT="세금계산서상태" CLASS=cboNormal TAG="11" style="width:170px"><OPTION VALUE=""></OPTION></SELECT>
											</TD> 											 											
											<TD CLASS="TD5"NOWRAP>수정사유</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboAmendCode" ALT="수정사유" CLASS=cboNormal TAG="11" style="width:150px"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
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
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
									</TR>
									<TR HEIGHT="40%">
										<TD WIDTH="100%" colspan="4"><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData2><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
							<TD WIDTH=10>&nbsp</TD>
							<TD><BUTTON NAME="btnPrint" CLASS="CLSSBTN" OnClick="VBScript:Call fnPrint()">계산서보기</BUTTON>&nbsp;
                                <BUTTON NAME="btnEmail" CLASS="CLSSBTN" OnClick="VBScript:Call fnEmail()">이메일 재발송</BUTTON>&nbsp;</TD>							  							    
							<TD WIDTH=10>&nbsp</TD>																												
						</TR>
  					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
		<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
		<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserDN" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="txtuserInfo" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="hbatchid" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="hconvid" tag="24" TABINDEX="-1">		
		<INPUT TYPE=HIDDEN NAME="hchangecode" tag="24" TABINDEX="-1">
		<INPUT TYPE=HIDDEN NAME="hcomregno" tag="24" TABINDEX="-1">
						
		<INPUT TYPE=HIDDEN NAME="hcboBillStatus"   TAG="24"  Tabindex="-1">
        <INPUT TYPE=HIDDEN NAME="htxtSupplierNm"   TAG="24"  Tabindex="-1">
        <INPUT TYPE=HIDDEN NAME="hcboAmendCode"   TAG="24"  Tabindex="-1">        
        <INPUT TYPE=HIDDEN NAME="htxtIssuedFromDt" TAG="24"  Tabindex="-1">
        <INPUT TYPE=HIDDEN NAME="htxtIssuedToDt"   TAG="24"  Tabindex="-1">

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
