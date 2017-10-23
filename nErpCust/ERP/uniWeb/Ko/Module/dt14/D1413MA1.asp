<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : D1413MA1
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
Const BIZ_PGM_ID  = "D1413MB1.asp"
Const BIZ_PGM_ID2 = "D1413MB2.asp"
Const BIZ_PGM_ID3 = "D1413MB3.asp"
Const BIZ_PGM_ID4 = "D1413MB4.asp"
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
Dim C1_gl_no             
Dim C1_vat_no            
Dim C1_inv_no  
Dim C1_inv_amend_type
Dim C1_inv_amend_type_nm  
Dim C1_remark            
Dim C1_remark2           
Dim C1_remark3         
Dim C1_issued_dt         
Dim C1_io_fg             
Dim C1_bp_cd             
Dim C1_bp_nm             
Dim C1_made_vat_fg       
Dim C1_doc_cur           
Dim C1_vat_type          
Dim C1_vat_type_nm       
Dim C1_vat_rate          
Dim C1_net_amt           
Dim C1_net_loc_amt       
Dim C1_vat_amt          
Dim C1_vat_loc_amt      
Dim C1_credit_cd         
Dim C1_report_biz_area_cd
Dim C1_tax_biz_area_nm   
Dim C1_biz_area_cd       
Dim C1_biz_area_nm       
Dim C1_ref_no            
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
Dim C2_item
Dim C2_item_std
Dim C2_item_prc
Dim C2_item_qty
Dim C2_item_date 
Dim C2_item_amt
Dim C2_item_tax
Dim C2_item_memo 
Dim C2_inv_no
Dim C2_vat_no
Dim C2_inv_item_seq_no


'Hidden Grid
Dim C3_item
Dim C3_item_std
Dim C3_item_prc
Dim C3_item_qty
Dim C3_item_date 
Dim C3_item_amt
Dim C3_item_tax
Dim C3_item_memo 
Dim C3_inv_no
Dim C3_vat_no
Dim C3_inv_item_seq_no
   

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
Dim lgSortKey1
Dim lgSortKey2

Const C_MaxKey = 3
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
'========================================================================================================= 
'========================================================================================================= 
Sub Form_Load()
   On Error Resume Next
   
   Call LoadInfTB19029
   Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
   
   Call ggoOper.LockField(Document, "N")                                   
   
   With frm1

      Call InitSpreadSheet("*")

      Call SetDefaultVal
      Call InitVariables
      Call InitComboBox()
      Call InitSpreadComboBox
 
      Call SetToolbar("110000000000111")										'⊙: 버튼 툴바 제어    	
 
      .popTaxRecipient.focus
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


Sub InitSpreadPosVariables(ByVal pvSpdNo)

    If pvSpdNo = "A" Or pvSpdNo = "*" Then
		C1_send_check        	=	1
		C1_proc_flag_nm      	=	2
		C1_is_send_nm        	=	3
		C1_gl_no             	=	4
		C1_vat_no            	=	5
		C1_inv_no  	=	6
		C1_inv_amend_type	=	7
		C1_inv_amend_type_nm  	=	8
		C1_remark            	=	9
		C1_remark2           	=	10
		C1_remark3         	=	11
		C1_issued_dt         	=	12
		C1_io_fg             	=	13
		C1_bp_cd             	=	14
		C1_bp_nm             	=	15
		C1_made_vat_fg       	=	16
		C1_doc_cur           	=	17
		C1_vat_type          	=	18
		C1_vat_type_nm       	=	19
		C1_vat_rate          	=	20
		C1_net_amt           	=	21
		C1_net_loc_amt       	=	22
		C1_vat_amt          	=	23
		C1_vat_loc_amt      	=	24
		C1_credit_cd         	=	25
		C1_report_biz_area_cd	=	26
		C1_tax_biz_area_nm   	=	27
		C1_biz_area_cd       	=	28
		C1_biz_area_nm       	=	29
		C1_ref_no            	=	30
		C1_disuse_reason     	=	31
		C1_legacy_pk         	=	32
		C1_sale_no             	=	33
		C1_proc_flag         	=	34
		C1_is_send           	=	35
		C1_issue_dt_fg       	=	36

        C1_name              = 37
        C1_name_B            = 38
        C1_email             = 39
        C1_bp_seq            = 40
        C1_name2             = 41
        C1_name2_B           = 42
        C1_email2            = 43
        C1_bp_seq2           = 44

	End If
	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
	'add tab1 detail datatable column
	
		C2_item	= 1
		C2_item_std	= 2
		C2_item_prc	= 3
		C2_item_qty	= 4
		C2_item_date 	= 5
		C2_item_amt	= 6
		C2_item_tax	= 7
		C2_item_memo 	= 8
		C2_inv_no	= 9
		C2_vat_no	= 10
		C2_inv_item_seq_no	= 11
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
	'add tab1 detail datatable column
	
		C3_item	= 1
		C3_item_std	= 2
		C3_item_prc	= 3
		C3_item_qty	= 4
		C3_item_date 	= 5
		C3_item_amt	= 6
		C3_item_tax	= 7
		C3_item_memo 	= 8
		C3_inv_no	= 9
		C3_vat_no	= 10
		C3_inv_item_seq_no	= 11
	End If
	
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
	frm1.popTaxRecipient.focus
	'lgGridPoupMenu          = GRID_POPUP_MENU_PRT
End Sub

'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    'On Error Resume Next
	
	Call initSpreadPosVariables(pvSpdNo)
	
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then

		With frm1.vspdData	
			.MaxCols = C1_bp_seq2 + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
			.Col = .MaxCols												'☆: 사용자 별 Hidden Column
			.ColHidden = True
			.MaxRows = 0
			ggoSpread.Source = frm1.vspdData
			.ReDraw = False
			ggoSpread.Spreadinit "V20090709",, parent.gAllowDragDropSpread
			
			Call GetSpreadColumnPos("A")

			' uniGrid1 setting
			
		    ggoSpread.SSSetCheck  C1_send_check        ,""                , 4,  -10, "", True, -1
		    ggoSpread.SSSetEdit   C1_proc_flag_nm      ,"계산서상태"      ,10
		    ggoSpread.SSSetEdit   C1_is_send_nm        ,"송신상태"        ,10
		    ggoSpread.SSSetEdit   C1_gl_no             ,"전표번호"        ,10
		    ggoSpread.SSSetEdit   C1_vat_no            ,"VAT No"          ,10
		    ggoSpread.SSSetEdit   C1_inv_no            ,"계산서번호"      ,10
		    ggoSpread.SSSetCombo   C1_inv_amend_type    ,"수정사유"       ,15   
		    ggoSpread.SSSetCombo   C1_inv_amend_type_nm ,"수정사유"       ,15 
		    ggoSpread.SSSetEdit   C1_remark            ,"비고1"           ,15, ,,150
		    ggoSpread.SSSetEdit   C1_remark2           ,"비고2"           ,15, ,,150
		    ggoSpread.SSSetEdit   C1_remark3           ,"비고3"           ,15, ,,150
		    ggoSpread.SSSetDate   C1_issued_dt         ,"발행일"          ,10, 2, parent.gDateFormat
		    ggoSpread.SSSetEdit   C1_io_fg             ,"매입매출구분"    ,10
		    ggoSpread.SSSetEdit   C1_bp_cd             ,"거래처"          ,10
		    ggoSpread.SSSetEdit   C1_bp_nm             ,"거래처명"        ,15
		    ggoSpread.SSSetEdit   C1_made_vat_fg       ,"파일생성여부"    ,10
		    ggoSpread.SSSetEdit   C1_doc_cur           ,"통화"            ,8
		    ggoSpread.SSSetEdit   C1_vat_type          ,"부가세유형"      ,8
		    ggoSpread.SSSetEdit   C1_vat_type_nm       ,"부가세유형명"    ,12
		    ggoSpread.SSSetFloat  C1_vat_rate          ,"VAT 율"          ,10, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C1_net_amt           ,"공급가액"        ,18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C1_net_loc_amt       ,"공급가액(자국)"  ,18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C1_vat_amt           ,"VAT 금액"        ,18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C1_vat_loc_amt       ,"VAT 금액(자국)"  ,18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetEdit   C1_credit_cd         ,"카드번호"        ,10
		    ggoSpread.SSSetEdit   C1_report_biz_area_cd,"세금신고사업장"  ,11
		    ggoSpread.SSSetEdit   C1_tax_biz_area_nm   ,"세금신고사업장명",15
		    ggoSpread.SSSetEdit   C1_biz_area_cd       ,"발생사업장"      ,10
		    ggoSpread.SSSetEdit   C1_biz_area_nm       ,"발생사업장명"    ,15
		    ggoSpread.SSSetEdit   C1_ref_no            ,"참조번호"        ,15
		    ggoSpread.SSSetEdit   C1_disuse_reason     ,"반려/폐기사유"   ,15
		    ggoSpread.SSSetEdit   C1_legacy_pk         ,"타시스템 PK"     ,15
		    ggoSpread.SSSetEdit   C1_sale_no           ,"거래명세서번호"  ,15
		    
		    ggoSpread.SSSetEdit   C1_proc_flag         ,"계산서상태"      ,10
		    ggoSpread.SSSetEdit   C1_is_send           ,"송신상태"        ,10
		    ggoSpread.SSSetEdit   C1_issue_dt_fg       ,"계산서발행여부"  ,10
		         
            ggoSpread.SSSetEdit     C1_name   ,         "담당자",               10, ,,30
            ggoSpread.SSSetButton   C1_name_B
            ggoSpread.SSSetEdit     C1_email    ,       "E-Mail",               20
            ggoSpread.SSSetEdit     C1_bp_seq   ,       "순번",                 10
            ggoSpread.SSSetEdit     C1_name2    ,       "담당자2",              10, ,,30
            ggoSpread.SSSetButton   C1_name2_B
            ggoSpread.SSSetEdit     C1_email2   ,       "E-Mail2",              20
            ggoSpread.SSSetEdit     C1_bp_seq2,         "순번2",                10
        
			Call ggoSpread.SSSetColHidden(C1_io_fg         , C1_io_fg         , True)
			Call ggoSpread.SSSetColHidden(C1_proc_flag     , C1_proc_flag     , True)
			Call ggoSpread.SSSetColHidden(C1_is_send       , C1_is_send       , True)
			Call ggoSpread.SSSetColHidden(C1_issue_dt_fg   , C1_issue_dt_fg   , True)
			Call ggoSpread.SSSetColHidden(C1_inv_amend_type   , C1_inv_amend_type   , True)
			
            Call ggoSpread.SSSetColHidden(C1_bp_seq, C1_bp_seq, True)
            Call ggoSpread.SSSetColHidden(C1_bp_seq2, C1_bp_seq2, True)

            Call ggoSpread.MakePairsColumn(C1_name, C1_name_B)
            Call ggoSpread.MakePairsColumn(C1_name2, C1_name2_B)

			.ReDraw = True
		End With

		Call SetSpreadLock("A")
		
	End If	'If pvSpdNo = "A" Or pvSpdNo = "*" Then
	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
	
		With frm1.vspdData2	
			.MaxCols = C2_inv_item_seq_no + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
			.Col = .MaxCols												'☆: 사용자 별 Hidden Column
			.ColHidden = True

			.MaxRows = 0
			ggoSpread.Source = frm1.vspdData2
			.ReDraw = False 
			ggoSpread.Spreadinit "V20090708",, parent.gAllowDragDropSpread

			Call GetSpreadColumnPos("B")
			
		    ggoSpread.SSSetEdit   C2_item     , "품목"      , 18
		    ggoSpread.SSSetEdit   C2_item_std , "규격"      , 20
		    ggoSpread.SSSetFloat  C2_item_prc , "단가"      , 15, parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C2_item_qty , "수량"      , 15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec
		    ggoSpread.SSSetDate   C2_item_date, "발행일"    , 10, 2, parent.gDateFormat
		    ggoSpread.SSSetFloat  C2_item_amt , "공급가액"  , 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C2_item_tax , "VAT 금액"  , 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetEdit   C2_item_memo, "비고"      , 15

		    ggoSpread.SSSetEdit   C2_inv_no   , "계산서번호", 15
		    ggoSpread.SSSetEdit   C2_vat_no   , "VAT No"    , 15
		    ggoSpread.SSSetEdit   C2_inv_item_seq_no   , "순번"    , 10
		    
			Call ggoSpread.SSSetColHidden(C2_inv_no         , C2_inv_no         , True)
			Call ggoSpread.SSSetColHidden(C2_vat_no         , C2_vat_no         , True)
		    Call ggoSpread.SSSetColHidden(C2_inv_item_seq_no         , C2_inv_item_seq_no         , True)
		    
			.ReDraw = True
		End With	
		'Call SetSpreadLock2()
	

	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then
	
		With frm1.vspdData3	
			.MaxCols = C3_inv_item_seq_no + 1								'☜: 최대 Columns의 항상 1개 증가시킴 
			.Col = .MaxCols												'☆: 사용자 별 Hidden Column
			.ColHidden = True
			.MaxRows = 0
			ggoSpread.Source = frm1.vspdData3
			.ReDraw = False 

			ggoSpread.Spreadinit
			
			Call GetSpreadColumnPos("C")
		
		    ggoSpread.SSSetEdit   C3_item     , "품목"      , 10
		    ggoSpread.SSSetEdit   C3_item_std , "규격"      , 10
		    ggoSpread.SSSetFloat  C3_item_prc , "단가"      , 15, parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C3_item_qty , "수량"      , 15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec
		    ggoSpread.SSSetDate   C3_item_date, "발행일"    , 10, 2, parent.gDateFormat
		    ggoSpread.SSSetFloat  C3_item_amt , "공급가액"  , 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat  C3_item_tax , "VAT 금액"  , 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetEdit   C3_item_memo, "비고"      , 10

		    ggoSpread.SSSetEdit   C3_inv_no   , "계산서번호", 10
		    ggoSpread.SSSetEdit   C3_vat_no   , "VAT No"    , 10
		    ggoSpread.SSSetEdit   C3_inv_item_seq_no   , "순번"    , 10
		    	    
			.ReDraw = True
		End With	
		'Call SetSpreadLock2()
	
	End If	
End Sub



'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	With frm1
		If pvSpdNo = "A" Then
			ggoSpread.Source = .vspdData
			.vspdData.ReDraw = False

			ggoSpread.SpreadLock 	C1_proc_flag_nm      ,  -1,C1_proc_flag_nm
			ggoSpread.SpreadLock 	C1_is_send_nm        ,  -1,C1_is_send_nm
			ggoSpread.SpreadLock 	C1_gl_no             ,  -1,C1_gl_no
			ggoSpread.SpreadLock 	C1_vat_no            ,  -1,C1_vat_no
			ggoSpread.SpreadLock 	C1_inv_no            ,  -1,C1_inv_no
			ggoSpread.SpreadUnLock 	C1_inv_amend_type    ,  -1,C1_inv_amend_type
			ggoSpread.SpreadUnLock 	C1_inv_amend_type_nm ,  -1,C1_inv_amend_type_nm
			ggoSpread.SpreadUnLock 	C1_remark            ,  -1,C1_remark
			ggoSpread.SpreadUnLock 	C1_remark2           ,  -1,C1_remark2
			ggoSpread.SpreadUnLock 	C1_remark3           ,  -1,C1_remark3
			ggoSpread.SpreadLock 	C1_issued_dt         ,  -1,C1_issued_dt
			ggoSpread.SpreadLock 	C1_io_fg             ,  -1,C1_io_fg
			ggoSpread.SpreadLock 	C1_bp_cd             ,  -1,C1_bp_cd
			ggoSpread.SpreadLock 	C1_bp_nm             ,  -1,C1_bp_nm
			ggoSpread.SpreadLock 	C1_made_vat_fg       ,  -1,C1_made_vat_fg
			ggoSpread.SpreadLock 	C1_doc_cur           ,  -1,C1_doc_cur
			ggoSpread.SpreadLock 	C1_vat_type          ,  -1,C1_vat_type
			ggoSpread.SpreadLock 	C1_vat_type_nm       ,  -1,C1_vat_type_nm
			ggoSpread.SpreadLock 	C1_vat_rate          ,  -1,C1_vat_rate
			ggoSpread.SpreadLock 	C1_net_amt           ,  -1,C1_net_amt
			ggoSpread.SpreadLock 	C1_net_loc_amt       ,  -1,C1_net_loc_amt
			ggoSpread.SpreadLock 	C1_vat_amt           ,  -1,C1_vat_amt
			ggoSpread.SpreadLock 	C1_vat_loc_amt       ,  -1,C1_vat_loc_amt
			ggoSpread.SpreadLock 	C1_credit_cd         ,  -1,C1_credit_cd
			ggoSpread.SpreadLock 	C1_report_biz_area_cd,  -1,C1_report_biz_area_cd
			ggoSpread.SpreadLock 	C1_tax_biz_area_nm   ,  -1,C1_tax_biz_area_nm
			ggoSpread.SpreadLock 	C1_biz_area_cd       ,  -1,C1_biz_area_cd
			ggoSpread.SpreadLock 	C1_biz_area_nm       ,  -1,C1_biz_area_nm
			ggoSpread.SpreadLock 	C1_ref_no            ,  -1,C1_ref_no
			ggoSpread.SpreadLock 	C1_disuse_reason     ,  -1,C1_disuse_reason
			ggoSpread.SpreadLock 	C1_legacy_pk         ,  -1,C1_legacy_pk
			ggoSpread.SpreadLock 	C1_sale_no           ,  -1,C1_sale_no
			ggoSpread.SpreadLock 	C1_proc_flag         ,  -1,C1_proc_flag
			ggoSpread.SpreadLock 	C1_is_send           ,  -1,C1_is_send
			ggoSpread.SpreadLock 	C1_issue_dt_fg       ,  -1,C1_issue_dt_fg
			
			ggoSpread.SSSetRequired	C1_inv_amend_type	,	-1	,-1
			ggoSpread.SSSetRequired	C1_inv_amend_type_nm,	-1	,-1
			ggoSpread.SSSetRequired	C1_remark	,	-1	,-1

            ggoSpread.SpreadUNLock    C1_name   ,  -1,  C1_name   
            ggoSpread.SpreadUNLock    C1_name_B  ,  -1,  C1_name_B 
            ggoSpread.SpreadLock    C1_email  ,  -1,  C1_email  
            ggoSpread.SpreadLock    C1_bp_seq ,  -1,  C1_bp_seq 
            ggoSpread.SpreadUNLock    C1_name2  ,  -1,  C1_name2 
            ggoSpread.SpreadUNLock    C1_name2_B  ,  -1,  C1_name2_B  
            ggoSpread.SpreadLock    C1_email2 ,  -1,  C1_email2 
            ggoSpread.SpreadLock    C1_bp_seq2,  -1,  C1_bp_seq2

			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1	,-1
			.vspdData.ReDraw = True
		End If
		
		If pvSpdNo = "B" Then
			ggoSpread.Source = .vspdData2
			.vspdData2.ReDraw = False
			
			ggoSpread.SpreadUnLock  	C2_item	,  -1,	C2_item
			ggoSpread.SpreadUnLock  	C2_item_std	,  -1,	C2_item_std
			ggoSpread.SpreadUnLock  	C2_item_prc	,  -1,	C2_item_prc
			ggoSpread.SpreadUnLock  	C2_item_qty	,  -1,	C2_item_qty
			ggoSpread.SpreadUnLock  	C2_item_date 	,  -1,	C2_item_date 
			ggoSpread.SpreadUnLock  	C2_item_amt	,  -1,	C2_item_amt
			ggoSpread.SpreadUnLock  	C2_item_tax	,  -1,	C2_item_tax
			ggoSpread.SpreadUnLock  	C2_item_memo 	,  -1,	C2_item_memo 
			
			.vspdData2.ReDraw = True
		End If
			
	End With
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pv1stRow, ByVal pvStartRow, ByVal pvEndRow)

	Dim BlnStatus, pvIssueDtFg, pvProcFlag
	Dim pvCnt
	
	With frm1
		.vspdData.Row = pv1stRow
		.vspdData.Col = C1_issue_dt_fg : pvIssueDtFg = Trim(.vspdData.text)
		.vspdData.Col = C1_proc_flag : pvProcFlag = Trim(.vspdData.Text)
		
				
	End With
    
    With frm1.vspdData2
		'.Redraw = False
		If pvIssueDtFg = "N" or pvIssueDtFg = ""  _ 
				Or(pvIssueDtFg = "Y" And (pvProcFlag = docStatusDelete or pvProcFlag = docStatusDisuse or pvProcFlag = docStatusReject)) then 
			ggoSpread.SpreadUnLock C2_item	,pvStartRow,	C2_item	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_std	,pvStartRow,	C2_item_std	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_prc	,pvStartRow,	C2_item_prc	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_qty	,pvStartRow,	C2_item_qty	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_date 	,pvStartRow,	C2_item_date 	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_amt	,pvStartRow,	C2_item_amt	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_tax	,pvStartRow,	C2_item_tax	,pvEndRow
			ggoSpread.SpreadUnLock C2_item_memo 	,pvStartRow,	C2_item_memo 	,pvEndRow
		Else
			ggoSpread.SpreadLock C2_item	,pvStartRow,	C2_item	,pvEndRow
			ggoSpread.SpreadLock C2_item_std	,pvStartRow,	C2_item_std	,pvEndRow
			ggoSpread.SpreadLock C2_item_prc	,pvStartRow,	C2_item_prc	,pvEndRow
			ggoSpread.SpreadLock C2_item_qty	,pvStartRow,	C2_item_qty	,pvEndRow
			ggoSpread.SpreadLock C2_item_date 	,pvStartRow,	C2_item_date 	,pvEndRow
			ggoSpread.SpreadLock C2_item_amt	,pvStartRow,	C2_item_amt	,pvEndRow
			ggoSpread.SpreadLock C2_item_tax	,pvStartRow,	C2_item_tax	,pvEndRow
			ggoSpread.SpreadLock C2_item_memo 	,pvStartRow,	C2_item_memo 	,pvEndRow
		
		
		End If
		'.Redraw = True
    
    End With

End Sub





'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C1_send_check        	=	iCurColumnPos(1)
			C1_proc_flag_nm      	=	iCurColumnPos(2)
			C1_is_send_nm        	=	iCurColumnPos(3)
			C1_gl_no             	=	iCurColumnPos(4)
			C1_vat_no            	=	iCurColumnPos(5)
			C1_inv_no  				=	iCurColumnPos(6)
			C1_inv_amend_type		=	iCurColumnPos(7)
			C1_inv_amend_type_nm  	=	iCurColumnPos(8)
			C1_remark            	=	iCurColumnPos(9)
			C1_remark2           	=	iCurColumnPos(10)
			C1_remark3				=	iCurColumnPos(11)
			C1_issued_dt         	=	iCurColumnPos(12)
			C1_io_fg             	=	iCurColumnPos(13)
			C1_bp_cd             	=	iCurColumnPos(14)
			C1_bp_nm             	=	iCurColumnPos(15)
			C1_made_vat_fg       	=	iCurColumnPos(16)
			C1_doc_cur           	=	iCurColumnPos(17)
			C1_vat_type          	=	iCurColumnPos(18)
			C1_vat_type_nm       	=	iCurColumnPos(19)
			C1_vat_rate          	=	iCurColumnPos(20)
			C1_net_amt           	=	iCurColumnPos(21)
			C1_net_loc_amt       	=	iCurColumnPos(22)
			C1_vat_amt          	=	iCurColumnPos(23)
			C1_vat_loc_amt      	=	iCurColumnPos(24)
			C1_credit_cd         	=	iCurColumnPos(25)
			C1_report_biz_area_cd	=	iCurColumnPos(26)
			C1_tax_biz_area_nm   	=	iCurColumnPos(27)
			C1_biz_area_cd       	=	iCurColumnPos(28)
			C1_biz_area_nm       	=	iCurColumnPos(29)
			C1_ref_no            	=	iCurColumnPos(30)
			C1_disuse_reason     	=	iCurColumnPos(31)
			C1_legacy_pk         	=	iCurColumnPos(32)
			C1_sale_no              =	iCurColumnPos(33)
			C1_proc_flag         	=	iCurColumnPos(34)
			C1_is_send           	=	iCurColumnPos(35)
			C1_issue_dt_fg       	=	iCurColumnPos(36)
            
            C1_name                 = iCurColumnPos(37)
            C1_name_B               = iCurColumnPos(38)
            C1_email                = iCurColumnPos(39)
            C1_bp_seq               = iCurColumnPos(40)
            C1_name2                = iCurColumnPos(41)
            C1_name2_B              = iCurColumnPos(42)
            C1_email2               = iCurColumnPos(43)
            C1_bp_seq2              = iCurColumnPos(44)
        Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)


			C2_item      = iCurColumnPos(  1)
			C2_item_std  = iCurColumnPos(  2)
			C2_item_prc  = iCurColumnPos(  3)
			C2_item_qty  = iCurColumnPos(  4)
            C2_item_date = iCurColumnPos(  5)
			C2_item_amt  = iCurColumnPos(  6)
			C2_item_tax  = iCurColumnPos(  7)
			C2_item_memo = iCurColumnPos(  8)
			C2_inv_no    = iCurColumnPos(  9)
			C2_vat_no    = iCurColumnPos( 10)
			C2_inv_item_seq_no = iCurColumnPos(11)    
			
		Case "C"
			ggoSpread.Source = frm1.vspdData3
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)


			C2_item      = iCurColumnPos(  1)
			C2_item_std  = iCurColumnPos(  2)
			C2_item_prc  = iCurColumnPos(  3)
			C2_item_qty  = iCurColumnPos(  4)
            C2_item_date = iCurColumnPos(  5)
			C2_item_amt  = iCurColumnPos(  6)
			C2_item_tax  = iCurColumnPos(  7)
			C2_item_memo = iCurColumnPos(  8)
			C2_inv_no    = iCurColumnPos(  9)
			C2_vat_no    = iCurColumnPos( 10)
			C2_inv_item_seq_no = iCurColumnPos(11)    	
			

	End Select    
End Sub


'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.popTaxRecipient.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발행처"					
	arrParam(1) = "b_biz_partner"				

	arrParam(2) = Trim(frm1.popTaxRecipient.Value)
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
		frm1.popTaxRecipient.focus
		Exit Function
	Else
		frm1.popTaxRecipient.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.popTaxRecipient.focus
	End If

	Set gActiveElement = document.activeElement 
End Function

'=========================================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.popTaxRecipient.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "세금신고사업장"
	arrParam(1) = "b_tax_biz_area"

	arrParam(2) = Trim(frm1.popTaxBizArea.Value)
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
		frm1.popTaxRecipient.focus
		Exit Function
	Else
		frm1.popTaxBizArea.Value    = arrRet(0)
		frm1.txtBizAreaNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True
		frm1.popTaxBizArea.focus
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
    Dim IntRetCD

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
    Dim IntRetCD

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

	StrSaveFlag = "FI"

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
					
					.vspdData.Col = C1_vat_no
					
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
	Dim lRow   , lngRows  , pvCnt   
	Dim lGrpCnt     
	Dim strVal, strDtlVal
	Dim strCheckSend, iSelectCnt
	Dim strCheckProc

	Dim net_loc_amt, fi_net_loc_amt, vat_loc_amt, fi_vat_loc_amt, proc_flag
	
	Dim strHndVatNo, strVatNo
	Dim boolExist
	
	Dim strAmendType, strRemark, strCheckMsg
	
	Dim arrayM, userDN, userInfo, userInfoSet
	
	Dim RetFlag
	
	Dim StrSaveFlag , StrMessageNo
	
	'⊙: Processing is NG
	Call LayerShowHide(1)

	StrSaveFlag = "FI"

	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = "" : strDtlVal = ""
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
			
				.vspdData.Col = C1_proc_flag
				proc_flag = Trim(.vspdData.text)
				

				If Not (proc_flag = docStatusMake or proc_flag = "")  Then
					
					.vspdData.Col = C1_vat_no
					
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
				
			
				.vspdData.Col = C1_net_loc_amt
				net_loc_amt = CDbl(.vspdData.text)
				
				.vspdData.Col = C1_vat_loc_amt
				vat_loc_amt = CDbl(.vspdData.text)
				
				fi_vat_loc_amt = CDbl(0)
				fi_net_loc_amt = CDbl(0)
				
				Call SortHSheet()
				
				.vspdData.Row = lRow
				.vspdData.Col = C1_vat_no
				strVatNo = Trim(.vspdData.text)
				
				boolExist = False

				'------------------------------------
				' Find First Row
				'------------------------------------ 
				For lngRows = 1 To .vspdData3.MaxRows
				    .vspdData3.Row = lngRows
				    .vspdData3.Col = C3_vat_no
				    strHndVatNo = .vspdData3.Text
				   
				    If strVatNo = strHndVatNo Then
				        boolExist = True
				        Exit For
				    End If
				Next

				'------------------------------------
				' Show Data
				'------------------------------------ 
				 .vspdData3.Row = lngRows
        
				If boolExist = True Then
					    
					For pvCnt = lngRows To .vspdData3.MaxRows

					     .vspdData3.Row = pvCnt
						        
					    .vspdData3.Col = C3_vat_no
						strHndVatNo = .vspdData3.Text
						        
					    If strVatNo = strHndVatNo Then
							If strVatNo = strHndVatNo Then

								.vspdData3.Col = C3_item_amt
								fi_net_loc_amt = fi_net_loc_amt + CDbl(.vspdData3.Text)
								
								.vspdData3.Col = C3_item_tax
								fi_vat_loc_amt = fi_vat_loc_amt + CDbl(.vspdData3.Text)
								
								.vspdData3.Col = 0
								
								'If .vspdData3.text = ggoSpread.InsertFlag or .vspdData3.text =  ggoSpread.UpdateFlag Then
									
									.vspdData3.Col = C3_vat_no :	strDtlVal = strDtlVal & Trim(.vspdData3.Text) & parent.gColSep
									.vspdData3.Col = C3_item :	strDtlVal = strDtlVal & Trim(.vspdData3.Text) & parent.gColSep
									.vspdData3.Col = C3_item_std :	strDtlVal = strDtlVal & Trim(.vspdData3.Text) & parent.gColSep
									.vspdData3.Col = C3_item_prc :	strDtlVal = strDtlVal & UNIConvNum(.vspdData3.Text,0) & parent.gColSep
									.vspdData3.Col = C3_item_qty :	strDtlVal = strDtlVal & UNIConvNum(.vspdData3.Text,0) & parent.gColSep
									.vspdData3.Col = C3_item_date :	strDtlVal = strDtlVal & UNIConvDate(.vspdData3.Text) & parent.gColSep
									.vspdData3.Col = C3_item_amt :	strDtlVal = strDtlVal & UNIConvNum(.vspdData3.Text,0) & parent.gColSep
									.vspdData3.Col = C3_item_tax :	strDtlVal = strDtlVal & UNIConvNum(.vspdData3.Text,0) & parent.gColSep
									.vspdData3.Col = C3_item_memo :	strDtlVal = strDtlVal & Trim(.vspdData3.Text) & parent.gColSep & parent.gRowSep
									 
								'End If
										
							End If
						Else 
							Exit For	
					    End If   
						
					Next
				
				End If

				If net_loc_amt <> fi_net_loc_amt Or vat_loc_amt <> fi_vat_loc_amt Then
            
					.vspdData.Col = C1_gl_no
					Call DisplayMsgBox("205910","X", .vspdData.text, "X")   '☜ 바뀐부분 
					Call LayerShowHide(0)
					Exit Function
				End If
				
				 strVal = strVal & "FI" & parent.gColSep								'BCDT03_IG1_where_flag
				strVal = strVal & "" & parent.gColSep									'BCDT03_IG1_tax_bill_no
				.vspdData.Col = C1_vat_no : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep					'BCDT03_IG1_vat_no
				strVal = strVal & "" & parent.gColSep									'BCDT03_IG1_iv_no
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
		.txtDtlSpread.value = strDtlVal

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
	StrSaveFlag = "FI"
	
	
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
					
					.vspdData.Col = C1_vat_no
					
					Call DisplayMsgBox("205906","X", .vspdData.text, pvMessageText)   '☜ 바뀐부분 
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
		.txtDtlSpread.value = pvChangeCode

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
Sub popTaxBizArea_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		exit sub
	ElseIf KeyAscii = 13 Then 
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub popTaxBizArea1_KeyPress(KeyAscii)
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
	
	
	Call SetPopupMenuItemInf("0000111111")

	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows <= 0 Then
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
		frm1.vspddata.Col = C1_vat_no
		
		Call ToolbarContol(frm1.vspdData.ActiveRow)
    
		frm1.vspddata2.MaxRows = 0
		
		lgOldRow = Row
		If DbDtlQuery(frm1.vspdData.ActiveRow) = False Then
			Call RestoreToolBar()
			Exit Sub
		End If

		lgOldRow = frm1.vspddata.Row
	Else
		If lgOldRow <> Row Then
 			'------ Developer Coding part (Start)
			frm1.vspddata.Row = Row
			frm1.vspddata.Col = C1_vat_no
			frm1.vspddata2.MaxRows = 0
			
			lgOldRow = Row
			
			Call ToolbarContol(Row)
			
			If DbDtlQuery(Row) = False Then
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



'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)    

	Dim pvVATRate, pvItemQty, pvItemPrc
	Dim pvNetAmt, pvVatAmt
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C1_vat_rate
	pvVATRate = UNICDbl(frm1.vspdData.text) * 0.01

	With frm1.vspdData2

		.Row = Row
		Select Case Col
		
			Case C2_item, C2_item_std, C2_item_date, C2_item_amt, C2_item_tax, C2_item_memo, C2_inv_no
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				CopyToHSheet Row
					
			Case C2_item_prc, C2_item_qty	
				.Col = C2_item_prc : pvItemPrc = UNICDbl(.text)
				.Col = C2_item_qty : pvItemQty = UNICDbl(.text)
				
				If not (pvItemPrc = 0) And Not (pvItemQty = 0) Then
					pvNetAmt = pvItemPrc * pvItemQty
					pvVatAmt = pvNetAmt * pvVATRate
					.Col = C2_item_amt : .Text = pvNetAmt
					.Col = C2_item_tax : .Text = pvVatAmt
				End If
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				CopyToHSheet Row
			
		End Select

	End With
	
	
End Sub


'#########################################################################################################
'												4. Common Function부 
'=========================================================================================================
Function FncQuery() 

    Dim IntRetCD 
    
    On Error Resume Next

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
	    
	    ggoSpread.Source = .vspdData3
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

		If frm1.popTaxRecipient.value = "" Then
           frm1.txtSupplierNm.value = ""
		End If
		
		'-----------------------
		'Erase contents area
		'-----------------------
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData3
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
Function FncInsertRow(ByVal pvRowCnt)

	Dim IntRetCD
	Dim imRow
	Dim pvRow
	Dim pvMaxSeq
	
	On Error Resume Next
	
	FncInsertRow = false
	
	If frm1.vspdData.MaxRows < 1 Then
		Call DisplayMsgBox("900002","X", .vspdData.text, "X")   '☜ 바뀐부분 
		Exit Function
	End If	
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If
	End If	
	
	With frm1
		.vspdData2.ReDraw = False
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
    	If frm1.vspdData2.selBlockRow = -1 Then
			ggoSpread.InsertRow 0, imRow
		Else
			ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
    	End If
    	
    	'최대값 얻기 
    	pvMaxSeq = GetMaxInvSeq()
    	
    	If .vspdData2.MaxRows >= 99 Then
    		Call DisplayMsgBox("205904","X", .vspdData.text, "X")   '☜ 바뀐부분 
			Exit Function
    	End If
    	
    	For pvRow = .vspdData2.ActiveRow To .vspdData2.ActiveRow + imRow -1
    		.vspdData.Row = .vspdData.ActiveRow
    		.vspdData2.Row = pvRow
    		.vspdData.Col = C1_vat_no
    		.vspdData2.Col = C2_vat_no
    		.vspdData2.Text =  Trim(.vspdData.Text)
    		
    		pvMaxSeq = pvMaxSeq + 1
    		.vspdData2.Col = C2_inv_item_seq_no
    		.vspdData2.Text =  pvMaxSeq
    		
    		.vspdData.Col = C1_issued_dt
    		.vspdData2.Col = C2_item_date
    		.vspdData2.Text =  Trim(.vspdData.Text)
    		
    	Next
    	
    	Call SetSpreadColor(.vspdData.ActiveRow, .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow -1, 0)
		.vspdData2.ReDraw = True
    End With
    
    FncInsertRow = true
    
    Set gActiveElement = document.ActiveElement    
    If Err.number = 0 Then FncInsertRow = True    

End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    
    Dim pvMaxSeq
     
	If frm1.vspdData2.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData2.focus
    Set gActiveElement = document.activeElement 
	frm1.vspdData2.EditMode = True
	frm1.vspdData2.ReDraw = False
	
	pvMaxSeq = GetMaxInvSeq() + 1
	
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.CopyRow

    Call frm1.vspdData2.SetText(C2_inv_item_seq_no,frm1.vspdData2.ActiveRow,pvMaxSeq)
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.CopyRow
    Call frm1.vspdData3.SetText(C3_inv_item_seq_no,frm1.vspdData3.ActiveRow,pvMaxSeq)
    Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow)
    frm1.vspdData2.ReDraw = True
   
   
End Function

'========================================================================================
Function FncCancel() 
	
Dim Row
Dim strMode
Dim	strVATNo,	strInvSeq
Dim strHdnVATNo, strHdnInvSeq
Dim lngRows


	If frm1.vspdData2.MaxRows < 1 Then Exit Function	

    ggoSpread.Source = frm1.vspdData2	
    Row = frm1.vspdData2.ActiveRow
    frm1.vspdData2.Row = Row
    frm1.vspdData2.Col = 0
    strMode = frm1.vspdData2.Text
    frm1.vspdData2.Col = C2_vat_no
    strVATNo = frm1.vspdData2.Text
    frm1.vspdData2.Col = C2_inv_item_seq_no
    strInvSeq = frm1.vspdData2.Text

	If strMode = ggoSpread.InsertFlag Then
		Call DeleteHSheet(strVATNo, strInvSeq)
		ggoSpread.Source = frm1.vspdData2
		frm1.vspdData2.Row = Row
		
	    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
	   
	Else
		'------------------------------------
		' Find First Row
		'------------------------------------ 
		For lngRows = 1 To frm1.vspdData3.MaxRows
		    frm1.vspdData3.Row = lngRows
		    frm1.vspdData3.Col = C3_vat_no
		    strHdnVATNo = frm1.vspdData3.Text
		    frm1.vspdData3.Col = C3_inv_item_seq_no
		    strHdnInvSeq = frm1.vspdData3.Text
		    If strVATNo = strHdnVATNo and strInvSeq = strHdnInvSeq Then
		        Exit For
		    End If
		Next
		
		ggoSpread.Source = frm1.vspdData3
	    ggoSpread.EditUndo lngRows
	    
	    Call CopyOneRowFromHSheet(lngRows, Row)
	End If
	
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
	Call InitSpreadSheet("C")      
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
		

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
		                      "&popTaxRecipient=" & Trim(.popTaxRecipient.value) & _
		                      "&popTaxBizArea=" & Trim(.popTaxBizArea.value) & _
		                      "&cboTaxDocumentType=" & Trim(.cboTaxDocumentType.value) & _
		                      "&cboTransmitStatus=" & Trim(.cboTransmitStatus.value) & _
		                      "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text) & _
		                      "&txtIssuedToDt=" & Trim(.txtIssuedToDt.text)
	
	End With
	
	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)																'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : Spread 2 And Spread 3 Data 조회 
'========================================================================================
Function DbDtlQuery(ByVal LngRow) 
	Dim strVal                                                        			'⊙: Processing is NG
	Dim strVatNo

	DbDtlQuery = False 
	
	frm1.vspdData2.MaxRows = 0
	frm1.vspdData.Row = LngRow
	frm1.vspdData.Col = C1_vat_no
	strVatNo = frm1.vspdData.Text
    
	If CopyFromHSheet(strVatNo) = True Then      
		Call SetSpreadColor(LngRow, 1,frm1.vspdData2.MaxRows)  
        Exit Function
    End If

	Call LayerShowHide(1)

	ggoSpread.Source = frm1.vspdData 

	frm1.vspddata.Row = LngRow
	frm1.vspddata.Col = C1_vat_no
	strVatNo          = frm1.vspddata.Text

	strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001 & _
								  "&txtVatNo=" & Trim(strVatNo) 
	
	Call RunMyBizASP(MyBizASP, strVal)

	DbDtlQuery = True                                                     
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
		
		Call ToolbarContol(1)
		
		If .vspdData.MaxRows > 0 Then
			.btnResend.disabled	= false
			.btnPublish.disabled	= false
			.btnApprovalDisuse.disabled	= false
			.btnCancelDisuse.disabled	= false

            ggoSpread.SSSetRequired    C1_name, -1, -1
            
			If DbDtlQuery(1) = False Then
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

Function DbDtlQueryOk()												'☆: 조회 성공후 실행로직 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2

	frm1.vspdData2.ReDraw = False
	
	
	Call SetSpreadColor(frm1.vspdData.ActiveRow, 1, frm1.vspdData2.MaxRows)
	
    'lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	'lgAfterQryFlg = True

	frm1.vspdData2.ReDraw = True

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
    
    ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
	
	Call MainQuery
	
End Function


'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData(Byval Row)

Dim strVATNo, strInvSeq
Dim strHndVATNo, strHndInvSeq
Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows

            .vspdData3.Row = lRows
            .vspdData3.Col = C3_inv_item_seq_no
            strHndInvSeq = .vspdData3.Text
            .vspdData3.Col = C3_vat_no
            strHndVATNo = .vspdData3.Text
            .vspdData2.Row = frm1.vspdData2.Row
            .vspdData2.Col = C2_inv_item_seq_no
            strInvSeq = .vspdData2.Text
            .vspdData2.Col = C2_vat_no
            strVATNo = .vspdData2.Text
           
            If Trim(strHndVATNo) = Trim(strVATNo) And Trim(strHndInvSeq) = Trim(strInvSeq) Then
				FindData = lRows
				Exit Function
            End If    
        Next
        
    End With
    
End Function


'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'====================================================================================================
Function CopyFromHSheet(ByVal strVatNo)

Dim lngRows, LngRow
Dim boolExist
Dim iCols
Dim strHndVatNo
Dim strStatus
Dim iCurColumnPos

    boolExist = False
    
    CopyFromHSheet = boolExist
    
    ggoSpread.Source = frm1.vspdData2
 			
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
    With frm1

        Call SortHSheet()

        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C3_vat_no
            strHndVatNo = .vspdData3.Text
           
            If strVatNo = strHndVatNo Then
                boolExist = True
                Exit For
            End If
        Next

        '------------------------------------
        ' Show Data
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            frm1.vspdData2.Redraw = False
            
            While lngRows <= .vspdData3.MaxRows

	             .vspdData3.Row = lngRows
                
                .vspdData3.Col = C3_vat_no
				strHndVatNo = .vspdData3.Text
                
                If strVatNo <> strHndVatNo Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
					If strVatNo = strHndVatNo Then
						.vspdData2.MaxRows = .vspdData2.MaxRows + 1
						.vspdData2.Row = .vspdData2.MaxRows
						.vspdData2.Col = 0
						.vspdData3.Col = 0
						.vspdData2.Text = .vspdData3.Text
						
						For iCols = 1 To .vspdData3.MaxCols
						    .vspdData2.Col = iCurColumnPos(iCols)
						    .vspdData3.Col = iCols
						    .vspdData2.Text = .vspdData3.Text
						Next
						
						LngRow = .vspdData2.MaxRows
						ggoSpread.Source = frm1.vspdData2
						
						'Call SetSpread2Color(LngRow)
					
					End If
                End If   
                
                
                lngRows = lngRows + 1
                
            Wend
            frm1.vspdData2.Redraw = True

        End If
            
    End With        
    
    CopyFromHSheet = boolExist
    
End Function


'=======================================================================================================
'   Function Name : CopyOneRowFromHSheet
'   Function Desc : 
'====================================================================================================
Function CopyOneRowFromHSheet(ByVal SourceRow, ByVal TargetRow)

Dim iCols
Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
    With frm1
        '------------------------------------
        ' Show Data
        '------------------------------------ 
		.vspdData3.Row = SourceRow
		frm1.vspdData2.Redraw = False
		.vspdData2.Row = TargetRow
		.vspdData2.Col = 0
		.vspdData3.Col = 0
		.vspdData2.Text = .vspdData3.Text
		For iCols = 1 To .vspdData3.MaxCols
		    .vspdData2.Col = iCurColumnPos(iCols)
		    .vspdData3.Col = iCols
		    .vspdData2.Text = .vspdData3.Text
		Next
		
		frm1.vspdData2.Redraw = True

    End With        
   
End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols
Dim LngRow
Dim Stype 'for hidden grid function
Dim Otype 'for hidden grid function
Dim iCurColumnPos
	
	ggoSpread.Source = frm1.vspdData2
 			
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
	With frm1 
                
	    lRow = FindData(Row)
	    
	    If lRow > 0 Then
			LngRow = lRow
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To .vspdData2.MaxCols 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text    
            Next
        Else
			.vspdData3.MaxRows = .vspdData3.MaxRows + 1
			LngRow = .vspdData3.MaxRows
            .vspdData3.Row = .vspdData3.MaxRows
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
       
            For iCols = 1 To .vspdData2.MaxCols 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
                
            Next
        
        End If
	
	End With
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strVATNo, Byval strInvSeq)

Dim boolExist
Dim lngRows
Dim strHdnVATNo, strHndInvSeq
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
        '------------------------------------
        ' Find First Row
        '------------------------------------ 
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C3_vat_no
			strHdnVATNo = .vspdData3.Text
            .vspdData3.Col = C3_inv_item_seq_no
			strHndInvSeq = .vspdData3.Text

            If strVATNo = strHdnVATNo and strInvSeq = strHndInvSeq Then
                boolExist = True
                Exit For
            End If    
        Next
       
        '------------------------------------
        ' Data Delete
        '------------------------------------ 
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
				.vspdData3.Col = C3_vat_no
				strHdnVATNo = .vspdData3.Text
				.vspdData3.Col = C3_inv_item_seq_no
				strHndInvSeq = .vspdData3.Text
                
                If (strVATNo <> strHdnVATNo) or (strInvSeq <> strHndInvSeq) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   

            Wend
            

        End If

    End With

    DeleteHSheet = True
End Function    

'======================================================================================================
' Function Name : SortHSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortHSheet()
    
    With frm1
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 'SS_SORT_BY_ROW
       
        .vspdData3.SortKey(1) = C3_vat_no
        .vspdData3.SortKey(2) = C3_inv_item_seq_no
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 'SS_ACTION_SORT
        .vspdData3.BlockMode = False
    End With        
    
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

Function GetMaxInvSeq()

	Dim pvMaxRow, pvNum
	Dim pvCnt

	GetMaxInvSeq = 0
	pvMaxRow = Cint(0)
	
	With frm1
		For pvCnt = 1 To .vspdData2.MaxRows
			
			.vspdData2.Row = pvCnt
			.vspdData2.Col = C2_inv_item_seq_no
			
			If Trim(.vspdData2.text) = "" Then
				pvNum = Cint(0)
			Else
				pvNum = Trim(.vspdData2.text)
			End If	
			
			If pvNum > pvMaxRow Then
				pvMaxRow = pvNum
			End If
		Next
	End With	

	
	GetMaxInvSeq =  pvMaxRow

End Function

Sub ToolbarContol(IRow)

	Dim BlnStatus, pvIssueDtFg, pvProcFlag
	
	
	
	With frm1.vspdData
		.Row = IRow
		.Col = C1_issue_dt_fg : pvIssueDtFg = Trim(.text)
		.Col = C1_proc_flag : pvProcFlag = Trim(.Text)
		
	End With

	If pvIssueDtFg = "N" or pvIssueDtFg = "" Then
		Call SetToolbar("110001010011111")
	Else
		IF (pvIssueDtFg = "Y" And (pvProcFlag = docStatusDelete or pvProcFlag = docStatusDisuse or pvProcFlag = docStatusReject)) then 
			Call SetToolbar("110001010011111")																'⊙: 버튼 툴바 제어 
		Else
			Call SetToolbar("110000000001111")																'⊙: 버튼 툴바 제어 
		End If	 
	End If
End Sub

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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>회계수정정발행등록</font></td>
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
 												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssuedToDt CLASS=FPDTYYYYMMDD   title=FPDATETIME tag="12" ALT="종료일자" class=required></OBJECT>');</SCRIPT>
 											</TD>
 											<TD CLASS="TD5" NOWRAP>발행처</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="popTaxRecipient" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
												<INPUT TYPE=TEXT AlT="발행처" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
										</TR>
										<TR>
 											<TD CLASS="TD5" NOWRAP>세금신고사업장</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="popTaxBizArea" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
												<INPUT TYPE=TEXT AlT="세금신고사업장" ID="txtBizAreaNm" NAME="txtBizAreaNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
 											<TD CLASS="TD5" NOWRAP></TD>
 											<TD CLASS="TD6" NOWRAP></TD>
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
									<TR HEIGHT="40%">
										<TD  WIDTH="100%" colspan=4><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData ID = "A" width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
									</TR>
									<TR HEIGHT="40%">
										<TD WIDTH="100%" colspan="4"><SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 ID = "B" WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" 2><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
							<TD><BUTTON NAME="btnReSend"         CLASS="CLSSBTN" OnClick="VBScript:Call fnResend()" title="재전송" >재전송</BUTTON>&nbsp; 
							    <BUTTON NAME="btnPublish"        CLASS="CLSSBTN" OnClick="VBScript:Call fnPublish()" title="발행">발행</BUTTON>&nbsp;
							    <BUTTON NAME="btnApprovalDisuse" CLASS="CLSSBTN" OnClick="VBScript:Call fnApprovalDisuse()" title="폐기승인">폐기승인</BUTTON>&nbsp; 
								<BUTTON NAME="btnCancelDisuse"   CLASS="CLSSBTN" OnClick="VBScript:Call fnCancelDisuse()" title="폐기취소">폐기취소</BUTTON>&nbsp;</TD>
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
		<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 ID = "C" WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"  ><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
