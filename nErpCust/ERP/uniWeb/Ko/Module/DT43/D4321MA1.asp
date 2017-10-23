<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : ���ڼ��ݰ�꼭(����Ʈ��(�����))
'*  2. Function Name        : 
'*  3. Program ID           : D4321MA1
'*  4. Program Name         : ���Լ��ݰ�꼭(������)
'*  5. Program Desc         : ���Լ��ݰ�꼭�� ���Ͽ� ������, �������, �������û   
'*  6. Component List       : 
'*  7. Modified date(First) : 2011/05/31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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
<% Call loadInfTB19029A("Q", "M","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "MA") %>
End Sub

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "D4321MB1.asp"  'Main��ȸ
Const BIZ_PGM_ID2 = "D4321MB2.asp"  'Dtl��ȸ
Const BIZ_PGM_ID3 = "D4321MB3.asp"  '��������, �������
Const BIZ_PGM_ID4 = "D4321MB4.asp"  '�������û(batch_id ä���� XXSB_DTI_MAIN ���̺� Update)
'Const BIZ_PGM_ID5 = "D4211MB5.asp"  '��������� XXSB_DTI_STATUS���̺� ���°� Update)

'==========================================  1.2.1 Global ��� ����  ======================================
'=                       4.2 Constant variables 
'========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"

'==========================================================================================================

'add header datatable column
Dim 	C1_send_check           '����
Dim     C1_iv_no                '���Թ�ȣ
Dim     C1_posted_flg           'Posting ����
Dim     C1_build_cd             '����ó
Dim     C1_bp_nm                '�ŷ�ó��
Dim     C1_issued_dt            '������
Dim     C1_dti_status           '��꼭����
Dim     C1_dti_status_nm        '��꼭���¸�
Dim 	C1_iv_cur               '��ȭ
Dim     C1_amend_code           '�����ڵ�
Dim     C1_amend_pop            '�����ڵ��˾�
Dim 	C1_net_doc_amt          '���ް���
Dim 	C1_fi_net_amt           '(ȸ��)���ް��� 
Dim 	C1_tot_vat_doc_amt      '�ΰ����ݾ�
Dim 	C1_fi_vat_amt           '(ȸ��)�ΰ����ݾ�
Dim 	C1_total_amt            '�հ�ݾ�
Dim 	C1_fi_total_amt         '(ȸ��)�հ�ݾ�    
Dim 	C1_net_loc_amt          '���ް���(�ڱ�)
Dim 	C1_fi_net_loc_amt       '(ȸ��)���ް���(�ڱ�)
Dim 	C1_tot_vat_loc_amt      '�ΰ����ݾ�(�ڱ�)
Dim 	C1_fi_vat_loc_amt       '(ȸ��)�ΰ����ݾ�(�ڱ�)
Dim     C1_total_loc_amt        '�հ�ݾ�(�ڱ�)
Dim     C1_fi_total_loc_amt     '(ȸ��)�հ�ݾ�(�ڱ�)
Dim     C1_vat_inc_flag         'VAT_INC_FLAG
Dim     C1_vat_inc_flag_nm      '�ΰ������Կ���
Dim     C1_vat_type             '�ΰ�������
Dim     C1_vat_type_nm          '�ΰ������¸�
Dim     C1_vat_rt               '�ΰ�����
Dim     C1_byr_emp_name         '�ŷ�ó�����
Dim     C1_byr_emp_pop          '�ŷ�ó����� �˾�
Dim     C1_byr_dept_name        '�ŷ�ó�μ���
Dim     C1_byr_tel_num          '�ŷ�ó ��ȭ��ȣ
Dim     C1_byr_email            '�ŷ�ó ����� E-Mail
Dim 	C1_tax_biz_area         '���ݽŰ�����
Dim 	C1_tax_biz_area_nm      '���ݽŰ������
Dim 	C1_pur_grp              '���ű׷�
Dim 	C1_pur_grp_nm           '���ű׷��
Dim	    C1_remark               '���
Dim     C1_issue_dt_flag        '���࿩��
Dim     C1_conversation_id      '���۰�����ȣ
Dim     C1_dti_wdate            '��������
Dim 	C1_where_flag           '������
Dim     C1_vat_loc_amt          '�ΰ����ݾ�(�ڱ�)


'add detail datatable column
Dim	C2_item_cd                  'ǰ���ڵ�
Dim	C2_item_nm                  'ǰ���
Dim	C2_spec                     '�԰�    
Dim	C2_iv_qty                   '����
Dim	C2_iv_unit                  '����
Dim	C2_iv_prc                 '�ܰ�
Dim	C2_total_amt                '�հ�ݾ�
Dim	C2_iv_doc_amt               '���ް���
Dim	C2_vat_doc_amt              '�ΰ����ݾ�
Dim	C2_total_amt_loc            '�հ�ݾ�(�ڱ�)          '    
Dim	C2_iv_loc_amt               '���ް���(�ڱ�)
Dim	C2_vat_loc_amt              '�ΰ����ݾ�(�ڱ�)
Dim	C2_iv_no                    '���Թ�ȣ
Dim	C2_iv_seq_no                '���Լ���



Dim lgStrPrevKeyTempGlNo
Dim lgStrPrevKeyTempGlDt
Dim lgQueryFlag					' �ű���ȸ �� �߰���ȸ ���� Flag
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

   Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field

   With frm1
      Call InitComboBox()
      Call InitSpreadSheet()
      Call InitSpreadSheet2
        
      Call SetDefaultVal
      Call InitVariables
 
      Call SetToolbar("111000000000111")										'��: ��ư ���� ����    	
 
      .txtSupplierCd.focus
      .btnSave.disabled	= true
      .btnSaveCancel.disabled	= true
      .btnPublishSD.disabled	= true
   End With		
End Sub

'========================================================================================================= 
Sub InitComboBox()
   Dim iCodeArr 
   Dim iNameArr
   Dim iDx
	
	'��꼭�� �������
    Call CommonQueryRs(" B.MINOR_CD , B.MINOR_NM "," B_CONFIGURATION A INNER JOIN B_MINOR B ON (A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B. MINOR_CD) ", _
                         " A.MAJOR_CD='DT409' and A.SEQ_NO = 2 and B.MINOR_CD IN ('X', 'A') ORDER BY A.REFERENCE ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
    Call SetCombo2(frm1.cboBillStatus ,lgF0  ,lgF1  ,Chr(11))
    
			
   
End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
		 	
 	 	C1_send_check           =  1  '����
        C1_iv_no                =  2  '���Թ�ȣ
        C1_posted_flg           =  3  'Posting ����
        C1_build_cd             =  4  '����ó
        C1_bp_nm                =  5  '�ŷ�ó��
        C1_issued_dt            =  6  '������
        C1_dti_status           =  7  '��꼭����
        C1_dti_status_nm        =  8  '��꼭���¸�
     	C1_iv_cur               =  9  '��ȭ
        C1_amend_code           =  10  '�����ڵ�
        C1_amend_pop            =  11  '�����ڵ��˾�
     	C1_net_doc_amt          =  12  '���ް���
     	C1_fi_net_amt           =  13  '(ȸ��)���ް��� 
     	C1_tot_vat_doc_amt      =  14  '�ΰ����ݾ�
     	C1_fi_vat_amt           =  15  '(ȸ��)�ΰ����ݾ�
     	C1_total_amt            =  16  '�հ�ݾ�
     	C1_fi_total_amt         =  17  '(ȸ��)�հ�ݾ�    
     	C1_net_loc_amt          =  18  '���ް���(�ڱ�)
     	C1_fi_net_loc_amt       =  19  '(ȸ��)���ް���(�ڱ�)
     	C1_tot_vat_loc_amt      =  20  '�ΰ����ݾ�(�ڱ�)
     	C1_fi_vat_loc_amt       =  21  '(ȸ��)�ΰ����ݾ�(�ڱ�)
        C1_total_loc_amt        =  22  '�հ�ݾ�(�ڱ�)
        C1_fi_total_loc_amt     =  23  '(ȸ��)�հ�ݾ�(�ڱ�)
        C1_vat_inc_flag         =  24  'VAT_INC_FLAG
        C1_vat_inc_flag_nm      =  25  '�ΰ������Կ���
        C1_vat_type             =  26  '�ΰ�������
        C1_vat_type_nm          =  27  '�ΰ������¸�
        C1_vat_rt               =  28  '�ΰ�����
        C1_byr_emp_name         =  29  '�ŷ�ó�����
        C1_byr_emp_pop          =  30  '�ŷ�ó����� �˾�
        C1_byr_dept_name        =  31  '�ŷ�ó�μ���
        C1_byr_tel_num          =  32  '�ŷ�ó ��ȭ��ȣ
        C1_byr_email            =  33  '�ŷ�ó ����� E-Mail
     	C1_tax_biz_area         =  34  '���ݽŰ�����
     	C1_tax_biz_area_nm      =  35  '���ݽŰ������
     	C1_pur_grp              =  36  '���ű׷�
     	C1_pur_grp_nm           =  37  '���ű׷��
    	C1_remark               =  38  '���
        C1_issue_dt_flag        =  39  '���࿩��
        C1_conversation_id      =  40  '���۰�����ȣ
        C1_dti_wdate            =  41  '��������
     	C1_where_flag           =  42  '������
        C1_vat_loc_amt          =  43  '�ΰ����ݾ�(�ڱ�)
	 	    	
End Sub

Sub InitSpreadPosVariables2()
	'add tab1 detail datatable column
			
		C2_item_cd              =  1    'ǰ���ڵ�
    	C2_item_nm              =  2    'ǰ���
    	C2_spec                 =  3    '�԰�    
    	C2_iv_qty               =  4    '����
    	C2_iv_unit              =  5    '����
    	C2_iv_prc               =  6    '�ܰ�
    	C2_total_amt            =  7    '�հ�ݾ�
    	C2_iv_doc_amt           =  8    '���ް���
    	C2_vat_doc_amt          =  9    '�ΰ����ݾ�
    	C2_total_amt_loc        =  10   '�հ�ݾ�(�ڱ�)            
    	C2_iv_loc_amt           =  11   '���ް���(�ڱ�)
    	C2_vat_loc_amt          =  12   '�ΰ����ݾ�(�ڱ�)
    	C2_iv_no                =  13   '���Թ�ȣ
    	C2_iv_seq_no            =  14   '���Լ���

End Sub

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
   
    lgStrPrevKeyTempGlNo = ""							'initializes Previous Key
    lgLngCurRows = 0									   'initializes Deleted Rows Count
    
    lgSortKey =  "1"
    lgSortKey1 =  "1"
    
    lgPageNo_B		= ""                          'initializes Previous Key for spreadsheet #2    
    'lgSortKey_B	= "1"

    lgOldRow = 0
    lgRow = 0
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	'������ ���ڴ� ���� ���Ϻ��� ����
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

	
	'If CommonQueryRs(" PUR_GRP_NM "," B_PUR_GRP "," USAGE_FLG = 'Y' and PUR_GRP = '" & parent.gSalesGrp & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
    '   frm1.txtSalesGrpNm.value =  Trim(Replace(lgF0,Chr(11),""))
    'else
    'End if 
	
	'frm1.txtSalesGrpCd.value = parent.gSalesGrp
	
	
	If CommonQueryRs(" top 1 minor_cd, minor_nm "," B_MINOR (nolock) "," MAJOR_CD = 'DT412'  order by minor_cd " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
       frm1.txtBizAreaCd.value =  Trim(Replace(lgF0,Chr(11),""))
       frm1.txtBizAreaNm.value =  Trim(Replace(lgF1,Chr(11),""))
    else
    End if 
	
	
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
        .MaxCols = C1_vat_loc_amt + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .Col = .MaxCols														'������Ʈ�� ��� Hidden Column
        .ColHidden = True
        .MaxRows = 0			

		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetCheck	C1_send_check,		"����",     				4,  -10, "", True, -1
		ggoSpread.SSSetEdit		C1_iv_no,	        "���Թ�ȣ", 				30, ,,50
		ggoSpread.SSSetEdit  	C1_posted_flg, 		"Posting ����",				10, 2,,1
		ggoSpread.SSSetEdit  	C1_build_cd,		"����ó",          			15, ,,18
		ggoSpread.SSSetEdit  	C1_bp_nm,			"�ŷ�ó��",       			15, ,,50				
		ggoSpread.SSSetDate  	C1_issued_dt,      	"������",     				13, 2, parent.gDateFormat		
		ggoSpread.SSSetEdit  	C1_dti_status,		 "��꼭����",				15, ,,18
		ggoSpread.SSSetEdit  	C1_dti_status_nm,	 "��꼭���¸�",			20, ,,40
		ggoSpread.SSSetEdit  	C1_iv_cur,   		 "��ȭ",					10, 2,,10		
		ggoSpread.SSSetEdit  	C1_amend_code,		 "�����ڵ�",				15, ,,18		
		ggoSpread.SSSetButton	C1_amend_pop	    		
		ggoSpread.SSSetFloat	C1_net_doc_amt,    	"���ް���",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_net_amt,		"(ȸ��)���ް���",     	    18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_tot_vat_doc_amt,	"�ΰ����ݾ�",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_fi_vat_amt,       "(ȸ��)�ΰ����ݾ�",		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec        				
		ggoSpread.SSSetFloat	C1_total_amt,		 "�հ�ݾ�",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetFloat	C1_fi_total_amt,     "(ȸ��)�հ�ݾ�",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec						
		ggoSpread.SSSetFloat	C1_net_loc_amt,      "���ް���(�ڱ�)",     	    18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetFloat	C1_fi_net_loc_amt,	"(ȸ��)���ް���(�ڱ�)",	    18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_tot_vat_loc_amt,	"�ΰ����ݾ�(�ڱ�)", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetFloat	C1_fi_vat_loc_amt,	"(ȸ��)�ΰ����ݾ�(�ڱ�)",	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec				
   		ggoSpread.SSSetFloat	C1_total_loc_amt,	"�հ�ݾ�(�ڱ�)",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec   		
		ggoSpread.SSSetFloat	C1_fi_total_loc_amt,"(ȸ��)�հ�ݾ�(�ڱ�)",	    18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetEdit  	C1_vat_inc_flag,	 "�ΰ������Կ���", 		    15, 2,,2
		ggoSpread.SSSetEdit  	C1_vat_inc_flag_nm,	 "�ΰ������Կ���", 		    15, 2,,15		
		ggoSpread.SSSetEdit		C1_vat_type,		 "�ΰ���Ÿ��",	  		    10, 2,,10
		ggoSpread.SSSetEdit		C1_vat_type_nm,		 "�ΰ������¸�",			20, ,,20		
		ggoSpread.SSSetFloat	C1_vat_rt,		     "�ΰ�����",    	     	18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  	C1_byr_emp_name,	 "�ŷ�ó�����",   			10, ,,50				
		ggoSpread.SSSetButton	C1_byr_emp_pop		
		ggoSpread.SSSetEdit  	C1_byr_dept_name,	 "�ŷ�ó�μ���",   			15, ,,50		
		ggoSpread.SSSetEdit  	C1_byr_tel_num,	     "�ŷ�ó��ȭ��ȣ", 			10, ,,50
		ggoSpread.SSSetEdit		C1_byr_email,		 "�ŷ�ó ����� E-Mail",	20, ,,40		
		ggoSpread.SSSetEdit		C1_tax_biz_area,	"���ݽŰ�����",			10, 2,,10
		ggoSpread.SSSetEdit		C1_tax_biz_area_nm,	"���ݽŰ������",			15, ,,20
		ggoSpread.SSSetEdit		C1_pur_grp,			"���ű׷�",					10, 2,,20
		ggoSpread.SSSetEdit		C1_pur_grp_nm,		"���ű׷��",				15, ,,20        
		ggoSpread.SSSetEdit		C1_remark,			"���",						30, ,,50
		ggoSpread.SSSetEdit  	C1_issue_dt_flag, 	"���࿩��",     			12, 2,,10
		ggoSpread.SSSetEdit		C1_conversation_id,	"���۰�����ȣ", 			30, ,,50
		ggoSpread.SSSetDate  	C1_dti_wdate,       "��������",     			13, 2, parent.gDateFormat		
		ggoSpread.SSSetEdit  	C1_where_flag, 		"������", 					8, ,,3
		ggoSpread.SSSetFloat	C1_vat_loc_amt,	   "�ΰ����ݾ�(�ڱ�)",			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

											        		                				
		'Call ggoSpread.MakePairsColumn(C1_change_reason_cd, C1_change_reason, "1")
      Call ggoSpread.SSSetColHidden(C1_issue_dt_flag, C1_issue_dt_flag, True)
      Call ggoSpread.SSSetColHidden(C1_conversation_id, C1_conversation_id, True)
      Call ggoSpread.SSSetColHidden(C1_dti_wdate, C1_dti_wdate, True)
      Call ggoSpread.SSSetColHidden(C1_dti_status, C1_dti_status, True)
      Call ggoSpread.SSSetColHidden(C1_where_flag, C1_where_flag, True)      
      Call ggoSpread.SSSetColHidden(C1_vat_inc_flag, C1_vat_inc_flag, True)
      Call ggoSpread.SSSetColHidden(C1_vat_type, C1_vat_type, True)      
      Call ggoSpread.SSSetColHidden(C1_vat_loc_amt, C1_vat_loc_amt, True)
            

		.ReDraw = True
	End With
	
	Call SetSpreadLock()
End Sub

Sub InitSpreadSheet2()
	Call initSpreadPosVariables2()
	With frm1.vspdData2	
	
	     ggoSpread.Source = frm1.vspdData2
		 ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	    .ReDraw = false
        .MaxCols = C2_iv_seq_no + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .Col = .MaxCols														'������Ʈ�� ��� Hidden Column
        .ColHidden = True
        .MaxRows = 0			 
	
					
		Call GetSpreadColumnPos2("A")
		
		ggoSpread.SSSetEdit  	C2_item_cd, 			"ǰ��", 			15, ,,18
		ggoSpread.SSSetEdit  	C2_item_nm, 			"ǰ���", 			30, ,,30
		ggoSpread.SSSetEdit  	C2_spec, 				"�԰�", 			15, ,,18

		ggoSpread.SSSetFloat	C2_iv_qty,	    		"����",				15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit  	C2_iv_unit, 			"����", 			15, ,,18
		ggoSpread.SSSetFloat  	C2_iv_prc, 		    "�ܰ�", 			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec	

		ggoSpread.SSSetFloat  	C2_total_amt, 			"�հ�ݾ�", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_iv_doc_amt, 			"���ް���", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_vat_doc_amt, 			"�ΰ����ݾ�",		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetFloat  	C2_total_amt_loc,		"�հ�ݾ�(�ڱ�)",   18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_iv_loc_amt, 	    	"���ް��׾�(�ڱ�)",   18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_vat_loc_amt, 		"�ΰ����ݾ�(�ڱ�)", 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		ggoSpread.SSSetEdit  	C2_iv_no, 			    "���Թ�ȣ", 		15, ,,18
		ggoSpread.SSSetEdit  	C2_iv_seq_no, 			"���Լ���",         15, ,,18
		
		Call ggoSpread.SSSetColHidden(C2_iv_no, C2_iv_no, True)
		Call ggoSpread.SSSetColHidden(C2_iv_seq_no, C2_iv_seq_no, True)

		.ReDraw = True
	End With	
	Call SetSpreadLock_B()
End Sub

'========================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

		ggoSpread.SpreadLock    C1_iv_no, 		    -1, C1_iv_no		
		ggoSpread.SpreadLock    C1_posted_flg, 		-1, C1_posted_flg
		ggoSpread.SpreadLock    C1_build_cd, 	    -1, C1_build_cd
		ggoSpread.SpreadLock	C1_bp_nm,			-1, C1_bp_nm				
		ggoSpread.SpreadLock    C1_issued_dt, 		-1, C1_issued_dt		
		ggoSpread.SpreadLock    C1_dti_status, 		-1, C1_dti_status		
		ggoSpread.SpreadLock    C1_dti_status_nm, 	-1, C1_dti_status_nm				
		ggoSpread.SpreadLock    C1_iv_cur, 		    -1, C1_iv_cur		
		
		ggoSpread.SpreadLock	C1_net_doc_amt,     -1, C1_net_doc_amt
		ggoSpread.SpreadLock	C1_fi_net_amt,		-1, C1_fi_net_amt
		ggoSpread.SpreadLock	C1_tot_vat_doc_amt,	-1, C1_tot_vat_doc_amt
		ggoSpread.SpreadLock	C1_fi_vat_amt,      -1, C1_fi_vat_amt	
		
		ggoSpread.SpreadLock	C1_total_amt,	    -1, C1_total_amt		
			
		ggoSpread.SpreadLock	C1_fi_total_amt,    -1, C1_fi_total_amt
		ggoSpread.SpreadLock	C1_net_loc_amt,     -1, C1_net_loc_amt
		ggoSpread.SpreadLock	C1_fi_net_loc_amt,	-1, C1_fi_net_loc_amt
		ggoSpread.SpreadLock	C1_tot_vat_loc_amt,	-1, C1_tot_vat_loc_amt
		ggoSpread.SpreadLock	C1_fi_vat_loc_amt,	-1, C1_fi_vat_loc_amt
		ggoSpread.SpreadLock	C1_total_loc_amt,	-1, C1_total_loc_amt	
		ggoSpread.SpreadLock	C1_fi_total_loc_amt,-1, C1_fi_total_loc_amt
		ggoSpread.SpreadLock	C1_total_loc_amt,	-1, C1_total_loc_amt		
		ggoSpread.SpreadLock	C1_fi_total_loc_amt,-1, C1_fi_total_loc_amt		
	    ggoSpread.SpreadLock    C1_vat_inc_flag,    -1, C1_vat_inc_flag
	    ggoSpread.SpreadLock    C1_vat_inc_flag_nm, -1, C1_vat_inc_flag_nm	    
	    ggoSpread.SpreadLock    C1_vat_type, 		-1, C1_vat_type		
		ggoSpread.SpreadLock    C1_vat_type_nm, 	-1, C1_vat_type_nm			    			
		ggoSpread.SpreadLock	C1_vat_rt,       	-1, C1_vat_rt
								
		ggoSpread.SSSetRequired	  C1_byr_emp_name,	-1, -1
		ggoSpread.SSSetRequired	  C1_byr_email,		-1, -1
				
		ggoSpread.SpreadLock	C1_tax_biz_area,	-1, C1_tax_biz_area
		ggoSpread.SpreadLock	C1_tax_biz_area_nm,	-1, C1_tax_biz_area_nm
		ggoSpread.SpreadLock	C1_pur_grp,		-1, C1_pur_grp
		ggoSpread.SpreadLock	C1_pur_grp_nm,	-1, C1_pur_grp_nm				
		'ggoSpread.SpreadLock    C1_remark, -1, C1_remark
						
		ggoSpread.SpreadLock    C1_issue_dt_flag,-1, C1_issue_dt_flag				
		ggoSpread.SpreadLock    C1_conversation_id, 		-1, C1_conversation_id		
		ggoSpread.SpreadLock    C1_dti_wdate, 	-1, C1_dti_wdate
		ggoSpread.SpreadLock	C1_where_flag,			-1, C1_where_flag														
		ggoSpread.SpreadLock    C1_vat_loc_amt, -1, C1_vat_loc_amt		
		
								
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
									
     	    C1_send_check           = iCurColumnPos(1)   '����
            C1_iv_no                = iCurColumnPos(2)   '���Թ�ȣ
            C1_posted_flg           = iCurColumnPos(3)   'Posting ����
            C1_build_cd             = iCurColumnPos(4)   '����ó
            C1_bp_nm                = iCurColumnPos(5)   '�ŷ�ó��
            C1_issued_dt            = iCurColumnPos(6)   '������
            C1_dti_status           = iCurColumnPos(7)   '��꼭����
            C1_dti_status_nm        = iCurColumnPos(8)   '��꼭���¸�
     	    C1_iv_cur               = iCurColumnPos(9)   '��ȭ
            C1_amend_code           = iCurColumnPos(10)   '�����ڵ�
            C1_amend_pop            = iCurColumnPos(11)   '�����ڵ��˾�
     	    C1_net_doc_amt          = iCurColumnPos(12)   '���ް���
     	    C1_fi_net_amt           = iCurColumnPos(13)   '(ȸ��)���ް��� 
     	    C1_tot_vat_doc_amt      = iCurColumnPos(14)   '�ΰ����ݾ�
     	    C1_fi_vat_amt           = iCurColumnPos(15)   '(ȸ��)�ΰ����ݾ�
     	    C1_total_amt            = iCurColumnPos(16)   '�հ�ݾ�
     	    C1_fi_total_amt         = iCurColumnPos(17)   '(ȸ��)�հ�ݾ�    
     	    C1_net_loc_amt          = iCurColumnPos(18)   '���ް���(�ڱ�)
     	    C1_fi_net_loc_amt       = iCurColumnPos(19)   '(ȸ��)���ް���(�ڱ�)
     	    C1_tot_vat_loc_amt      = iCurColumnPos(20)   '�ΰ����ݾ�(�ڱ�)
     	    C1_fi_vat_loc_amt       = iCurColumnPos(21)   '(ȸ��)�ΰ����ݾ�(�ڱ�)
            C1_total_loc_amt        = iCurColumnPos(22)   '�հ�ݾ�(�ڱ�)
            C1_fi_total_loc_amt     = iCurColumnPos(23)   '(ȸ��)�հ�ݾ�(�ڱ�)
            C1_vat_inc_flag         = iCurColumnPos(24)   'VAT_INC_FLAG
            C1_vat_inc_flag_nm      = iCurColumnPos(25)   '�ΰ������Կ���
            C1_vat_type             = iCurColumnPos(26)   '�ΰ�������
            C1_vat_type_nm          = iCurColumnPos(27)   '�ΰ������¸�
            C1_vat_rt               = iCurColumnPos(28)   '�ΰ�����
            C1_byr_emp_name         = iCurColumnPos(29)   '�ŷ�ó�����
            C1_byr_emp_pop          = iCurColumnPos(30)   '�ŷ�ó����� �˾�
            C1_byr_dept_name        = iCurColumnPos(31)   '�ŷ�ó�μ���
            C1_byr_tel_num          = iCurColumnPos(32)   '�ŷ�ó ��ȭ��ȣ
            C1_byr_email            = iCurColumnPos(33)   '�ŷ�ó ����� E-Mail
     	    C1_tax_biz_area         = iCurColumnPos(34)   '���ݽŰ�����
     	    C1_tax_biz_area_nm      = iCurColumnPos(35)   '���ݽŰ������
     	    C1_pur_grp              = iCurColumnPos(36)   '���ű׷�
     	    C1_pur_grp_nm           = iCurColumnPos(37)   '���ű׷��
    	    C1_remark               = iCurColumnPos(38)   '���
            C1_issue_dt_flag        = iCurColumnPos(39)   '���࿩��
            C1_conversation_id      = iCurColumnPos(40)   '���۰�����ȣ
            C1_dti_wdate            = iCurColumnPos(41)   '��������
     	    C1_where_flag           = iCurColumnPos(42)   '������
            C1_vat_loc_amt          = iCurColumnPos(43)   '�ΰ����ݾ�(�ڱ�)
	End Select    
End Sub

'========================================================================================
Sub GetSpreadColumnPos2(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
									
			C2_item_cd              = iCurColumnPos(1)     'ǰ���ڵ�
    	    C2_item_nm              = iCurColumnPos(2)     'ǰ���
    	    C2_spec                 = iCurColumnPos(3)     '�԰�    
    	    C2_iv_qty               = iCurColumnPos(4)     '����
    	    C2_iv_unit              = iCurColumnPos(5)     '����
    	    C2_iv_prc               = iCurColumnPos(6)     '�ܰ�
    	    C2_total_amt            = iCurColumnPos(7)     '�հ�ݾ�
    	    C2_iv_doc_amt           = iCurColumnPos(8)     '���ް���
    	    C2_vat_doc_amt          = iCurColumnPos(9)     '�ΰ����ݾ�
    	    C2_total_amt_loc        = iCurColumnPos(10)    '�հ�ݾ�(�ڱ�)            
    	    C2_iv_loc_amt           = iCurColumnPos(11)    '���ް���(�ڱ�)
    	    C2_vat_loc_amt          = iCurColumnPos(12)    '�ΰ����ݾ�(�ڱ�)
    	    C2_iv_no                = iCurColumnPos(13)    '���Թ�ȣ
    	    C2_iv_seq_no            = iCurColumnPos(14)    '���Լ���
						
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

	arrParam(0) = "����ó"					
	arrParam(1) = "b_biz_partner"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE IN ('S','CS')"	
	arrParam(5) = "����ó"						

	arrField(0) = "bp_cd"					
	arrField(1) = "bp_nm"	
	arrField(2) = "bp_rgst_no"				

	arrHeader(0) = "����ó"				
	arrHeader(1) = "����ó��"	
	arrHeader(2) = "����ڵ�Ϲ�ȣ"		

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

	If IsOpenPop = True  Or UCase(frm1.txtSalesGrpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"
	arrParam(1) = "B_PUR_GRP"

	arrParam(2) = Trim(frm1.txtSalesGrpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "usage_flg = 'Y'"
	arrParam(5) = "���ű׷�"

	arrField(0) = "PUR_GRP"
	arrField(1) = "PUR_GRP_NM"

	arrHeader(0) = "���ű׷�"				
	arrHeader(1) = "���ű׷��"	

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

	If IsOpenPop = True  Or UCase(frm1.txtBizAreaCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ݽŰ�����"
	arrParam(1) = "b_tax_biz_area"

	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = ""
	arrParam(5) = "���ݽŰ�����"

	arrField(0) = "tax_biz_area_cd"
	arrField(1) = "tax_biz_area_nm"

	arrHeader(0) = "���ݽŰ�����"				
	arrHeader(1) = "���ݽŰ������"	

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
                                                          
         arrParam(0) = "�����ڵ��˾�"
         arrParam(1) = "B_MINOR (nolock) " ' TABLE ��Ī 
         arrParam(2) = strcode      ' Code Condition
         arrParam(3) = ""       ' Name Cindition
         arrParam(4) = " MAJOR_CD = 'DT408' "       ' Where Condition
         arrParam(5) = "�����ڵ�"    ' �����ʵ��� �� ��Ī 
         

         arrField(0) = "MINOR_CD"     ' Field��(0)
         arrField(1) = "MINOR_NM"     ' Field��(1)

         arrHeader(0) = "�ڵ�"    ' Header��(0)
         arrHeader(1) = "�ڵ��"     ' Header��(1)
	    		
		Case 2
        
        
         frm1.vspddata.Col = C1_build_cd
         strBpCd = Trim(frm1.vspddata.value) 
         
         arrParam(0) = "�ŷ�ó�����"
         arrParam(1) = "XXSB_DTI_BP_USER (nolock) " ' TABLE ��Ī 
         arrParam(2) = strcode      ' Code Condition
         arrParam(3) = ""       ' Name Cindition
         arrParam(4) = " FND_BP_CD = '" & strBpCd & "'"    ' Where Condition
         arrParam(5) = "�ŷ�ó�����"    ' �����ʵ��� �� ��Ī 
         

         arrField(0) = "FND_USER_NAME"          ' Field��(0)
         arrField(1) = "FND_BP_CD"              ' Field��(1)
         arrField(2) = "FND_USER_DEPT_NAME"     ' Field��(2)
         arrField(3) = "FND_USER_TEL_NUM"       ' Field��(3)
         arrField(4) = "FND_USER_EMAIL"         ' Field��(4)
         

         arrHeader(0) = "�ŷ�ó����ڸ�"        ' Header��(0)
         arrHeader(1) = "�ŷ�ó"                ' Header��(1)
         arrHeader(2) = "�ŷ�ó�μ���"          ' Header��(2)
         arrHeader(3) = "�ŷ�ó��ȭ��ȣ"        ' Header��(3)
         arrHeader(4) = "�ŷ�ó�����E-Mail"    ' Header��(4)
         
            
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
            .vspdData.Col = C1_byr_emp_name
           .vspdData.Text = arrRet(0)
           
            .vspdData.Col = C1_byr_dept_name
           .vspdData.Text = arrRet(2)
           
           .vspdData.Col = C1_byr_tel_num
           .vspdData.Text = arrRet(3)
           
           .vspdData.Col = C1_byr_email
           .vspdData.Text = arrRet(4)
           
'           Call vspdData_Change(C_PuNo, .vspdData.Row)	     	     
      End Select
   End With
End Function

'��������
Function fnSave()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lGrpCnt, lRow
	Dim strWhere
	Dim IntRetCD
	Dim strVal
	Dim Status
	Dim report_biz_area
	Dim strBpCd
	Dim strAmend
	Dim totamount
	Dim messageNo
	Dim messageAmend
	Dim strRemark
	Dim strConverid
	Dim strByrEmpName
	Dim strByrEmail
	Dim RetFlag
	Dim iSelectCnt
	Dim saveFlag
	Dim net_loc_amt,  fi_net_loc_amt
	Dim fi_vat_loc_amt, tot_vat_loc_amt
	Dim strDtistatus


    fnSave = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
    saveFlag = "MM"
      	      	
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
				
				  .vspdData.Col = C1_dti_status
				  strDtistatus = Ucase(Trim(.vspdData.text))
				  
				  .vspdData.Col = C1_iv_no
				  messageNo = Trim(.vspdData.text)
				  
				  
				  if strDtistatus <> "X" then
				        Call DisplayMsgBox("W70001","X", "�������� ����� �ƴѰ��� �ֽ��ϴ�.\n- ���Թ�ȣ : " & messageNo, "X")	            		
				        'W70001: %1	
				         Call LayerShowHide(0)
				        Exit Function
				  
				  end if
																																 
				 '���ݽŰ������ڵ�
				 '.vspdData.Col = C1_report_biz_area
				 'report_biz_area = Trim(.vspdData.text)
				 
				 report_biz_area = Trim(frm1.txtBizAreaCd.value)
				 
				 				 				 				 
				if report_biz_area <> "" then 
					 '���ݽŰ����� �ڵ�
					 
					If CommonQueryRs("  ind_type, ind_class ", "  B_BIZ_partner (nolock) ", " bp_cd = '" & report_biz_area & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then						 
					 					 					 					 					 					 					    
					else
					    Call DisplayMsgBox("DT4120","X", report_biz_area,"X")	            		
					    'DT4120: ����� �̷µ��(B1263MA1)���� %1��  ������ Ȯ���ϼ���.
					    Call LayerShowHide(0)
					    Exit Function
					END IF
				end if
				
				
				'�ŷ�ó�ڵ�
				 .vspdData.Col = C1_build_cd
				 strBpCd = Trim(.vspdData.text)
				
				if	strBpCd <> "" then
				
					If CommonQueryRs("  ind_type, ind_class ", "  B_BIZ_partner (nolock) ", " bp_cd = '" & strBpCd & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then						 
					 					 					 					 					 					 					    
					else
					    Call DisplayMsgBox("DT4120","X", report_biz_area,"X")	            		
					    'DT4120: ����� �̷µ��(B1263MA1)���� %1��  ������ Ȯ���ϼ���.
					    Call LayerShowHide(0)
					    Exit Function
					END IF
															
				end if	
				
				
				.vspdData.Col = C1_amend_code
				strAmend = Trim(.vspdData.text)
				
				if	strAmend <> "" then
				   if (strAmend = "03" or  strAmend = "04") then
				        
				        .vspdData.Col = total_amt				        				        
				        totamount = CDbl(.vspdData.text)
				        
				        if totamount > 0 then
				        
				            .vspdData.Col = C1_iv_no
				            messageNo = Trim(.vspdData.text)
				            
				            .vspdData.Col = C1_remark
				            strRemark =  Trim(.vspdData.text)
				        
				            messageAmend =  strAmend & ":" &  strRemark
				        
				            Call DisplayMsgBox("DT4116","X", messageNo, messageAmend)	            		
					        'DT4116: %1�� %2�� ��쿡�� (-)���ݰ�꼭�� ��ȿ�մϴ�.	
					         Call LayerShowHide(0)
					        Exit Function			        
				       end if	
				   end if    			        
				 end if	
				
				'----------------------------------
				.vspdData.Col = C1_conversation_id
				strConverid = Trim(.vspdData.text) 
				
				if strConverid <> "" then  
				    .vspdData.Col = C1_iv_no
				    messageNo = Trim(.vspdData.text)
				
				     Call DisplayMsgBox("DT4101","X", messageNo, "X")	            		
				        'DT4101: %1�� ������ �� �ִ� ����� �ƴմϴ�.	
				         Call LayerShowHide(0)
				        Exit Function
                else
                
                    .vspdData.Col = C1_net_loc_amt
				    net_loc_amt = CDbl(.vspdData.text)
    				
				    .vspdData.Col = C1_fi_net_loc_amt
				    fi_net_loc_amt = CDbl(.vspdData.text)
    				
				    '.vspdData.Col = C1_vat_loc_amt
				    'vat_loc_amt = CDbl(.vspdData.text)
    			
				    .vspdData.Col = C1_fi_vat_loc_amt
				    fi_vat_loc_amt = CDbl(.vspdData.text)
				    
				    .vspdData.Col = C1_tot_vat_loc_amt
				    tot_vat_loc_amt = CDbl(.vspdData.text)

                    If (net_loc_amt <> fi_net_loc_amt) Or (tot_vat_loc_amt <> fi_vat_loc_amt) Then
              
				        .vspdData.Col = C1_iv_no
            	        RetFlag = DisplayMsgBox("205911", parent.VB_YES_NO, .vspdData.text, "X")   '�� �ٲ�κ� 
            	        '205911: [%1]�� ȸ���� �ݾװ� �ٸ��ϴ�. ����Ͻðڽ��ϱ�?

					        If RetFlag = VBNO Then
						        Call LayerShowHide(0)
						        Exit Function
					        End If
			        End If
                
                
                    .vspdData.Col = C1_byr_emp_name
				    strByrEmpName = Trim(.vspdData.text) 
    				
				    if strByrEmpName = "" then  

				        strByrEmpName = Trim(.vspdData.text)
    				
				         Call DisplayMsgBox("DT4102","X", strByrEmpName, "X")	            		
				            'DT4102: %1 ���޹޴� ���� �̸��� �����ϴ�.	
				             Call LayerShowHide(0)
				            Exit Function	
				    else
    				                				            
				        .vspdData.Col = C1_byr_email
			            strByrEmail =  Trim(.vspdData.text)
        				
			            if strByrEmail = "" then
			                
			                .vspdData.Col = C1_tax_bill_no
			                 messageNo = Trim(.vspdData.text)
			            
				            Call DisplayMsgBox("DT4103","X",messageNo,"X")	            		
				            'DT4103: %1 ���޹޴� ���� e-mail ������ �����ϴ�.
				            Call LayerShowHide(0)
				            Exit Function
			            end if    				                				                				                				        		        
				    end if                				        			        
				end if
				'----------------------------------------																																				
								
				 strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_where_flag      :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'������                           '2			     
  			     .vspdData.Col = C1_iv_no           :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'���Թ�ȣ                         '3			     
  			     .vspdData.Col = C1_byr_emp_name    :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'�ŷ�ó�����                     '4			       			     
  			     .vspdData.Col = C1_byr_dept_name   :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'�ŷ�ó�μ���                     '5			       			     
  			     .vspdData.Col = C1_byr_tel_num     :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'�ŷ�ó ��ȭ��ȣ                  '6			       			     
  			     .vspdData.Col = C1_byr_email       :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'�ŷ�ó ����� E-Mail             '7
  			     .vspdData.Col = C1_amend_code      :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'�����ڵ�                         '8
  													:	strVal = strVal & saveFlag & parent.gRowSep				    '����           			      '9	
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1
       End If
     Next

		if iSelectCnt = 0 then
			Call DisplayMsgBox("DT4205","X", "X","X")	
			'DT4205: ���õ� ���� �����ϴ�.            		
			Call LayerShowHide(0)
			Exit Function
		end if

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID3)
	
	fnSave = True	

End Function

'�������
'========================================================================================
Function fnSaveCancel()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lGrpCnt, lRow
	Dim strWhere
	Dim IntRetCD
	Dim strVal
	Dim strDtiStatus
	Dim messageNo
	Dim iSelectCnt
	Dim SaveCancelFlag

    fnSaveCancel = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
    SaveCancelFlag = "SD"
      	      	
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
																								 
				 								
				'������üũ
				 .vspdData.Col = C1_dti_status
				 strDtiStatus = Trim(.vspdData.text)
				
				if	(strDtiStatus <> "A" and strDtiStatus <> "W" and strDtiStatus <> "T" and strDtiStatus <> "R" and strDtiStatus <> "O" ) then
				
				    .vspdData.Col = C1_iv_no
				    messageNo = Trim(.vspdData.text)
											
				    Call DisplayMsgBox("DT4115","X", messageNo,"X")	            		
				    'DT4115:  %1�� ���� ��� ����� �ƴմϴ�.
				    Call LayerShowHide(0)
				    Exit Function
																			
				end if	
				
																				
				'----------------------------------------																																				
								
				 strVal = strVal & "U" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_where_flag      :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'������                           '2			     
  			     .vspdData.Col = C1_conversation_id :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID                  '3			     
  													:	strVal = strVal & SaveCancelFlag & parent.gRowSep			'����           			      '4	
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1
          End If
       Next

		if iSelectCnt = 0 then
			Call DisplayMsgBox("DT4205","X", "X","X")	
			'DT4205: ���õ� ���� �����ϴ�.            		
			Call LayerShowHide(0)
			Exit Function
		end if

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID3)
	
	fnSaveCancel = True	
End Function


'���Թ���
Function fnPublish()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lGrpCnt, lRow
	Dim strWhere
	Dim IntRetCD
	Dim strVal
	Dim Status
	Dim strDtiStatus
	Dim strBpCd
	Dim strAmend
	Dim totamount
	Dim messageNo
	Dim messageAmend
	Dim strRemark
	Dim strConverid
	Dim strByrEmpName
	Dim strByrEmail
	Dim RetFlag
	Dim iSelectCnt
	Dim issueFlag
	Dim net_loc_amt,  fi_net_loc_amt
	Dim tot_vat_loc_amt,  fi_vat_loc_amt

    fnPublish = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
    issueFlag = "SD"
      	      	
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
																								 
				 '������üũ
				 .vspdData.Col = C1_dti_status
				 strDtiStatus = Trim(.vspdData.text)
				 				 				 				 
				if strDtiStatus <> "A" then 
					 
					 .vspdData.Col = C1_iv_no
					 messageNo = Trim(.vspdData.text)
					 
				    Call DisplayMsgBox("DT4104","X", messageNo,"X")	            		
				    'DT4104: %1�� ���� �� �� �ִ� ����� �ƴմϴ�.
				    Call LayerShowHide(0)
				    Exit Function
				else
				
				    .vspdData.Col = C1_net_loc_amt
				    net_loc_amt = CDbl(.vspdData.text)
    				
				    .vspdData.Col = C1_fi_net_loc_amt
				    fi_net_loc_amt = CDbl(.vspdData.text)
    				
				    '.vspdData.Col = C1_vat_loc_amt
				    'vat_loc_amt = CDbl(.vspdData.text)
    			
				    .vspdData.Col = C1_tot_vat_loc_amt
				    tot_vat_loc_amt = CDbl(.vspdData.text)
				    
				    .vspdData.Col = C1_fi_vat_loc_amt
				    fi_vat_loc_amt = CDbl(.vspdData.text)
				    
				    
                    If (net_loc_amt <> fi_net_loc_amt) Or (tot_vat_loc_amt <> fi_vat_loc_amt) Then
              
				        .vspdData.Col = C1_iv_no
            	        RetFlag = DisplayMsgBox("205911", parent.VB_YES_NO, .vspdData.text, "X")   '�� �ٲ�κ� 
            	        '205911: [%1]�� ȸ���� �ݾװ� �ٸ��ϴ�. ����Ͻðڽ��ϱ�?

					        If RetFlag = VBNO Then
						        Call LayerShowHide(0)
						        Exit Function
					        End If
			        End If
				    
				end if
																												
								
				'----------------------------------------																																				
								
				 strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_where_flag       :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'������                           '2			     
  			     .vspdData.Col = C1_conversation_id  :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID                  '3			       			     
  													 :	strVal = strVal & issueFlag & parent.gRowSep				    '����            		      '4	
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1
       End If
     Next

		if iSelectCnt = 0 then
			Call DisplayMsgBox("DT4205","X", "X","X")	
			'DT4205: ���õ� ���� �����ϴ�.            		
			Call LayerShowHide(0)
			Exit Function
		end if

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID4)
	
	fnPublish = True	

End Function


Function WebControl(batchid, status)
	DIm IntRetCD
	DIm strURL
    Dim StrRegNo
    Dim StrComName
    Dim strExpDate
    Dim NowDate
    Dim strSmartid,  strSmartPW
    Dim strUrlinfo
    Dim strIssueASP
    Dim strYear, strMonth, strDay
    Dim arrRet
    
    
	DIm lRow
	
    'If lgIntFlgMode <> parent.OPMD_UMODE Then												'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                                       
    '    Exit Function
    'End If
    
     
     If CommonQueryRs(" TOP 1 A.SMART_ID, A.SMART_PASSWORD "," XXSB_DTI_SM_USER A (nolock) ", _
              " A.FND_USER = '" & parent.gUsrID & "'  AND A.FND_REGNO = (SELECT TOP 1 REPLACE(OWN_RGST_NO,'-','') FROM B_TAX_BIZ_AREA WHERE TAX_BIZ_AREA_CD = '" & Trim(frm1.txtBizAreaCd.value) & "')" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then        
        strSmartid = Trim(Replace(lgF0,Chr(11),""))
        strSmartPW = Trim(Replace(lgF1,Chr(11),""))
     else
        Call DisplayMsgBox("DT4118","X", "X","X")	               		
        'DT4118:���ڼ��ݰ�꼭�� ����ڸ� Ȯ���ϼ���
        Exit Function 	  
     End if
    
      If CommonQueryRs(" TOP 1 Convert(varchar(10),EXPIRATION_DATE,120) as EXPIRATION_DATE "," XXSB_DTI_CERT (nolock) "," CERT_REGNO IN ( SELECT REPLACE(B.BP_RGST_NO,'-','') FROM B_TAX_BIZ_AREA A INNER JOIN B_BIZ_PARTNER B ON (A.TAX_BIZ_AREA_CD = B.BP_CD) WHERE A.TAX_BIZ_AREA_CD = '" & Trim(frm1.txtBizAreaCd.value) & "')" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
        strExpDate =  Trim(Replace(lgF0,Chr(11),""))
        
        if  strExpDate = "" then
            Call DisplayMsgBox("DT4102","X", "X","X")	               		
            'DT4102:%1 ���޹޴� ���� �̸��� �����ϴ�.
            Exit Function 	  
        end if      
      else
         Call DisplayMsgBox("DT4102","X", "X","X")	               		
         'DT4102:%1 ���޹޴� ���� �̸��� �����ϴ�.
         Exit Function 	  
      End if
      
      Call	ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

      NowDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)

      Call ExtractDateFrom(strExpDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

      strExpDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
      

      NowDate = replace(NowDate, "-", "")
      strExpDate = replace(strExpDate, "-", "")

  
      if (strExpDate < NowDate) then
         Call DisplayMsgBox("DT4202","X", "X","X")	               		
         'DT4202: �������� ����Ǿ����ϴ�.
         Exit Function 	        
      end if
      
       If CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = 'DT400' AND MINOR_CD = '01' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
            strUrlinfo =  Trim(Replace(lgF0,Chr(11),""))
            
            if strUrlinfo = "" then
                 Call DisplayMsgBox("DT4108","X", "X","X")	               		
           'DT4108: URL ������ �����ϴ�.
           Exit Function 	  
            end if            
       else
           Call DisplayMsgBox("DT4108","X", "X","X")	               		
           'DT4108: URL ������ �����ϴ�.
           Exit Function 	  
       End if
      
      
      if status <> "C" then
            if  frm1.chkShowBiz.checked = true then
                strIssueASP = "XXSB_DTI_ISSUE_T.asp"
            else
                strIssueASP = "XXSB_DTI_ISSUE.asp"
            end if    
      
            if (frm1.vspdData2.maxRows <= 4 ) then
                strURL =  strUrlinfo & strIssueASP & "?batch_id=" + batchid + "&ID=" + strSmartid + "&PASS=" + strSmartPW + ""
            else
                strURL =  strUrlinfo & strIssueASP & "?batch_id=" + batchid + "&ID=" + strSmartid + "&PASS=" + strSmartPW + ""
            end if
      else        
         strURL =  strUrlinfo & "XXSB_DTI_PRINT.asp?batch_id=" + batchid +  "&SORTFIELD=A&SORTORDER=1 "
          
      end if


       arrRet =  window.showModalDialog(strUrl ,, "dialogWidth=810px; dialogHeight=480px; center: Yes; help: No; resizable: No; status: no; scroll:Yes;")          
      
      'frm1.target = "legacy"	
      'frm1.action =  strURL
      'frm1.submit()
      
       DbSaveOk
                                                 	
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

Sub vspdData_Change(ByVal Col , ByVal Row )

   Dim strAmend
   Dim strEmpName
   Dim strBpCd


    Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        
        Case  C1_amend_code
            strAmend = Trim(Frm1.vspdData.value)
    
            If strAmend = "" Then
			        Frm1.vspdData.Col = C1_remark
			        Frm1.vspdData.text = ""
            Else					
				If CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR (nolock)  "," MAJOR_CD = 'DT408' and minor_cd = '" & strAmend & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					Frm1.vspdData.Col = C1_amend_code
				    Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
				    Frm1.vspdData.Col = C1_remark
				    Frm1.vspdData.text = Trim(Replace(lgF1,Chr(11),""))
				else
				    Call DisplayMsgBox("971001","X",strAmend,"X")	               		
				    '971001: %1 ��(��) �������� �ʽ��ϴ�.
				    Frm1.vspdData.Col = C1_amend_code
				    Frm1.vspdData.text = ""
				    Frm1.vspdData.Col = C1_remark
				    Frm1.vspdData.text = ""
				END IF					
            End if    
            
        Case  C1_byr_emp_name
            strEmpName = Trim(Frm1.vspdData.value)
            
            Frm1.vspdData.Col = C1_build_cd
            strBpCd = Trim(Frm1.vspdData.value)
            
    
            If strEmpName = "" Then
  	            Frm1.vspdData.Col = C1_byr_dept_name
                Frm1.vspdData.value = ""
                Frm1.vspdData.Col = C1_byr_tel_num
                Frm1.vspdData.value = ""
                Frm1.vspdData.Col = C1_byr_email
                Frm1.vspdData.value = ""
            Else					
				If CommonQueryRs(" FND_USER_NAME,  FND_BP_CD, FND_USER_DEPT_NAME, FND_USER_TEL_NUM, FND_USER_EMAIL "," XXSB_DTI_BP_USER (nolock)"," FND_BP_CD = '" & strBpCd & "' and FND_USER_NAME = '" & strEmpName & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					Frm1.vspdData.Col = C1_byr_emp_name
				    Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
				    Frm1.vspdData.Col = C1_byr_dept_name
                    Frm1.vspdData.value = Trim(Replace(lgF2,Chr(11),""))
                    Frm1.vspdData.Col = C1_byr_tel_num
                    Frm1.vspdData.value = Trim(Replace(lgF3,Chr(11),""))
                    Frm1.vspdData.Col = C1_byr_email
                    Frm1.vspdData.value = Trim(Replace(lgF4,Chr(11),""))
				else
				    Call DisplayMsgBox("970000","X",strEmpName,"X")	               		
				    '970000:%1 ��(��) �������� �ʽ��ϴ�.
				   Frm1.vspdData.Col = C1_byr_dept_name
                    Frm1.vspdData.value = ""
                    Frm1.vspdData.Col = C1_byr_tel_num
                    Frm1.vspdData.value = ""
                    Frm1.vspdData.Col = C1_byr_email
                    Frm1.vspdData.value = ""
				END IF					
            End if    
                                                                                        
    End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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
		frm1.vspddata.Col = C1_iv_no
    
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
            frm1.vspddata.Col = C1_iv_no
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

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
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
	' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
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
     lgBlnFlgChgValue = True													'��: Indicates that value changed
End Sub


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
		    Select Case Col
			    Case C1_amend_pop
				   .Row = Row
		           .Col = C1_amend_code
		            Call OpenPopup(.Text, 1)
			    Case C1_byr_emp_pop
			        frm1.vspddata.Col = C1_byr_emp_name			    
				    Call OpenPopup(.text, 2)        
		    End Select    
	    End If								
    End With                     
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

    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then '��: ������ üũ'
        If lgPageNo_B <> "" Then                                                    '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ����
           'Call DbQuery("2",frm1.vspddata.ActiveRow)
        End If
   End if
End Sub

'#########################################################################################################
'												4. Common Function�� 
'=========================================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False																		'��: Processing is NG

    Err.Clear																				'��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    With frm1
	    ggoSpread.Source = .vspdData
	    If  ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
	    	If IntRetCD = vbNo Then
		      	Exit Function
	    	End If
	    End If

		'-----------------------
	    'Check condition area
	    '-----------------------
		   If Not chkField(Document, "1") Then									         '��: This function check required field
               Exit Function
            End If

		
		'If Not chkFieldByCell(.txtIssuedFromDt, "A", "1") Then Exit Function
		'If Not chkFieldByCell(.txtIssuedToDt, "A", "1") Then Exit Function

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
		'	    Call ggoOper.ClearField(Document, "2")												'��: Clear Contents  Field
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

		Call InitVariables 																	'��: Initializes local global variables

		FncQuery = True	
	End With
	
	Call DBquery()
End Function

'========================================================================================
Function FncNew() 
	Dim IntRetCD 

	FncNew = False																	'��: Processing is NG

	Err.Clear																			'��: Protect system from crashing
	'On Error Resume Next															'��: Protect system from crashing

	'-----------------------
	'Check previous data area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") '�� �ٲ�κ�    

		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")												'��: Clear Condition Field
	Call ggoOper.ClearField(Document, "2")												'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call ggoOper.LockField(Document, "N")									'Lock  Suitable  Field    
    Call InitVariables()															'Initializes local global va    Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "R")
    
	Call SetDefaultVal
    
    
	FncNew = True																		'��: Processing is OK
End Function


'========================================================================================
Function FncSave() 
End Function

'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo																	'��: Protect system from crashing    
End Function

'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()																'��: Protect system from crashing
End Function

'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)											'��:ȭ�� ����, Tab ���� 
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")								'����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
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
			'Call InitComboBoxGrid 
			Call ggoSpread.ReOrderingSpreadData()
			'Call InitData(1)
			'Call SetSpreadColor2       		         
		Case "vspdData2"
			Call InitSpreadSheet2		
'			Call InitComboBox     		
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.ReOrderingSpreadData()	
	End Select     
End Sub

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtStatusflag
	Dim txtKeyNo

	DbQuery = False
	
	'frm1.btnRetransfer.disabled = false
    
    Call SetButtonDisable

	With frm1
		

		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
		                      "&txtSupplierCd=" & Trim(.txtSupplierCd.value) & _
		                      "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value) & _
		                      "&txtBizAreaCd=" & Trim(.txtBizAreaCd.value) & _
		                      "&cboBillStatus=" & Trim(.cboBillStatus.value) & _
		                      "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text) & _
		                      "&txtIssuedToDt=" & Trim(.txtIssuedToDt.text)
	End With

	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)																'��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 And Spread 3 Data ��ȸ 
'========================================================================================
Function DbQuery2() 
	DbQuery2 = False 

	Dim strVal                                                        			'��: Processing is NG
	Dim iTaxBillNo
	Dim strWhereFlag

	ggoSpread.Source = frm1.vspdData 
	frm1.vspddata.Row = lgOldRow
	frm1.vspddata.Col = C1_iv_no : iTaxBillNo = Trim(frm1.vspddata.Text)
	
	strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001 & _
								  "&txtTaxBillNo=" & Trim(iTaxBillNo) 

	Call RunMyBizASP(MyBizASP, strVal)

	DbQuery2 = True                                                     
End Function

'========================================================================================
Function DbQueryOk()																		'��: ��ȸ ������ ������� 

    Dim strConid
    Dim strDtistatus
    Dim iRow
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE																'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")	
    'Call ggoOper.SetReqAttr(frm1.txtBizAreaCd, "Q")
    
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		
		For iRow= 1 to .Maxrows
		    .Row = iRow
		    .Col = C1_conversation_id
		    strConid = Trim(.Text)
		    
		    .Col = C1_dti_status
		    strDtistatus = Ucase(Trim(.Text))
		    

            If (strConid <> "")  then    'Conversation ID�� �ִ� ���(��������� �ڷ�)
    			
			    ggoSpread.SSSetProtected	C1_amend_code, iRow, iRow 
			    ggoSpread.SpreadLock		C1_amend_pop, iRow, iRow 
    			
			    ggoSpread.SSSetProtected	C1_byr_emp_name, iRow, iRow 
			    ggoSpread.SpreadLock		C1_byr_emp_pop, iRow, iRow 
			    ggoSpread.SSSetProtected	C1_byr_dept_name, iRow, iRow 
			    ggoSpread.SSSetProtected	C1_byr_tel_num, iRow, iRow 
			    ggoSpread.SSSetProtected	C1_byr_email, iRow, iRow 
		    end if	
		
		    if  strDtistatus <> "X"  then
			    ggoSpread.SSSetProtected	C1_remark, iRow, iRow 
		    End if	
		    
	   Next 
	   .ReDraw = True						
    End With
                
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
		Call SetToolbar("111000000001111")																'��: ��ư ���� ���� 
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
'Function DbSave2()
'    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
'    Dim lGrpCnt, lRow
'	Dim strWhere
'	Dim IntRetCD
'	Dim strVal
'	Dim strDtiStatus
'	Dim messageNo
'	Dim iSelectCnt
'	Dim strReturn

'    DbSave2 = False
    
'  	If LayerShowHide(1) = False then
'   	Exit Function 
'   End if 
    

'    If CommonQueryRs(" Top 1 isnull(a.RETURN_CODE,'')  RETURN_CODE "," XXSB_DTI_STATUS  A (nolock) inner join XXSB_DTI_MAIN b(nolock)  on b.CONVERSATION_ID = a.CONVERSATION_ID "," b.INTERFACE_BATCH_ID = '" & Trim(frm1.hbatchid.value) & "' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
'        strReturn = Trim(Replace(lgF0,Chr(11),""))
'    else   
'    end if



'    if  strReturn = "30000" then
         
'   else
'        LayerShowHide(0)
'        exit function        
'    end if    
        
'	With frm1
		'-----------------------
		'Data manipulate area
		'-----------------------
'		lGrpCnt = 1
'		strVal = ""
        
'        lRow = 1
        
		'-----------------------
		'Data manipulate area
		'-----------------------		
																																
'		 strVal = strVal & "U" & parent.gColSep								'0
'	     strVal = strVal & lRow & parent.gColSep							'1
'											:	strVal = strVal & Trim(frm1.hbatchid.value) & parent.gRowSep			'interface_batch_id           			      
											  													  																						  			       										  			       										  			       										  			       	
'		lGrpCnt = lGrpCnt + 1

'	   .txtMode.value        = parent.UID_M0002
'	   .txtMaxRows.value     = lGrpCnt-1	
'	   .txtSpread.value      = strVal

'	End With

'	Call ExecMyBizASP(frm1, BIZ_PGM_ID5)
	
'	DbSave2 = True	
   
'End Function


'========================================================================================
Function SaveResult()
	Call ExecMyBizASP(frm1, BIZ_PGM_ID4)			' ��: �����Ͻ� ASP �� ���� 
End Function

'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
	
	ggoSpread.Source = Frm1.vspdData2    
	ggoSpread.ClearSpreadData      
	
    'Call InitVariables															'��: Initializes local global variables
	Call MainQuery()
End Function


'========================================================================================
Function DbSaveOk2()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
	
	ggoSpread.Source = Frm1.vspdData2    
	ggoSpread.ClearSpreadData      
    'Call InitVariables															'��: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================
Function DbSaveNotOk2()													'��: ���� ������ ���� ���� 

    
End Function



'=======================================================================================================
'   Event Name : txtYr1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtIssuedToDt.Focus
        Set gActiveElement = document.activeElement
    End If
End Sub

Sub SetButtonDisable()
    
   With frm1 
    if UCase(Trim(.cboBillStatus.value)) = "X" then
        .btnSave.disabled = false
        .btnSaveCancel.disabled = true
        .btnPublishSD.disabled = true
    elseif  UCase(Trim(.cboBillStatus.value)) = "" then   
        .btnSave.disabled = true
        .btnSaveCancel.disabled = true
        .btnPublishSD.disabled = true
    else
        .btnSave.disabled = true
        .btnSaveCancel.disabled = false
        .btnPublishSD.disabled = false    
    end if
  End With  
End Sub


Function  ExeNumOk()
    Dim dti_status 
'    Call DisableToolBar(parent.TBC_QUERY)
    dti_status = "A"


    Call WebControl(Trim(frm1.hbatchid.value), dti_status)
    
	'If DbSave2 = False Then
	'	Call RestoreToolBar()
    '   Exit Function
    'End If

End Function


Function  ExeNumNot()

    'Call DisplayMsgBox("120705","X","X","X")	'%1 �ڵ�ä���� �����Ͽ����ϴ�.
	Call DisplayMsgBox("800407","X","X","X")	'�۾������� �����Դϴ�.
		
		
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
			
			 vspdData.Row = lRow
			.vspdData.Col = C1_send_check

			If .vspdData.text = "1" Then			
			   
			   iSelectCnt = iSelectCnt + 1				    															
			end if	 					 		 					 					 		    
		Next
	End With
	
	if iSelectCnt = 0 then
		IntRetCD = DisplayMsgBox("181216","X","X","X")        
		'181216: ���õ� ���� �����ϴ�.
		Exit Function
	end if
    
    Check1 = True
    	
End Function


Function ExeReflect()
	Dim intRow
	Dim intIndex
	Dim IntRetCD
	Dim DtFlag
	
	
	if frm1.vspddata.maxrows = 0 then 
	   exit Function
	end if   
	
'	if frm1.cboConfirm.value = "C" then
'		Call DisplayMsgBox("YA1299","X", "X","X") 
'		'YA299: �̹� ���εǾ��� ������ ��ü����,��ü��� ����� �ƴմϴ�
'		Exit Sub
'	end if
	
'	if frm1.vspddata.maxrows > 100 then
'	 IntRetCD = DisplayMsgBox("DA0058",parent.VB_YES_NO, "X", "X")
	 'DA0058: ȭ����  �ǳ����� ��ũ���߽��ϱ�?
	
'		If IntRetCD = vbNo Then
'			Exit Sub
'		End If
'	end if	
		
	With frm1.vspdData
	    .Redraw = False

		For intRow = 1 To .MaxRows
			.Row = intRow
			'.Col = C1_issue_dt_flag
			'DtFlag = Trim(.text)
			
			.Col = C1_send_check
			'if .value = "0" and DtFlag = "N" then
			if .value = "0" then
			.value = "1"
			
'			ggoSpread.Source = frm1.vspdData
'			ggoSpread.UpdateRow intRow	
		    else 
		    end if
'		ggoSpread.Source = frm1.vspdData1
'        ggoSpread.UpdateRow intRow			
		Next
                
	    .Redraw = True
	End With
End Function

Function ExeReflect2()
	Dim intRow
	Dim intIndex
	Dim IntRetCD
	Dim Flag
	Dim DtFlag
	
	if frm1.vspddata.maxrows = 0 then 
	   exit Function
	end if   
	
'	if frm1.cboConfirm.value = "C" then
'		Call DisplayMsgBox("YA1299","X", "X","X") 
'		'YA299: �̹� ���εǾ��� ������ ��ü����,��ü��� ����� �ƴմϴ�
'		Exit Sub
'	end if
	
'	if frm1.vspddata.maxrows > 100 then
'	 IntRetCD = DisplayMsgBox("DA0058",parent.VB_YES_NO, "X", "X")
	 'DA0058: ȭ����  �ǳ����� ��ũ���߽��ϱ�?
	
'		If IntRetCD = vbNo Then
'			Exit Sub
'		End If
'	end if	
	
	With frm1.vspdData
	    .Redraw = False

		For intRow = 1 To .MaxRows
			.Row = intRow           
			'.Col = C1_issue_dt_flag
			'DtFlag = Trim(.text)

			.Col = C1_send_check
			
			'if .value = "1" and DtFlag = "N" then
			if .value = "1"  then
				.value = "0" 				
				 
'				ggoSpread.Source = frm1.vspdData
'				ggoSpread.UpdateRow intRow
'				ggoSpread.EditUndo intRow
'				.Col = C_PostFlag
'				.text = "N"
			else 
			
			end if
			
'		ggoSpread.Source = frm1.vspdData1
'        ggoSpread.UpdateRow intRow	
		Next

	    .Redraw = True
	End With
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag�� 
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
							<TD WIDTH="*" ALIGN=RIGHT><BUTToN NAME="btnExeReflect" CLASS="CLSSBTNCALC" ONCLICK="vbscript:Call ExeReflect()" >��ü����</BUTToN>&nbsp;<BUTToN NAME="btnExeReflect2" CLASS="CLSSBTNCALC" ONCLICK="vbscript:Call ExeReflect2()" >��ü���</BUTToN></TD>
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
											<TD CLASS="TD5"NOWRAP>������</TD>
											<TD CLASS="TD6"NOWRAP>
												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssuedFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="�����������" class=required></OBJECT>');</SCRIPT> ~
 												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssuedToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������������" class=required></OBJECT>');</SCRIPT>
 											</TD>
 											<TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="13XXXU" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
												<INPUT TYPE=TEXT AlT="���ݽŰ�����" ID="txtBizAreaNm" NAME="txtBizAreaNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD> 											
										</TR>
										<TR>
 											<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSalesGrpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ű׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSalesGrp()">
												<INPUT TYPE=TEXT AlT="���ű׷�" ID="txtSalesGrpNm" NAME="txtSalesGrpNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>
											<TD CLASS="TD5" NOWRAP>����ó</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
												<INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD>											 											
										</TR>
										<TR>
											<TD CLASS="TD5"NOWRAP>��꼭����</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboBillStatus" ALT="���ݰ�꼭����" CLASS=cboNormal TAG="11" style="width:120px"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
											<TD CLASS=TD5 nowrap>�ŷ���������</TD>											
											<TD CLASS=TD6 nowrap><INPUT type="checkbox" CLASS="STYLE CHECK" NAME=chkShowBiz ID=chkShowBiz  tag="11" ></TD>
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
							<TD><BUTTON NAME="btnSave" CLASS="CLSSBTN" OnClick="VBScript:Call fnSave()">��������</BUTTON>&nbsp;
                                <BUTTON NAME="btnSaveCancel" CLASS="CLSSBTN" OnClick="VBScript:Call fnSaveCancel()">�������</BUTTON>&nbsp;
							    <BUTTON NAME="btnPublishSD" CLASS="CLSSBTN" OnClick="VBScript:Call fnPublish()">�������û&nbsp;</TD>
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
