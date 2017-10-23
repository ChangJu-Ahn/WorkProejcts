<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : ���ڼ��ݰ�꼭(����Ʈ��(�����))
'*  2. Function Name        : 
'*  3. Program ID           : D4311MA1
'*  4. Program Name         : ���Լ��ݰ�꼭(������)
'*  5. Program Desc         :    
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
Const BIZ_PGM_ID  = "D4311MB1.asp"  'Main��ȸ
Const BIZ_PGM_ID2 = "D4311MB2.asp"  'Dtl��ȸ
Const BIZ_PGM_ID3 = "D4311MB3.asp"  '���Ž���, ���Űź�, ������ҿ�û
Const BIZ_PGM_ID4 = "D4311MB4.asp"  '������ҽ���
Const BIZ_PGM_ID5 = "D4311MB5.asp"  '������Ұź�
Const BIZ_PGM_ID6 = "D4311MB6.asp"  '��꼭����


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
Dim     C1_dti_wdate            '��꼭����
Dim     C1_conversation_id      '��꼭��ȣ
Dim 	C1_dti_type             '��꼭����
Dim 	C1_dti_type_nm          '��꼭����
Dim     C1_sup_com_regno        '������ ����ڵ�Ϲ�ȣ
Dim     C1_sup_com_name         '������ ��ȣ
Dim 	C1_sup_amount           '���ް���
Dim 	C1_tax_amount           '�ΰ���
Dim 	C1_total_amount         '�հ�ݾ�
Dim     C1_dti_status           '��������
Dim     C1_dti_status_nm        '��������
Dim     C1_sup_emp_name         '������ ����ڸ�
Dim     C1_sup_tel_num          '������ ����� ��ȭ��ȣ
Dim     C1_sup_email            '������ ����� Email
Dim     C1_sbdescription        '��Ұźλ���
Dim     C1_amend_code           '�����ڵ�
Dim     C1_amend_code_name      '�����ڵ��
Dim     C1_remark               '���
Dim     C1_byr_com_regno        '���޹޴��� ����ڵ�Ϲ�ȣ



'add detail datatable column
Dim	    C2_item_code            'ǰ���ڵ�
Dim	    C2_item_name            'ǰ���
Dim	    C2_item_size            '�԰�  
Dim	    C2_item_qty             '����
Dim	    C2_unit_price           '�ܰ�
Dim	    C2_sup_amount           '���ް���
Dim     C2_tax_amount           '�ΰ�����
Dim     C2_total_amount         '�հ�ݾ�
Dim     C2_item_md              '��������
Dim     C2_remark               '���
Dim     C2_conversation_id      '��꼭��ȣ


Dim lgStrPrevKey1    
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
      Call initcombobox  
      
      Call InitSpreadSheet()
      Call InitSpreadSheet2
      
      
      Call SetDefaultVal
      Call InitVariables
 
      Call SetToolbar("111000000000111")										'��: ��ư ���� ����    	
 
      .txtSupplierCd.focus
      
      .btnApprove.disabled	= false
      .btnReceiveReject.disabled	= false
      
      .btnCancelRequest.disabled	= true      
      .btnAccept.disabled	= true      
      .btnReject.disabled	= true      
      .btnPrint.disabled	= true
      
      
   End With		
End Sub

'========================================================================================================= 
Sub InitComboBox()
   Dim iCodeArr 
   Dim iNameArr
   Dim iDx
	
	'��꼭�� �������
    Call CommonQueryRs(" B.MINOR_CD , B.MINOR_NM "," B_CONFIGURATION A INNER JOIN B_MINOR B ON (A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B. MINOR_CD) ", _
                         " A.MAJOR_CD='DT409' and A.SEQ_NO = 1 and B.MINOR_CD not in ('X', 'S') ORDER BY A.REFERENCE ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    'Call CommonQueryRs(" MINOR_CD , MINOR_NM "," B_MINOR (NOLOCK)  ", " MAJOR_CD='DT409' and MINOR_CD in('I','O','T','V', 'C','R','W','N','M','O')  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
    Call SetCombo2(frm1.cboBillStatus ,lgF0  ,lgF1  ,Chr(11))
        
    frm1.cboBillStatus.Value = "I"
    
    
	'��������
    Call CommonQueryRs(" MINOR_CD , MINOR_NM "," B_MINOR (NOLOCK)  ", " MAJOR_CD='DT408'  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
    Call SetCombo2(frm1.cboAmendCode ,lgF0  ,lgF1  ,Chr(11))		
    		   
End Sub

Sub InitSpreadPosVariables()
	'add tab1 header datatable column
		 	
 	 	C1_send_check           = 1  '����
        C1_dti_wdate            = 2  '��꼭����
        C1_conversation_id      = 3  '��꼭��ȣ
     	C1_dti_type             = 4  '��꼭����
     	C1_dti_type_nm          = 5  '��꼭����
        C1_sup_com_regno        = 6  '������ ����ڵ�Ϲ�ȣ
        C1_sup_com_name         = 7  '������ ��ȣ
     	C1_sup_amount           = 8  '���ް���
     	C1_tax_amount           = 9  '�ΰ���
     	C1_total_amount         = 10  '�հ�ݾ�
        C1_dti_status           = 11  '��������
        C1_dti_status_nm        = 12  '��������
        C1_sup_emp_name         = 13  '������ ����ڸ�
        C1_sup_tel_num          = 14  '������ ����� ��ȭ��ȣ
        C1_sup_email            = 15  '������ ����� Email
        C1_sbdescription        = 16  '��Ұźλ���
        C1_amend_code           = 17  '�����ڵ�
        C1_amend_code_name      = 18  '�����ڵ��
        C1_remark               = 19  '���
        C1_byr_com_regno        = 20  '���޹޴��� ����ڵ�Ϲ�ȣ
	 		 		 	       	 		 	    	    	
End Sub

Sub InitSpreadPosVariables2()
	'add tab1 detail datatable column
	
	    C2_item_code            = 1   'ǰ���ڵ�
	    C2_item_name            = 2   'ǰ���
	    C2_item_size            = 3   '�԰�  
	    C2_item_qty             = 4   '����
	    C2_unit_price           = 5   '�ܰ�
	    C2_sup_amount           = 6   '���ް���
        C2_tax_amount           = 7   '�ΰ�����
        C2_total_amount         = 8   '�հ�ݾ�
        C2_item_md              = 9   '��������
        C2_remark               = 10  '���
        C2_conversation_id      = 11  '��꼭��ȣ

End Sub

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE				'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
   
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgStrPrevKey1      = ""                                      '��: initializes Previous Key
    lgLngCurRows = 0									   'initializes Deleted Rows Count
    
    lgSortKey       = 1                                       '��: initializes sort direction
    lgSortKey1	    = "1"
    lgPageNo_B		= ""                          'initializes Previous Key for spreadsheet #2    
    
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

	
	'If CommonQueryRs(" SALES_GRP_NM "," B_SALES_GRP "," SALES_GRP = '" & parent.gSalesGrp & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
    '   frm1.txtSalesGrpNm.value =  Trim(Replace(lgF0,Chr(11),""))
    'else
    'End if 
	
	'frm1.txtSalesGrpCd.value = parent.gSalesGrp

	If CommonQueryRs(" top 1 minor_cd, minor_nm "," B_MINOR (nolock) "," MAJOR_CD = 'DT412'  order by minor_cd " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
       frm1.txtBizAreaCd.value =  Trim(Replace(lgF0,Chr(11),""))
       frm1.txtBizAreaNm.value =  Trim(Replace(lgF1,Chr(11),""))
    else
    End if 
	
	frm1.cboBillStatus.Value = "I"
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
        .MaxCols = C1_byr_com_regno + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .Col = .MaxCols														'������Ʈ�� ��� Hidden Column
        .ColHidden = True
        .MaxRows = 0


		Call GetSpreadColumnPos("A")

		' uniGrid1 setting
		ggoSpread.SSSetCheck	C1_send_check,		"����",     			 	 4,  -10, "", True, -1
		ggoSpread.SSSetDate  	C1_dti_wdate,       "��꼭����",     			13, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C1_conversation_id,	"��꼭��ȣ",				30, ,,50
		ggoSpread.SSSetEdit  	C1_dti_type,		"��꼭����",				15, ,,18		
		ggoSpread.SSSetEdit  	C1_dti_type_nm,		"��꼭����",				15, ,,18		
		ggoSpread.SSSetEdit  	C1_sup_com_regno,   "������ ����ڵ�Ϲ�ȣ",	15, ,,30
		ggoSpread.SSSetEdit  	C1_sup_com_name ,   "������ ��ȣ",  	        25, ,,50
		ggoSpread.SSSetFloat	C1_sup_amount ,    	"���ް���",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_tax_amount ,    	"�ΰ����ݾ�",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat	C1_total_amount,	"�հ�ݾ�",     			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetEdit  	C1_dti_status,		"��������",	    			15, ,,18
		ggoSpread.SSSetEdit  	C1_dti_status_nm,	"��������",		    	    20, ,,40				
		ggoSpread.SSSetEdit  	C1_sup_emp_name,   "������ ����ڸ�",       	10, ,,50
		ggoSpread.SSSetEdit  	C1_sup_emp_name,   "������ ����ڸ�",       	10, ,,50
		ggoSpread.SSSetEdit  	C1_sup_tel_num,    "������ ����� ��ȭ��ȣ",	25, ,,30		
		ggoSpread.SSSetEdit  	C1_sup_email,      "������ ����� Email",       20, ,,40		
		ggoSpread.SSSetEdit  	C1_sbdescription,	 "���/�źλ���",			15, ,,50
		ggoSpread.SSSetEdit  	C1_amend_code,		 "�����ڵ�",				15, ,,18		
		ggoSpread.SSSetEdit  	C1_amend_code_name,	 "��������",			    15, ,,40						
		ggoSpread.SSSetEdit		C1_remark,			"���",				    	20, ,,50		
		ggoSpread.SSSetEdit  	C1_byr_com_regno,   "���޹޴��� ����ڵ�Ϲ�ȣ",	25, ,,30
		
		                        		
		'Call ggoSpread.MakePairsColumn(C1_change_reason_cd, C1_change_reason, "1")
      Call ggoSpread.SSSetColHidden(C1_dti_type, C1_dti_type, True)
      Call ggoSpread.SSSetColHidden(C1_dti_status, C1_dti_status, True)
      Call ggoSpread.SSSetColHidden(C1_byr_com_regno, C1_byr_com_regno, True)


		.ReDraw = True
	End With

	Call SetSpreadLock()
End Sub

Sub InitSpreadSheet2()

	Call initSpreadPosVariables2()
	With frm1.vspdData2	
		.MaxCols = C2_conversation_id + 1								'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols												'��: ����� �� Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData2
		.ReDraw = False 
		ggoSpread.Spreadinit "V20090707",, parent.gAllowDragDropSpread
		.ReDraw = False

		Call GetSpreadColumnPos("B")
		
		ggoSpread.SSSetEdit  	C2_item_code, 			"ǰ��", 			15, ,,18
		ggoSpread.SSSetEdit  	C2_item_name, 			"ǰ���", 			30, ,,30		
		ggoSpread.SSSetEdit  	C2_item_size,			"�԰�", 			15, ,,40
        ggoSpread.SSSetFloat	C2_item_qty ,			"����",				15, parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"		
        ggoSpread.SSSetFloat  	C2_unit_price, 		    "�ܰ�", 			18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec	
        ggoSpread.SSSetFloat  	C2_sup_amount, 			"���ް���", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  	C2_tax_amount, 			"�ΰ�����", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec				
		ggoSpread.SSSetFloat  	C2_total_amount, 		"�հ�ݾ�", 		18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetDate  	C2_item_md,             "��������",     	13, 2, parent.gDateFormat
		ggoSpread.SSSetEdit  	C2_remark, 		        "���", 		    20, ,,50				
		ggoSpread.SSSetEdit		C2_conversation_id,	    "��꼭��ȣ",		30, ,,50		
						        						
		Call ggoSpread.SSSetColHidden(C2_conversation_id, C2_conversation_id, True)

		.ReDraw = True
	End With	
	Call SetSpreadLock_B()
End Sub

'========================================================================================
Sub SetSpreadLock()
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

	
		'ggoSpread.SpreadLock    C_send_check,		-1, C_send_check
		ggoSpread.SpreadLock    C1_dti_wdate, 		-1, C1_dti_wdate						
		ggoSpread.SpreadLock    C1_conversation_id, -1, C1_conversation_id		
		ggoSpread.SpreadLock    C1_dti_type, 		-1, C1_dti_type		
		ggoSpread.SpreadLock    C1_dti_type_nm, 	-1, C1_dti_type_nm						
		ggoSpread.SpreadLock    C1_sup_com_regno, 	-1, C1_sup_com_regno		
		ggoSpread.SpreadLock    C1_sup_com_name, 	-1, C1_sup_com_name		
		ggoSpread.SpreadLock    C1_sup_amount , 	-1, C1_sup_amount 		
		ggoSpread.SpreadLock    C1_tax_amount , 	-1, C1_tax_amount 						
		ggoSpread.SpreadLock    C1_total_amount , 	-1, C1_total_amount 		
		ggoSpread.SpreadLock    C1_dti_status, 		-1, C1_dti_status		
		ggoSpread.SpreadLock    C1_dti_status_nm, 	-1, C1_dti_status_nm		
		ggoSpread.SpreadLock	C1_sup_emp_name,	-1, C1_sup_emp_name
		ggoSpread.SpreadLock	C1_sup_tel_num,	    -1, C1_sup_tel_num		
		ggoSpread.SpreadLock	C1_sup_email,	    -1, C1_sup_email		
		ggoSpread.SpreadLock    C1_sbdescription,   -1, C1_sbdescription
		ggoSpread.SpreadLock    C1_amend_code, 		-1, C1_amend_code
		ggoSpread.SpreadLock    C1_amend_code_name, -1, C1_amend_code_name
		ggoSpread.SpreadLock	C1_remark,		    -1, C1_remark		
		ggoSpread.SpreadLock    C1_byr_com_regno,   -1, C1_byr_com_regno
																			
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
																	
            C1_send_check           = iCurColumnPos(1)    '����
            C1_dti_wdate            = iCurColumnPos(2)    '��꼭����
            C1_conversation_id      = iCurColumnPos(3)    '��꼭��ȣ
     	    C1_dti_type             = iCurColumnPos(4)    '��꼭����
     	    C1_dti_type_nm          = iCurColumnPos(5)    '��꼭����
            C1_sup_com_regno        = iCurColumnPos(6)    '������ ����ڵ�Ϲ�ȣ
            C1_sup_com_name         = iCurColumnPos(7)    '������ ��ȣ
     	    C1_sup_amount           = iCurColumnPos(8)    '���ް���
     	    C1_tax_amount           = iCurColumnPos(9)    '�ΰ���
     	    C1_total_amount         = iCurColumnPos(10)   '�հ�ݾ�
            C1_dti_status           = iCurColumnPos(11)   '��������
            C1_dti_status_nm        = iCurColumnPos(12)   '��������
            C1_sup_emp_name         = iCurColumnPos(13)   '������ ����ڸ�
            C1_sup_tel_num          = iCurColumnPos(14)   '������ ����� ��ȭ��ȣ
            C1_sup_email            = iCurColumnPos(15)   '������ ����� Email
            C1_sbdescription        = iCurColumnPos(16)   '��Ұźλ���
            C1_amend_code           = iCurColumnPos(17)   '�����ڵ�
            C1_amend_code_name      = iCurColumnPos(18)   '�����ڵ��
            C1_remark               = iCurColumnPos(19)   '���
            C1_byr_com_regno        = iCurColumnPos(20)   '���޹޴��� ����ڵ�Ϲ�ȣ

        Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C2_item_code            = iCurColumnPos(1)    'ǰ���ڵ�
	        C2_item_name            = iCurColumnPos(2)    'ǰ���
	        C2_item_size            = iCurColumnPos(3)    '�԰�  
	        C2_item_qty             = iCurColumnPos(4)    '����
	        C2_unit_price           = iCurColumnPos(5)    '�ܰ�
	        C2_sup_amount           = iCurColumnPos(6)    '���ް���
            C2_tax_amount           = iCurColumnPos(7)    '�ΰ�����
            C2_total_amount         = iCurColumnPos(8)    '�հ�ݾ�
            C2_item_md              = iCurColumnPos(9)    '��������
            C2_remark               = iCurColumnPos(10)   '���
            C2_conversation_id      = iCurColumnPos(11)   '��꼭��ȣ
													    
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

	arrParam(0) = "�ŷ�ó"					
	arrParam(1) = "b_biz_partner"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE IN ('S','CS')"	
	arrParam(5) = "�ŷ�ó"						

	arrField(0) = "bp_cd"					
	arrField(1) = "BP_FULL_NM"	
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

	arrParam(0) = "�����׷�"
	arrParam(1) = "b_sales_grp"

	arrParam(2) = Trim(frm1.txtSalesGrpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "usage_flag = 'Y'"
	arrParam(5) = "�����׷�"

	arrField(0) = "sales_grp"
	arrField(1) = "sales_grp_nm"

	arrHeader(0) = "�����׷�"				
	arrHeader(1) = "�����׷��"	

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
        
        
         frm1.vspddata.Col = C1_bp_cd
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

'���Ž���
Function fnApprove()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Dim lGrpCnt, lRow
	Dim strWhere
	Dim IntRetCD
	Dim strVal
	Dim DtiStatus
	Dim RetFlag
	Dim iSelectCnt
    Dim sbdescription
    Dim ApproveFlag
    Dim strConvid
    Dim strDtiStatus

    fnApprove = False

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    
    If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
    
   '����� ID �� URL, ������ üũ
    if Check9() = false then
      Call LayerShowHide(0)
      exit function
    end if  
                  
    ApproveFlag = "SD"
    DtiStatus = "C"
    sbdescription  = Trim(frm1.txtIsusse.value)
     
             	      	      	
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

			If Trim(.vspdData.text) = "1"  Then
			
			     
			     '������üũ
			      .vspdData.Col = C1_dti_status
			      strDtiStatus = Trim(.vspdData.text)
				  
				  
				  .vspddata.Col = C1_conversation_id
				  strConvid = Trim(.vspdData.text)
				  
			     if	(strDtiStatus <> "I") then
					Call DisplayMsgBox("DT4104","X", strConvid,"X")	            		
			         'DT4104:  %1�� ���� �� �� �ִ� ����� �ƴմϴ�.
			         Call LayerShowHide(0)
			         Exit Function                                                               																					         
                 else
											
			         
			     end if	  
			
				
				'--------------------------------------------------------------------------																																				
								
				 strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_conversation_id     :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID              '2			     
                                                        :	strVal = strVal & sbdescription & parent.gColSep		                                  '3  			     			       			     
  													    :	strVal = strVal & DtiStatus & parent.gColSep				     			      '4	
                                                        :	strVal = strVal & ApproveFlag & parent.gRowSep		                   '5			       													    
  													  													  															
															  			       			       	  			       			       			     
				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1
            End If
     Next

		if iSelectCnt = 0 then			
			Call DisplayMsgBox("DT4205","X", "X","X")	
			'DT4205: ���õ� ���� �����ϴ�.            		
			Call LayerShowHide(0)
			Exit Function
		Elseif iSelectCnt > 1 then 
			Call DisplayMsgBox("DT4201","X", "X","X")	            		
			'DT4201: �������� ���ݰ�꼭�� ���ÿ� ó���Ҽ� �����ϴ�.
			Call LayerShowHide(0)
			Exit Function
		end if
        
	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID3)
	
	fnApprove = True	

End Function


'���Űź�
Function fnReceiveReject()

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
	Dim strType
    Dim DtiStatus
    Dim sbdescription
    Dim ReceieveRejectFlag

    fnReceiveReject = False

     DtiStatus = "R"
     ReceieveRejectFlag = "SD"
     
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
     
   '����� ID �� URL, ������ üũ
    if Check9() = false then
      Call LayerShowHide(0)
      exit function
    end if   
         
    sbdescription  = Trim(frm1.txtIsusse.value)
    
     if sbdescription = "" then
        RetFlag = DisplayMsgBox("DT4204", parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ� 
        'DT4204: �ź�/��� ������ �����ϴ�. ��������Ͻðڽ��ϱ�?

        If RetFlag = VBNO Then
            Call LayerShowHide(0)
            Exit Function
        End If
    End if                      
      	      	
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
				  strDtiStatus = Trim(.vspdData.text) 
    				
				  if strDtiStatus <>  "I" then  
			        .vspdData.Col = C1_conversation_id
			         messageNo = Trim(.vspdData.text)
				
			         Call DisplayMsgBox("W70001","X", messageNo & "�� ���Űź� ����� �ƴմϴ�.", "X")	            		
			            'W70001: %1	
			             Call LayerShowHide(0)
			            Exit Function												
			      end if
			
						                                                                                																																				     
				'--------------------------------------------------------------------------																																				
								
				 strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_conversation_id     :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID              '2			     
                                                        :	strVal = strVal & sbdescription & parent.gColSep		                                  '3  			     			       			     
  													    :	strVal = strVal & DtiStatus & parent.gColSep				     			      '4	
                                                        :	strVal = strVal & ReceieveRejectFlag & parent.gRowSep		                   '5			       					       													      													    
  													    
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
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
	
	fnReceiveReject = True	

End Function


'������� ��û
Function fnCancelRequest()

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
	Dim strType
    Dim DtiStatus
    Dim sbdescription
    Dim CancelRequestFlag

    fnCancelRequest = False

    'If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                               
    '    Exit Function
    'End If
    
            
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 

    
   '����� ID �� URL, ������ üũ
    if Check9() = false then
      Call LayerShowHide(0)
      exit function
    end if   

        
    CancelRequestFlag = "SD"    
    DtiStatus = "M"
        
    sbdescription  = Trim(frm1.txtIsusse.value)

    
     if sbdescription = "" then
        RetFlag = DisplayMsgBox("DT4204", parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ� 
        'DT4204: �ź�/��� ������ �����ϴ�. ��������Ͻðڽ��ϱ�?

        If RetFlag = VBNO Then
            Call LayerShowHide(0)
            Exit Function
        End If
    End if                      
      	      	
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
				  strDtiStatus = Trim(.vspdData.text) 
    				
				  if strDtiStatus <>  "C"   then  
			        .vspdData.Col = C1_conversation_id
			         messageNo = Trim(.vspdData.text)
				
			         Call DisplayMsgBox("W70001","X", messageNo & "�� ������ҿ�û ����� �ƴմϴ�.", "X")	            		
			            'W70001: %1	
			             Call LayerShowHide(0)
			            Exit Function												
			      end if
			
			
						                                                                                																																				     
				'--------------------------------------------------------------------------																																				
								
				strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_conversation_id     :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID              '2			     
                                                        :	strVal = strVal & sbdescription & parent.gColSep		                                  '3  			     			       			     
  													    :	strVal = strVal & DtiStatus & parent.gColSep				     			      '4	
                                                        :	strVal = strVal & CancelRequestFlag & parent.gRowSep		                   '5			       		
  													    
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
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
	
	fnCancelRequest = True	

End Function


'������� ����
Function fnAccept()

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
	Dim strType
    Dim DtiStatus
    Dim sbdescription
    Dim AcceptFlag

    fnAccept = False

    'If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                               
    '    Exit Function
    'End If
    
            
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
        
   '����� ID �� URL, ������ üũ
    if Check9() = false then
      Call LayerShowHide(0)
      exit function
    end if   
   
    
    AcceptFlag = "SD"    
    DtiStatus = "O"
        
    sbdescription  = Trim(frm1.txtIsusse.value)

                              	      	
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
				  strDtiStatus = Trim(.vspdData.text) 
    				
				  if strDtiStatus <>  "N"   then  
			        .vspdData.Col = C1_conversation_id
			         messageNo = Trim(.vspdData.text)
				
			         Call DisplayMsgBox("DT4104","X", messageNo, "X")	            		
			            'DT4104: %1�� ���� �� �� �ִ� ����� �ƴմϴ�.	
			             Call LayerShowHide(0)
			            Exit Function												
			      end if
			
			
						                                                                                																																				     
				'--------------------------------------------------------------------------																																				
								
				strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_conversation_id     :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID              '2			     
                                                        :	strVal = strVal & sbdescription & parent.gColSep		                                  '3  			     			       			     
  													    :	strVal = strVal & DtiStatus & parent.gColSep				     			      '4	
                                                        :	strVal = strVal & AcceptFlag & parent.gRowSep		                   '5			    
  													    
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
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
	
	fnAccept = True	

End Function


'������� �ź�
Function fnReject()

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
	Dim strType
    Dim DtiStatus
    Dim sbdescription
    Dim RejectFlag

    fnReject = False

    'If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                               
    '    Exit Function
    'End If
    
            
   	If LayerShowHide(1) = False then
    	Exit Function 
    End if 
    
        
   '����� ID �� URL, ������ üũ
    if Check9() = false then
      Call LayerShowHide(0)
      exit function
    end if   

    
    RejectFlag = "SD"    
    DtiStatus = "C"
        
    sbdescription  = Trim(frm1.txtIsusse.value)
            
     if sbdescription = "" then
        RetFlag = DisplayMsgBox("DT4204", parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ� 
        'DT4204: �ź�/��� ������ �����ϴ�. ��������Ͻðڽ��ϱ�?

        If RetFlag = VBNO Then
            Call LayerShowHide(0)
            Exit Function
        End If
    End if                      
    
    
                              	      	
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
				  strDtiStatus = Trim(.vspdData.text) 
    				
				  if strDtiStatus <>  "N"   then  
			        .vspdData.Col = C1_conversation_id
			         messageNo = Trim(.vspdData.text)
				
			         Call DisplayMsgBox("DT4104","X", messageNo, "X")	            		
			            'DT4104: %1�� ���� �� �� �ִ� ����� �ƴմϴ�.	
			             Call LayerShowHide(0)
			            Exit Function												
			      end if
			
			
						                                                                                																																				     
				'--------------------------------------------------------------------------																																				
								
				strVal = strVal & "C" & parent.gColSep								'0
  			     strVal = strVal & lRow & parent.gColSep							'1
  			     .vspdData.Col = C1_conversation_id     :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep		'Conversation ID              '2			     
                                                        :	strVal = strVal & sbdescription & parent.gColSep		                                  '3  			     			       			     
  													    :	strVal = strVal & DtiStatus & parent.gColSep				     			      '4	
                                                        :	strVal = strVal & RejectFlag & parent.gRowSep		                   '5			    
  													    
  													  													  																						  			       										  			       										  			       										  			       	
				  			       			       	  			       			       			     
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

	Call ExecMyBizASP(frm1, BIZ_PGM_ID5)
	
	fnReject = True	

End Function




'��꼭 ����
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
    
    
   '����� ID �� URL, ������ üũ
    if Check9() = false then
      Call LayerShowHide(0)
      exit function
    end if   
    
    
    DtiStatus = "CP"
    sbdescription  = ""
    PrintFlag = ""
      	      	
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
				  strDtiStatus = Trim(.vspdData.text) 
    				
				  if strDtiStatus <>  "C"   then  
			        .vspdData.Col = C1_conversation_id
			         messageNo = Trim(.vspdData.text)
				
			         Call DisplayMsgBox("DT4111","X", messageNo, "X")	            		
			            'DT4111: %1�� ��� �� �� �����ϴ�.
			             Call LayerShowHide(0)
			            Exit Function												
			      end if
			
			
			
						                                                                                																																				     
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
			'DT4205: ���õ� ���� �����ϴ�.            		
			Call LayerShowHide(0)
			Exit Function
		end if

	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID6)
	
	fnPrint = True	

End Function



Function WebControl(batchid,  status, sbdescription, signal)
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
    Dim strBatchid
    Dim strConvid
    
    
	DIm lRow
	
    'If lgIntFlgMode <> parent.OPMD_UMODE Then												'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                                       
    '    Exit Function
    'End If
    
    strBatchid = ""

     If CommonQueryRs(" TOP 1 A.SMART_ID, A.SMART_PASSWORD "," XXSB_DTI_SM_USER A (nolock) ", _
              " A.FND_USER = '" & parent.gUsrID & "'  AND A.FND_REGNO = (SELECT TOP 1 REPLACE(OWN_RGST_NO,'-','') FROM B_TAX_BIZ_AREA WHERE TAX_BIZ_AREA_CD = '" & Trim(frm1.txtBizAreaCd.value) & "')" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then               
        strSmartid = Trim(Replace(lgF0,Chr(11),""))
        strSmartPW = Trim(Replace(lgF1,Chr(11),""))
     else
        Call DisplayMsgBox("DT4118","X", "X","X")	               		
        'DT4118:���ڼ��ݰ�꼭�� ����ڸ� Ȯ���ϼ���
        Exit Function 	  
     End if
    
      If CommonQueryRs(" TOP 1 Convert(varchar(10),EXPIRATION_DATE,120) as EXPIRATION_DATE "," XXSB_DTI_CERT (nolock) "," CERT_REGNO IN ( SELECT REPLACE(B.BP_RGST_NO,'-','') FROM B_TAX_BIZ_AREA A INNER JOIN B_BIZ_PARTNER B ON (A.TAX_BIZ_AREA_CD = B.BP_CD) where A.TAX_BIZ_AREA_CD = '" & Trim(frm1.txtBizAreaCd.value) & "')" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
        strExpDate =  Trim(Replace(lgF0,Chr(11),""))
        
        if  strExpDate = "" then
            Call DisplayMsgBox("DT4206","X", "X","X")	               		
            'DT4206: �������� Ȯ���ϼ���
            Exit Function 	  
        end if      
      else
         Call DisplayMsgBox("DT4206","X", "X","X")	               		
         'DT4206: �������� Ȯ���ϼ���
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
      
      'If CommonQueryRs(" Top 1 convert(numeric(15),substring(replace(replace(replace(replace(convert(varchar(25), getdate(), 121), '-',''),' ',''), ':',''),'.',''),1,15)) AS Batch_Id "," b_minor (nolock) "," 1 = 1"  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
      '    strBatchid = Trim(Replace(lgF0,Chr(11),""))
      'else      
      'end if    
      
      'if strBatchid = "" then
      '      Call DisplayMsgBox("W70001","X", "BatchID ä���� �����߽��ϴ�.","X")	            		
	  '       'W70001:   %1
	  '       Exit Function
      'end if
                       
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
      
                
      if UCase(status) <> "CP" then 
         strIssueASP = "XXSB_DTI_STATUS_CHANGE.asp"
            strURL =  strUrlinfo & strIssueASP & "?batch_id=" + batchid + "&STATUS=" + status + "&SIGNAL=" + signal + "&ID=" + strSmartid + "&PASS=" + strSmartPW + "&SBDESCRIPTION=" + sbdescription + ""   
      else  
             strURL =  strUrlinfo & "XXSB_DTI_PRINT.asp?batch_id=" + batchid +  "&SORTFIELD=A&SORTORDER=1 "
              
            'strIssueASP = "XXSB_DTI_RARISSUE.asp"
            'strURL =  strUrlinfo & strIssueASP & "?CONVERSATION_ID=" + strConvid + "&ID=" + strSmartid + "&PASS=" + strSmartPW + ""            
      end if                
                                
                                        
       arrRet =  window.showModalDialog(strUrl ,, "dialogWidth=680px; dialogHeight=600px; center: Yes; help: No; resizable: No; status: no; scroll:Yes;")       
           

        DbSaveOk
                                                 	
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
           'DT4108: URL ������ �����ϴ�.
           Exit Function 	  
            end if            
       else
           Call DisplayMsgBox("DT4108","X", "X","X")	               		
           'DT4108: URL ������ �����ϴ�.
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
'												4. Common Function�� 
'=========================================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False																		'��: Processing is NG


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
        
        frm1.txtIsusse.value = ""
        
		'-----------------------
		'Erase contents area
		'-----------------------
		'	    Call ggoOper.ClearField(Document, "2")												'��: Clear Contents  Field
		
	End With
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

		Call InitVariables 																	'��: Initializes local global variables
        
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


	Call ggoOper.LockField(Document, "N")												'��: Lock  Suitable  Field
	Call SetDefaultVal
	Call InitVariables																	'��: Initializes local global variables
    
    
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

	Call SetButtonDisable()
	 

    If LayerShowHide(1) = False then
		Exit Function 
   	End if
		
	With Frm1
	    If lgIntFlgMode = parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'��: 
		    strVal = strVal & "&txtIssuedFromDt=" & Trim(.htxtIssuedFromDt.value)		
		    strVal = strVal & "&txtIssuedToDt=" & Trim(.htxtIssuedToDt.value)				    
		    strVal = strVal & "&cboBillStatus=" & Trim(.hcboBillStatus.value)						
		    strVal = strVal & "&txtSupplierNm=" & Trim(.htxtSupplierNm.value)
		    strVal = strVal & "&txtBizAreaCd=" & Trim(.htxtBizAreaCd.value)	
		    strVal = strVal & "&cboAmendCode=" & Trim(.hcboAmendCode.Value)						                										
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        Else
		    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
		    strVal = strVal     &  "&txtIssuedFromDt=" & Trim(.txtIssuedFromDt.text)
		    strVal = strVal     &  "&txtIssuedToDt=" & Trim(.txtIssuedToDt.text)		    
		    strVal = strVal     &  "&cboBillStatus=" & Trim(.cboBillStatus.Value)						    
		    strVal = strVal     &  "&txtSupplierNm=" & Replace(Trim(.txtSupplierNm.Value),"-","")				
		    strVal = strVal     &  "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)						    
            strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
            strVal = strVal     &  "&cboAmendCode=" & Trim(.cboAmendCode.Value)						                
            strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
        End if
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
	Dim iConvid
    
    if LayerShowHide(1) = False then
	   Exit Function
	end if                                                         '��: Show Processing Message
    
	ggoSpread.Source = frm1.vspdData 
	frm1.vspddata.Row = lgOldRow
	frm1.vspddata.Col = C1_conversation_id : iConvid = Trim(frm1.vspddata.Text)
									  
    With Frm1
	    strVal = BIZ_PGM_ID2 & "?txtMode="            & parent.UID_M0001						         
	    strVal = strVal     &  "&txtConvid=" & Trim(iConvid) 
        strVal = strVal     &  "&txtMaxRows="         & .vspdData2.MaxRows
        strVal = strVal     &  "&lgStrPrevKey1=" & lgStrPrevKey1                 '��: Next key tag
   End With    								  
								  	
	Call RunMyBizASP(MyBizASP, strVal)

	DbQuery2 = True                                                     
End Function

'========================================================================================
Function DbQueryOk()																		'��: ��ȸ ������ ������� 

    Dim iRow
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE																'��: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q") 
                    
	lgOldRow = 1
	frm1.vspdData.Col = 1
	frm1.vspdData.Row = 1
                
	With frm1 
       Select Case Trim(frm1.cboBillStatus.value)
            Case  "I"

                .btnApprove.disabled = false
                .btnReceiveReject.disabled = false
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = true       
   
            Case "C"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = false
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = false

            Case "N"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = false
                .btnReject.disabled = false
                .btnPrint.disabled = true
                
            Case "M"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = true

            Case else

                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = false
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
		Call SetToolbar("111000000001111")
	End With
End Function

Function DbQueryNotOk()

    With frm1
       Select Case Trim(frm1.cboBillStatus.value)
            Case  "I"

                .btnApprove.disabled = false
                .btnReceiveReject.disabled = false
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = true       
   
            Case "C"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = false
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = false

            Case "N"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = false
                .btnReject.disabled = false
                .btnPrint.disabled = true
                
            Case "M"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = true

            Case else

                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = false
       End Select      
     End With   

End Function

'========================================================================================================
' Function Name : DbQueryOk2
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk2()
    DIm Row
													     
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
	frm1.vspdData.focus	
	
End Function


'========================================================================================================
' Function Name : DbQueryOk2
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryNotOk2()
    DIm Row
													     
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
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
	Call ExecMyBizASP(frm1, BIZ_PGM_ID4)			' ��: �����Ͻ� ASP �� ���� 
End Function

'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
	
	 ggoSpread.Source = Frm1.vspdData2    
	ggoSpread.ClearSpreadData      
    'Call InitVariables															'��: Initializes local global variables
    
    frm1.txtIsusse.value = ""
    
	Call MainQuery()
End Function


'========================================================================================
Function DbSaveOk2()													'��: ���� ������ ���� ���� 
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
	
	frm1.txtIsusse.value = ""
	
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
       Select Case Trim(frm1.cboBillStatus.value)
            Case  "I"

                .btnApprove.disabled = false
                .btnReceiveReject.disabled = false
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = true       
   
            Case "C"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = false
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = false

            Case "N"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = false
                .btnReject.disabled = false
                .btnPrint.disabled = true
                
            Case "M"
                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = true

            Case else

                .btnApprove.disabled = true
                .btnReceiveReject.disabled = true
                .btnCancelRequest.disabled = true
                .btnAccept.disabled = true
                .btnReject.disabled = true
                .btnPrint.disabled = false
       End Select      
   End With  

End Sub



'���Ž���, ���Űź�, ������ҿ�û, ������ҽ��� (�������� ȣ���� ���ѻ���)
Function  ExeNumOk()
    Dim dti_status 
    Dim signal
    Dim sbdescription
'    Call DisableToolBar(parent.TBC_QUERY)

    dti_status =  Trim(frm1.hdtistatus.value)  
    
    if dti_status = "C" then   '���Ž���
        signal = "APPROVE"
        sbdescription = ""
        
    elseif  dti_status = "R"   then    '���Űź�
        signal = "REJECT"
        sbdescription = Trim(frm1.txtIsusse.value)
        
    elseif  dti_status = "M"   then    '������ҿ�û
        signal = "CANCELREQ"
        sbdescription = Trim(frm1.txtIsusse.value)
    
    elseif  dti_status = "O"   then    '������ҽ��� 
         signal = "SRADAPPRP"
        sbdescription = ""
                
    end if    
    
            
    Call WebControl(Trim(frm1.hbatchid.value), dti_status, sbdescription, signal)
    
End Function


'������Ұź� (�������� ȣ���� ���ѻ���)
Function  ExeNumOk5()
    Dim dti_status 
    Dim signal
    Dim sbdescription
'    Call DisableToolBar(parent.TBC_QUERY)

    dti_status =  Trim(frm1.hdtistatus.value)  
    
    
     dti_status = "C"   
     signal = "REQREJECTP"
     sbdescription = Trim(frm1.txtIsusse.value)
            
                
    Call WebControl(Trim(frm1.hbatchid.value), dti_status, sbdescription, signal)
    
End Function


'��꼭 ����(�������� ȣ���� ���ѻ���)
Function  ExeNumOk6()

 Dim dti_status 
    Dim signal
    Dim sbdescription
'    Call DisableToolBar(parent.TBC_QUERY)

    dti_status =  "CP" 
            

    Call WebControl(Trim(frm1.hbatchid.value), dti_status, sbdescription, signal)
    
End Function



Function  ExeNumNot()

    Call DisplayMsgBox("800407","X","X","X")	'�۾������� �����Դϴ�.
		
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
		'DT4201: �������� ���ݰ�꼭�� ���ÿ� ó���Ҽ� �����ϴ�.
		Exit Function
	end if
    
    Check1 = True
    	
End Function


'����� ID, URL,  ������ üũ
Function Check9()

    Dim strExpDate
    Dim NowDate
    Dim strUrlinfo
    Dim strYear, strMonth, strDay
    Dim IntRetCD,imRow
    Dim iRow
                     
    Check9 = False

         
     If CommonQueryRs(" TOP 1 A.SMART_ID, A.SMART_PASSWORD "," XXSB_DTI_SM_USER A (nolock) ", _
              " A.FND_USER = '" & parent.gUsrID & "'  AND A.FND_REGNO = (SELECT TOP 1 REPLACE(OWN_RGST_NO,'-','') FROM B_TAX_BIZ_AREA WHERE TAX_BIZ_AREA_CD = '" & Trim(frm1.txtBizAreaCd.value) & "')" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then               
     else
        Call DisplayMsgBox("DT4118","X", "X","X")	               		
        'DT4118:���ڼ��ݰ�꼭�� ����ڸ� Ȯ���ϼ���
        Exit Function 	  
     End if
    
      If CommonQueryRs(" TOP 1 Convert(varchar(10),EXPIRATION_DATE,120) as EXPIRATION_DATE "," XXSB_DTI_CERT (nolock) "," CERT_REGNO IN ( SELECT REPLACE(B.BP_RGST_NO,'-','') FROM B_TAX_BIZ_AREA A INNER JOIN B_BIZ_PARTNER B ON (A.TAX_BIZ_AREA_CD = B.BP_CD) where A.TAX_BIZ_AREA_CD = '" & Trim(frm1.txtBizAreaCd.value) & "')" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
        strExpDate =  Trim(Replace(lgF0,Chr(11),""))
        
        if  strExpDate = "" then
            Call DisplayMsgBox("DT4206","X", "X","X")	               		
            'DT4206: �������� Ȯ���ϼ���
            Exit Function 	  
        end if      
      else
         Call DisplayMsgBox("DT4206","X", "X","X")	               		
         'DT4206: �������� Ȯ���ϼ���
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

    Check9 = True

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
												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtIssuedFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��������" class=required></OBJECT>');</SCRIPT> ~
 												<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtIssuedToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��������" class=required></OBJECT>');</SCRIPT>
 											</TD>
 											<TD CLASS="TD5" NOWRAP>����ó��</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSupplierCd" SIZE=15 MAXLENGTH=60 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
												<INPUT TYPE=TEXT AlT="����ó��" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD> 											
										</TR>
										<TR> 											
											<TD CLASS="TD5"NOWRAP>��꼭����</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboBillStatus" ALT="���ݰ�꼭����" CLASS=cboNormal TAG="12" style="width:170px"></SELECT>
											</TD> 											 											
											<TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="13XXXU" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxareaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
												<INPUT TYPE=TEXT AlT="���ݽŰ�����" ID="txtBizAreaNm" NAME="txtBizAreaNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 >
											</TD> 											
										</TR>
										<TR>
										    <TD CLASS="TD5"NOWRAP>��������</TD>
											<TD CLASS="TD6"NOWRAP>
												<SELECT NAME="cboAmendCode" ALT="��������" CLASS=cboNormal TAG="11" style="width:150px"><OPTION VALUE=""></OPTION></SELECT>
											</TD>
											<TD CLASS="TD5"NOWRAP>&nbsp;</TD>
											<TD CLASS="TD6"NOWRAP>&nbsp;</TD>
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
							<TD><BUTTON NAME="btnApprove" CLASS="CLSSBTN" OnClick="VBScript:Call fnApprove()">���Ž���</BUTTON>&nbsp;
                                <BUTTON NAME="btnReceiveReject" CLASS="CLSSBTN" OnClick="VBScript:Call fnReceiveReject()">���Űź�</BUTTON>&nbsp;
							    <BUTTON NAME="btnCancelRequest" CLASS="CLSSBTN" OnClick="VBScript:Call fnCancelRequest()">������� ��û</BUTTON>&nbsp;							    
							    <BUTTON NAME="btnAccept" CLASS="CLSSBTN" OnClick="VBScript:Call fnAccept()">������� ����</BUTTON>&nbsp;							    
							    <BUTTON NAME="btnReject" CLASS="CLSSBTN" OnClick="VBScript:Call fnReject()">������� �ź�</BUTTON>&nbsp;
							    <BUTTON NAME="btnPrint" CLASS="CLSSBTN" OnClick="VBScript:Call fnPrint()">��꼭����</BUTTON>&nbsp							    
							    ���/�źλ���&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtIsusse" SIZE=35 MAXLENGTH=35 STYLE="TEXT-ALIGN: Left" tag="21XXXU" ALT="���/�źλ���"></TD>
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
		<INPUT TYPE=HIDDEN NAME="hdtistatus" tag="24" TABINDEX="-1">
						
		<INPUT TYPE=HIDDEN NAME="hcboBillStatus"   TAG="24"  Tabindex="-1">
        <INPUT TYPE=HIDDEN NAME="htxtSupplierNm"   TAG="24"  Tabindex="-1">
        <INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"   TAG="24"  Tabindex="-1"> 
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
