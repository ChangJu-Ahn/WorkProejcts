<%@ LANGUAGE="VBSCRIPT"%>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m6111ma2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : ���԰����()															*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 																			*
'*  8. Modified date(Last)  : 2003-06-04																*
'*  9. Modifier (First)     : Jin-hyun Shin																*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must cccccchange"								*
'* 13. History              : 																			*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************IV000060************************************************ !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "m6111mb2.asp"												'��: �����Ͻ� ���� ASP�� 

Dim interface_Account

Dim C_posting_flag
Dim C_charge_no
Dim C_glref_pop
Dim C_charge_type
Dim C_charge_type_pop
Dim c_charge_type_nm
Dim C_bp_cd 		
Dim C_bp_cd_pop 	
Dim C_bp_cd_Nm 		

Dim C_BuildCd       
Dim C_BuildCd_pop   
Dim C_Build_Nm      

Dim C_charge_dt 	
Dim C_vat_type 		
Dim C_vat_type_pop 	
Dim C_vat_type_Nm 	
Dim C_tax_biz_area	
Dim C_tax_biz_area_pop
Dim C_tax_biz_area_nm
Dim C_currency 		
Dim C_currency_pop 	
Dim C_charge_doc_amt
Dim C_charge_loc_amt
Dim C_xch_rate	 	
Dim C_Vat_rate 		
Dim C_vat_doc_amt	
Dim C_vat_loc_amt 	
Dim C_pay_type 		
Dim C_pay_type_pop 	
Dim C_pay_type_Nm 	
Dim C_pay_doc_amt
Dim C_pay_loc_amt		'�����ڱ��ݾ� �߰�(2003.08.14)
Dim C_pay_due_dt	
Dim C_charge_rate 	
Dim C_cost_flag 	
Dim C_bank_cd 		
Dim C_bank_pop 		
Dim C_bank_Nm 		
Dim C_bank_acct 	
Dim C_bank_acct_pop 
Dim C_note_no		
Dim C_note_no_pop	
Dim C_prpaym_no		
Dim C_prpaym_no_pop	
Dim C_pp_xch_rt
Dim C_remark 		
Dim C_bas_no		
Dim C_calcd			
Dim C_pay_type_seq4	
Dim C_GlType        
Dim C_GlNo          
Dim C_old_posting_flg

'Dim lgBlnFlgChgValue    '������ ���� 
'Dim lgIntGrpCount       'data count
'Dim lgIntFlgMode        '�ű�,����,���� mode �� 

'Dim lgStrPrevKey        '���� data key �� 
'Dim lgLngCurRows
Dim gblnWinEvent	    '�̺�Ʈ ���� 
'Dim lgSortKey

Dim IsOpenPop          
Dim gChangeOpt
Dim arrCollectType

Dim iDBSYSDate
Dim EndDate, StartDate

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)
'==============================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE   
    lgBlnFlgChgValue = False    
    lgIntGrpCount = 0           
    
    lgStrPrevKey = ""           
    lgLngCurRows = 0       
    lgSortKey    = 1      
End Sub
'==============================================================================================================================
Sub initSpreadPosVariables()  
	C_posting_flag 		= 1        'Ȯ�� 
	C_charge_no 		= 2        '��������ȣ 
	C_glref_pop         = 3        '��ǥ��ȸ �˾� 
	C_charge_type 		= 4        '����׸� 
	C_charge_type_pop 	= 5        '����׸� �˾� 
	c_charge_type_nm	= 6        '����׸�� 
	C_bp_cd 			= 7        '����ó 
	C_bp_cd_pop 		= 8        '����ó �˾� 
	C_bp_cd_Nm 			= 9        '����ó�� 

	C_BuildCd           = 10       '��꼭����ó 
	C_BuildCd_pop       = 11       '��꼭����ó �˾� 
	C_Build_Nm          = 12       '��꼭����ó�� 

	C_charge_dt 		= 13       '�߻��� 
	C_vat_type 			= 14       'VAT
	C_vat_type_pop 		= 15       'VAT �˾� 
	C_vat_type_Nm 		= 16       'VAT�� 
	C_tax_biz_area		= 17       '���ݽŰ����� 
	C_tax_biz_area_pop	= 18       '���ݽŰ����� �˾� 
	C_tax_biz_area_nm	= 19       '���ݽŰ������ 
	C_currency 			= 20       'ȭ�� 
	C_currency_pop 		= 21       'ȭ�� �˾� 
	C_charge_doc_amt 	= 22       '�߻��ݾ� 
	C_charge_loc_amt 	= 23       '�߻��ڱ��ݾ� 
	C_xch_rate	 		= 24       'ȯ�� 
	C_Vat_rate 			= 25       'VAT�� 
	C_vat_doc_amt		= 26       'VAT�ݾ� 
	C_vat_loc_amt 		= 27       'VAT�ڱ��ݾ� 
	C_pay_type 			= 28       '�������� 
	C_pay_type_pop 		= 29       '�������� �˾� 
	C_pay_type_Nm 		= 30       '���������� 
	C_pay_doc_amt		= 31       '���ޱݾ� 
	C_pay_loc_amt		= 32       '�����ڱ��ݾ� �߰�(2003.08.14)
	C_pay_due_dt		= 33       '������ 
	C_charge_rate 		= 34       '�����(%)
	C_cost_flag 		= 35       '�������Կ��� 
	C_bank_cd 			= 36       '������� 
	C_bank_pop 			= 37       '������� �˾� 
	C_bank_Nm 			= 38       '�������� 
	C_bank_acct 		= 39       '��ݰ��� 
	C_bank_acct_pop 	= 40       '��ݰ��� �˾� 
	C_note_no			= 41       '������ȣ 
	C_note_no_pop		= 42       '������ȣ �˾� 
	C_prpaym_no			= 43       '���ޱݹ�ȣ 
	C_prpaym_no_pop		= 44       '���ޱݹ�ȣ �˾� 
	C_pp_xch_rt			= 45	   '���ޱ�ȯ�� 
	C_remark 			= 46       '��Ÿ���� 
	C_bas_no			= 47       '�߻��ٰ� ������ȣ 
	C_calcd				= 48       'calcd(hidden)	�ݾװ��� �ʿ��� * or / �ڵ� 
	C_pay_type_seq4		= 49       'seq4 (hidden)
	C_GlType            = 50       '��ǥ type
	C_GlNo              = 51       '��ǥ��ȣ 
	C_old_posting_flg	= 52       'oldpostingflg(hidden)

End Sub
'==============================================================================================================================
Sub SetDefaultVal()
	Call SetToolBar("1110110100101111")

	frm1.txtChargeFrDt.Text	= StartDate
	frm1.txtChargeToDt.Text	= EndDate
    
	interface_Account = GetSetupMod(parent.gSetupMod, "a")
    frm1.txtprocess_step.focus                        '���౸�� 
	
    Set gActiveElement = document.activeElement
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==============================================================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()   
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	'patch version
    ggoSpread.Spreadinit "V20031009",,parent.gAllowDragDropSpread 

	.ReDraw = false
		
    .MaxCols = C_old_posting_flg + 1	
    .MaxRows = 0
    
    Call GetSpreadColumnPos("A")
    
    ggoSpread.SSSetCheck 	C_posting_flag		, "Ȯ��", 15,,,True
    ggoSpread.SSSetEdit 	C_charge_no			, "��������ȣ", 18,,,18,2
    ggoSpread.SSSetButton 	C_glref_pop
    ggoSpread.SSSetEdit 	C_charge_type		, "����׸�", 20,,,20,2
    ggoSpread.SSSetButton 	C_charge_type_pop
    ggoSpread.SSSetEdit 	C_charge_type_nm	, "����׸��", 20
    ggoSpread.SSSetEdit 	C_bp_cd				, "����ó", 10,,,10,2
    ggoSpread.SSSetButton 	C_bp_cd_pop
    ggoSpread.SSSetEdit 	C_bp_cd_Nm			, "����ó��", 20
    
    ggoSpread.SSSetEdit 	C_BuildCd			, "���ݰ�꼭����ó", 18,,,10,2
    ggoSpread.SSSetButton 	C_BuildCd_pop
    ggoSpread.SSSetEdit 	C_Build_Nm			, "���ݰ�꼭����ó��", 20
    
    ggoSpread.SSSetDate 	C_charge_dt			, "�߻���", 10, 2, parent.gDateFormat
    ggoSpread.SSSetEdit 	C_vat_type			, "VAT", 10,,,5,2
    ggoSpread.SSSetButton 	C_vat_type_pop
    ggoSpread.SSSetEdit 	C_vat_type_Nm		, "VAT��", 20
	ggoSpread.SSSetEdit 	C_tax_biz_area		, "���ݽŰ�����", 20,,,10,2
	ggoSpread.SSSetButton 	C_tax_biz_area_pop
	ggoSpread.SSSetEdit 	C_tax_biz_area_nm	, "���ݽŰ������", 20
    ggoSpread.SSSetEdit 	C_currency			, "ȭ��", 8,,,3,2
    ggoSpread.SSSetButton 	C_currency_pop
    ggoSpread.SSSetFloat    C_charge_doc_amt	, "�߻��ݾ�"		, 15    ,"A"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
    SetSpreadFloatLocal 	C_charge_loc_amt	, "�߻��ڱ��ݾ�"	, 15, 1, 2
    ggoSpread.SSSetFloat	C_xch_rate			, "ȯ��"				, 10, parent.ggExchRateNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, 1,,"Z"
    ggoSpread.SSSetFloat    C_Vat_rate			, "VAT��"			, 10    ,"D"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
    ggoSpread.SSSetFloat    C_vat_doc_amt		, "VAT�ݾ�"			, 15    ,"A"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
    SetSpreadFloatLocal 	C_vat_loc_amt		, "VAT�ڱ��ݾ�"		, 15, 1, 2
    ggoSpread.SSSetEdit 	C_pay_type			, "��������"		, 10,,,5,2
    ggoSpread.SSSetButton 	C_pay_type_pop
    ggoSpread.SSSetEdit 	C_pay_type_Nm		, "����������"		, 20
    ggoSpread.SSSetDate 	C_pay_due_dt		, "������"			, 10, 2, parent.gDateFormat
	ggoSpread.SSSetFloat    C_pay_doc_amt		, "���ޱݾ�"		, 15    ,"A"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
	'�����ڱ��ݾ� �߰�(2003.08.14)
	SetSpreadFloatLocal 	C_pay_loc_amt		, "�����ڱ��ݾ�"	, 15, 1, 2
	ggoSpread.SSSetFloat    C_charge_rate		, "�����(%)"		, 15    ,"D"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
  	ggoSpread.SSSetCheck 	C_cost_flag			, "�������Կ���"	, 15,,,True
    ggoSpread.SSSetEdit 	C_bank_acct			, "��ݰ���"		, 20,,,30,2
    ggoSpread.SSSetButton 	C_bank_acct_pop
    ggoSpread.SSSetEdit 	C_bank_cd			, "�������"		, 10,,,10,2
    ggoSpread.SSSetButton 	C_bank_pop
    ggoSpread.SSSetEdit 	C_bank_Nm			, "��������"		, 20
    ggoSpread.SSSetEdit 	C_note_no			, "������ȣ/��ǥ��ȣ", 20,,,30,2
    ggoSpread.SSSetButton 	C_note_no_pop
    ggoSpread.SSSetEdit 	C_prpaym_no			, "���ޱݹ�ȣ"		, 20,,,30,2
    ggoSpread.SSSetButton 	C_prpaym_no_pop
    ggoSpread.SSSetFloat    C_pp_xch_rt			, "���ޱ�ȯ��"			, 10    ,parent.ggExchRateNo  ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
    ggoSpread.SSSetEdit 	C_remark			, "��Ÿ����"		, 20,,,50
	ggoSpread.SSSetEdit 	C_bas_no			, "�߻��ٰ� ������ȣ", 20
    ggoSpread.SSSetEdit 	C_calcd				, "XCH_RATE_OP"			, 10		'�ݾװ��� �ʿ��� * or / �ڵ� 
    ggoSpread.SSSetEdit 	C_pay_type_seq4		, "seq4", 10
    ggoSpread.SSSetEdit 	C_GlType			, "C_GlType"			, 10
    ggoSpread.SSSetEdit 	C_GlNo				, "C_GlNo"				, 10
    ggoSpread.SSSetEdit 	C_old_posting_flg	, "oldpostingflg"		, 10
  
    Call ggoSpread.MakePairsColumn(C_charge_no,C_glref_pop)
    Call ggoSpread.MakePairsColumn(C_charge_type,C_charge_type_pop)
    Call ggoSpread.MakePairsColumn(C_bp_cd,C_bp_cd_pop)
    Call ggoSpread.MakePairsColumn(C_BuildCd,C_BuildCd_pop)
    Call ggoSpread.MakePairsColumn(C_vat_type,C_vat_type_pop)
    Call ggoSpread.MakePairsColumn(C_tax_biz_area,C_tax_biz_area_pop)
    Call ggoSpread.MakePairsColumn(C_currency,C_currency_pop)
    Call ggoSpread.MakePairsColumn(C_pay_type,C_pay_type_pop)
    Call ggoSpread.MakePairsColumn(C_bank_acct,C_bank_acct_pop)
    Call ggoSpread.MakePairsColumn(C_bank_cd,C_bank_pop)
    Call ggoSpread.MakePairsColumn(C_note_no,C_note_no_pop)
    Call ggoSpread.MakePairsColumn(C_prpaym_no,C_prpaym_no_pop)

    Call ggoSpread.SSSetColHidden(C_calcd,C_old_posting_flg,True)
    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    If interface_Account = "N" Then
		Call ggoSpread.SSSetColHidden(C_posting_flag,C_posting_flag,True)		'Ȯ�� 
		Call ggoSpread.SSSetColHidden(C_bank_acct_pop,C_bank_acct_pop,True)		'��ݰ��� �˾� 
		Call ggoSpread.SSSetColHidden(C_note_no_pop,C_note_no_pop,True)			'������ȣ �˾� 
		Call ggoSpread.SSSetColHidden(C_pay_type_seq4,C_pay_type_seq4,True)
	End If
  
	.ReDraw = true
	
    End With
End Sub
'==============================================================================================================================
Sub SetSpreadLock()
    
	ggoSpread.spreadUnlock 		C_posting_flag, -1                             'Ȯ�� 
	ggoSpread.spreadUnlock 		C_charge_no, -1, C_charge_no                   '��������ȣ 
	ggoSpread.spreadUnlock 		C_charge_type, -1,C_charge_type, -1            '����׸� 
	ggoSpread.SSSetRequired		C_charge_type,-1                               
	ggoSpread.spreadUnlock 		C_charge_type_pop, -1,C_charge_type_pop, -1    '����׸� �˾� 
	ggoSpread.SpreadLock 		C_charge_type_nm, -1,C_bp_cd_Nm, -1            '����׸��,����ó�� 
	ggoSpread.SSSetProtected	C_charge_type_nm, -1                           
	ggoSpread.spreadUnlock 		C_bp_cd, -1,C_bp_cd, -1                        '����ó 
	ggoSpread.SSSetRequired		C_bp_cd,-1                                    
	ggoSpread.spreadUnlock 		C_bp_cd_pop, -1,C_bp_cd_pop, -1                '����ó �˾� 
	ggoSpread.SpreadLock 		C_bp_cd_Nm, -1,C_bp_cd_Nm, -1                  '����ó�� 
	ggoSpread.SSSetProtected	C_bp_cd_Nm, -1                                 
    
	ggoSpread.spreadUnlock 		C_BuildCd, -1,C_BuildCd, -1                    '��꼭����ó 
	ggoSpread.spreadUnlock 		C_BuildCd_pop, -1,C_BuildCd_pop, -1            '��꼭����ó �˾� 
	ggoSpread.SSSetProtected	C_Build_Nm, -1                                 '��꼭����ó�� 
    
	ggoSpread.spreadUnlock 		C_charge_dt, -1,C_charge_dt, -1                '�߻��� 
	ggoSpread.SSSetRequired		C_charge_dt,-1                                 
	ggoSpread.spreadUnlock 		C_vat_type, -1,C_vat_type, -1                  'VAT
	ggoSpread.spreadUnlock 		C_vat_type_pop, -1,C_vat_type_pop, -1          'VAT �˾� 
	ggoSpread.SpreadLock 		C_vat_type_Nm, -1,C_vat_type_Nm, -1            'VAT�� 
	ggoSpread.SSSetProtected	C_vat_type_Nm, -1                              
	ggoSpread.spreadUnlock 		C_tax_biz_area, -1,C_tax_biz_area, -1          '���ݽŰ����� 
	ggoSpread.spreadUnlock 		C_tax_biz_area_pop, -1,C_tax_biz_area_pop, -1  '���ݽŰ����� �˾� 
	ggoSpread.spreadlock 		C_tax_biz_area_nm, -1,C_tax_biz_area_nm, -1    '���ݽŰ������ 
	ggoSpread.SSSetProtected	C_tax_biz_area_nm, -1                          
	ggoSpread.spreadUnlock 		C_currency, -1,C_currency, -1                  'ȭ�� 
	ggoSpread.SSSetRequired		C_currency,-1                                  
	ggoSpread.spreadUnlock 		C_currency_pop, -1,C_currency_pop, -1          'ȭ�� �˾� 
	ggoSpread.spreadUnlock 		C_charge_doc_amt, -1,C_charge_doc_amt, -1      '�߻��ݾ� 
	ggoSpread.SSSetRequired		C_charge_doc_amt,-1                            
	ggoSpread.spreadUnlock 		C_xch_rate, -1,C_xch_rate, -1                  'ȯ�� 
	ggoSpread.SpreadLock 		C_Vat_rate, -1,C_Vat_rate, -1                  'VAT�� 
	ggoSpread.SSSetProtected	C_Vat_rate, -1                                 
	ggoSpread.SpreadLock 		C_vat_loc_amt, -1,C_vat_loc_amt, -1            'VAT�ڱ��ݾ� 
	ggoSpread.spreadUnlock 		C_pay_type, -1,C_pay_type, -1                  '�������� 
	ggoSpread.spreadUnlock 		C_pay_type_pop, -1,C_pay_type_pop, -1          '���������˾� 
	ggoSpread.SpreadLock 		C_pay_type_Nm, -1,C_pay_type_Nm, -1            '���������� 
	ggoSpread.SSSetProtected	C_pay_type_Nm, -1                              
	ggoSpread.spreadUnlock 		C_pay_due_dt, -1,C_pay_due_dt, -1              '������ 
	ggoSpread.spreadUnlock 		C_pay_doc_amt, -1,C_pay_doc_amt, -1            '���ޱݾ� 
	
	ggoSpread.SpreadLock 		C_pay_loc_amt, -1,C_pay_loc_amt, -1            '�����ڱ��ݾ� (2003.08.14)
	ggoSpread.SSSetProtected	C_pay_loc_amt, -1  
	ggoSpread.SpreadLock 		C_pp_xch_rt, -1,C_pp_xch_rt, -1					'���ޱ�ȯ�� (2003.08.14)
	ggoSpread.SSSetProtected	C_pp_xch_rt, -1  
	
	ggoSpread.spreadUnlock 		C_charge_rate, -1,C_charge_rate, -1            '�����(%)
	ggoSpread.spreadUnlock 		C_cost_flag, -1,C_cost_flag, -1                '�������Կ��� 
	ggoSpread.spreadUnlock 		C_bank_acct, -1,C_bank_acct, -1                '��ݰ��� 
	ggoSpread.spreadUnlock 		C_bank_acct_pop, -1,C_bank_acct_pop, -1        '��ݰ��� �˾� 
	ggoSpread.spreadUnlock 		C_bank_cd, -1,C_bank_cd, -1                    '������� 
	ggoSpread.spreadUnlock 		C_bank_pop, -1,C_bank_pop, -1                  '������� �˾� 
	ggoSpread.SpreadLock 		C_bank_Nm, -1,C_bank_Nm, -1                    '�������� 
	ggoSpread.SSSetProtected	C_bank_Nm, -1                                  
	ggoSpread.spreadUnlock 		C_remark, -1,C_remark, -1                      '��Ÿ���� 
	ggoSpread.spreadUnlock 		C_old_posting_flg, -1,C_old_posting_flg, -1    'hidden
	ggoSpread.SSSetProtected	C_old_posting_flg + 1, -1                                  
    
End Sub
'==============================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    frm1.vspdData.ReDraw = False
	ggoSpread.SSSetRequired			C_charge_type		, pvStartRow, pvEndRow      '����׸� 
	ggoSpread.SSSetProtected	    C_charge_type_nm	, pvStartRow, pvEndRow		'����׸�� 
	ggoSpread.SSSetRequired			C_bp_cd				, pvStartRow, pvEndRow      '����ó 
	ggoSpread.SSSetProtected		C_bp_cd_Nm			, pvStartRow, pvEndRow      '����ó�� 
	ggoSpread.SSSetProtected		C_Build_Nm			, pvStartRow, pvEndRow      '��꼭����ó�� 
	ggoSpread.SSSetRequired			C_charge_dt			, pvStartRow, pvEndRow      '�߻��� 
	ggoSpread.SSSetProtected		C_vat_type_nm		, pvStartRow, pvEndRow      'VAT�� 
	ggoSpread.SSSetProtected		C_tax_biz_area_nm	, pvStartRow, pvEndRow      '���ݽŰ������ 
	ggoSpread.SSSetRequired			C_currency			, pvStartRow, pvEndRow      'ȭ�� 
	ggoSpread.SSSetRequired			C_charge_doc_amt	, pvStartRow, pvEndRow      '�߻��ݾ� 
	ggoSpread.SSSetProtected		C_charge_loc_amt	, pvStartRow, pvEndRow      '�߻��ڱ��ݾ� 
	ggoSpread.SSSetProtected		C_vat_rate			, pvStartRow, pvEndRow      'VAT �� 
	ggoSpread.SSSetProtected		C_vat_doc_amt		, pvStartRow, pvEndRow      'VAT �ݾ� 
	'#####
	'ggoSpread.SSSetRequired			C_vat_loc_amt		, pvStartRow, pvEndRow      'VAT �ڱ��ݾ� 
	ggoSpread.SSSetProtected		C_vat_loc_amt		, pvStartRow, pvEndRow      'VAT �ڱ��ݾ� 
	ggoSpread.SSSetProtected		C_bank_cd			, pvStartRow, pvEndRow      '������� 
	ggoSpread.SSSetProtected		C_bank_pop			, pvStartRow, pvEndRow      '������� �˾� 
	ggoSpread.SSSetProtected		C_bank_acct			, pvStartRow, pvEndRow      '��ݰ��� 
	ggoSpread.SSSetProtected		C_bank_acct_pop		, pvStartRow, pvEndRow      '��ݰ��� �˾� 
	ggoSpread.SSSetProtected		C_note_no			, pvStartRow, pvEndRow      '������ȣ 
	ggoSpread.SSSetProtected		C_note_no_pop		, pvStartRow, pvEndRow      '������ȭ �˾� 
	ggoSpread.SSSetProtected		C_prpaym_no			, pvStartRow, pvEndRow      '���ޱݹ�ȣ 
	ggoSpread.SSSetProtected		C_prpaym_no_pop		, pvStartRow, pvEndRow      '���ޱݹ�ȣ �˾� 
	ggoSpread.SSSetProtected		C_pp_xch_rt			, pvStartRow, pvEndRow      '���ޱ�ȯ�� 
	ggoSpread.SSSetProtected		C_pay_type_nm		, pvStartRow, pvEndRow      '���������� 
	ggoSpread.SSSetProtected		C_bank_nm			, pvStartRow, pvEndRow      '�������� 
	ggoSpread.SSSetProtected		C_old_posting_flg + 1	, pvStartRow, pvEndRow      'MaxCol
    
	frm1.vspdData.ReDraw = True
End Sub
'==============================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_posting_flag 		= iCurColumnPos(1)
			C_charge_no 		= iCurColumnPos(2)
			C_glref_pop         = iCurColumnPos(3)
			C_charge_type 		= iCurColumnPos(4)
			C_charge_type_pop 	= iCurColumnPos(5)
			c_charge_type_nm	= iCurColumnPos(6)
			C_bp_cd 			= iCurColumnPos(7)
			C_bp_cd_pop 		= iCurColumnPos(8)
			C_bp_cd_Nm 			= iCurColumnPos(9)
			C_BuildCd           = iCurColumnPos(10)
			C_BuildCd_pop       = iCurColumnPos(11)
			C_Build_Nm          = iCurColumnPos(12)
			C_charge_dt 		= iCurColumnPos(13)
			C_vat_type 			= iCurColumnPos(14)
			C_vat_type_pop 		= iCurColumnPos(15)
			C_vat_type_Nm 		= iCurColumnPos(16)
			C_tax_biz_area		= iCurColumnPos(17)
			C_tax_biz_area_pop	= iCurColumnPos(18)
			C_tax_biz_area_nm	= iCurColumnPos(19)
			C_currency 			= iCurColumnPos(20)
			C_currency_pop 		= iCurColumnPos(21)
			C_charge_doc_amt 	= iCurColumnPos(22)
			C_charge_loc_amt 	= iCurColumnPos(23)
			C_xch_rate	 		= iCurColumnPos(24)
			C_Vat_rate 			= iCurColumnPos(25)
			C_vat_doc_amt		= iCurColumnPos(26)
			C_vat_loc_amt 		= iCurColumnPos(27)
			C_pay_type 			= iCurColumnPos(28)
			C_pay_type_pop 		= iCurColumnPos(29)
			C_pay_type_Nm 		= iCurColumnPos(30)
			C_pay_doc_amt		= iCurColumnPos(31)
			C_pay_loc_amt		= iCurColumnPos(32)	'�����ڱ��ݾ� �߰� (2003.08.14)
			C_pay_due_dt		= iCurColumnPos(33)
			C_charge_rate 		= iCurColumnPos(34)
			C_cost_flag 		= iCurColumnPos(35)
			C_bank_cd 			= iCurColumnPos(36)
			C_bank_pop 			= iCurColumnPos(37)
			C_bank_Nm 			= iCurColumnPos(38)
			C_bank_acct 		= iCurColumnPos(39)
			C_bank_acct_pop 	= iCurColumnPos(40)
			C_note_no			= iCurColumnPos(41)
			C_note_no_pop		= iCurColumnPos(42)
			C_prpaym_no			= iCurColumnPos(43)
			C_prpaym_no_pop		= iCurColumnPos(44)
			C_pp_xch_rt			= iCurColumnPos(45)
			C_remark 			= iCurColumnPos(46)
			C_bas_no			= iCurColumnPos(47)
			C_calcd				= iCurColumnPos(48)
			C_pay_type_seq4		= iCurColumnPos(49)
			C_GlType            = iCurColumnPos(50)
			C_GlNo              = iCurColumnPos(51)
			C_old_posting_flg	= iCurColumnPos(52)

    End Select    
End Sub
'==============================================================================================================================
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "���ݽŰ�����"	
	arrParam(1) = "B_Tax_Biz_Area"
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_tax_biz_area
	arrParam(2) = Trim(frm1.vspdData.text)
	
	'arrParam(4) = "Tax_Flag = 'Y'"
	arrParam(5) = "���ݽŰ�����"			
	
    arrField(0) = "Tax_Biz_Area_Cd"
    arrField(1) = "Tax_Biz_Area_Nm"
    
    arrHeader(0) = "���ݽŰ�����"
    arrHeader(1) = "���ݽŰ������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_tax_biz_area,		frm1.vspdData.ActiveRow, arrRet(0))
		Call frm1.vspdData.SetText(C_tax_biz_area_nm,	frm1.vspdData.ActiveRow, arrRet(1))
		Call vspdData_Change(C_tax_biz_area , frm1.vspdData.Row)
		lgBlnFlgChgValue = True
	End If	
End Function
'==============================================================================================================================
Function Openprocess_step()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���౸��"					
	arrParam(1) = "B_minor"						
	arrParam(2) = frm1.txtprocess_step.value	
	arrParam(3) = ""							
	arrParam(4) = "major_cd=" & FilterVar("M9014", "''", "S") & ""			
	arrParam(5) = "���౸��"			
	
    arrField(0) = "minor_cd"					
    arrField(1) = "minor_nm"					
    
    arrHeader(0) = "���౸��"				
    arrHeader(1) = "���౸�и�"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtprocess_step.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtprocess_step.value		= arrRet(0)
		frm1.txtprocess_stepNm.value	= arrRet(1)
		frm1.txtbas_no.value = ""          '�߻��ٰ� ������ȣ "" setting
		frm1.txtprocess_step.focus
		Set gActiveElement = document.activeElement
	End If	

End Function
'==============================================================================================================================
Function Openprocess_step1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtprocess_step1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���౸��"					
	arrParam(1) = "B_minor"						
	arrParam(2) = frm1.txtprocess_step1.value	
	arrParam(3) = ""							
	arrParam(4) = "major_cd=" & FilterVar("M9014", "''", "S") & ""			
	arrParam(5) = "���౸��"			
	
    arrField(0) = "minor_cd"					
    arrField(1) = "minor_nm"					
    
    arrHeader(0) = "���౸��"				
    arrHeader(1) = "���౸�и�"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtprocess_step1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtprocess_step1.value		= arrRet(0)
		frm1.txtprocess_stepNm1.value	= arrRet(1)
		lgBlnFlgChgValue = True 
		frm1.txtbas_no1.value = ""           '�߻��ٰ� ������ȣ "" setting
		frm1.txtprocess_step1.focus
		Set gActiveElement = document.activeElement
	End If	

End Function
'==============================================================================================================================
Function OpenBasNoPop(ByVal strPath)
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
		
    '�߻��ٰ� ������ȣ�� �����Ұ��� �Լ��� ���� ������ 
	If gblnWinEvent = True Or UCase(frm1.txtBas_No.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
		
	iCalledAspName = AskPRAspName(strPath)
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, strPath, "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	gblnWinEvent = False
		
	if UCase(Trim(frm1.txtprocess_step.value)) = "PO" then   '���౸�� 
			
		If strRet(0) = "" Then
			Call SetBasNo("BasNo")
			Exit Function
		Else
			frm1.txtbas_no.value = strRet(0)                 '�߻��ٰŹ�ȣ 
			Call SetBasNo("BasNo")
		End If
	else
		If strRet = "" Then
			Call SetBasNo("BasNo")
			Exit Function
		Else
			frm1.txtbas_no.value = strRet
			Call SetBasNo("BasNo")
		End If

	End if

End Function
'==============================================================================================================================
Function OpenBasNoPop1(ByVal strPath)
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
		
	If gblnWinEvent = True Or UCase(frm1.txtBas_No.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
		
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
		
	iCalledAspName = AskPRAspName(strPath)
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, strPath, "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	if UCase(Trim(frm1.txtprocess_step1.value)) = "PO" then
			
		If strRet(0) = "" Then
			Call SetBasNo("BasNo1")
			Exit Function
		Else
			frm1.txtbas_no1.value = strRet(0)
			Call SetBasNo("BasNo1")
		End If
	else
		If strRet = "" Then
			Call SetBasNo("BasNo1")
			Exit Function
		Else
			frm1.txtbas_no1.value = strRet
			lgBlnFlgChgValue = True
			Call SetBasNo("BasNo1")
			
		End If
	End if
End Function
'==============================================================================================================================
Function SetBasNo(ByVal sTag)
	If sTag = "BasNo" Then
		frm1.txtbas_no.focus
	Else	
		frm1.txtbas_no1.focus
	End If
	Set gActiveElement = document.activeElement	
End Function
'==============================================================================================================================
Function ChangeBp(ByVal strBpCd, ByVal Row)
	Dim strVal
    
    Err.Clear                                   

	If CheckRunningBizProcess = True Then
		Exit Function
	End If	
    
    if Trim(GetSpreadText(frm1.vspdData,C_bp_cd, Row,"X","X")) = "" then
    	Exit Function
    End if
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"
    strVal = strVal & "&txtBpCd=" & Trim(strBpCd)
    
    If  LayerShowHide(1) = False Then
      	Exit Function
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)						
	
End Function
'==============================================================================================================================
function changeProcess_step()   '���౸�� setting�� �߻��ٰŹ�ȣ�� "" setting
	frm1.txtbas_no.value = ""
end function 
'==============================================================================================================================
function changeProcess_step1()
	frm1.txtbas_no1.value = ""
end function 
'==============================================================================================================================
Function Openpur_grp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"			
	arrParam(1) = "B_pur_grp"			
	arrParam(2) = frm1.txtpur_grp.value	
	arrParam(3) = ""					
	arrParam(4) = ""					
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "pur_grp"				
    arrField(1) = "pur_grp_nm"			
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtpur_grp.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtpur_grp.value	= arrRet(0)
		frm1.txtpur_grpNm.value	= arrRet(1)
		frm1.txtpur_grp.focus
		Set gActiveElement = document.activeElement
	End If	

End Function
'==============================================================================================================================
Function Openpur_grp1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtpur_grp1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"			
	arrParam(1) = "B_pur_grp"			
	arrParam(2) = frm1.txtpur_grp1.value     	
	arrParam(3) = ""							
	arrParam(4) = ""							
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "pur_grp"					
    arrField(1) = "pur_grp_nm"				
    
    arrHeader(0) = "���ű׷�"			
    arrHeader(1) = "���ű׷��"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtpur_grp1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtpur_grp1.value	= arrRet(0)
		frm1.txtpur_grpNm1.value	= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtpur_grp1.focus
		Set gActiveElement = document.activeElement
	End If	

End Function
'==============================================================================================================================
Function OpenChargeType()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����׸�"			    
	arrParam(1) = "A_JNL_ITEM,b_trade_charge"	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_charge_type,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(3) = ""							
	arrParam(4) = "b_trade_charge.charge_cd=A_JNL_ITEM.JNL_CD And A_JNL_ITEM.JNL_TYPE=" & FilterVar("EC", "''", "S") & " and b_trade_charge.module_type=" & FilterVar("M", "''", "S") & " "
	arrParam(5) = "����׸�"			
		
	arrField(0) = "A_JNL_ITEM.JNL_CD"			
	arrField(1) = "A_JNL_ITEM.JNL_NM"			
	    
	arrHeader(0) = "����׸�"				
	arrHeader(1) = "����׸��"				
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_charge_type,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(C_charge_type_nm,	frm1.vspdData.ActiveRow,	arrRet(1))
		Call vspdData_Change(C_charge_type , frm1.vspdData.Row)
	End If	
End Function
'==============================================================================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ȭ�����"					
	arrParam(1) = "B_CURRENCY"		    	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_currency,frm1.vspdData.ActiveRow,"X","X"))          
	arrParam(3) = ""						  
	arrParam(4) = ""						  
	arrParam(5) = "ȭ�����"			
	
    arrField(0) = "CURRENCY"				  
    arrField(1) = "CURRENCY_DESC"			
    
    arrHeader(0) = "ȭ�����"			
    arrHeader(1) = "ȭ�������"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_currency, frm1.vspdData.ActiveRow,	arrRet(0))
		Call vspdData_Change(C_currency, frm1.vspdData.ActiveRow)
	End If	

End Function
'==============================================================================================================================
Function OpenPay_Type()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"
	arrParam(1) = "B_CONFIGURATION Config, B_MINOR Minor"
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_pay_type,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(3) = ""
	arrParam(4) = "Config.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " AND Config.SEQ_NO = " & FilterVar("1", "''", "S") & "  " _
				& "AND Config.MINOR_CD = Minor.MINOR_CD AND Config.MAJOR_CD = Minor.MAJOR_CD " _
				& "AND Config.REFERENCE IN(" & FilterVar("RP", "''", "S") & "," & FilterVar("P", "''", "S") & " )"
	arrParam(5) = "��������"			
	
	arrField(0) = "Config.MINOR_CD"			
	arrField(1) = "Minor.MINOR_NM"			
    
    arrHeader(0) = "��������"
    arrHeader(1) = "����������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_pay_type,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(C_pay_type_nm,	frm1.vspdData.ActiveRow,	arrRet(1))
		Call vspdData_Change(C_pay_type , frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenVat_Type()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT����"				
	arrParam(1) = "B_minor,b_configuration"	
    arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_vat_type,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(3) = ""						
	arrParam(4) = "b_minor.MAJOR_CD=" & FilterVar("b9001", "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd "				
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT����"			
	
    arrField(0) = "B_minor.minor_cd"   				         	
    arrField(1) = "B_minor.minor_nm"					        
    arrField(2) = "b_configuration.REFERENCE"					
    
    arrHeader(0) = "VAT����"				
    arrHeader(1) = "VAT������"			
    arrHeader(2) = "VAT��"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_vat_type,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(C_vat_type_Nm,	frm1.vspdData.ActiveRow,	arrRet(1))
		Call frm1.vspdData.SetText(C_Vat_rate,		frm1.vspdData.ActiveRow,	arrRet(2))	
		Call vspdData_Change(C_vat_type , frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenBank()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�������"					
	arrParam(1) = "F_DPST,B_Bank"		    	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X"))    
	arrParam(3) = ""						    
	arrParam(4) = "F_DPST.BANK_CD = B_BANK.BANK_CD"
	arrParam(5) = "�������"			
	
    arrField(0) = "F_DPST.bank_cd"				   
    arrField(1) = "B_Bank.bank_nm"					
    
    arrHeader(0) = "�������"			     	
    arrHeader(1) = "��������"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_Bank_cd,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(C_Bank_nm,		frm1.vspdData.ActiveRow,	arrRet(1))
		Call vspdData_Change(C_Bank_cd , frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenBp_Cd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"					
	arrParam(1) = "B_BIZ_PARTNER"		    	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X"))           
	arrParam(3) = ""						    
	'arrParam(4) = ""
	'����ó�̰ų� ����/����ó�� �ŷ�ó�� ��ȸ��(2003.09.19)	
	arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "					    
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"				        
    arrField(1) = "bp_nm"					    
    
    arrHeader(0) = "����ó"			     	
    arrHeader(1) = "����ó��"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_bp_cd,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(C_bp_cd_nm,	frm1.vspdData.ActiveRow,	arrRet(1))
		Call vspdData_Change(C_bp_cd , frm1.vspdData.Row)
		'Call ChangeBp()
	End If	

End Function
'==============================================================================================================================
Function OpenBank_Acct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	arrParam(0) = "��ݰ���"					
	arrParam(1) = "F_DPST,B_BANK, B_BANK_ACCT"		    	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_bank_acct,frm1.vspdData.ActiveRow,"X","X"))           
	arrParam(3) = ""						    
	arrParam(4) = "F_DPST.Bank_Cd = B_BANK.Bank_Cd AND F_DPST.BANK_ACCT_NO = B_BANK_ACCT.BANK_ACCT_NO AND  F_DPST.BANK_CD = B_BANK_ACCT.BANK_CD"
	
	if Trim(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X")) <> "" then
		arrParam(4) = arrParam(4) & " And B_BANK.Bank_Cd =  " & FilterVar(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & " "
	end if
	
	if Trim(GetSpreadText(frm1.vspdData,C_currency,frm1.vspdData.ActiveRow,"X","X")) <> "" then
		arrParam(4) = arrParam(4) & " And F_DPST.DOC_CUR =  " & FilterVar(GetSpreadText(frm1.vspdData,C_currency,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & " "
	end if
	
	arrParam(4) = arrParam(4) & " AND (F_DPST.DPST_FG = " & FilterVar("SV", "''", "S") & " OR F_DPST.DPST_FG = " & FilterVar("ET", "''", "S") & ") " '����, ��Ÿ 
	arrParam(5) = "��ݰ���"			
	
    arrField(0) = "F_DPST.BANK_ACCT_NO"				
    arrField(1) = "B_BANK.BANK_CD"					
    arrField(2) = "B_BANK.BANK_NM"					
    
    arrHeader(0) = "��ݰ���"			     	
    arrHeader(1) = "�������"			     	
    arrHeader(2) = "��������"			     	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_bank_acct,		frm1.vspdData.ActiveRow,	arrRet(0))	
		
		If Trim(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X")) = "" then
			Call frm1.vspdData.SetText(C_bank_cd,		frm1.vspdData.ActiveRow,	arrRet(1))
			Call frm1.vspdData.SetText(C_bank_Nm,		frm1.vspdData.ActiveRow,	arrRet(2))
		End if
		
		Call vspdData_Change(C_bank_acct , frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenNoteNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������ȣ"					
	arrParam(1) = "F_NOTE,B_BANK"		    	
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_note_no,frm1.vspdData.ActiveRow,"X","X"))            
	arrParam(3) = ""						    
	arrParam(4) = "B_BANK.BANK_CD = F_NOTE.BANK_CD AND F_NOTE.NOTE_FG=" & FilterVar("D3", "''", "S") & " AND F_NOTE.NOTE_STS = " & FilterVar("BG", "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND F_NOTE.NOTE_AMT = " & UNICDbl(GetSpreadText(frm1.vspdData,C_pay_loc_amt,frm1.vspdData.ActiveRow,"X","X"))
	
	if Trim(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X")) = "" then
		Call DisplayMsgBox("17A002","X" , "����ó","X")  '%1�� �Է��ϼ���.
		IsOpenPop = False
		Exit Function
	end if
	
	arrParam(4) = arrParam(4) & " AND F_NOTE.BP_CD =  " & FilterVar(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & "  "
	
	if Trim(GetSpreadText(frm1.vspdData,C_charge_dt,frm1.vspdData.ActiveRow,"X","X")) = "" then
		Call DisplayMsgBox("17A002","X" , "�߻���","X")
		IsOpenPop = False
		Exit Function
	end if
	arrParam(4) = arrParam(4) & " AND F_NOTE.ISSUE_DT <=  " & FilterVar(UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_charge_dt,frm1.vspdData.ActiveRow,"X","X"))), "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND F_NOTE.DUE_DT >=  " & FilterVar(UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_charge_dt,frm1.vspdData.ActiveRow,"X","X"))), "''", "S") & " "
	
	if Trim(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X")) <> "" then
		arrParam(4) = arrParam(4) & " And B_BANK.Bank_Cd =  " & FilterVar(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & "  "
	end if
	arrParam(5) = "������ȣ"
	
	
    arrField(0) = "F_NOTE.NOTE_NO"				
    arrField(1) = "B_BANK.BANK_CD"				
    arrField(2) = "B_BANK.BANK_NM"					    
    arrField(3) = "F2" & parent.gColSep & "F_NOTE.NOTE_AMT"
    arrField(4) = "DD" & parent.gColSep & "F_NOTE.ISSUE_DT"
    arrField(5) = "DD" & parent.gColSep & "F_NOTE.DUE_DT"
    
    arrHeader(0) = "������ȣ"			     		
    arrHeader(1) = "�������"			     		
    arrHeader(2) = "��������"			     		
    arrHeader(3) = "�����ݾ�"			     		
    arrHeader(4) = "�߻���"			     			
    arrHeader(5) = "������"			     			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_note_no,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call vspdData_Change(C_note_no, frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenFnoteNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ǥ��ȣ"					
	arrParam(1) = "F_NOTE_NO,B_BANK"						
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_note_no,frm1.vspdData.ActiveRow,"X","X"))  	
	arrParam(3) = ""							
	arrParam(4) = "F_NOTE_NO.BANK_CD = B_BANK.BANK_CD "			
	arrParam(4) = arrParam(4) & "AND F_NOTE_NO.NOTE_KIND=" & FilterVar("CH", "''", "S") & " "
	arrParam(4) = arrParam(4) & "AND F_NOTE_NO.STS = " & FilterVar("NP", "''", "S") & " "
	
	if Trim(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X")) <> "" then
	    arrParam(4) = arrParam(4) & "and  B_BANK.BANK_CD =  " & FilterVar(GetSpreadText(frm1.vspdData,C_bank_cd,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & "  "
	end if
	
	arrParam(5) = "��ǥ��ȣ"			
	arrField(0) = "F_NOTE_NO.NOTE_NO"					
    arrField(1) = "B_BANK.BANK_CD"					
    arrField(2) = "B_BANK.BANK_NM"	
    
    arrHeader(0) = "��ǥ��ȣ"				
    arrHeader(1) = "�������"
    arrHeader(2) = "��������"				
    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_note_no,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call vspdData_Change(C_note_no , frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenPpNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	frm1.vspdData.row = frm1.vspdData.ActiveRow
	
	If Trim(GetSpreadText(frm1.vspdData,C_currency,frm1.vspdData.ActiveRow,"X","X")) = "" then
		Call DisplayMsgBox("17A002","X" , "ȭ��","X")
		Exit Function
	Elseif Trim(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X")) = "" then
		Call DisplayMsgBox("17A002","X" , "����ó","X")
		Exit Function
	End if
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "���ޱݹ�ȣ"	
	arrParam(1) = "F_PRPAYM, B_MINOR"
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_note_no,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(4) = "DOC_CUR =  " & FilterVar(GetSpreadText(frm1.vspdData,C_currency,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & "  "
	arrParam(4) = arrParam(4) & " And BP_CD =  " & FilterVar(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & "  AND BAL_AMT > 0"
	arrParam(4) = arrParam(4) & " AND B_MINOR.MINOR_CD = F_PRPAYM.CONF_FG AND B_MINOR.MAJOR_CD = " & FilterVar("F1012", "''", "S") & " "
	arrParam(5) = "���ޱݹ�ȣ"			
	
    arrField(0) = "PRPAYM_NO"
    arrField(1) = "F2" & parent.gColSep & "PRPAYM_AMT"
    arrField(2) = "DOC_CUR"
    arrField(3) = "F2" & parent.gColSep & "BAL_AMT"
    arrField(4) = "F5" & parent.gColSep & "XCH_RATE"  
    arrField(5) = "ED10" & parent.gColSep & "B_MINOR.MINOR_NM"
    
    arrHeader(0) = "���ޱݹ�ȣ"		
    arrHeader(1) = "���ޱ�"		
    arrHeader(2) = "���ޱ�ȭ��"
    arrHeader(3) = "���ޱ��ܾ�"
    arrHeader(4) = "ȯ��"
    arrHeader(5) = "����"
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_prpaym_no,	frm1.vspdData.ActiveRow, arrRet(0))
		Call frm1.vspdData.SetText(C_pp_xch_rt,	frm1.vspdData.ActiveRow, arrRet(4))
		Call vspdData_Change(C_prpaym_no , frm1.vspdData.Row)
	End If	

End Function
'==============================================================================================================================
Function OpenBuild()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
    
    if Trim(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X")) = "" then
		Call DisplayMsgBox("17A002","X" , "����ó","X")  '%1�� �Է��ϼ���.
		IsOpenPop = False
		exit Function
	end if
    
    arrHeader(0) = "���ݰ�꼭����ó"
    arrHeader(1) = "���ݰ�꼭����ó��"
    
    arrField(0) = "bpftn.partner_bp_cd"
    arrField(1) = "ptner.bp_nm"
    
	arrParam(0) = "���ݰ�꼭����ó"
	arrParam(1) = "b_biz_partner_ftn bpftn,b_biz_partner bp, b_biz_partner ptner"
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_BuildCd,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(4) = "bpftn.bp_cd=bp.bp_cd  And bpftn.partner_ftn=" & FilterVar("MBI", "''", "S") & " and ptner.bp_cd= bpftn.partner_bp_cd "
	arrParam(4) = arrParam(4) & " and bpftn.bp_cd= " & FilterVar(GetSpreadText(frm1.vspdData,C_bp_cd,frm1.vspdData.ActiveRow,"X","X"), "''", "S") & " "
	arrParam(5) = "���ݰ�꼭����ó��"


	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(C_BuildCd,	frm1.vspdData.ActiveRow, arrRet(0))
		Call frm1.vspdData.SetText(C_Build_Nm,	frm1.vspdData.ActiveRow, arrRet(1))
		Call vspdData_Change(C_BuildCd , frm1.vspdData.Row)
	End If	
    
End Function
'==============================================================================================================================
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True
	
	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_GlNo,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(1) = ""
	
    If UCase(Trim(GetSpreadText(frm1.vspdData,C_GlType,frm1.vspdData.ActiveRow,"X","X"))) = "A" Then               'ȸ����ǥ�˾� 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif UCase(Trim(GetSpreadText(frm1.vspdData,C_GlType,frm1.vspdData.ActiveRow,"X","X"))) = "T" Then          '������ǥ�˾� 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif UCase(Trim(GetSpreadText(frm1.vspdData,C_GlType,frm1.vspdData.ActiveRow,"X","X"))) = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '���� ��ǥ�� �������� �ʾҽ��ϴ�. 
    End if

	IsOpenPop = False
	
End Function
'==============================================================================================================================
Sub Getglno()
    Dim strFrom,strrefno
    Dim strGlNo,strTempGlNo
    Dim iCurRow
    
    Err.Clear
    iCurRow = frm1.vspdData.ActiveRow
    strFrom =  " ufn_a_GetGlNo( " & FilterVar(GetSpreadText(frm1.vspdData,C_charge_no,iCurRow,"X","X"), "''", "S") & " )"
    
    Call CommonQueryRs(" TEMP_GL_NO, GL_NO ", strFrom, "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    If lgF0 <> "" then 
		strTempGlNo = Split(lgF0, Chr(11))
		strGlNo		= Split(lgF1, Chr(11))
					
		If strGlNo(0) = "" and strTempGlNo(0) = "" then 
			Call frm1.vspdData.SetText(C_GlType,	iCurRow,	"B")
			Call frm1.vspdData.SetText(C_GlNo,		iCurRow,	"")
		Elseif strGlNo(0) = "" and strTempGlNo(0) <> "" then
			Call frm1.vspdData.SetText(C_GlType,	iCurRow,	"T")
			Call frm1.vspdData.SetText(C_GlNo,		iCurRow,	strTempGlNo(0))
		Elseif strGlNo(0) <> "" then 
			Call frm1.vspdData.SetText(C_GlType,	iCurRow,	"A")
			Call frm1.vspdData.SetText(C_GlNo,		iCurRow,	strGlNo(0))
		End If
	Else 
		Call frm1.vspdData.SetText(C_GlType,	iCurRow,	"B")
		Call frm1.vspdData.SetText(C_GlNo,		iCurRow,	"")
	End if	 

End Sub
'==============================================================================================================================
Function GetTaxBizArea(Byval strFlag, ByVal Row)
   	
   	Dim strSelectList, strFromList, strWhereList
	Dim strBilltoParty, strSalesGrp, strTaxBizArea
	Dim strRs
	Dim arrTaxBizArea(2), arrTemp
	
	Err.Clear
	'**�ٸ� ������ �̵��ϴ� ��� ���� ���� �����ְ��� �ߴ� �࿡ ��Ÿ���� ��.(2003.08.27)
	'frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = Row
     
	If strFlag = "NM" Then                                 '���ݽŰ����� ����� �̸����� �����´� 
		strTaxBizArea = UCase(Trim(GetSpreadText(frm1.vspdData,C_tax_biz_area,Row,"X","X")))
	Else
		strBilltoParty = UCase(Trim(GetSpreadText(frm1.vspdData,C_BuildCd,Row,"X","X")))
		strSalesGrp    = frm1.txtpur_grp1.value            '���ű׷� 
		'����ó�� ���� �׷��� ��� ��ϵǾ� �ִ� ��� �����ڵ忡 ������ rule�� ������ 
		If Len(strBillToParty) > 0 And Len(strSalesGrp) > 0	Then strFlag = "*"
	End if
	
	strSelectList = " * "
	strFromList = " dbo.ufn_m_GetTaxBizArea ( " & FilterVar(strBilltoParty, "''", "S") & " ,  " & FilterVar(strSalesGrp, "''", "S") & " ,  " & FilterVar(strTaxBizArea, "''", "S") & " ,  " & FilterVar(strFlag, "''", "S") & " ) "
	strWhereList = ""
	
	If CommonQueryRs2by2(strSelectList,strFromList,strWhereList,strRs) Then
		arrTemp = Split(strRs, Chr(11))
		Call frm1.vspdData.SetText(C_tax_biz_area,		Row,	arrTemp(1))
		Call frm1.vspdData.SetText(C_tax_biz_area_nm,	Row,	arrTemp(2))
	Else
		If Err.number <> 0 Then
			MsgBox Err.Description,vbInformation,parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		
		Call frm1.vspdData.SetText(C_tax_biz_area,		Row,	"")
		Call frm1.vspdData.SetText(C_tax_biz_area_nm,	Row,	"")
	End if
End Function
'==============================================================================================================================
Function CheckPayType(ByVal PayType)
    Dim iRow
	For iRow = 0 To UBound(arrCollectType,1)
	    If arrCollectType(iRow,0) = PayType and arrCollectType(iRow,1) <> "" Then
	       CheckPayType = arrCollectType(iRow,1)
	       Exit Function
	    End if
    Next
    CheckPayType = ""
End Function
'==============================================================================================================================
Sub InitCollectType()
	Dim iCodeArr, iRateArr
	Dim i
	
    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD=" & FilterVar("A1006", "''", "S") & " And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = Split(lgF0, Chr(11))
    iRateArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation, parent.gLogoName 
		Err.Clear 
		Exit Sub
	End If
	
	Redim arrCollectType(UBound(iCodeArr) - 1, 1)
	
	For i = 0 to UBound(iCodeArr) - 1
		arrCollectType(i, 0) = iCodeArr(i)
		arrCollectType(i, 1) = iRateArr(i)
	Next
	
End Sub	
'==============================================================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select
         
End Sub
'==============================================================================================================================
Function setPayDueDt(ByVal Row)

	'�߻��ݾ� + vat�ݾ� > ���ޱݾ� 
	if (UNICDbl(GetSpreadText(frm1.vspdData,C_charge_doc_amt,Row,"X","X")) + UNICDbl(GetSpreadText(frm1.vspdData,C_vat_doc_amt,Row,"X","X"))) > UNICDbl(GetSpreadText(frm1.vspdData,C_pay_doc_amt,Row,"X","X")) then
		ggoSpread.SSSetRequired		C_pay_due_dt, Row, Row '������ 
	else
		ggoSpread.spreadUnlock 		C_pay_due_dt, Row,C_pay_due_dt, Row
	end if

End Function
'==============================================================================================================================
Function ChangeCurOrDt(ByVal Row)

    Err.Clear                                                               '��: Protect system from crashing
	
	Dim strVal
    
    With frm1
		
		If Trim(GetSpreadText(frm1.vspdData,C_currency,Row,"X","X")) = "" Or _
			Trim(GetSpreadText(frm1.vspdData,C_charge_dt,Row,"X","X")) = "" then
			Exit Function
		End If

		If UCase(Trim(GetSpreadText(frm1.vspdData,C_currency,Row,"X","X"))) = UCase(Trim(parent.gCurrency)) then                '�ڱ� ȭ���ΰ�� 
			Call .vspdData.SetText(C_xch_rate,	Row, "1")'ȯ�� 
			Call .vspdData.SetText(C_calcd,		Row, "*")'�ڵ尪 
			ggoSpread.SSSetProtected	C_xch_rate, Row,Row
			ggoSpread.SSSetProtected	C_pp_xch_rt, Row,Row
			Call ChangeVatAmt(Row)	'*����*
			Call ChangeChargeLocAmt(Row)
			'�����ڱ��ݾ� ���(2003.08.14)
			frm1.vspdData.row = Row
			frm1.vspdData.col = C_pay_type
			If Trim(frm1.vspdData.text) <> "" Then
				Call ChangePayLocAmt(Row, frm1.vspdData.text)
			End If
			Exit Function
		Else
			ggoSpread.spreadUnlock 		C_xch_rate, Row,C_xch_rate, Row
		End If
		
   		strVal = BIZ_PGM_ID & "?txtMode=" & "LookupDailyExRt"	
		strVal = strVal & "&Currency=" & Trim(GetSpreadText(frm1.vspdData,C_currency,Row,"X","X"))                             'ȭ�� 
		strVal = strVal & "&ChargeDt=" & Trim(GetSpreadText(frm1.vspdData,C_charge_dt,Row,"X","X"))                        '�߻��� 
		strVal = strVal & "&gChangeOpt=" & gChangeOpt
				
    End With
	
    If  LayerShowHide(1) = False Then
      	Exit Function
    End If

    
	Call RunMyBizASP(MyBizASP, strVal)
        
End Function
'==============================================================================================================================
Function ChangeVatType(ByVal Row)

    Dim strVal
    Dim VatType, Cur
    
    Err.Clear
    
    With frm1
		
		.vspdData.ReDraw = false
		frm1.vspdData.Row = Row
		
		If Trim(GetSpreadText(frm1.vspdData,C_Vat_Type,Row,"X","X")) = "" then          'vat Ÿ�� 
			Call .vspdData.SetText(C_Vat_rate,		Row, "0")
			Call .vspdData.SetText(C_vat_doc_amt,	Row, "0")
			Call .vspdData.SetText(C_vat_loc_amt,	Row, "0")
		
			ggoSpread.SpreadLock 		C_vat_loc_amt, Row, C_vat_loc_amt, Row
			ggoSpread.SSSetProtected	C_vat_loc_amt, Row, Row
			Exit Function
		Else
			frm1.vspdData.Col = C_currency		'ȭ�� 
			Cur = frm1.vspdData.Text		 
'20040220 vat type ������ 
			If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) Then

				ggoSpread.SpreadLock 		C_vat_doc_amt, Row, C_vat_doc_amt, Row
				ggoSpread.SSSetProtected	C_vat_doc_amt, Row, Row
			
				ggoSpread.spreadUnlock 		C_vat_loc_amt, Row, C_vat_loc_amt, Row
				ggoSpread.SSSetRequired		C_vat_loc_amt, Row, Row
			Else
				ggoSpread.spreadUnlock 		C_vat_loc_amt, Row, C_vat_loc_amt, Row
				ggoSpread.SSSetRequired		C_vat_loc_amt, Row, Row

				ggoSpread.spreadUnlock 		C_vat_doc_amt, Row, C_vat_doc_amt, Row
				ggoSpread.SSSetRequired		C_vat_doc_amt, Row, Row
			End If
		End If
		
		.vspdData.ReDraw = true
		
		strVal = BIZ_PGM_ID & "?txtMode=" & "LookupVatType"
		strVal = strVal & "&VatType=" & Trim(GetSpreadText(frm1.vspdData,C_Vat_Type,Row,"X","X"))
		
    End With

    If  LayerShowHide(1) = False Then
      	Exit Function
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)
        
End Function
'==============================================================================================================================
Function ChangeVatAmt(ByVal Row)

	Dim ChargeDocAmt, VatDocAmt, VatLocAmt, VatRt, XchRt
	Dim cur
	Dim icdArr
    Dim swhere
	
	With frm1.vspdData
		gChangeOpt = ""
		
		.Row = Row
		
		cur = Trim(GetSpreadText(frm1.vspdData,C_currency,Row,"X","X"))
		ChargeDocAmt = UniCdbl(GetSpreadText(frm1.vspdData,C_charge_doc_amt,Row,"X","X"))
		VatRt = UniCdbl(GetSpreadText(frm1.vspdData,C_Vat_rate,Row,"X","X"))
		XchRt = UniCdbl(GetSpreadText(frm1.vspdData,C_xch_rate,Row,"X","X"))	'ȯ�� 
		
		VatDocAmt = ChargeDocAmt * VatRt / 100      'vat �ݾ� = �߻��ݾ� * vat �� * 0.01
		
		'[VAT�ڱ��ݾ� ���]+++++++++++++++++++++++++++++++
		If UCase(cur) = UCase(Trim(parent.gCurrency)) then
			VatLocAmt = VatDocAmt
		ElseIf Trim(GetSpreadText(frm1.vspdData,C_calcd,Row,"X","X")) = "*" then
			VatLocAmt = VatDocAmt * XchRt 
		ElseIf Trim(GetSpreadText(frm1.vspdData,C_calcd,Row,"X","X")) = "/" then
			VatLocAmt = VatDocAmt / XchRt			'vat �ڱ��ݾ� = vat �ݾ� C_calcd( / �Ǵ� *) ȯ�� 
		Else
			VatLocAmt = 0
		End If
		
		'.Col = C_Vat_rate:			.text = UNIFormatNumber(VatRt, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
	    .Col = C_vat_doc_amt:		.Text = UNIConvNumPCToCompanyByCurrency(VatDocAmt, cur, parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")
		.Col = C_vat_loc_amt:		.Text = UNIConvNumPCToCompanyByCurrency(CStr(VatLocAmt),parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")
		'++++++++++++++++++++++++++++++++++++++++++++++++
	End With
	
End Function
'==============================================================================================================================
Function CookieOp()

    frm1.txtprocess_step.Value 	= ReadCookie("Process_Step")
	frm1.txtbas_no.Value 		= ReadCookie("Po_No")
	frm1.txtpur_grp.Value 		= ReadCookie("Pur_Grp")
		
    WriteCookie "Process_Step" , ""
	WriteCookie "Po_No" ,""
	WriteCookie "Pur_Grp",""
	WriteCookie "Po_Cur",""
	WriteCookie "Po_Xch",""
	
	if frm1.txtprocess_step.Value <> "" And frm1.txtpur_grp.Value <> ""  then
			Call MainQuery()
	End if
	
End Function
'==============================================================================================================================
Sub SetSpreadLockAfterQuery()

	Dim index 
	Dim sPayType, Cur

    With frm1
	
		.vspdData.ReDraw = False
	
		For index = Cint(.hdnmaxrow.value) + 1	to .vspdData.MaxRows 
			.vspdData.Col = C_posting_flag   'Ȯ������ 
			.vspdData.Row = index
			
			if Trim(.vspdData.Text) <> "0" then	 'Ȯ���̸� ��ü lock
				ggoSpread.spreadUnlock 		C_glref_pop, index,C_glref_pop,index
				ggoSpread.SpreadLock 		C_charge_no, index,C_charge_no, index
				ggoSpread.SpreadLock		C_charge_type, index, .vspdData.MaxCols, index
			else
				ggoSpread.spreadUnlock 		C_posting_flag, index ,C_posting_flag,index
				ggoSpread.SpreadLock 		C_glref_pop, index,C_glref_pop,index
				ggoSpread.SpreadLock 		C_charge_no, index,C_charge_no, index
				ggoSpread.spreadUnlock 		C_charge_type, index,C_charge_type,index
				ggoSpread.SSSetRequired		C_charge_type,index,index
				ggoSpread.spreadUnlock 		C_charge_type_pop, index,C_charge_type_pop,index
				ggoSpread.SpreadLock 		C_charge_type_nm, index,C_bp_cd_Nm, index
				ggoSpread.SSSetProtected	C_charge_type_nm, index,index
				ggoSpread.spreadUnlock 		C_bp_cd, index,C_bp_cd, index
				ggoSpread.SSSetRequired		C_bp_cd,index,index
				ggoSpread.spreadUnlock 		C_bp_cd_pop, index,C_bp_cd_pop, index
				ggoSpread.SpreadLock 		C_bp_cd_Nm, index,C_bp_cd_Nm, index
				ggoSpread.SSSetProtected	C_bp_cd_Nm, index,index
				ggoSpread.SSSetProtected	C_Build_Nm, index,index			
				ggoSpread.spreadUnlock 		C_charge_dt,index,C_charge_dt,index
				ggoSpread.SSSetRequired		C_charge_dt,index,index
				ggoSpread.spreadUnlock 		C_vat_type, index,C_vat_type, index
				ggoSpread.spreadUnlock 		C_vat_type_pop, index,C_vat_type_pop, index
				ggoSpread.SpreadLock 		C_vat_type_Nm, index,C_vat_type_Nm, index
				ggoSpread.SSSetProtected	C_vat_type_Nm, index,index  
				ggoSpread.SSSetProtected	C_tax_biz_area_nm, index,index
				ggoSpread.spreadUnlock 		C_currency, index,C_currency, index
				ggoSpread.SSSetRequired		C_currency,index,index
				ggoSpread.spreadUnlock 		C_currency_pop, index,C_currency_pop, index
				ggoSpread.spreadUnlock 		C_charge_doc_amt, index,C_charge_doc_amt, index
				ggoSpread.SSSetRequired		C_charge_doc_amt,index,index
				ggoSpread.SpreadLock 		C_charge_loc_amt, index,C_charge_loc_amt, index
				
				.vspdData.Col = C_currency  
				Cur = frm1.vspdData.Text     
			
				If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then	 'ȭ�� �ڱ��ΰܿ�(KRW)			
					ggoSpread.SSSetProtected	C_xch_rate, .vspdData.Row,.vspdData.Row 'ȯ�� �����Ұ� 
				Else
					ggoSpread.spreadUnlock 		C_xch_rate, .vspdData.Row,C_xch_rate, .vspdData.Row
				End If
				ggoSpread.SSSetProtected	C_pp_xch_rt, .vspdData.Row,.vspdData.Row '���ޱ�ȯ�� �����Ұ� 
				
				ggoSpread.SpreadLock 		C_Vat_rate, index,C_Vat_rate, index
				ggoSpread.SSSetProtected	C_Vat_rate, index,index
				ggoSpread.SSSetProtected	C_vat_doc_amt, index,index
				ggoSpread.spreadUnlock 		C_vat_loc_amt, index,C_vat_loc_amt, index
				.vspdData.Col = C_vat_type
				if Trim(.vspdData.Text) <> "" And UCase(Trim(Cur)) <> UCase(Trim(parent.gCurrency)) then                 'VAT type
					ggoSpread.SSSetRequired		C_vat_loc_amt, index, index  'VAT �ڱ��ݾ� 
				else
					ggoSpread.spreadlock 		C_vat_loc_amt, index,C_vat_loc_amt, index
					ggoSpread.SSSetProtected	C_vat_loc_amt, index,index
				end if
				ggoSpread.spreadUnlock 		C_pay_type, index,C_pay_type, index
				ggoSpread.spreadUnlock 		C_pay_type_pop, index,C_pay_type_pop, index
				ggoSpread.SpreadLock 		C_pay_type_Nm, index,C_pay_type_Nm, index
				ggoSpread.SSSetProtected	C_pay_type_Nm, index,index
				ggoSpread.spreadUnlock 		C_charge_rate, index,C_charge_rate, index
				ggoSpread.spreadUnlock 		C_cost_flag, index,C_cost_flag, index
				frm1.vspdData.Col = C_pay_type
				sPayType = CheckPayType(Frm1.vspdData.text)      '�������� 
				if sPayType = "DP" then      '������(����,���¹�ȣ ��������)                        
					ggoSpread.spreadUnlock 		C_bank_cd, index,C_bank_acct_pop, index
					ggoSpread.SSSetRequired		C_bank_cd,index,index
					ggoSpread.SSSetRequired		C_bank_acct,index,index
				else
					ggoSpread.spreadlock 		C_bank_cd, index,C_bank_acct_pop, index
				end if

				frm1.vspdData.Col = C_pay_type   '�������� 
				if Trim(frm1.vspdData.Text) <> "" then   
				ggoSpread.spreadUnlock		C_pay_doc_amt,index,C_pay_doc_amt, index  '���ޱݾ� 
				ggoSpread.SSSetRequired		C_pay_doc_amt,index,index  '���ޱݾ� 
			Else
				ggoSpread.spreadUnlock		C_pay_doc_amt,index,C_pay_doc_amt, index  '���ޱݾ� 
			End If
				ggoSpread.spreadlock 		C_pay_loc_amt, index,C_pay_loc_amt, index
				ggoSpread.SSSetProtected	C_pay_loc_amt, index,index
				
				ggoSpread.SpreadLock 		C_bank_Nm, index,C_bank_Nm, index
				ggoSpread.SSSetProtected	C_bank_Nm, index,index
				frm1.vspdData.Col = C_pay_type_seq4    '���������� = C_pay_type_seq4
				
				if sPayType = "NO" then '���޾���(������ȣ ��������)
					ggoSpread.spreadUnlock 		C_note_No, index,C_note_no_pop, index
					ggoSpread.SSSetRequired		C_note_No,index,index
					ggoSpread.SpreadLock 		C_prpaym_No, index,C_prpaym_no_pop, index
				elseif sPayType = "CK" then
				    ggoSpread.spreadUnlock 		C_note_No, index, index
				    ggoSpread.SSSetProtected    C_note_no_pop, index, index
				    ggoSpread.SpreadLock 		C_prpaym_No, index,C_prpaym_no_pop, index
				elseif sPayType = "PP" then '���ޱ�(���ޱݹ�ȣ ��������)
					ggoSpread.spreadUnlock 		C_prpaym_No, index,C_prpaym_no_pop, index
					ggoSpread.SSSetRequired		C_prpaym_No,index,index
					ggoSpread.SpreadLock 		C_note_No, index,C_note_no_pop, index
				elseif sPayType = "" then   'type�� ������(����,����,���ޱݹ�ȣ,������ȣ �����Ұ�)
					ggoSpread.spreadUnlock 		C_note_No, index,C_prpaym_no_pop, index
					ggoSpread.spreadUnlock 		C_bank_cd, index,C_bank_pop, index
					ggoSpread.spreadUnlock 		C_bank_acct, index,C_bank_acct_pop, index
				else
					ggoSpread.SpreadLock 		C_note_No, index,C_prpaym_no_pop, index
				end if
				ggoSpread.spreadUnlock 		C_remark, index,C_remark, index
				ggoSpread.spreadlock 		C_bas_no, index,C_bas_no, index
				ggoSpread.spreadUnlock 		C_old_posting_flg, index,C_old_posting_flg, index
				ggoSpread.SSSetProtected    C_old_posting_flg + 1, index, index
				
				Call setPayDueDt(index)     '������ setting (�߻��ݾ� + vat�ݾ� > ���ޱݾ� �� ��������)
			End if    
			
		Next
		.vspdData.ReDraw = True
	End With
End Sub
'����Ŀ� spreadLock ���� (2003.08.27) - Lee, Eun Hee
'===============================================================================================
Sub SetSpreadLockAfterCancel(ByVal index)

	Dim sPayType, Cur

    With frm1
	
		.vspdData.ReDraw = False
	
		.vspdData.Col = C_posting_flag   'Ȯ������ 
		.vspdData.Row = index
			
		If Trim(.vspdData.Text) <> "0" Then	 'Ȯ���̸� ��ü lock
			ggoSpread.spreadUnlock 		C_glref_pop, index,C_glref_pop,index
			ggoSpread.SpreadLock 		C_charge_no, index,C_charge_no, index
			ggoSpread.SpreadLock		C_charge_type, index, .vspdData.MaxCols, index
		Else
			ggoSpread.spreadUnlock 		C_posting_flag, index ,C_posting_flag,index
			ggoSpread.SpreadLock 		C_glref_pop, index,C_glref_pop,index
			ggoSpread.SpreadLock 		C_charge_no, index,C_charge_no, index
			ggoSpread.spreadUnlock 		C_charge_type, index,C_charge_type,index
			ggoSpread.SSSetRequired		C_charge_type,index,index
			ggoSpread.spreadUnlock 		C_charge_type_pop, index,C_charge_type_pop,index
			ggoSpread.SpreadLock 		C_charge_type_nm, index,C_bp_cd_Nm, index
			ggoSpread.SSSetProtected	C_charge_type_nm, index,index
			ggoSpread.spreadUnlock 		C_bp_cd, index,C_bp_cd, index
			ggoSpread.SSSetRequired		C_bp_cd,index,index
			ggoSpread.spreadUnlock 		C_bp_cd_pop, index,C_bp_cd_pop, index
			ggoSpread.SpreadLock 		C_bp_cd_Nm, index,C_bp_cd_Nm, index
			ggoSpread.SSSetProtected	C_bp_cd_Nm, index,index
			ggoSpread.SSSetProtected	C_Build_Nm, index,index			
			ggoSpread.spreadUnlock 		C_charge_dt,index,C_charge_dt,index
			ggoSpread.SSSetRequired		C_charge_dt,index,index
			ggoSpread.spreadUnlock 		C_vat_type, index,C_vat_type, index
			ggoSpread.spreadUnlock 		C_vat_type_pop, index,C_vat_type_pop, index
			ggoSpread.SpreadLock 		C_vat_type_Nm, index,C_vat_type_Nm, index
			ggoSpread.SSSetProtected	C_vat_type_Nm, index,index  
			ggoSpread.SSSetProtected	C_tax_biz_area_nm, index,index
			ggoSpread.spreadUnlock 		C_currency, index,C_currency, index
			ggoSpread.SSSetRequired		C_currency,index,index
			ggoSpread.spreadUnlock 		C_currency_pop, index,C_currency_pop, index
			ggoSpread.spreadUnlock 		C_charge_doc_amt, index,C_charge_doc_amt, index
			ggoSpread.SSSetRequired		C_charge_doc_amt,index,index
			ggoSpread.SpreadLock 		C_charge_loc_amt, index,C_charge_loc_amt, index
				
			.vspdData.Col = C_currency       
			Cur = frm1.vspdData.Text
			
			If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then	 'ȭ�� �ڱ��ΰܿ�(KRW)			
				ggoSpread.SSSetProtected	C_xch_rate, .vspdData.Row,.vspdData.Row 'ȯ�� �����Ұ� 
			Else
				ggoSpread.spreadUnlock 		C_xch_rate, .vspdData.Row,C_xch_rate, .vspdData.Row
			End If
			ggoSpread.SSSetProtected	C_pp_xch_rt, .vspdData.Row,.vspdData.Row '���ޱ�ȯ�� �����Ұ� 
				
			ggoSpread.SpreadLock 		C_Vat_rate, index,C_Vat_rate, index
			ggoSpread.SSSetProtected	C_Vat_rate, index,index
			ggoSpread.SSSetProtected	C_vat_doc_amt, index,index
			ggoSpread.spreadUnlock 		C_vat_loc_amt, index,C_vat_loc_amt, index
			.vspdData.Col = C_vat_type
			if Trim(.vspdData.Text) <> "" And UCase(Trim(Cur)) <> UCase(Trim(parent.gCurrency)) then                 'VAT type
				ggoSpread.SSSetRequired		C_vat_loc_amt, index, index  'VAT �ڱ��ݾ� 
			else
				ggoSpread.spreadlock 		C_vat_loc_amt, index,C_vat_loc_amt, index
				ggoSpread.SSSetProtected	C_vat_loc_amt, index,index
			end if
			ggoSpread.spreadUnlock 		C_pay_type, index,C_pay_type, index
			ggoSpread.spreadUnlock 		C_pay_type_pop, index,C_pay_type_pop, index
			ggoSpread.SpreadLock 		C_pay_type_Nm, index,C_pay_type_Nm, index
			ggoSpread.SSSetProtected	C_pay_type_Nm, index,index
			ggoSpread.spreadUnlock 		C_charge_rate, index,C_charge_rate, index
			ggoSpread.spreadUnlock 		C_cost_flag, index,C_cost_flag, index
			frm1.vspdData.Col = C_pay_type
			sPayType = CheckPayType(Frm1.vspdData.text)      '�������� 
			if sPayType = "DP" then      '������(����,���¹�ȣ ��������)                        
				ggoSpread.spreadUnlock 		C_bank_cd, index,C_bank_acct_pop, index
				ggoSpread.SSSetRequired		C_bank_cd,index,index
				ggoSpread.SSSetRequired		C_bank_acct,index,index
			else
				ggoSpread.spreadlock 		C_bank_cd, index,C_bank_acct_pop, index
			end if

			frm1.vspdData.Col = C_pay_type   '�������� 
			if Trim(frm1.vspdData.Text) <> "" then   
				ggoSpread.spreadUnlock		C_pay_doc_amt,index,C_pay_doc_amt, index  '���ޱݾ� 
				ggoSpread.SSSetRequired		C_pay_doc_amt,index,index  '���ޱݾ� 
			Else
				ggoSpread.spreadUnlock		C_pay_doc_amt,index,C_pay_doc_amt, index  '���ޱݾ� 
			End If
			ggoSpread.spreadlock 		C_pay_loc_amt, index,C_pay_loc_amt, index
			ggoSpread.SSSetProtected	C_pay_loc_amt, index,index
				
			ggoSpread.SpreadLock 		C_bank_Nm, index,C_bank_Nm, index
			ggoSpread.SSSetProtected	C_bank_Nm, index,index
			frm1.vspdData.Col = C_pay_type_seq4    '���������� = C_pay_type_seq4
				
			if sPayType = "NO" then '���޾���(������ȣ ��������)
				ggoSpread.spreadUnlock 		C_note_No, index,C_note_no_pop, index
				ggoSpread.SSSetRequired		C_note_No,index,index
				ggoSpread.SpreadLock 		C_prpaym_No, index,C_prpaym_no_pop, index
			elseif sPayType = "CK" then
			    ggoSpread.spreadUnlock 		C_note_No, index, index
			    ggoSpread.SSSetProtected    C_note_no_pop, index, index
			    ggoSpread.SpreadLock 		C_prpaym_No, index,C_prpaym_no_pop, index
			elseif sPayType = "PP" then '���ޱ�(���ޱݹ�ȣ ��������)
				ggoSpread.spreadUnlock 		C_prpaym_No, index,C_prpaym_no_pop, index
				ggoSpread.SSSetRequired		C_prpaym_No,index,index
				ggoSpread.SpreadLock 		C_note_No, index,C_note_no_pop, index
			elseif sPayType = "" then   'type�� ������(����,����,���ޱݹ�ȣ,������ȣ �����Ұ�)
				ggoSpread.spreadUnlock 		C_note_No, index,C_prpaym_no_pop, index
				ggoSpread.spreadUnlock 		C_bank_cd, index,C_bank_pop, index
				ggoSpread.spreadUnlock 		C_bank_acct, index,C_bank_acct_pop, index
			else
				ggoSpread.SpreadLock 		C_note_No, index,C_prpaym_no_pop, index
			end if
			ggoSpread.spreadUnlock 		C_remark, index,C_remark, index
			ggoSpread.spreadlock 		C_bas_no, index,C_bas_no, index
			ggoSpread.spreadUnlock 		C_old_posting_flg, index,C_old_posting_flg, index
			ggoSpread.SSSetProtected    C_old_posting_flg + 1, index, index
				
			Call setPayDueDt(index)     '������ setting (�߻��ݾ� + vat�ݾ� > ���ޱݾ� �� ��������)
		End If    
			
		.vspdData.ReDraw = True
	End With
End Sub
'==============================================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")               
    
    Call SetDefaultVal
	Call InitSpreadSheet                                
    Call InitVariables

    Call CookieOp()
	Call InitCollectType
        
End Sub
'==============================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	If lgIntFlgMode <> Parent.OPMD_UMODE Or frm1.vspdData.MaxRows < 1 Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	
	gMouseClickStatus = "SPC"   
   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If       
       Exit Sub
    End If   

End Sub
'==============================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'==============================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'==============================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_posting_flag Or NewCol <= C_posting_flag Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'==============================================================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function
'==============================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim sPayType,sVatDocAmt
	Dim sBpCd, Cur
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
	frm1.vspdData.Col = C_currency		'ȭ�� 
	Cur = frm1.vspdData.Text
				
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        '  <------����� ǥ�� ���� 

	Select Case Col
	
		Case C_bp_cd                                          '����ó 
			frm1.vspdData.Col = C_bp_cd
			sBpCd = frm1.vspdData.text
			Call ChangeBp(sBpCd, Row)                                   '����ó��,��������,����������,ȭ��,VAT,VAT��,VAT�������� 
		Case C_pay_type                                       '�������� 
			frm1.vspdData.ReDraw = false
			frm1.vspdData.Col = C_pay_type
	    	sPayType = CheckPayType(Frm1.vspdData.text)       'checkpaytype type�� ������ "" �Ѿ�´� 
            
			if sPayType <> "" then
				if sPayType = "NO" then  '���޾��� 
					ggoSpread.spreadUnlock 	C_note_no, Row,C_note_no_pop, Row
					if interface_Account = "Y" then
						ggoSpread.SSSetRequired	C_note_no,Row,Row
					end if
				
					Call frm1.vspdData.SetText(C_prpaym_no,	Row, "")'���ޱݹ�ȣ 
					Call frm1.vspdData.SetText(C_pp_xch_rt,	Row, "")'���ޱ�ȯ�� 
				
					ggoSpread.spreadlock 	C_prpaym_no, Row,C_prpaym_no_pop, Row
					ggoSpread.spreadlock 	C_pp_xch_rt, Row,C_pp_xch_rt, Row
					
				elseif sPayType = "PP" then  '���ޱ� 
					ggoSpread.spreadUnlock 	C_prpaym_no, Row, C_prpaym_no_pop, Row
					if interface_Account = "Y" then
						ggoSpread.SSSetRequired	C_prpaym_no,Row,Row
					end if
				
					Call frm1.vspdData.SetText(C_note_no,	Row, "")'������ȣ 
					ggoSpread.spreadlock 	C_note_no, Row,C_note_no_pop, Row
					
				elseif sPayType = "CK" then
				    ggoSpread.spreadUnlock 	C_note_no, Row,C_note_no_pop, Row
				    
				    Call frm1.vspdData.SetText(C_prpaym_no,	Row, "")'��ǥ�ΰ�� 
				    Call frm1.vspdData.SetText(C_pp_xch_rt,	Row, "")'���ޱ�ȯ�� 
				    ggoSpread.spreadlock 	C_prpaym_no, Row,C_prpaym_no_pop, Row
					ggoSpread.spreadlock 	C_pp_xch_rt, Row,C_pp_xch_rt, Row
					
					
				else
					Call frm1.vspdData.SetText(C_note_no,	Row, "")
					ggoSpread.spreadlock 	C_note_no, Row,C_note_no_pop, Row
					
					Call frm1.vspdData.SetText(C_prpaym_no,	Row, "")
					Call frm1.vspdData.SetText(C_pp_xch_rt,	Row, "")'���ޱ�ȯ�� 
					ggoSpread.spreadlock 	C_prpaym_no, Row,C_prpaym_no_pop, Row
					ggoSpread.spreadlock 	C_pp_xch_rt, Row,C_pp_xch_rt, Row
					
				end if
		
				frm1.vspdData.Col = C_pay_type  ''������ 
				if sPayType = "DP" then
					ggoSpread.spreadUnlock 		C_bank_cd, Row,C_bank_acct_pop, Row
					ggoSpread.SSSetRequired		C_bank_cd,Row,Row
					if interface_Account = "Y" then
						ggoSpread.SSSetRequired		C_bank_acct,Row,Row
					end if
					ggoSpread.SSSetProtected	C_bank_nm,Row,Row
				else
					Call frm1.vspdData.SetText(C_bank_cd,	Row, "")
					Call frm1.vspdData.SetText(C_bank_Nm,	Row, "")
					Call frm1.vspdData.SetText(C_bank_acct,	Row, "")
					ggoSpread.spreadlock 	C_bank_cd, Row,C_bank_acct_pop, Row
				end if	
				
				if Trim(GetSpreadText(frm1.vspdData,C_pay_type,Row,"X","X")) <> "" then
					ggoSpread.spreadUnlock 		C_pay_doc_amt, Row,C_pay_doc_amt, Row
					ggoSpread.SSSetRequired		C_pay_doc_amt,Row,Row
					
				else
					Call frm1.vspdData.SetText(C_pay_doc_amt,	Row, "0")'���������� ������ ���ޱݾ��� 0 
					ggoSpread.SSSetProtected	C_pay_doc_amt, Row, Row
				end if
				
				Call frm1.vspdData.SetText(C_pay_doc_amt,	Row, "0")
				Call frm1.vspdData.SetText(C_pay_loc_amt,	Row, "0")
				
				
			else    '���������� arrCollectType�� ���� ��� 
				'����, ��ݰ��� , ������ȣ, ���ޱݹ�ȣ ���� Optionaló�� 
				Call frm1.vspdData.SetText(C_bank_cd,	Row, "")
				Call frm1.vspdData.SetText(C_bank_Nm,	Row, "")
				Call frm1.vspdData.SetText(C_bank_acct,	Row, "")
				Call frm1.vspdData.SetText(C_note_no,	Row, "")
				Call frm1.vspdData.SetText(C_prpaym_no,	Row, "")
				
				ggoSpread.spreadUnlock  C_pay_doc_amt,Row,C_pay_doc_amt, Row
				ggoSpread.spreadUnlock 	C_bank_cd, Row,C_bank_pop, Row
				ggoSpread.spreadUnlock 	C_bank_acct, Row,C_bank_acct_pop, Row
				ggoSpread.spreadUnlock 	C_note_no, Row,C_note_no_pop, Row
				ggoSpread.spreadUnlock 	C_prpaym_no, Row, C_prpaym_no_pop, Row
				
				Call frm1.vspdData.SetText(C_pay_doc_amt,	Row, "0")
				Call frm1.vspdData.SetText(C_pay_loc_amt,	Row, "0")
				
			end if
			Call setPayDueDt(Row)
			frm1.vspdData.ReDraw = true
		
		Case C_prpaym_no
			'--[�����ڱ��ݾ� ���]-----------------(2003.08.14)
			frm1.vspdData.Col = C_pay_type
			Call ChangePayLocAmt(Row, frm1.vspdData.Text)
			'---------------------------------------
		
		Case C_currency, C_charge_dt                    'ȭ��,�߻��� 
			
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_charge_doc_amt,   "A" ,"X","X")
            Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_Vat_rate,"D" ,"X","X")
            Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_vat_doc_amt,"A" ,"X","X")
            Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_pay_doc_amt,"A" ,"X","X")
            Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_charge_rate,"D" ,"X","X")
			
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_currency,C_charge_doc_amt,   "A" ,"I","X","X")
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_currency,C_Vat_rate,"D" ,"I","X","X")         
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_currency,C_vat_doc_amt,"A" ,"I","X","X")         
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_currency,C_pay_doc_amt,"A" ,"I","X","X")         
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_currency,C_charge_rate,"D" ,"I","X","X")                       
			'�̺�Ʈ ������� ����(2003.08.14)
			gChangeOpt = ""
			Call ChangeCurOrDt(Row)                        'ȭ�󺯵��� ȯ����,�߻��Ϻ��� 
			
			If  UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency))  then
			   	ggoSpread.SpreadLock 		C_vat_doc_amt, frm1.vspdData.Row, C_vat_doc_amt, frm1.vspdData.Row
				ggoSpread.SSSetProtected    C_vat_doc_amt, frm1.vspdData.Row, frm1.vspdData.Row
				
				ggoSpread.spreadUnlock 		C_vat_loc_amt, frm1.vspdData.Row, C_vat_loc_amt, frm1.vspdData.Row
				ggoSpread.SSSetRequired		C_vat_loc_amt, frm1.vspdData.Row, frm1.vspdData.Row
			Else
				' �ڱ�ȭ��(krw)�� �ƴ� ����  vat �ݾ�/�ڱ� �ݾ� required(20040204)
				ggoSpread.spreadUnlock 		C_vat_doc_amt, frm1.vspdData.Row, C_vat_doc_amt, frm1.vspdData.Row
			    ggoSpread.SSSetRequired		C_vat_doc_amt, frm1.vspdData.Row, frm1.vspdData.Row
				 
			    ggoSpread.spreadUnlock 		C_vat_loc_amt, frm1.vspdData.Row, C_vat_loc_amt, frm1.vspdData.Row
			    ggoSpread.SSSetRequired		C_vat_loc_amt, frm1.vspdData.Row, frm1.vspdData.Row
			End If
			
			
		Case C_charge_doc_amt, C_Vat_rate   '�߻��ݾ�,VAT�� 
			'�߻��ڱ��ݾ� ���(2003.08.14)
			Call ChangeChargeLocAmt(Row)
			
			Call ChangeVatAmt(Row)                         'vat �ݾ� ������ ȣ�� 
			If Col = C_charge_doc_amt then
				frm1.vspdData.ReDraw = False
				Call setPayDueDt(Row)                   '������ setting (�߻��ݾ�,vat�ݾ�,���ޱݾ� ������ ȣ��)
				
				'�ڱ�ȭ��(KRW)�� ���� vat�ڱ��ݾ��� protected(2003.09.22)
				' �ڱ�ȭ��(krw)�� ����  vat �ݾ� protected,vat �ڱ� �ݾ� required(20040204)
				'If UniCdbl(GetSpreadText(frm1.vspdData,C_vat_doc_amt,Row,"X","X")) = 0 Or UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency))  then
				If  UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency))  then
				   	ggoSpread.SpreadLock 		C_vat_doc_amt, frm1.vspdData.Row, C_vat_doc_amt, frm1.vspdData.Row
					ggoSpread.SSSetProtected    C_vat_doc_amt, frm1.vspdData.Row, frm1.vspdData.Row
				
					ggoSpread.spreadUnlock 		C_vat_loc_amt, frm1.vspdData.Row, C_vat_loc_amt, frm1.vspdData.Row
					ggoSpread.SSSetRequired		C_vat_loc_amt, frm1.vspdData.Row, frm1.vspdData.Row
				Else
					' �ڱ�ȭ��(krw)�� �ƴ� ����  vat �ݾ�/�ڱ� �ݾ� required(20040204)
					ggoSpread.spreadUnlock 		C_vat_doc_amt, frm1.vspdData.Row, C_vat_doc_amt, frm1.vspdData.Row
				    ggoSpread.SSSetRequired		C_vat_doc_amt, frm1.vspdData.Row, frm1.vspdData.Row
				 
				    ggoSpread.spreadUnlock 		C_vat_loc_amt, frm1.vspdData.Row, C_vat_loc_amt, frm1.vspdData.Row
				    ggoSpread.SSSetRequired		C_vat_loc_amt, frm1.vspdData.Row, frm1.vspdData.Row
				End If
				
				frm1.vspdData.ReDraw = True
			End If
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_charge_doc_amt,   "A" ,"X","X")
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_Vat_rate,"D" ,"X","X")
		Case C_vat_type 
			frm1.vspdData.ReDraw = False
			Call ChangeVatType(Row)						'vat������ ȣ�� (vat��,vat��,vat�ݾ�,vat �ڱ��ݾ� ����)

			' vat ������ vat��=0�� ��� vat�ڱ��ݾ��� protected, vat��<>0�� ��� vat�ڱ��ݾ��� required
			' 20040204 �ּ�ó�� 
		'	If UNICDbl(Trim(GetSpreadText(frm1.vspdData,C_vat_rate,Row,"X","X"))) = 0 Or UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then
		'		ggoSpread.SpreadLock 		C_vat_loc_amt, frm1.vspdData.Row, C_vat_loc_amt, frm1.vspdData.Row
		'		ggoSpread.SSSetProtected C_vat_loc_amt, frm1.vspdData.Row, frm1.vspdData.Row
		'	Else 
		'		ggoSpread.spreadUnlock	C_vat_loc_amt, frm1.vspdData.Row, C_vat_loc_amt, frm1.vspdData.Row
		'		ggoSpread.SSSetRequired	C_vat_loc_amt, frm1.vspdData.Row, frm1.vspdData.Row
		'	End If
			frm1.vspdData.ReDraw = True
		Case C_xch_rate
			If UNICDbl(Trim(GetSpreadText(frm1.vspdData,C_xch_rate,Row,"X","X"))) = 0 Then
				Call ChangeCurOrDt(Row)
			Else
				gChangeOpt = "XCH"
				Call ChangeCurOrDt(Row)
			End If
		Case C_pay_doc_amt                              '���ޱݾ� 
			frm1.vspdData.ReDraw = False              
			Call setPayDueDt(Row)
			'--[�����ڱ��ݾ� ���]-----------------(2003.08.14)
			Dim LgPayType
			frm1.vspdData.Col = C_pay_type
			LgPayType = frm1.vspdData.Text 
								
			Call ChangePayLocAmt(Row, LgPayType)
			'-------------------------------------
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_pay_doc_amt,"A" ,"X","X")
			frm1.vspdData.ReDraw = True
		Case C_BuildCd                              '����ó 
			frm1.vspdData.ReDraw = False    
            Call GetTaxBizArea("*", Row)
            frm1.vspdData.ReDraw = True
        Case C_tax_biz_area							'���ݽŰ����� 
			frm1.vspdData.ReDraw = False 
			Call GetTaxBizArea("NM", Row)
            frm1.vspdData.ReDraw = True
        '----------------------------------------
        Case C_charge_rate
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_currency,C_charge_rate,"D" ,"X","X")
        '20040204
        Case c_vat_loc_amt
		 	If  UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency))  then
				 frm1.vspdData.Col = C_vat_loc_amt:		sVatDocAmt = frm1.vspdData.text
				 frm1.vspdData.Col = C_vat_doc_amt:		frm1.vspdData.Text = UNIConvNumPCToCompanyByCurrency(CStr(sVatDocAmt),parent.gCurrency,parent.ggAmtOfMoneyNo,parent.gTaxRndPolicyNo,"X")
			End If
			
        End Select
    
End Sub
'==========================================================================================
'   Event Name : ChangeChargeLocAmt()
'   Event Desc : �߻��ڱ��ݾ� ������ (�߻��ݾ� ����� ȣ��)
'==========================================================================================	
Function ChangeChargeLocAmt(Byval LRow)

    Err.Clear

    Dim Cur,DocAmt,XchRt
    
    With frm1
			
		.vspdData.Row = LRow
		frm1.vspdData.Col = C_currency		'ȭ�� 
		Cur = .vspdData.Text		 

		.vspdData.Col = C_charge_doc_amt		'�߻��ݾ� 
		DocAmt = .vspdData.Text
		
		frm1.vspdData.Col = C_xch_rate
		XchRt = .vspdData.Text

		If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) Then
			.vspdData.Col = C_xch_rate
			.vspdData.Text = "1"
			.vspdData.Col = C_charge_loc_amt
			.vspdData.Text = DocAmt
			Exit Function
		End If	
		
		.vspdData.Col = C_calcd
		If Trim(.vspdData.Text) = "*" Then
			.vspdData.Col = C_charge_loc_amt
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) * UNICDbl(XchRt),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		Elseif Trim(.vspdData.Text) = "/" Then
			.vspdData.Col = C_charge_loc_amt
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) / UNICDbl(XchRt),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		End If

    End With
    
        
End Function
'==========================================================================================
'   Event Name : ChangePayLocAmt()
'   Event Desc : �����ڱ��ݾ� ������ (���ޱݾ� ����� ȣ��)
'==========================================================================================	
Function ChangePayLocAmt(Byval LRow, ByVal iPayType)

    Err.Clear                                                               '��: Protect system from crashing

    Dim Cur,DocAmt,XchRt
    
    With frm1
			
		.vspdData.Row = LRow
		frm1.vspdData.Col = C_currency		'ȭ�� 
		Cur = .vspdData.Text		 

		.vspdData.Col = C_pay_doc_amt		'���ޱݾ� 
		DocAmt = .vspdData.Text
		
		frm1.vspdData.Col = C_pp_xch_rt
		If  UCase(Trim(iPayType)) <> "PP" Then		'���ޱ��� ��츦 �����ϰ�� ��������� ȯ���� ������.
			frm1.vspdData.Col = C_xch_rate
			XchRt = .vspdData.Text
		ElseIf (UCase(Trim(iPayType)) = "PP" and UNICDbl(.vspdData.Text) <> 0) Then  '���ޱ��ϰ��� ȯ���� �־�� ��.
			frm1.vspdData.Col = C_pp_xch_rt		'���ޱ�ȯ�� 
			XchRt = .vspdData.Text
		End If
			
		If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then    'ȭ�� KRW�̸� �����ڱ��ݾ� = ���ޱݾ� * 1
			.vspdData.Col = C_xch_rate
			.vspdData.Text = "1"
			.vspdData.Col = C_pay_loc_amt                      '�����ڱ��ݾ� 
			.vspdData.Text = DocAmt
			Exit Function
		End If	
		
		.vspdData.Col = C_calcd
		If Trim(.vspdData.Text) = "*" Then
			.vspdData.Col = C_pay_loc_amt
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) * UNICDbl(XchRt),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		Elseif Trim(.vspdData.Text) = "/" Then
			.vspdData.Col = C_pay_loc_amt
		    .vspdData.text = UNIConvNumPCToCompanyByCurrency(UNICDbl(DocAmt) / UNICDbl(XchRt),parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo , "X")
		End If

    End With
    
        
End Function
'==============================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
    Dim sPayType
	Dim sCurCurrency
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		.Row = Row

		If Row > 0 then

			if  Col = C_charge_type_pop Then
				Call OpenChargeType()
			elseIf Col = C_glref_pop Then
			    Call Getglno()
			    Call OpenGLRef()	
			elseIf Col = C_currency_pop Then
				Call OpenCurrency()
			elseIf  Col = C_pay_type_pop Then
				Call OpenPay_Type()
			elseIf  Col = C_vat_type_pop Then
				Call OpenVat_Type()
			elseIf  Col = C_bp_cd_pop Then
				Call OpenBp_Cd()
			elseIf  Col = C_bank_acct_pop Then
				Call OpenBank_Acct()
			elseIf  Col = C_bank_pop Then
				Call OpenBank()
			elseif  Col = C_note_no_pop then
				frm1.vspdData.Col = C_pay_type
				sPayType = CheckPayType(Frm1.vspdData.text) 
				frm1.vspdData.Col = C_pay_type
				Call OpenNoteNo()
			elseif  Col = C_prpaym_no_pop then
				Call OpenPpNo()
			elseif  Col = C_tax_biz_area_pop then
				Call OpenBizArea()
			elseif  Col = C_BuildCd_pop then
				Call OpenBuild()
			end if

		End If
    
    End With
End Sub
'==============================================================================================================================
Sub btnbas_noOnClick()
	Dim strChargeType
		
	strChargeType = Trim(frm1.txtprocess_step.value)
		
	If strChargeType <> "" Then
		Select Case UCase(strChargeType)
		Case "PO"							'Count Offer
			Call OpenBasNoPop("m3111pa1")			
		Case "VL"							'���� L/C
			Call OpenBasNoPop("m3211pa1")	
		Case "VA"							'���� L/C Amend
			Call OpenBasNoPop("m3221pa1")			
		Case "VO"							'���� Local L/C
			Call OpenBasNoPop("m3211pa2")		
		Case "VF"							'���� Local L/C Amend
			Call OpenBasNoPop("m3211pa2")		
		Case "VD"							'������� 
			Call OpenBasNoPop("m4211pa1")			
		Case "VB"							'���Լ��� 
			Call OpenBasNoPop("m5211pa1")
		End Select
	Else
		Call DisplayMsgBox("17A002","X" , "���౸��","X")
	End If
End Sub
'==============================================================================================================================	
Sub btnbas_no1OnClick()
  	Dim strChargeType
		
  	if UCase(frm1.txtpur_grp1.className) = UCase(parent.UCN_PROTECTED) then Exit sub
		 
  	strChargeType = Trim(frm1.txtprocess_step1.value)
		
  	If strChargeType <> "" Then
  		Select Case UCase(strChargeType)
  		Case "PO"							'Count Offer
  			Call OpenBasNoPop1("m3111pa1")			
  		Case "VL"							'���� L/C
  			Call OpenBasNoPop1("m3211pa1")	
  		Case "VA"							'���� L/C Amend
  			Call OpenBasNoPop1("m3221pa1")			
  		Case "VO"							'���� Local L/C
  			Call OpenBasNoPop1("m3211pa2")		
  		Case "VF"							'���� Local L/C Amend
  			Call OpenBasNoPop1("m3211pa2")		
  		Case "VD"							'������� 
  			Call OpenBasNoPop1("m4211pa1")			
  		Case "VB"							'���Լ��� 
  			Call OpenBasNoPop1("m5211pa1")		
  		End Select
  	Else
  		Call DisplayMsgBox("17A002","X" , "���౸��","X")
  	End If
End Sub
'==============================================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  1
				.Col = Col
				intIndex = .Value
				.Col = C_BillFG
				.Value = intIndex
		End Select
	End With
End Sub
'==============================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then							
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			DbQuery
		End If
    End if
    
End Sub
'==============================================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_charge_doc_amt
            Call EditModeCheck(frm1.vspdData, Row, C_currency, C_charge_doc_amt,    "A" ,"I", Mode, "X", "X")
        'Case C_xch_rate
         '   Call EditModeCheck(frm1.vspdData, Row, parent.gCurrency, C_xch_rate, "D" ,"I", Mode, "X", "X")          
        Case C_pay_doc_amt
            Call EditModeCheck(frm1.vspdData, Row, C_currency, C_pay_doc_amt, "A" ,"I", Mode, "X", "X")  
        Case C_charge_rate
            Call EditModeCheck(frm1.vspdData, Row, C_currency, C_charge_rate, "D" ,"I", Mode, "X", "X")               
    End Select
End Sub
'==============================================================================================================================
Sub txtChargeFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChargeFrDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtChargeFrDt.focus
	End If
End Sub
'==============================================================================================================================
Sub txtChargeToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtChargeToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtChargeToDt.focus
	End If
End Sub
'==============================================================================================================================
Sub txtChargeFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==============================================================================================================================
Sub txtChargeToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==============================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                               
	
	ggoSpread.Source = frm1.vspdData
	
    with frm1
		  If CompareDateByFormat(.txtChargeFrDt.text,.txtChargeToDt.text,.txtChargeFrDt.Alt,.txtChargeToDt.Alt, _
                   "970025",.txtChargeFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtChargeFrDt.text) <> "" And Trim(.txtChargeToDt.text) <> "" Then
			Call DisplayMsgBox("17a003","X","�߻�����","X")	
			Exit Function
		End if   
	End with

    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If Not chkField(Document, "1") Then						
       Exit Function
    End If
    
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    Call ggoOper.ClearField(Document, "A")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
    
    FncNew = True                                           
	Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear            
    
    If CheckRunningBizProcess = True Then
	  	Exit Function
    End If	                                       
    
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")      
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If   
    
    If DbSave  = False Then Exit Function				                                           
    
    FncSave = True                                                     
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncCopy() 
	Dim IntRetCD
	Dim cur
	
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCopy = False                                                               '��: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then Exit Function
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_currency,C_charge_doc_amt,   "A" ,"I","X","X")
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_currency,C_Vat_rate,"D" ,"I","X","X")         
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_currency,C_vat_doc_amt,"A" ,"I","X","X")         
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_currency,C_pay_doc_amt,"A" ,"I","X","X")         
            Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.ActiveRow,.ActiveRow,C_currency,C_charge_rate,"D" ,"I","X","X")                       

			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	Call frm1.vspdData.SetText(C_charge_no,	frm1.vspdData.ActiveRow, "")
	Call frm1.vspdData.SetText(C_old_posting_flg,	frm1.vspdData.ActiveRow, "0")
    ggoSpread.SpreadLock C_glref_pop, frm1.vspdData.ActiveRow ,C_glref_pop,frm1.vspdData.ActiveRow 
	
	with frm1
		call vspdData_Change(C_pay_type , .vspdData.ActiveRow)  '�������� 
		
		Cur = GetSpreadText(frm1.vspdData,C_currency,frm1.vspdData.ActiveRow,"X","X")
		
		If UCase(Trim(Cur)) = UCase(Trim(parent.gCurrency)) then
			Call .vspdData.SetText(C_xch_rate,	frm1.vspdData.ActiveRow, "1")
			Call .vspdData.SetText(C_calcd,	frm1.vspdData.ActiveRow, "*")
			
			ggoSpread.SSSetProtected	C_xch_rate, .vspdData.Row,.vspdData.Row
		End if
	end with
	
	If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncCancel() 
	Dim iOrgRow, iNewRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCancel = False
    
    frm1.vspdData.Redraw = False
    
    iOrgRow = Frm1.vspdData.Maxrows
    
	If frm1.vspdData.Maxrows < 1	Then Exit Function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo 
    
	'����Ŀ� spreadLock ���� (2003.08.27) - Lee, Eun Hee
	iNewRow = Frm1.vspdData.Maxrows
	
	If iOrgRow = iNewRow Then
		Call SetSpreadLockAfterCancel(Frm1.vspdData.ActiveRow)
	
	
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_charge_doc_amt,   "A" ,"I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_Vat_rate,"D" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_vat_doc_amt,"A" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_pay_doc_amt,"A" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_charge_rate,"D" ,"I","X","X")                       
	
	End If
	
    frm1.vspdData.Redraw = True
    
    If Err.number = 0 Then	
       FncCancel = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                 
End Function
'==============================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
 	
 	Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End If
	
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
	    ggoSpread.Source = .vspdData

	    ggoSpread.InsertRow .vspdData.ActiveRow, imRow

	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    
	    Call ReFormatSpreadCellByCellByCurrency2(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1,Parent.gCurrency,C_charge_doc_amt,   "A" ,"I","X","X")
        Call ReFormatSpreadCellByCellByCurrency2(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1,Parent.gCurrency,C_Vat_rate,"D" ,"I","X","X")         
        Call ReFormatSpreadCellByCellByCurrency2(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1,Parent.gCurrency,C_vat_doc_amt,"A" ,"I","X","X")         
        Call ReFormatSpreadCellByCellByCurrency2(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1,Parent.gCurrency,C_pay_doc_amt,"A" ,"I","X","X")         
        Call ReFormatSpreadCellByCellByCurrency2(.vspdData,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1,Parent.gCurrency,C_charge_rate,"D" ,"I","X","X")
            
	    '----------------------
	    '200404 ������ġ 
	  '  if Trim(frm1.txtprocess_step1.Value) = "" then
		'	Call DisplayMsgBox("17A002", "X", "���౸��", "X")
		'	frm1.txtprocess_step1.focus
		'	Set gActiveElement = document.activeElement
		'	Exit Function
		'end if
		
		'if Trim(frm1.txtpur_grp1.Value) = "" then
		'	Call DisplayMsgBox("17A002", "X", "���ű׷�", "X")
		'	frm1.txtpur_grp1.focus
		'	Set gActiveElement = document.activeElement
		'	Exit Function
		'end if
		
		Dim iInsRow
		For iInsRow = .vspdData.ActiveRow to  .vspdData.ActiveRow + imRow - 1                 '����,��¥ �ʵ� �⺻�� setting 
			
			ggoSpread.SpreadLock C_glref_pop, iInsRow ,C_glref_pop,iInsRow 
			
			Call .vspdData.SetText(C_charge_doc_amt,	iInsRow,	"0")
			Call .vspdData.SetText(C_pay_doc_amt,		iInsRow,	"0")
			Call .vspdData.SetText(C_charge_rate,		iInsRow,	"0")
			Call .vspdData.SetText(C_posting_flag,		iInsRow,	"0")
			Call .vspdData.SetText(C_old_posting_flg,	iInsRow,	"0")
			Call .vspdData.SetText(C_charge_dt,			iInsRow,	UNIFormatDate("<%=GetSvrDate%>"))
			
			If Trim(.txtbas_no1.value) <> "" Then        '�߻��ٰ� ������ȣ�� ������ �ڵ� setting
				Call .vspdData.SetText(C_bas_no,	iInsRow,	Trim(frm1.txtbas_no1.value))
			End If
		Next
		.vspdData.ReDraw = True
		 '200404 ������ġ ��ġ �̵� 
		if Trim(frm1.txtprocess_step1.Value) = "" then
			Call DisplayMsgBox("17A002", "X", "���౸��", "X")
			frm1.txtprocess_step1.focus
			Set gActiveElement = document.activeElement
			Exit Function
		end if
		
		if Trim(frm1.txtpur_grp1.Value) = "" then
			Call DisplayMsgBox("17A002", "X", "���ű׷�", "X")
			frm1.txtpur_grp1.focus
			Set gActiveElement = document.activeElement
			Exit Function
		end if
		
    End With

    Set gActiveElement = document.ActiveElement
	
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If
End Function
'==============================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncDeleteRow = False                                                          '��: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then Exit function	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'-------------------------
	If Err.number = 0 Then	
       FncDeleteRow = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
'==============================================================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
	Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function FncExcel() 
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncExport(parent.C_SINGLEMULTI)		
    Set gActiveElement = document.ActiveElement						
End Function
'==============================================================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)   
    Set gActiveElement = document.ActiveElement	                        
End Function
'==============================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'==============================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadLockAfterQuery()
	
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_currency,C_charge_doc_amt,   "A" ,"I","X","X")
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_currency,C_Vat_rate,"D" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_currency,C_vat_doc_amt,"A" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_currency,C_pay_doc_amt,"A" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_currency,C_charge_rate,"D" ,"I","X","X")                       

End Sub
'==============================================================================================================================
Function FncExit()
	
	Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")           
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    Set gActiveElement = document.ActiveElement
End Function
'==============================================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                           

	Dim strVal
    
    With frm1
    
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtprocess_step=" & .hprocecc_step.value	
			strVal = strVal & "&txtbas_no=" & .hbas_no.value
			strVal = strVal & "&txtpur_grp=" & .hpur_grp.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtChargeFrDt=" & .hdnChargeFrDt.value
			strVal = strVal & "&txtChargeToDt=" & .hdnChargeToDt.value
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey				
			strVal = strVal & "&txtprocess_step=" & Trim(.txtprocess_step.value)
			strVal = strVal & "&txtbas_no=" & Trim(.txtbas_no.value)
			strVal = strVal & "&txtpur_grp=" & Trim(.txtpur_grp.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtChargeFrDt=" & .txtChargeFrDt.text
			strVal = strVal & "&txtChargeToDt=" & .txtChargeToDt.text
		End If
	
		.hdnmaxrow.value = .vspdData.MaxRows
	
		If  LayerShowHide(1) = False Then
		  	Exit Function
		End If

		Call RunMyBizASP(MyBizASP, strVal)										
        
    End With
    
    DbQuery = True
    
End Function
'==============================================================================================================================
Function DbQueryOk()														
	
    lgIntFlgMode = parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									
	If frm1.vspdData.MaxRows < 1 Then
		Call SetToolBar("11101101001111")
		frm1.txtprocess_step.focus
	Else
		Call SetToolBar("11101111001111")
		frm1.vspdData.focus										
	End If
	
	Call RemovedivTextArea
	
	Call SetSpreadLockAfterQuery()
	
End Function
'==============================================================================================================================
Function dbNotFoundOK()
	
    frm1.txtprocess_step1.Value		= frm1.txtprocess_step.Value
	frm1.txtbas_no1.Value 			= frm1.txtbas_no.Value
	frm1.txtpur_grp1.Value 			= frm1.txtpur_grp.Value
	
    frm1.txtprocess_stepNm1.Value	= frm1.txtprocess_stepNm.Value
	frm1.txtpur_grpNm1.Value 		= frm1.txtpur_grpNm.Value
	
	Call ggoOper.LockField(Document, "N")
	
End Function
'==============================================================================================================================
Function dbquerysupplierok(ByVal Row)
	frm1.vspddata.row = Row
	frm1.vspddata.col = C_pay_type
	If Trim(frm1.vspddata.text) <> "" Then
		Call vspdData_Change(C_pay_type, Row)
	End If
	gChangeOpt = ""
	Call ChangeCurOrDt(Row)
	
End Function
'==============================================================================================================================
Function ChangeCurOrDtOk(ByVal Row)
	'�߻��ڱ��ݾ� ���(2003.08.14)
	Call ChangeChargeLocAmt(Row)
	Call ChangeVatType(Row)	'*����*
	'�����ڱ��ݾ� ���(2003.08.14)
	frm1.vspdData.row = Row
	frm1.vspdData.col = C_pay_type
	If Trim(frm1.vspdData.text) <> "" Then
		Call ChangePayLocAmt(Row, frm1.vspdData.text)
	End If
End Function
'==============================================================================================================================
Function DbSave() 
    Dim lRow        
    Dim strVal, strDel
	Dim iColSep, iRowSep
	
	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen  '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
	
	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size

	Dim iTmpDBuffer          '������ ���� [����] 
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size
	
    DbSave = False                                                          
    
	Call DisableToolBar(Parent.TBC_SAVE)                                          '��: Disable Save Button Of ToolBar

    If LayerShowHide(1) = False Then
		Exit Function
	End If 
	
	iColSep = Parent.gColSep													
	iRowSep = Parent.gRowSep													
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]
	
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	
	frm1.txtMode.value = Parent.UID_M0002
	strVal = ""
	strDel = ""
    
	With frm1
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
					'PreCheck
					If UsrFnPreCheck(lRow) = False Then
						Call RemovedivTextArea
						Exit Function
					End If
					
					.vspdData.Row = lRow
					.vspdData.Col = 0
					If .vspdData.Text=ggoSpread.InsertFlag then
						strVal = "C" & iColSep		'0		
					ElseIf .vspdData.Text=ggoSpread.UpdateFlag then
						strVal = "U" & iColSep
					End If 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_posting_flag,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_charge_no,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_charge_type,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_bp_cd,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_BuildCd,lRow, "X","X"))					& iColSep 
					strVal = strVal & UNIConvDate(GetSpreadText(frm1.vspdData,C_charge_dt,lRow, "X","X"))			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_vat_type,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_tax_biz_area,lRow, "X","X"))				& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_currency,lRow, "X","X"))					& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_charge_doc_amt,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_xch_rate,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_Vat_rate,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_vat_doc_amt,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_vat_loc_amt,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_pay_type,lRow, "X","X"))					& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_pay_doc_amt,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & UNIConvDate(GetSpreadText(frm1.vspdData,C_pay_due_dt,lRow, "X","X"))			& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_charge_rate,lRow, "X","X"),0)		& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_cost_flag,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_bank_cd,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_bank_acct,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_note_no,lRow, "X","X"))					& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_prpaym_no,lRow, "X","X"))					& iColSep 
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_pp_xch_rt,lRow, "X","X"),0)			& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_calcd,lRow, "X","X"))						& iColSep
					'�߻��ڱ��ݾ�, �����ڱ��ݾ� �߰�(2003.08.14)
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_charge_loc_amt,lRow, "X","X"),0)		& iColSep
					strVal = strVal & UNIConvNum(GetSpreadText(frm1.vspdData,C_pay_loc_amt,lRow, "X","X"),0)		& iColSep    
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_remark,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_bas_no,lRow, "X","X"))						& iColSep 
					strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_old_posting_flg,lRow, "X","X"))			& iColSep 
					strVal = strVal & lRow & iRowSep
				Case ggoSpread.DeleteFlag
					strDel = "D" & iColSep
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_posting_flag,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_charge_no,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_charge_type,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_bp_cd,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_BuildCd,lRow, "X","X"))					& iColSep 
					strDel = strDel & UNIConvDate(GetSpreadText(frm1.vspdData,C_charge_dt,lRow, "X","X"))			& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_vat_type,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_tax_biz_area,lRow, "X","X"))				& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_currency,lRow, "X","X"))					& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_charge_doc_amt,lRow, "X","X"),0)		& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_xch_rate,lRow, "X","X"),0)			& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_Vat_rate,lRow, "X","X"),0)			& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_vat_doc_amt,lRow, "X","X"),0)		& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_vat_loc_amt,lRow, "X","X"),0)		& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_pay_type,lRow, "X","X"))					& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_pay_doc_amt,lRow, "X","X"),0)		& iColSep 
					strDel = strDel & UNIConvDate(GetSpreadText(frm1.vspdData,C_pay_due_dt,lRow, "X","X"))			& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_charge_rate,lRow, "X","X"),0)		& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_cost_flag,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_bank_cd,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_bank_acct,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_note_no,lRow, "X","X"))					& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_prpaym_no,lRow, "X","X"))					& iColSep 
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_pp_xch_rt,lRow, "X","X"),0)			& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_calcd,lRow, "X","X"))						& iColSep 
					'�߻��ڱ��ݾ�, �����ڱ��ݾ� �߰�(2003.08.14)
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_charge_loc_amt,lRow, "X","X"),0)		& iColSep
					strDel = strDel & UNIConvNum(GetSpreadText(frm1.vspdData,C_pay_loc_amt,lRow, "X","X"),0)		& iColSep  
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_remark,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_bas_no,lRow, "X","X"))						& iColSep 
					strDel = strDel & Trim(GetSpreadText(frm1.vspdData,C_old_posting_flg,lRow, "X","X"))			& iColSep 
					strDel = strDel & lRow & iRowSep
		    End Select
			
			.vspdData.Row = lRow
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
					                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
					 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If
					       
			         iTmpCUBufferCount = iTmpCUBufferCount + 1
					      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   
			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
					          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
					       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
					         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			End Select
		Next
	End With
	
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If  
	
	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If
	
	'------ Developer Coding part (End ) -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	If Err.number = 0 Then	 
	   DbSave = True                                                             '��: Processing is OK
	End If

	Set gActiveElement = document.ActiveElement                    
End Function
'==============================================================================================================================
Function DbSaveOk()												
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()

End Function
'==============================================================================================================================
Function RemovedivTextArea()
	Dim ii
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
'==============================================================================================================================
Function UsrFnPreCheck(Byval lRow)
	Dim chargedt, payduedt
	Dim paydocamt, chargedocamt, vatdocamt
	Dim paylocamt, chargelocamt, vatlocamt
	
	UsrFnPreCheck = False
	With frm1
		
		chargedt = GetSpreadText(frm1.vspdData,C_charge_dt,lRow,"X","X")
		payduedt = GetSpreadText(frm1.vspdData,C_pay_due_dt,lRow,"X","X")
						
		if payduedt <> "" then
			if UniConvDateToYYYYMMDD(chargedt,parent.gDateFormat,"") > UniConvDateToYYYYMMDD(payduedt,parent.gDateFormat,"") then 
				Call DisplayMsgBox("970023", "X",lRow&"Row,"&"������","�߻���")
				Call LayerShowHide(0)
				Exit Function
			end if
		end if
						
		paydocamt = GetSpreadText(frm1.vspdData,C_pay_doc_amt,lRow,"X","X")
		chargedocamt = GetSpreadText(frm1.vspdData,C_charge_doc_amt,lRow,"X","X")
		vatdocamt = GetSpreadText(frm1.vspdData,C_vat_doc_amt,lRow,"X","X")
						
		If UNICDbl(paydocamt) > UNICDbl(chargedocamt) + UNICDbl(vatdocamt) Then 
			Call DisplayMsgBox("970023", "X",lRow&"Row,"&"�߻��ݾ�","���ޱݾ�")
			Call LayerShowHide(0)
			Exit Function
		End If
						
		If Trim(GetSpreadText(frm1.vspdData,C_pay_type,lRow,"X","X")) <> "" Then
			If (UNICDbl(paydocamt) = "" Or UNICDbl(paydocamt) = 0) Then
				Call DisplayMsgBox("970021","X" ,lRow&"Row,"&"���ޱݾ�", "X")
				Call LayerShowHide(0)
				Exit Function
			End If
		End If
						
		If (Trim(GetSpreadText(frm1.vspdData,C_charge_doc_amt,lRow,"X","X")) = "" Or UNICDBl(GetSpreadText(frm1.vspdData,C_charge_doc_amt,lRow,"X","X")) = 0) Then
			Call DisplayMsgBox("970021","X" ,lRow&"Row,"&"�߻��ݾ�","X")
			Call LayerShowHide(0)
			Exit Function
		End If
		
		'-- issue for 8550 by Byun Jee Hyun 2004-08-10
		If Trim(GetSpreadText(frm1.vspdData,C_pay_type,lRow,"X","X")) = "" Then
			If (UNICDbl(paydocamt) <> "" and UNICDbl(paydocamt) <> 0) Then
				Call DisplayMsgBox("17A003","X" ,lRow&"Row,"&"���ޱݾ�", "X")
				Call LayerShowHide(0)
				Exit Function
			End If
		End If
	End With
	
	UsrFnPreCheck = True
End Function
'==============================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���Ű��</font></td>
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
									<TD CLASS="TD5" NOWRAP>���౸��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ALT="���౸��" TYPE=TEXT NAME="txtprocess_step" SIZE=10 MAXLENGTH=5 tag="12NXXU" onChange="vbscript:changeProcess_step()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnprocess_step" align=top TYPE="BUTTON" ONCLICK="vbscript:Openprocess_step()">
														   <INPUT ALT="���౸��" TYPE=TEXT NAME="txtprocess_stepNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ALT="���ű׷�" TYPE=TEXT NAME="txtpur_grp" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnpur_grp" align=top TYPE="BUTTON"ONCLICK="vbscript:Openpur_grp()">
    													   <INPUT ALT="���ű׷�" TYPE=TEXT NAME="txtpur_grpNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�߻��ٰ� ������ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT ALT="�߻��ٰ� ������ȣ" CLASS="clstxt" TYPE=TEXT NAME="txtbas_no" SIZE=32 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnbas_no" align=top TYPE="BUTTON" ONCLICK="vbscript:btnbas_noOnClick()"></TD>
									<TD CLASS="TD5" NOWRAP>�߻�����</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m6111ma2_fpDateTime1_txtChargeFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m6111ma2_fpDateTime1_txtChargeToDt.js'></script>
												</td>
											</tr>
										</table>
									</TD>
								</TR> 
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>���౸��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT ALT="���౸��" TYPE=TEXT NAME="txtprocess_step1" SIZE=10 MAXLENGTH=5 tag="23NXXU" onChange="vbscript:changeProcess_step1()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnprocess_step" align=top TYPE="BUTTON" ONCLICK="vbscript:Openprocess_step1()">
													   <INPUT ALT="���౸��" TYPE=TEXT NAME="txtprocess_stepNm1" SIZE=20 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT ALT="���ű׷�" TYPE=TEXT NAME="txtpur_grp1" SIZE=10 MAXLENGTH=4 tag="23NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnpur_grp" align=top TYPE="BUTTON"ONCLICK="vbscript:Openpur_grp1()">
													   <INPUT ALT="���ű׷�" TYPE=TEXT NAME="txtpur_grpNm1" SIZE=20 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�߻��ٰ� ������ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT ALT="�߻��ٰ� ������ȣ" CLASS="clstxt" TYPE=TEXT NAME="txtbas_no1" SIZE=32 MAXLENGTH=18 tag="25NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnbas_no1" align=top TYPE="BUTTON" ONCLICK="vbscript:btnbas_no1OnClick()"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/m6111ma2_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hprocecc_step" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hReqStatus" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hbas_no" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hpur_grp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hbas_doc_no" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdninterface_Account" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnChargeFrDt"  tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hdnChargeToDt"  tag="14" TABINDEX="-1">
<P ID="divTextArea"></P>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
