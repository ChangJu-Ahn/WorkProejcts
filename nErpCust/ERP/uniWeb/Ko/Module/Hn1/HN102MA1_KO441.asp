<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : HN102MA1_KO441
*  4. Program Name         : HN102MA1_KO441
*  5. Program Desc         : ��/�󿩳����ݿ�
*  6. Comproxy List        :
*  7. Modified date(First) : 2008/01/17
*  8. Modified date(Last)  :
*  9. Modifier (First)     : 
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "HN102MB1_KO441.asp"	
Const BIZ_PGM_ID1     = "HN102MB1_KO441.asp"						           '��: Biz Logic ASP Name
Const CookieSplit = 1233
Const C_SHEETMAXROWS    =   21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop 
Dim lgStrComDateType		'Company Date Type�� ����(��� Mask�� �����.)
Dim lgType

Dim C_EMP_NO				'���
Dim C_EMP_NM				'����
Dim C_DEPT_NM				'�μ���
Dim C_DEPT_CD				'�μ��ڵ�
Dim C_PROV_TYPE				'���ޱ��� 
Dim C_PROV_TYPE_HIDDEN		'���ޱ����ڵ�
Dim C_PROV_DT				'������			
Dim	C_PAY_BONUS_TOT_AMT		'�޿��Ѿ�,���Ѿ�
Dim C_TAX_AMT				'�޿������Ѿ�,�󿩰����Ѿ�
Dim C_NONTAX_TOT_AMT		'������Ѿ�
Dim C_PROV_TOT_AMT			'�����Ѿ�
Dim C_SUB_TOT_AMT			'�����Ѿ�
Dim C_REAL_PROV_AMT			'�����޾�
Dim C_INCOME_TAX			'�ҵ漼	
Dim C_RES_TAX				'�ֹμ�
Dim C_AUNT					'���ο���
Dim C_MED_INSURE			'�ǰ�����
Dim C_EMP_INSURE			'��뺸��
Dim C_PAY_YYMM				'�ش���	

Dim C_OCPT_TYPE				'�����ڵ�
Dim C_PAY_GRD1				'����
Dim C_PAY_GRD2				'ȣ��
Dim C_PAY_CD				'�޿�����
Dim C_TAX_CD				'���ױ���
Dim C_INTERNAL_CD			'���κμ��ڵ�

'------------------------------------
Dim C_ALLOW_CD				'�����ڵ�
Dim C_ALLOW_NM				'�����
Dim C_ALLOW					'����ݾ�
'----------------------------------
Dim C_SUB_CD				'�����ڵ�
Dim C_SUB_NM				'������
Dim C_SUB_AMT				'�����ݾ�

'lgType = "A" 'hanc


'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)	 

	Select Case pvSpdNo
		   Case "A"
				C_EMP_NO				=1		'���
				C_EMP_NM				=2		'����

				C_DEPT_CD				=3		'�μ��ڵ�
				C_DEPT_NM				=4		'�μ���				
				
				C_PROV_TYPE				=5		'���ޱ��� 
				C_PROV_TYPE_HIDDEN		=6		'���ޱ����ڵ�
				C_PROV_DT				=7		'������			
				C_PAY_BONUS_TOT_AMT		=8		'�޿��Ѿ�,���Ѿ�
				C_TAX_AMT				=9		'�޿������Ѿ�,�󿩰����Ѿ�
				C_NONTAX_TOT_AMT		=10		'������Ѿ�
				C_PROV_TOT_AMT			=11		'�����Ѿ�
				C_SUB_TOT_AMT			=12		'�����Ѿ�
				C_REAL_PROV_AMT			=13		'�����޾�
				C_INCOME_TAX			=14		'�ҵ漼	
				C_RES_TAX				=15		'�ֹμ�
				C_AUNT					=16		'���ο���
				C_MED_INSURE			=17		'�ǰ�����
				C_EMP_INSURE			=18		'��뺸��				
				
				C_OCPT_TYPE				=19		'�����ڵ�
				C_PAY_GRD1				=20		'����
				C_PAY_GRD2				=21		'ȣ��
				C_PAY_CD				=22		'�޿�����
				C_TAX_CD				=22		'���ױ���
				C_INTERNAL_CD			=23		'���κμ��ڵ�				

		   Case "B"				
				C_PAY_YYMM			=1			'�ش���
				C_EMP_NO			=2			'���
				C_PROV_TYPE			=3			'��������
				C_PROV_TYPE_HIDDEN	=4			'���������ڵ�
				C_ALLOW_CD			=5			'�����ڵ�
				C_ALLOW_NM			=6			'�����
				C_ALLOW				=7			'����ݾ�

		   Case "C"
				C_PAY_YYMM			=1			'�ش���
				C_EMP_NO			=2			'���
				C_PROV_TYPE			=3			'��������
				C_PROV_TYPE_HIDDEN	=4			'���������ڵ�
				C_SUB_CD			=5			'�����ڵ�
				C_SUB_NM			=6			'������
				C_SUB_AMT			=7			'�����ݾ�

	End Select	

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False											'��: Indicates that no value changed
	lgIntGrpCount     = 0												'��: Initializes Group View Size
    lgStrPrevKey      = ""												'��: initializes Previous Key
    lgSortKey         = 1												'��: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	
	Dim strYear
	Dim strMonth
	Dim strDay
	
	
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
	
	With frm1
		.txtYYMM.Year = strYear
		.txtYYMM.Month = strMonth
		Call  ggoOper.FormatDate(.txtYYMM,  parent.gDateFormat, 2)					
		.txtYYMM.focus
	End With

End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

    Dim iCodeArr 
    Dim iNameArr
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0040", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
	iCodeArr = lgF0
    iNameArr = lgF1
    
	Call SetCombo2(frm1.txtProv_cd, iCodeArr, iNameArr, Chr(11))

End Sub


'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
 
	With Frm1		
		lgKeyStream  = Trim(.txtYYMM.text) & parent.gColSep   
		lgKeyStream  = lgKeyStream & Trim(.txtProv_cd.Value) & parent.gColSep 
		lgKeyStream  = lgKeyStream & Trim(.txtEmp_No.Value) & parent.gColSep  	
	End With
	
End Sub        



'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables(lgType)  

	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	   .ReDraw = false
		Select Case lgType
			   Case "A"
					.MaxCols = C_INTERNAL_CD + 1											 ' ��:��: Add 1 to Maxcols
			   Case "B"
					.MaxCols = C_ALLOW + 1                                                      
			   Case "C"
					.MaxCols = C_SUB_AMT + 1     
		End Select

		Call ggoSpread.ClearSpreadData()
		Call GetSpreadColumnPos(lgType)  

		Select Case lgType
			   Case "A"	
					'Call AppendNumberPlace("6","2","0")								
					
					ggoSpread.SSSetEdit  C_EMP_NO				, "���"			, 10,,,13,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_EMP_NM				, "����"			, 10,,,13,2		'Lock/ Edit	
					ggoSpread.SSSetEdit  C_DEPT_CD				, "�μ��ڵ�"		, 10,,,40,2		'Lock/ Edit		
					ggoSpread.SSSetEdit  C_DEPT_NM				, "�μ���"			, 10,,,40,2		'Lock/ Edit							
					ggoSpread.SSSetEdit  C_PROV_TYPE			, "��������"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PROV_TYPE_HIDDEN		, "��������Code"	, 12,,,1,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PROV_DT				, "������"			, 12,,,13,2		'Lock/ Edit
					ggoSpread.SSSetFloat C_PAY_BONUS_TOT_AMT	, "��/���Ѿ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_TAX_AMT				, "�����Ѿ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_NONTAX_TOT_AMT		, "������Ѿ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec					
					ggoSpread.SSSetFloat C_PROV_TOT_AMT			, "�����Ѿ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_SUB_TOT_AMT			, "�����Ѿ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_REAL_PROV_AMT		, "�����޾�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec					
					ggoSpread.SSSetFloat C_INCOME_TAX			, "�ҵ漼"			,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_RES_TAX				, "�ֹμ�"			,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_AUNT					, "���ο���"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_MED_INSURE			, "�ǰ�����"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
					ggoSpread.SSSetFloat C_EMP_INSURE			, "��뺸��"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
					
					ggoSpread.SSSetEdit  C_OCPT_TYPE			, "�����ڵ�"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PAY_GRD1				, "����"			, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PAY_GRD2				, "ȣ��"			, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PAY_CD				, "�޿�����"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_TAX_CD				, "���ױ���"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_INTERNAL_CD			, "���κμ��ڵ�"	, 12,,,50,2		'Lock/ Edit

					Call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)
					Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)
					Call ggoSpread.SSSetColHidden(C_OCPT_TYPE,C_DEPT_CD,True)
					Call ggoSpread.SSSetColHidden(C_PAY_GRD1,C_PROV_TYPE_HIDDEN,True)
					Call ggoSpread.SSSetColHidden(C_PAY_GRD2,C_DEPT_CD,True)
					Call ggoSpread.SSSetColHidden(C_PAY_CD,C_PROV_TYPE_HIDDEN,True)
					Call ggoSpread.SSSetColHidden(C_TAX_CD,C_DEPT_CD,True)
					Call ggoSpread.SSSetColHidden(C_INTERNAL_CD,C_PROV_TYPE_HIDDEN,True)
					
		
			   Case "B"									
					ggoSpread.SSSetEdit  C_EMP_NO			, "���"			, 10,,,13,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PAY_YYMM			, "�ش���"		, 12,,,13,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PROV_TYPE		, "��������"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PROV_TYPE_HIDDEN , "��������Code"	, 12,,,1,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_ALLOW_CD			, "�����ڵ�"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_ALLOW_NM			, "�����"			, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetFloat C_ALLOW			, "����ݾ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
					
					Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)
					
			   Case "C"
		   
					ggoSpread.SSSetEdit  C_EMP_NO			, "���"			, 10,,,13,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PAY_YYMM			, "�ش���"		, 12,,,13,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PROV_TYPE		, "��������"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_PROV_TYPE_HIDDEN , "��������Code"	, 12,,,1,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_SUB_CD			, "�����ڵ�"		, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetEdit  C_SUB_NM			, "������"			, 12,,,50,2		'Lock/ Edit
					ggoSpread.SSSetFloat C_SUB_AMT			, "�����ݾ�"		,12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	
					
					Call ggoSpread.SSSetColHidden(C_PROV_TYPE_HIDDEN,C_PROV_TYPE_HIDDEN,True)

		End Select
		
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		.Redraw = True 
		 
      
    End With
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
	
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
	Dim iCurColumnPos
    
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
    Select Case UCase(pvSpdNo)
		   Case "A"
				C_EMP_NO				=iCurColumnPos(1)		'���
				C_EMP_NM				=iCurColumnPos(2)		'����
				
				C_DEPT_CD				=iCurColumnPos(3)		'�μ��ڵ�
				C_DEPT_NM				=iCurColumnPos(4)		'�μ���
				
				C_PROV_TYPE				=iCurColumnPos(5)		'���ޱ��� 
				C_PROV_TYPE_HIDDEN		=iCurColumnPos(6)		'���ޱ����ڵ�
				C_PROV_DT				=iCurColumnPos(7)		'������			
				C_PAY_BONUS_TOT_AMT		=iCurColumnPos(8)		'�޿��Ѿ�,���Ѿ�
				C_TAX_AMT				=iCurColumnPos(9)		'�޿������Ѿ�,�󿩰����Ѿ�
				C_NONTAX_TOT_AMT		=iCurColumnPos(10)		'������Ѿ�
				C_PROV_TOT_AMT			=iCurColumnPos(11)		'�����Ѿ�
				C_SUB_TOT_AMT			=iCurColumnPos(12)		'�����Ѿ�
				C_REAL_PROV_AMT			=iCurColumnPos(13)		'�����޾�
				C_INCOME_TAX			=iCurColumnPos(14)		'�ҵ漼	
				C_RES_TAX				=iCurColumnPos(15)		'�ֹμ�
				C_AUNT					=iCurColumnPos(16)		'���ο���
				C_MED_INSURE			=iCurColumnPos(17)		'�ǰ�����
				C_EMP_INSURE			=iCurColumnPos(18)		'��뺸��			

				C_OCPT_TYPE				=iCurColumnPos(19)		'�����ڵ�
				C_PAY_GRD1				=iCurColumnPos(20)		'����
				C_PAY_GRD2				=iCurColumnPos(21)		'ȣ��
				C_PAY_CD				=iCurColumnPos(22)		'�޿�����
				C_TAX_CD				=iCurColumnPos(23)		'���ױ���
				C_INTERNAL_CD			=iCurColumnPos(24)		'���κμ��ڵ�				

		   Case "B"
				C_PAY_YYMM			=iCurColumnPos(1)			'�ش���
				C_EMP_NO			=iCurColumnPos(2)			'���
				C_PROV_TYPE			=iCurColumnPos(3)			'��������
				C_PROV_TYPE_HIDDEN	=iCurColumnPos(4)			'���������ڵ�
				C_ALLOW_CD			=iCurColumnPos(5)			'�����ڵ�
				C_ALLOW_NM			=iCurColumnPos(6)			'�����
				C_ALLOW				=iCurColumnPos(7)			'����ݾ�

		   Case "C"
				C_PAY_YYMM			=iCurColumnPos(1)			'�ش���
				C_EMP_NO			=iCurColumnPos(2)			'���
				C_PROV_TYPE			=iCurColumnPos(3)			'��������
				C_PROV_TYPE_HIDDEN	=iCurColumnPos(4)			'���������ڵ�
				C_SUB_CD			=iCurColumnPos(5)			'�����ڵ�
				C_SUB_NM			=iCurColumnPos(6)			'������
				C_SUB_AMT			=iCurColumnPos(7)			'�����ݾ�

			
    End Select    
End Sub
'======================================================================================================
' Function Name : vspdData_ScriptLeaveCell
' Function Desc : ��(YYYY).��(MM) check
'======================================================================================================
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear																			'��: Clear err status
	Call LoadInfTB19029																	'��: Load table , B_numeric_format

    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")												'��: Lock Field
	   
	lgType = "A"
    Call  InitSpreadSheet																'Setup the Spread sheet
		
    Call  InitVariables																	'Initializes local global variables
	
    Call  FuncGetAuth(gStrRequestMenuID ,  parent.gUsrID, lgUsrIntCd)					' �ڷ����:lgUsrIntCd ("%", "1%")
	
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")													'��ư ���� ����

	Call InitComboBox
	Call CookiePage (0)
   
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()	

    Dim IntRetCD 
    Dim strwhere

    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData	

	If  txtEmp_no_Onchange() then
        Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables	
    Call MakeKeyStream("X")
	
    If DbQuery = False Then
       Exit Function	
    End If																		 '��: Query db data
       
    FncQuery = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Name : FncIFQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncIFQuery()	

    Dim IntRetCD 
    Dim strwhere

    FncIFQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData	

	If  txtEmp_no_Onchange() then
        Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables	
    Call MakeKeyStream("X")
	
    If DbIFQuery = False Then
       Exit Function	
    End If																		 '��: Query db data
    	
    FncIFQuery = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()

	Dim IntRetCD 
    
    FncDelete = False																'��: Processing is NG
    
    Err.Clear																		'��: Clear err status
    	
	IF frm1.txtYYMM.Text = "" then 
	    Call  DisplayMsgBox("970021","X","��/�󿩳��","X")							'��/�󿩳���� Ȯ���Ͻʽÿ�.    
	    frm1.txtYYMM.focus 
	    Exit Function    
	End If

	If lgIntFlgMode <>  parent.OPMD_UMODE Then										'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")									'��:
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
	If IntRetCD = vbNo Then															'------ Delete function call area ------ 
		Exit Function	
	End If
       
    Call MakeKeyStream("X")
	
	'Call DisableToolBar( parent.TBC_DELETE)
	
	If DbDelete = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																		'��: Query db data
    
    FncDelete = True                                                              '��: Processing is OK

End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
	
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
    ggoSpread.Source = frm1.vspdData
	
	
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If     	 		

	IF frm1.txtYYMM.Text = "" then 
	    Call  DisplayMsgBox("970021","X","��/�󿩳��","X")						'��/�󿩳���� Ȯ���Ͻʽÿ�.    
	    frm1.txtYYMM.focus 
	    Exit Function    
	End If

	Dim lRow

	'���Ǻ��� ��/�����ڿ� ��/�󿩳������Ÿ�� Excel ����Ÿ�� ��ġ���� ������ ������� ����
	With Frm1
		
        For lRow = 1 To .vspdData.MaxRows
			
            .vspdData.Row = lRow
            .vspdData.Col = 0
			
			Select Case lgType
				   Case "A"
'						.vspdData.Col = C_PAY_YYMM
'						If Trim(Replace(.txtYYMM,"-","")) <> Trim(.vspdData.Text) Then	
'							MsgBox "��/�󿩳�� ����Ÿ�� ��ġ���� �׽��ϴ�.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If											

						.vspdData.Col = C_PROV_TYPE
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","��������","X")
							Exit Function
						End If

						.vspdData.Col = C_EMP_NO
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","�����ȣ","X")
							Exit Function
						End If	

				   Case "B"
						
'						.vspdData.Col = C_PAY_YYMM
'						If Trim(Replace(.txtYYMM,"-","")) <> Trim(.vspdData.Text) Then				
'							MsgBox "��/�󿩳�� ����Ÿ�� ��ġ���� �׽��ϴ�.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If
					
						.vspdData.Col = C_PROV_TYPE
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","��������","X")
							Exit Function
						End If

						.vspdData.Col = C_EMP_NO
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","�����ȣ","X")
							Exit Function
						End If	

						.vspdData.Col = C_ALLOW_CD
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","�����ڵ�","X")
							Exit Function
						End If	

				   Case "C"

'						.vspdData.Col = C_PAY_YYMM
'						If Trim(Replace(.txtYYMM,"-","")) <> Trim(.vspdData.Text) Then				
'							MsgBox "��/�󿩳�� ����Ÿ�� ��ġ���� �׽��ϴ�.    ", vbExclamation, "uniERPII[Warning]"							
'							Exit Function
'						End If
					
						.vspdData.Col = C_PROV_TYPE
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","��������","X")
							Exit Function
						End If

						.vspdData.Col = C_EMP_NO
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","�����ȣ","X")
							Exit Function
						End If	

						.vspdData.Col = C_SUB_CD
						If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
							Call DisplayMsgBox("970000","X","�����ڵ�","X")
							Exit Function
						End If	

			End Select
			
					            
        Next

	End With
       
    Call MakeKeyStream("X")

	'Call DisableToolBar( parent.TBC_SAVE)
	
	If DbSAVE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																		'��: Query db data
    
    FncSave = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
     ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
             ggoSpread.CopyRow
			 SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
     ggoSpread.Source = frm1.vspdData	
     ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
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
        
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	 ggoSpread.Source = frm1.vspdData 
    	lDelRows =  ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_MULTI, False)
End Function
'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()          
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub
'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
    
	Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 	

    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	if LayerShowHide(1) = false then
		exit Function
	end if

	Dim strVal	
   
    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	
	if LayerShowHide(1) = false then		
		exit Function
	end if
		    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey                 '��: Next key tag
		strVal = strVal     & "&htxtFileGubun="		 & lgType	
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If
	
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbIFQuery
' Desc : This function is called by FncIFQuery
'========================================================================================================
Function DbIFQuery() 	

    DbIFQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	if LayerShowHide(1) = false then
		exit Function
	end if

	Dim strVal	
   
    DbIFQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	
	if LayerShowHide(1) = false then		
		exit Function
	end if
		    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0004						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey                 '��: Next key tag
		strVal = strVal     & "&htxtFileGubun="		 & lgType	
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If
	
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbIFQuery = True
    
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave() 
	
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lStartRow   
    Dim lEndRow     
	Dim strVal, strDel
	
    DbSave = False                                                          
	
	If LayerShowHide(1) = false then		
		exit Function
	End if
	
	With frm1
		.txtMode.value      = parent.UID_M0002                                        '��: Save
		.txtFlgMode.value   = lgIntFlgMode
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	
	With Frm1

       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
     
           Select Case .vspdData.Text
                  Case  ggoSpread.InsertFlag																	'��: Insert                  
															strVal = strVal & "C" & parent.gColSep				'array(0)
															strVal = strVal & lRow & parent.gColSep
						Select Case lgType
							   Case "A"									

									.vspdData.Col = C_EMP_NO				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_DEPT_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep				
									.vspdData.Col = C_PROV_TYPE_HIDDEN		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_PROV_DT				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_PAY_BONUS_TOT_AMT		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_TAX_AMT				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_NONTAX_TOT_AMT		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep				
									.vspdData.Col = C_PROV_TOT_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_SUB_TOT_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_REAL_PROV_AMT			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_INCOME_TAX			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_RES_TAX				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_AUNT					: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_MED_INSURE			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_EMP_INSURE			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									
									.vspdData.Col = C_OCPT_TYPE				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_PAY_GRD1				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_PAY_GRD2				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_PAY_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_TAX_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_INTERNAL_CD			: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep	'>>AIR
									
									.vspdData.Col = C_DEPT_NM				: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep	'>>AIR
							   Case "B"
									
									.vspdData.Col = C_PAY_YYMM				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_EMP_NO				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep				
									.vspdData.Col = C_PROV_TYPE_HIDDEN		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_ALLOW_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_ALLOW					: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep									
							   Case "C"
									.vspdData.Col = C_PAY_YYMM				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_EMP_NO				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep				
									.vspdData.Col = C_PROV_TYPE_HIDDEN		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_SUB_CD				: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
									.vspdData.Col = C_SUB_AMT				: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep									
						End Select
                   
               
                    lGrpCnt = lGrpCnt + 1 
			End Select
       Next
	   	  
	   .htxtYYMM.value		= .txtYYMM.text
	   .htxtProvCD.value	= .txtProv_CD.Value
	   .htxtFileGubun.value = lgType		
	   
	   .txtMaxRows.value    = lGrpCnt-1	
	   .txtSpread.value     = strDel & strVal
			
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	
    DbSave = True                                                           
    
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()   
    
	Dim strVal

	DbDelete = False																'��: Processing is NG
    
	If LayerShowHide(1) = false then		
		Exit Function
	End if

	With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0003						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                  '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey					'��: Next key tag
		strVal = strVal     & "&htxtFileGubun="		 & lgType	
    End With	

	Call RunMyBizASP(MyBizASP, strVal)
																					'��: Query db data
    DbDelete = True																	'��: Processing is OK

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
    
	lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'��: Lock field
    'Call InitData()
	Call SetToolbar("110011110011111")	 	
	frm1.vspdData.focus
	
End Function

'======================================================================================================
'	Name : DBAutoQueryOk()
'	Description : HN101BB2_KO441.asp ���� Query OK�� ��
'=======================================================================================================
Sub DBAutoQueryOk()
    Dim lRow
	Dim intIndex
	Dim daytimeVal 
	Dim strSub_type 
    
    With Frm1
        .vspdData.ReDraw = false
         ggoSpread.Source = .vspdData
   
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            .vspdData.Text =  ggoSpread.InsertFlag
        Next
            .vspdData.ReDraw = TRUE
        
    End With 
    ggoSpread.ClearSpreadData "T"
     Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = frm1.vspdData	
	ggoSpread.ClearSpreadData
	Call RemovedivTextArea	
    Call InitVariables															'��: Initializes local global variables
	ggoSpread.ClearSpreadData
    
    Call DisplayMsgBox("183114","X","X","X")
		
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	ggoSpread.Source = frm1.vspdData	
	ggoSpread.ClearSpreadData
	Call RemovedivTextArea	
    Call InitVariables															'��: Initializes local global variables
	ggoSpread.ClearSpreadData
    
    Call DisplayMsgBox("183114","X","X","X")
End Function

'----------------------------------------  OpenEmptName()  ------------------------------------------
'	Name : OpenEmptName()                                                         <==== ����/��� �˾� 
'	Description : Employee PopUp
'------------------------------------------------------------------------------------------------
Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	End If
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EmpNo
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetEmp()  ------------------------------------------------
'	Name : SetEmp()
'	Description : Employee Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_EmpNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EmpNo
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()       
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
		Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	   
        Case "2"
            arrParam(0) = "���ޱ��� �˾�"				' �˾� ��Ī 
	        arrParam(1) = "B_MINOR"				 		' TABLE ��Ī 
	        arrParam(2) = frm1.txtprov_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtprov_nm.value		' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD NOT IN (" & FilterVar("B", "''", "S") & " ," & FilterVar("C", "''", "S") & " )"			' Where Condition							' Where Condition
	        arrParam(5) = "���ޱ���"					' TextBox ��Ī 
	
            arrField(0) = "minor_cd"					' Field��(0)
            arrField(1) = "minor_nm"				    ' Field��(1)
    
            arrHeader(0) = "���ޱ����ڵ�"				' Header��(0)
            arrHeader(1) = "���ޱ��и�"
	
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtprov_cd.focus
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function



'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split( parent.gDateFormat, parent.gComDateType)
	
	If  parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType =  parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 
	
End Function


'========================================================================================================
'   Event Name : txtEmp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()

    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
         IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
                
        If  IntRetCd < 0 then
            If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")			'�ش����� �������� �ʽ��ϴ�.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if  
End Function


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
      gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row
     
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   
    
End Sub

'=======================================
'   Event Name :txtYYMM_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================
Sub txtYYMM_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("D")    
        frm1.txtYYMM.Action = 7
        frm1.txtYYMM.focus
    End If
End Sub

'==========================================================================================
'   Event Name : txtYYMM_KeyDown()
'   Event Desc : ��ȸ���Ǻ��� txtYYMM_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
Sub txtYYMM_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call mainQuery()
End Sub

'==========================================================================================
'   Event Name : rbo_type1_OnClick()
'   Event Desc : radio button Click�� Grid Setting
'==========================================================================================
Sub rdoCase1_OnClick()
    lgType = "A"   
'    call Form_Load()
    Call InitSpreadSheet
End Sub

Sub rdoCase2_OnClick()
    lgType = "B"
'    call Form_Load()
    Call InitSpreadSheet
End Sub

Sub rdoCase3_OnClick()
    lgType = "C"
'    call Form_Load()
    Call InitSpreadSheet
End Sub


'========================================================================================
 ' Function Name : RemovedivTextArea
 ' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
 Function RemovedivTextArea()
 
 	Dim ii
 		
 	For ii = 1 To divTextArea.children.length
 	    divTextArea.removeChild(divTextArea.children(0))
 	Next
 
 End Function




'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub


'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�޻󿩳����ݿ�</font></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>width=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
					    
						    <TR>								
								<TD CLASS="TD5" NOWRAP>�ݿ�����</TD>
								<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X" checked>
								<LABEL FOR="rdoCase1">��/�󿩳���</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X">
								<LABEL FOR="rdoCase2">���系��</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase3" TAG="1X">
								<LABEL FOR="rdoCase3">��������</LABEL></TD>

								<TD CLASS=TD5 NOWRAP>��/�󿩳��</TD>
								<TD CLASS=TD6 NOWRAP>
								<OBJECT classid=<%=gCLSIDFPDT%> id=txtYYMM NAME="txtYYMM" CLASS=FPDTYYYYMM title=FPDATETIME  ALT="��/�󿩳��" tag="12X1" VIEWASTEXT> </OBJECT></TD>
								
							</TR>
							<TR>
								
								<TD CLASS="TD5" NOWRAP>���ޱ���</TD>
			            	    <TD CLASS="TD6"><SELECT NAME="txtProv_cd" CLASS ="cbonormal" tag="11" ALT="���ޱ���"><OPTION Value=""></OPTION></SELECT></TD>
								
								<!-- <TD CLASS="TD5" NOWRAP>���ޱ���</TD>
	                        	<TD CLASS="TD6" NOWRAP><INPUT NAME="txtProv_cd" MAXLENGTH="1" SIZE="10" ALT ="���ޱ���" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup(2)">
	                        	<INPUT NAME="txtProv_nm" MAXLENGTH="20" SIZE="20" ALT ="���ޱ��и�" tag="14XXXU"></TD>  -->

							   	<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 NOWRAP>
								<INPUT NAME="txtEmp_no" MAXLENGTH="13" SIZE="13" ALT ="���" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName(0)">
								<INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="����" tag="14XXXU"></TD>
								
							</TR>																													  
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread>
										<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
      <TD <%=HEIGHT_TYPE_01%>></TD>
    <TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
					<TD WIDTH=10>&nbsp;</TD>
	                <TD WIDTH=10><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: FncSave"   >�ݿ�</BUTTON>&nbsp;</TD>
					<TD WIDTH=10><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: FncDelete" >���</BUTTON>&nbsp;</TD>
					<TD WIDTH=10>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp;</TD>
	                <TD WIDTH=10><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: FncIFQuery">I/F ����Ÿ �ҷ�����</BUTTON>&nbsp;</TD>
	                <TD Width=*>&nbsp;</TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>   
	
	</TR>
		<TD WIDTH=100% HEIGHT=<%=Bizsize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=Bizsize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>

</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtProvCD" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtFileGubun" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
