<%@ LANGUAGE="VBSCRIPT" %>
<!--===================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : RECEIPT
'*  3. Program ID		    : f5121ma1
'*  4. Program Name         : �ε�����ó�� 
'*  5. Program Desc         : �ε�����ó�� ��� ���� ���� ��ȸ 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2003/04/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Soo Min, Oh
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************
'==========================================  1.1.1 Style Sheet  ==========================================
'========================================================================================================== -->

<!--========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance
																			'��: indicates that All variables must be declared in advance 

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Const BIZ_PGM_ID  = "f5121mb1.asp"										'��: �����Ͻ� ���� ASP�� 
'Const BIZ_PGM_ID2 = "f5121mb2.asp"										'��: �����Ͻ� ���� ASP�� 
'Const JUMP_PGM_ID_NOTE_CHG = "f5107ma1"									'���������� 

Dim C_SEQ
Dim C_STTL_TYPE
Dim C_STTL_TYPE_NM
Dim C_RCPT_TYPE
Dim C_RCPT_TYPE_BT
Dim C_RCPT_TYPE_NM
Dim C_REF_NOTE_NO
Dim C_REF_NOTE_BT
Dim C_ACCT_CD
Dim C_ACCT_BT
Dim C_ACCT_NM
Dim C_BANK_ACCT
Dim C_BANK_ACCT_BT
Dim C_BANK_CD
Dim C_BANK_BT
Dim C_BANK_NM
Dim C_STTL_AMT
Dim C_NOTE_ITEM_DESC

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
 '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE								'��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False										'��: Indicates that no value changed
    lgIntGrpCount = 0												'��: Initializes Group View Size

	lgStrPrevKey = ""
	lgLngCurRows = 0												'initializes Deleted Rows Count
    IsOpenPop = False												'��: ����� ���� �ʱ�ȭ 

    lgSortKey = 1
	lgPageNo  = ""

    lgBlnFlgChgValue = False
End Sub

Sub initSpreadPosVariables()
	C_SEQ = 1
	C_STTL_TYPE = 2
	C_STTL_TYPE_NM = 3
	C_RCPT_TYPE = 4
	C_RCPT_TYPE_BT = 5
	C_RCPT_TYPE_NM = 6
	C_REF_NOTE_NO = 7
	C_REF_NOTE_BT = 8	
	C_ACCT_CD = 9
	C_ACCT_BT = 10
	C_ACCT_NM = 11
	C_BANK_ACCT = 12
	C_BANK_ACCT_BT = 13
	C_BANK_CD = 14
	C_BANK_BT = 15
	C_BANK_NM = 16
	C_STTL_AMT = 17
	C_NOTE_ITEM_DESC = 18
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtStsDt.Text	= UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
	
	frm1.hOrgChangeId.value = Parent.gChangeOrgId
	frm1.txtNoteNoQry.focus
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()
    Dim sList

    With frm1
		.vspdData.MaxCols = C_NOTE_ITEM_DESC + 1
		.vspdData.Col = .vspdData.MaxCols	:	.vspdData.ColHidden = True				'��: ������Ʈ�� ��� Hidden Column
		.vspdData.MaxRows = 0
		ggoSpread.Source = .vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread
        Call GetSpreadColumnPos("A")
	
		ggoSpread.SSSetFloat	C_SEQ,			"����", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec		
		ggoSpread.SSSetCombo	C_STTL_TYPE,	"ó������", 12
		ggoSpread.SSSetCombo	C_STTL_TYPE_NM, "ó��������", 12
		ggoSpread.SSSetEdit		C_RCPT_TYPE,	"�Ա�����", 15, , , 20
		ggoSpread.SSSetButton	C_RCPT_TYPE_BT
		ggoSpread.SSSetEdit		C_RCPT_TYPE_NM, "�Ա�������", 15, , , 30
		ggoSpread.SSSetEdit		C_REF_NOTE_NO,	"����������ȣ", 15, , , 20
		ggoSpread.SSSetButton	C_REF_NOTE_BT				
		ggoSpread.SSSetEdit		C_ACCT_CD,		"�����ڵ�", 15, , , 20
		ggoSpread.SSSetButton	C_ACCT_BT
		ggoSpread.SSSetEdit		C_ACCT_NM,		"�����ڵ��", 15, , , 50
		ggoSpread.SSSetEdit		C_BANK_ACCT,	"���¹�ȣ", 20, , , 30
		ggoSpread.SSSetButton	C_BANK_ACCT_BT
		ggoSpread.SSSetEdit		C_BANK_CD,		"�����ڵ�", 15, , , 30
		ggoSpread.SSSetButton	C_BANK_BT
		ggoSpread.SSSetEdit		C_BANK_NM,		"�����", 15, , , 30
		ggoSpread.SSSetFloat	C_STTL_AMT,		"ó���ݾ�", 17, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit		C_NOTE_ITEM_DESC, "���", 35, , , 128
	
'2003/09/09 ���� �ʿ�	 
		Call ggoSpread.MakePairsColumn(C_STTL_TYPE,C_STTL_TYPE_NM,"1")
		Call ggoSpread.MakePairsColumn(C_RCPT_TYPE,C_RCPT_TYPE_NM,"1")
		Call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_NM)
		Call ggoSpread.MakePairsColumn(C_BANK_ACCT,C_BANK_ACCT_BT)
		Call ggoSpread.MakePairsColumn(C_BANK_CD,C_BANK_NM)
		
		Call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
		Call ggoSpread.SSSetColHidden(C_STTL_TYPE,C_STTL_TYPE,True)

		Call SetSpreadLock                                              '�ٲ�κ� 
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
	Dim RowCnt

	ggoSpread.Source = frm1.vspdData	

    With frm1
		.vspdData.ReDraw = False      
'		ggoSpread.SpreadLock		1 ,     -1  

		ggoSpread.SpreadLock		C_SEQ				, -1	, C_SEQ
		ggoSpread.SpreadUnLock		C_STTL_TYPE_NM		, -1    , C_STTL_TYPE_NM
		ggoSpread.SSSetRequired		C_STTL_TYPE_NM		, -1    

		ggoSpread.SpreadLock		C_RCPT_TYPE_NM,			-1		, C_ACCT_NM
		ggoSpread.SpreadLock		C_ACCT_NM,			-1		, C_ACCT_NM
		ggoSpread.SSSetRequired		C_STTL_TYPE_NM,		-1
		ggoSpread.SSSetRequired		C_ACCT_CD,			-1
		ggoSpread.SpreadUnLock		C_STTL_AMT			, -1    , C_STTL_AMT		
		ggoSpread.SSSetRequired		C_STTL_AMT,			-1
		ggoSpread.SpreadUnLock		C_NOTE_ITEM_DESC	, -1    , C_NOTE_ITEM_DESC

		For RowCnt = 1 To .vspdData.MaxRows			
			.vspdData.Col = C_STTL_TYPE
			.vspdData.Row = RowCnt	

			If UCase(Trim(.vspdData.text)) = "RI" Then						'��ȯ 
				ggoSpread.SpreadUnLock		C_RCPT_TYPE,		RowCnt,	C_RCPT_TYPE	,RowCnt			
				ggoSpread.SSSetRequired		C_RCPT_TYPE,		RowCnt,	RowCnt			
				ggoSpread.SpreadUnLock		C_RCPT_TYPE_BT,		RowCnt,	C_RCPT_TYPE_BT	

				ggoSpread.SpreadLock		C_REF_NOTE_NO,		RowCnt,	C_REF_NOTE_NO,RowCnt			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		RowCnt,	RowCnt			

				.vspdData.Col = C_RCPT_TYPE
				.vspdData.Row = RowCnt		

				If Trim(.vspdData.text) = "DP" Then						
					ggoSpread.SpreadUnLock		C_BANK_ACCT,		RowCnt,	C_BANK_ACCT	,RowCnt			
					ggoSpread.SSSetRequired		C_BANK_ACCT,		RowCnt,	RowCnt			
					ggoSpread.SpreadUnLock		C_BANK_ACCT_BT,		RowCnt,	C_BANK_ACCT_BT

					ggoSpread.SpreadUnLock		C_BANK_CD,			RowCnt,	C_BANK_CD	,RowCnt			
					ggoSpread.SSSetRequired		C_BANK_CD,			RowCnt,	RowCnt			
					ggoSpread.SpreadUnLock		C_BANK_BT,			RowCnt,	C_BANK_BT					
				Else
					ggoSpread.SpreadLock		C_BANK_ACCT,		-1		, C_BANK_ACCT_BT
					ggoSpread.SpreadLock		C_BANK_CD,			-1		, C_BANK_NM
				End If
			ElseIf  UCase(Trim(.vspdData.text)) = "NR" Then	
				ggoSpread.SpreadLock		C_RCPT_TYPE,		RowCnt, C_RCPT_TYPE			,RowCnt			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		RowCnt, RowCnt
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		RowCnt,	C_RCPT_TYPE_BT	
				
				ggoSpread.SpreadUnLock		C_REF_NOTE_NO,		RowCnt, C_REF_NOTE_NO			,RowCnt			
				ggoSpread.SSSetRequired		C_REF_NOTE_NO,		RowCnt, RowCnt
				ggoSpread.SpreadUnLock		C_REF_NOTE_BT,		RowCnt,	C_RCPT_TYPE_BT		
				
				ggoSpread.SpreadLock		C_BANK_ACCT,		-1		, C_BANK_ACCT_BT
				ggoSpread.SpreadLock		C_BANK_CD,			-1		, C_BANK_NM
			ElseIf UCase(Trim(.vspdData.text)) <> "AL" Or UCase(Trim(.vspdData.text)) <> "EP" Then
				ggoSpread.SpreadLock		C_RCPT_TYPE,		RowCnt, C_RCPT_TYPE			,RowCnt			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		RowCnt, RowCnt
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		RowCnt,	C_RCPT_TYPE_BT	
				
				ggoSpread.SpreadLock		C_REF_NOTE_NO,		RowCnt,	C_REF_NOTE_NO,RowCnt			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		RowCnt,	RowCnt			
				
				ggoSpread.SpreadLock		C_BANK_ACCT,		-1		, C_BANK_ACCT_BT
				ggoSpread.SpreadLock		C_BANK_CD,			-1		, C_BANK_NM
			End If 
		Next 

		.vspdData.ReDraw = True
   End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal lRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired		C_STTL_TYPE_NM,		lRow,	lRow
		ggoSpread.SSSetProtected	C_RCPT_TYPE,		lRow,	lRow
		ggoSpread.SSSetProtected	C_RCPT_TYPE_BT,		lRow,	lRow
		ggoSpread.SSSetProtected	C_RCPT_TYPE_NM,		lRow,	lRow
		ggoSpread.SSSetProtected	C_REF_NOTE_NO,		lRow,	lRow
		ggoSpread.SSSetProtected	C_REF_NOTE_BT,		lRow,	lRow
		ggoSpread.SSSetRequired		C_ACCT_CD,			lRow,	lRow
		ggoSpread.SSSetProtected	C_ACCT_NM,			lRow,	lRow
		ggoSpread.SSSetProtected	C_BANK_ACCT,		lRow,	lRow
		ggoSpread.SSSetProtected	C_BANK_ACCT_BT,		lRow,	lRow		
		ggoSpread.SpreadLock		C_BANK_CD,			lRow,	lRow
		ggoSpread.SSSetRequired		C_STTL_AMT,			lRow,	lRow
		
		.vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Function InitCombo()
	ggoSpread.Source = frm1.vspdData
	                   'Select                 From        Where                Return value list  
    Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("F1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_STTL_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_STTL_TYPE_NM       

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))    
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))    
End Function

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenNoteInfo()
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("f5121ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f5121ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		frm1.txtNoteNoQry.focus
		Exit Function
	Else
		frm1.txtNoteNoQry.value  = arrRet(0)
		frm1.txtNoteNoQry.focus
	End If	
End Function

'------------------------------------------  OpenPopUpNoteNo()  ---------------------------------------------
'	Name : OpenPopUpNoteNo()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function  OpenPopUpNoteNo()
	Dim strNoteFg
	Dim IntRetCd
	
	strNoteFg = frm1.cboNoteFg.Value
	
	If strNoteFg = "" Then
	    IntRetCD = DisplayMsgBox("141327","x","x","x")	'���������� ���� �Է��Ͻʽÿ�.
		Exit Function		      	
	ElseIf strNoteFg = "D3" Then 
		Call OpenPopUp(frm1.txtNoteNo.Value, 1) '���޾��� 
	Else  
	    IntRetCD = DisplayMsgBox("141220","x","x","x")	'������ȣ�� ���� �Է����ֽʽÿ�.
		Exit Function		
    End If	
End Function  

'==================================================================================
'	Name : OpenPopUp()
'	Description : �����˾� ���� 
'==================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 2		'�Ա����� 
 			arrParam(0) = "�Ա������˾�"
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " "
			arrParam(4) = arrParam(4) & "AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD "
			arrParam(4) = arrParam(4) & "AND B.SEQ_NO = 1 AND B.REFERENCE = " & FilterVar("RP", "''", "S") & "  "
			arrParam(5) = strCode

			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"

			arrHeader(0) = "�Ա�����"
			arrHeader(1) = "�Ա�������"
		Case 3		' ���� 
			arrParam(0) = "���� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"						' TABLE ��Ī 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "									' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(5) = strCode													' �����ʵ��� �� ��Ī 
			
			arrField(0) = "A.BANK_CD"						' Field��(0)
			arrField(1) = "A.BANK_NM"						' Field��(1)
			arrField(2) = "B.BANK_ACCT_NO"					' Field��(2)
			
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)
			arrHeader(2) = "���¹�ȣ"					' Header��(2)				
		Case 4		' ���¹�ȣ 
'			If frm1.txtBankAcct1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "���¹�ȣ �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"				' TABLE ��Ī 
			arrParam(2) = strCode							' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "												' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "	
			arrParam(5) = "���¹�ȣ"					' �����ʵ��� �� ��Ī 
			arrField(0) = "B.BANK_ACCT_NO"					' Field��(0)
			arrField(1) = "A.BANK_CD"						' Field��(0)
			arrField(2) = "A.BANK_NM"						' Field��(0)
			arrHeader(0) = "���¹�ȣ"					' Header��(0)
			arrHeader(1) = "�����ڵ�"					' Header��(0)
			arrHeader(2) = "�����"						' Header��(0)	
		Case 5		' �Ա����������ڵ� 
			arrParam(0) = "�Աݰ����˾�"								' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & " " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  "
			arrParam(4) = arrParam(4) & " AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ "
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  and D.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "			
			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = C_STTL_TYPE
			
			If Trim(frm1.vspdData.Text) <> "" Then
				arrParam(4) = arrParam(4) & " AND D.JNL_CD = " & FilterVar(frm1.vspdData.Text, "''", "S") 					 				 		
			End If 
			
			frm1.vspdData.Col = C_RCPT_TYPE
			
			If Trim(frm1.vspdData.Text) <> "" Then
			arrParam(4) = arrParam(4) & " AND D.EVENT_CD = " & FilterVar(frm1.vspdData.Text, "''", "S")
			End If

			arrParam(5) = strCode											' �����ʵ��� �� ��Ī 

			arrField(0) = "A.ACCT_CD"										' Field��(0)
			arrField(1) = "A.ACCT_NM"										' Field��(1)
			arrField(2) = "B.GP_CD"											' Field��(2)
			arrField(3) = "B.GP_NM"					 						' Field��(3)

			arrHeader(0) = "�Ա����������ڵ�"							' Header��(0)
			arrHeader(1) = "�Ա�����������"								' Header��(1)
			arrHeader(2) = "�׷��ڵ�"									' Hea der��(2)
			arrHeader(3) = "�׷��"										' Header��(3)	
		Case 6																'������ȣ POPUP
 			arrParam(0) = "������ȣ�˾�"
			arrParam(1) = "F_NOTE A, B_BANK	B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.NOTE_STS = " & FilterVar("BG", "''", "S") & " "
			arrParam(4) = arrParam(4) & " AND A.BANK_CD = B.BANK_CD "
			arrParam(5) = strCode

			arrField(0) = "A.NOTE_NO"			
			arrField(1) = "A.NOTE_AMT"
			arrField(2) = "B.BANK_NM"

			arrHeader(0) = "������ȣ"
			arrHeader(1) = "�����ݾ�"
			arrHeader(2) = "��������"
		Case 7		' ���ڼ��Ͱ����ڵ� 
			arrParam(0) = "���ڼ��Ͱ����˾�"							' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN004", "''", "S") & " " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  "
			arrParam(4) = arrParam(4) & " AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ "
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  and D.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  "						
			arrParam(4) = arrParam(4) & " AND D.JNL_CD = " & FilterVar("IR", "''", "S") & "  " 					 				 					
			arrParam(5) = strCode											' �����ʵ��� �� ��Ī 

			arrField(0) = "A.ACCT_CD"										' Field��(0)
			arrField(1) = "A.ACCT_NM"										' Field��(1)
			arrField(2) = "B.GP_CD"											' Field��(2)
			arrField(3) = "B.GP_NM"					 						' Field��(3)

			arrHeader(0) = "���ڼ��Ͱ����ڵ�"							' Header��(0)
			arrHeader(1) = "���ڼ��Ͱ�����"								' Header��(1)
			arrHeader(2) = "�׷��ڵ�"									' Hea der��(2)
			arrHeader(3) = "�׷��"										' Header��(3)
	End Select
  
	IsOpenPop = True
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  EscPopUp()  --------------------------------------------------
'	Name : EscPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 2		' �Ա�����		
				Call SetActiveCell(.vspdData,C_RCPT_TYPE,.vspdData.ActiveRow ,"M","X","X")
				Call SetActiveCell(.vspdData,C_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 3		' ���� 
				.txtBankCD.focus
			Case 4		' �ŷ�ó 
				.txtBpCd.focus
			Case 5		' �Ա����������ڵ� 
				Call SetActiveCell(.vspdData,C_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 6		' �������� 
				Call SetActiveCell(.vspdData,C_REF_NOTE_NO,.vspdData.ActiveRow ,"M","X","X")
			Case 7		' ���ڼ��� 
				.txtIntAcctCd.focus
		End Select
	End With
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 2		' �Ա�����										
				.vspdData.Col = C_RCPT_TYPE_NM
				.vspdData.Text = arrRet(1)					
				.vspdData.Col = C_RCPT_TYPE
				.vspdData.Text = arrRet(0)

				If UCase(Trim(.vspdData.Text)) = "DP" Then												
					ggoSpread.SpreadUnLock		C_BANK_ACCT,		.vspdData.ActiveRow,	C_BANK_ACCT	,.vspdData.ActiveRow			
					ggoSpread.SSSetRequired		C_BANK_ACCT,		.vspdData.ActiveRow,	.vspdData.ActiveRow			
					ggoSpread.SpreadUnLock		C_BANK_ACCT_BT,		.vspdData.ActiveRow,	C_BANK_ACCT_BT										
				
					ggoSpread.SpreadUnLock		C_BANK_CD,			.vspdData.ActiveRow,	C_BANK_CD	,.vspdData.ActiveRow			
					ggoSpread.SSSetRequired		C_BANK_CD,			.vspdData.ActiveRow,	.vspdData.ActiveRow			
					ggoSpread.SpreadUnLock		C_BANK_BT,			.vspdData.ActiveRow,	C_BANK_BT
					ggoSpread.SpreadLock		C_BANK_NM,			.vspdData.ActiveRow,	C_BANK_NM					
				Else
					ggoSpread.SpreadLock		C_BANK_ACCT,		.vspdData.ActiveRow,	C_BANK_ACCT			,.vspdData.ActiveRow			
					ggoSpread.SSSetProtected	C_BANK_ACCT,		.vspdData.ActiveRow,	.vspdData.ActiveRow
					ggoSpread.SpreadLock		C_BANK_CD,			.vspdData.ActiveRow,	C_BANK_CD			,.vspdData.ActiveRow			
					ggoSpread.SSSetProtected	C_BANK_CD,			.vspdData.ActiveRow,	.vspdData.ActiveRow				
					ggoSpread.SpreadLock		C_BANK_NM,			.vspdData.ActiveRow,	C_BANK_NM
				End If
				
				.vspdData.Col = C_ACCT_CD
				.vspdData.Text = ""
				.vspdData.Col = C_ACCT_NM
				.vspdData.Text = ""
				.vspdData.Col = C_BANK_ACCT
				.vspdData.Text = ""
				.vspdData.Col = C_BANK_CD
				.vspdData.Text = ""
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = ""				
				
				ggoSpread.SpreadUnLock		C_ACCT_CD,			.vspdData.ActiveRow,	C_ACCT_CD	,.vspdData.ActiveRow			
				ggoSpread.SSSetRequired		C_ACCT_CD,			.vspdData.ActiveRow,	.vspdData.ActiveRow			
				ggoSpread.SpreadUnLock		C_ACCT_BT,			.vspdData.ActiveRow,	C_ACCT_BT
				ggoSpread.SpreadLock		C_ACCT_NM,			.vspdData.ActiveRow,	C_ACCT_NM									

				Call SetActiveCell(.vspdData,C_RCPT_TYPE,.vspdData.ActiveRow ,"M","X","X")
			Case 3		' ���� 
				.vspdData.Col = C_BANK_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_BANK_ACCT
				.vspdData.Text = arrRet(2)
			Case 4		' ���¹�ȣ 
				.vspdData.Col = C_BANK_ACCT
				.vspdData.Text = arrRet(0)				
				.vspdData.Col = C_BANK_CD
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_BANK_NM
				.vspdData.Text = arrRet(2)
			Case 5		' �Ա����������ڵ� 
				.vspdData.Col = C_ACCT_CD
				.vspdData.Text = arrRet(0)				
				.vspdData.Col = C_ACCT_NM
				.vspdData.Text = arrRet(1)				
				Call SetActiveCell(.vspdData,C_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 6		' �������� 
				.vspdData.Col = C_REF_NOTE_NO
				.vspdData.Text = arrRet(0)				
				.vspdData.Col = C_STTL_AMT
				.vspdData.Text = arrRet(1)	
				Call SetActiveCell(.vspdData,C_REF_NOTE_NO,.vspdData.ActiveRow ,"M","X","X")
			Case 7		' ���ڼ��� 
				.txtIntAcctCd.value = arrRet(0)
				.txtIntAcctNM.value = arrRet(1)
				.txtIntAcctCd.focus
		End Select

		lgBlnFlgChgValue = True
	End With
End Function

'------------------------------------------  OpenPopupDept()  --------------------------------------------
'	Name : OpenPopupDept()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'�μ��ڵ� 
	arrParam(1) = frm1.txtStsDt.Text			'��¥(Default:������)
	arrParam(2) = "1"							'�μ�����(lgUsrIntCd)
	arrParam(3) = "F"							'�μ�����(lgUsrIntCd)

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeptCD.focus
		Exit Function
	End If

	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	frm1.txtStsDt.text = arrRet(3)
	Call txtDeptCD_Change()
	frm1.txtDeptCD.focus

	lgBlnFlgChgValue = True
End Function

'------------------------------------------  OpenPopupTempGL()  --------------------------------------------
'	Name : OpenPopupTempGL()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
    Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	With frm1		
		arrParam(0) = Trim(.hTempGlNo.value)	'��ǥ��ȣ 
		arrParam(1) = ""						'Reference��ȣ		
	End With
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'------------------------------------------  OpenPopupGL()  --------------------------------------------
'	Name : OpenPopupGL()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
		
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	With frm1		
		arrParam(0) = Trim(.hGlNo.value)	'��ǥ��ȣ 
		arrParam(1) = ""			'Reference��ȣ		
	End With
	
	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function CookiePage(ByVal Kubun)
	Dim strTemp

	Select Case Kubun		
		Case "FORM_LOAD"		
			strTemp = ReadCookie("NOTE_NO")
			
			Call WriteCookie("NOTE_NO", "")
			
			If strTemp = "" Then Exit Function

			frm1.txtNoteNoQry.value = strTemp
	
			If Err.number <> 0 Then
				Err.Clear
				Call WriteCookie("NOTE_NO", "")
				Exit Function 
			End If
					
			Call MainQuery()
		Case Else
			Exit Function
	End Select
End Function	

'========================================================================================================
'	Desc : ȭ���̵� 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD

	'-----------------------
	'Check previous data area
	'----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ����Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call CookiePage(strPgmId)
    Call PgmJump(strPgmId)
End Function

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029															'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.ClearField(Document, "1")										'��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")										'��: Lock  Suitable  Field

	Call InitSpreadSheet                                                        'Setup the Spread sheet			
	Call InitCombo    	
    Call SetDefaultVal    
    Call InitVariables															'��: Initializes local global variables

	Call SetToolbar("1110110000001111")
'    Call CookiePage("FORM_LOAD")
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SEQ           =   iCurColumnPos(1)
			C_STTL_TYPE		=	iCurColumnPos(2)
			C_STTL_TYPE_NM	= 	iCurColumnPos(3)
			C_RCPT_TYPE		=	iCurColumnPos(4)
			C_RCPT_TYPE_BT	=	iCurColumnPos(5)
			C_RCPT_TYPE_NM	=	iCurColumnPos(6)
			C_REF_NOTE_NO	=	iCurColumnPos(7)
			C_REF_NOTE_BT	=	iCurColumnPos(8)			
			C_ACCT_CD		=	iCurColumnPos(9)
			C_ACCT_BT		=	iCurColumnPos(10)
			C_ACCT_NM		=	iCurColumnPos(11)
			C_BANK_ACCT		=	iCurColumnPos(12)
			C_BANK_ACCT_BT	=	iCurColumnPos(13)
			C_BANK_CD		=	iCurColumnPos(14)
			C_BANK_BT		=	iCurColumnPos(15)
			C_BANK_NM		=	iCurColumnPos(16)
			C_STTL_AMT		=	iCurColumnPos(17)
			C_NOTE_ITEM_DESC=	iCurColumnPos(18)
    End Select    
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow
		frm1.vspdData.Col = C_STTL_TYPE
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_STTL_TYPE_NM
		frm1.vspdData.value = intindex
	Next
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStsDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStsDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStsDt.Focus 
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStsDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStsDt_Change()
    Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2

	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtStsDt.Text <> "") Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStsDt.Text, gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
								
			For ii = 0 to Ubound(arrVal1,1) - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				If Trim(arrVal2(2)) <> Trim(frm1.hOrgChangeId.value) Then
					frm1.txtDeptCd.value = ""
					frm1.txtDeptNm.value = ""
					frm1.hOrgChangeId.value = Trim(arrVal2(2))
				End If
			Next
		End If
	End If
	
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDeptCD_Change()
'   Event Desc : Vlidation Check of Department Code
'=======================================================================================================
Sub txtDeptCD_Change()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii

	If Trim(frm1.txtDeptCd.value) = "" And Trim(frm1.txtStsDt.Text = "") Then Exit Sub

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStsDt.Text, gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			

		For ii = 0 to Ubound(arrVal1,1) - 1
			arrVal2 = Split(arrVal1(ii), chr(11))
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next	
	End If

     lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtEndDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtStsDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCashRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtNoteAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtSttlAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtIntRevAmt_Change()
    lgBlnFlgChgValue = True
    If unicdbl(frm1.txtIntRevAmt.Text) > 0 Then   		
		Call ggoOper.SetReqAttr(frm1.txtIntAcctCd, "N")		
	Else 
		frm1.txtIntAcctCd.value = ""
		frm1.txtIntAcctNm.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtIntAcctCd, "Q")			
	End If		
End Sub

 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************
Sub cboNoteFg_OnChange1()							'dbqueryok ���� event (field not clear)
	with frm1
		Select Case frm1.cboNoteFg.value
			Case "D1"	'��������				
				Call ggoOper.SetReqAttr(.txtCashRate, "N")	'N:Required, Q:Protected, D:Default				
			Case "D3"	'���޾���			
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default				
			Case Else				
				Call ggoOper.SetReqAttr(.txtCashRate, "Q")	'N:Required, Q:Protected, D:Default
		End Select
	End with
End Sub

Sub cboPlace_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub cboRcptFg_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtNoteNo_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtDeptCD_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtPublisher_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBpCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtNoteDesc_OnChange()
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'   Event Name : vspd	
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    lgBlnFlgChgValue = True
End Sub 
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			.Col = Col
			.Row = Row
			Select Case Col
			Case C_RCPT_TYPE_BT
				Call OpenPopup(.Text, 2)
			Case C_REF_NOTE_BT			
				Call OpenPopup(.Text, 6)
			Case C_ACCT_BT
				Call OpenPopup(.Text, 5)
			Case C_BANK_ACCT_BT
				Call OpenPopup(.Text, 4)
			Case C_BANK_BT
				Call OpenPopup(.Text, 3)
			Case Else
			End Select
		End If
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data clicked
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
    Set gActiveSpdSheet = frm1.vspdData
    
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
	
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
    End If     
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgPageNo <> "" Then                         
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData
	
	With frm1.vspdData
		.ReDraw = False
		.Row = Row
    
		Select Case Col
			Case  C_STTL_TYPE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_STTL_TYPE
				.Value = intIndex
				varData = .text							
			Case C_STTL_TYPE
				.Col = Col
				intIndex = .Value
				.Col = C_STTL_TYPE_NM
				.Value = intIndex
				varData = .text				
		End Select
		
		ggoSpread.Source = frm1.vspdData												
		
		Select Case UCase(Trim(.text))		
			Case "RI"						'��ȯ(ReImbursement)												
				ggoSpread.SpreadUnLock		C_RCPT_TYPE,		Row, C_RCPT_TYPE	,Row			
				ggoSpread.SSSetRequired		C_RCPT_TYPE,		Row, Row			
				ggoSpread.SpreadUnLock		C_RCPT_TYPE_BT,		Row, C_RCPT_TYPE_BT	,Row
						
				ggoSpread.SpreadLock		C_REF_NOTE_NO,		Row, C_REF_NOTE_NO	,Row			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		Row, Row				
			Case "NR"						'�űԹ�������(Note Receivable)			
				ggoSpread.SpreadLock		C_RCPT_TYPE,		Row, C_RCPT_TYPE	,Row			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		Row, Row
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		Row, C_RCPT_TYPE_BT	,Row
						
				ggoSpread.SpreadUnLock		C_REF_NOTE_NO,		Row, C_REF_NOTE_NO	,Row			
				ggoSpread.SSSetRequired		C_REF_NOTE_NO,		Row, Row
				ggoSpread.SpreadUnLock		C_REF_NOTE_BT,		Row, C_RCPT_TYPE_BT	,Row	
						
				ggoSpread.SpreadLock		C_BANK_ACCT,		Row, C_BANK_ACCT	,Row			
				ggoSpread.SSSetProtected	C_BANK_ACCT,		Row, Row	
				ggoSpread.SpreadLock		C_BANK_CD,			Row, C_BANK_CD		,Row			
				ggoSpread.SSSetProtected	C_BANK_CD,			Row, Row		
			Case Else
				ggoSpread.SpreadLock		C_RCPT_TYPE,		Row, C_RCPT_TYPE	,Row			
				ggoSpread.SSSetProtected	C_RCPT_TYPE,		Row, Row
				ggoSpread.SpreadLock		C_RCPT_TYPE_BT,		Row, C_RCPT_TYPE_BT	,Row			
			
				ggoSpread.SpreadLock		C_REF_NOTE_NO,		Row, C_REF_NOTE_NO	,Row			
				ggoSpread.SSSetProtected	C_REF_NOTE_NO,		Row, Row							
			
				ggoSpread.SpreadLock		C_BANK_ACCT,		Row, C_BANK_ACCT	,Row			
				ggoSpread.SSSetProtected	C_BANK_ACCT,		Row, Row	
				ggoSpread.SpreadLock		C_BANK_CD,			Row, C_BANK_CD		,Row			
				ggoSpread.SSSetProtected	C_BANK_CD,			Row, Row					
		End Select
		
		ggoSpread.SpreadLock		C_BANK_ACCT,		Row, C_BANK_ACCT			,Row			
		ggoSpread.SSSetProtected	C_BANK_ACCT,		Row, Row	
		ggoSpread.SpreadLock		C_BANK_CD,			Row, C_BANK_CD			,Row			
		ggoSpread.SSSetProtected	C_BANK_CD,			Row, Row
		ggoSpread.SpreadLock		C_BANK_NM,			Row, C_BANK_NM			,Row			
		ggoSpread.SSSetProtected	C_BANK_NM,			Row, Row				
		
'		ggoSpread.SpreadLock		C_ACCT_CD,			Row, C_ACCT_CD			,Row			
'		ggoSpread.SSSetProtected	C_ACCT_CD,			Row, Row	
		
		.Col = C_RCPT_TYPE			
		.Text = ""
		.Col = C_RCPT_TYPE_NM			
		.Text = ""
		.Col = C_REF_NOTE_NO			
		.Text = ""
		.Col = C_ACCT_CD			
		.Text = ""
		.Col = C_ACCT_NM			
		.Text = ""
		.Col = C_BANK_ACCT			
		.Text = ""
		.Col = C_BANK_CD			
		.Text = ""
		.Col = C_BANK_NM			
		.Text = ""
			
		.ReDraw = True	
	End With
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    	
    '-----------------------
    'Check previous data area
    '----------------------- 
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")		'�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables														'��: Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then										'��: This function check indispensable field
		Exit Function
    End If
    
    Call ggoOper.LockField(Document, "N")									'��: This function lock the suitable field

	'-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery															'��: Query db data
       
    FncQuery = True															'��: Processing is OK
	Set gActiveElement = document.activeElement          
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False														'��: Processing is NG
    
	'-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
	    IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")	'�� �ٲ�κ� 
	     If IntRetCD = vbNo Then
	         Exit Function
	     End If
	End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")								'��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")								'��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")								'��: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables													'��: Initializes local global variables

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call SetToolbar("1110100000000011")									'��: ��ư ���� ���� 

    FncNew = True														'��: Processing is OK

	frm1.txtNoteNoQry.focus 
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False														'��: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        intRetCD = DisplayMsgBox("900002","x","x","x")						'�� �ٲ�κ� 
        'Call MsgBox("��ȸ���Ŀ� ������ �� �ֽ��ϴ�.", vbInformation)
        Exit Function
    End If    
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")			'�� �ٲ�κ� 
    If IntRetCD = vbNo Then
        Exit Function
    End If    
    
	'-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001","x","x","x")  '�� �ٲ�κ� 
		Exit Function
	End If
	
	'-----------------------
	  'Check content area
	'-----------------------
    If Not chkField(Document, "2") Then										'��: Check contents area
		Exit Function
    End If
   
	'-----------------------
	'sum(single amt) = sum(multi amt) check
	'----------------------- 
	Call DoSum()
	 
	If chkSttlAmt() = False Then
		DisplayMsgBox "113119","X","X","X"
		Exit Function 
	End If 		    
  
	'-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                '��: Save db data
       
    FncSave = True                                                          '��: Processing is OK
    Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	With frm1
   		.vspdData.ReDraw = False

		If .vspdData.MaxRows < 1 Then Exit Function

		ggoSpread.Source = .vspdData
		ggoSpread.CopyRow

		MaxSpreadVal .vspdData, C_SEQ , .vspdData.ActiveRow

		Call SetSpreadColor(.vspdData.ActiveRow)

		.vspdData.ReDraw = True
	End With

	Set gActiveElement = document.activeElement    
End Function

'==========================================================================================
'   Event Desc : Grid�� Max Count �� ã�´�.
'==========================================================================================
Function MaxSpreadVal(ByVal objSpread, ByVal intCol, byval Row)

	Dim iRows
	Dim MaxValue
	Dim tmpVal

	MAxValue = 0

	For iRows = 1 to  objSpread.MaxRows
		objSpread.row = iRows
	    objSpread.col = intCol

		If objSpread.Text = "" Then
		   tmpVal = 0
		Else
  		   tmpVal = cdbl(objSpread.value)
		End If

		If tmpval > MaxValue   Then
		   MaxValue = cdbl(tmpVal)
		End If
	Next

	MaxValue = MaxValue + 1

	objSpread.row	= row
	objSpread.col	= intCol
	objSpread.Text	= MaxValue

end Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function

	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo

	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim imRow
    Dim ii,iCurRowPos
    
    On Error Resume Next															'��: If process fails
    Err.Clear																		'��: Clear error status
    
    FncInsertRow = False															'��: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
            Exit Function
        End If
    End If

	With frm1
		iCurRowPos = .vspdData.ActiveRow
        .vspdData.Redraw = False
        ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		
		For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
			Call MaxSpreadVal(.vspdData, C_SEQ, ii)
		Next
		
		.Col = 2																	' �÷��� ���� ��ġ�� �̵�      
		.Row = 	ii - 1
		.Action = 0
		
        Call SetSpreadColor(iCurRowPos + 1)
        .ReDraw = True
	End With        

    If Err.number = 0 Then
		FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    If Frm1.vspdData.MaxRows < 1 Then
       Exit function
	End if	

	lgBlnFlgChgValue = True

    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With

    Set gActiveElement = document.ActiveElement  
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'��: ȭ�� ���� 
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                     '��:ȭ�� ����, Tab ���� 
	Set gActiveElement = document.activeElement  
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
	Call InitCombo()
    Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadLock()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003				'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtNoteNo=" & Trim(frm1.txtNoteNoQry.value)		'��: ���� ���� ����Ÿ 
    strVal = strVal & "&hGlNo=" & Trim(frm1.hGlNo.value)				'��: ���� ���� ����Ÿ 
    strVal = strVal & "&hTempGlNo=" & Trim(frm1.hTempGlNo.value)		'��: ���� ���� ����Ÿ 
       
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncQuery()
End Function

'========================================================================================
' Function Name : DoSum() 
' Function Desc : ���������� ���� ���Ѵ�.
'========================================================================================
Sub DoSum()
	Dim tmpSttlSum		
	DIm Row	
	
	tmpSttlSum = 0 	
	
	With frm1
		For row = 1 To .vspdData.maxRows
			.vspdData.Col = 0
			.vspdData.Row = row
				
			If .vspdData.Text <> ggoSpread.DeleteFlag Then
				.vspdData.Col = C_STTL_AMT
				tmpSttlSum = CDbl(tmpSttlSum) + unicdbl(.vspdData.text) 
				
				'UNIConvNumPCToCompanyByCurrency(tmpSttlSum,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				.htxtSumSttlAmt.text = tmpSttlSum				
			End If	
		Next	
	End With 
End Sub

Function chkSttlAmt()
	chkSttlAmt = True  
	
	With frm1		
		If uniCdbl(.htxtSumSttlAmt.text) <> uniCdbl(.txtNoteAmt.text) +  uniCdbl(.txtIntRevAmt.text) Then 
			chkSttlAmt = False 
			Exit Function 
		End If 		
	End With
End Function 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    
    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG
    
	Call LayerShowHide(1)

	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtNoteNoQry=" & Trim(.hNoteNo.value)		'��: ��ȸ ���� ����Ÿ 
		Else		
			strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001			'��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtNoteNoQry=" & Trim(.txtNoteNoQry.value)	'��: ��ȸ ���� ����Ÿ 
		End If
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo	=" & lgPageNo         
			strVal = strVal & "&txtMaxRows	=" & .vspdData.MaxRows
	End With

    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
 	Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call InitData
	Call SetToolbar("1111111100111111")
	
	Call SetSpreadLock()		
	
	If frm1.vspdData.MaxRows > 0 Then  
		lgIntFlgMode = Parent.OPMD_UMODE									'��: Indicates that current mode is Update mode
	Else 
		lgIntFlgMode = Parent.OPMD_CMODE									'��: Indicates that current mode is Update mode
	End If
	
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 
	Dim strVal
	Dim lRow
	Dim lGrpCnt	
	Dim	intRetCd			

    Err.Clear																'��: Protect system from crashing
	DbSave = False															'��: Processing is NG	
	
	With frm1
		.txtMode.value = Parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode	
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
		    .vspdData.Col = 0		    			
			If  .vspdData.Text <> ggoSpread.DeleteFlag Then 			
				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep	'��: C=Create, ����		0,1 
									
				.vspdData.Col = C_SEQ
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ����	2
				.vspdData.Col = C_STTL_TYPE
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ó������	3
				.vspdData.Col = C_RCPT_TYPE
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' �Ա�����	4
				.vspdData.Col = C_ACCT_CD 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' �Աݰ����ڵ�	5
				.vspdData.Col = C_REF_NOTE_NO 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ����������ȣ	6
				.vspdData.Col = C_BANK_ACCT 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' ���¹�ȣ		7
				.vspdData.Col = C_BANK_CD
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep		' �����ڵ�		8
				.vspdData.Col = C_STTL_AMT
				strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & Parent.gColSep	' ó���ݾ�		9
				.vspdData.Col = C_NOTE_ITEM_DESC 
				strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep		' ���			10						
				
				lGrpCnt = lGrpCnt + 1
			End If 				
		Next			
		
		frm1.txtSpread.Value = strVal						
		
		If frm1.txtSpread.Value = "" Then
		'�� spread��ü delete��[�ε����������� ����Ͻðڽ��ϱ�?]
			intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")   
			If intRetCd = VBNO Then
				Exit Function
			End If
		
			If  DbDelete = False Then
				Exit Function
			End If
			Exit Function
		End If 	
		
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal				
		
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)											
	End With		
		
    DbSave = True																'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk(Byval ptxtNoteNo)												'��: ���� ������ ���� ���� 
    Select Case lgIntFlgMode
		Case Parent.OPMD_CMODE
			frm1.txtNoteNoQry.value = ptxtNoteNo
    End Select

    Call InitVariables
    Call MainQuery()
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ε�����ó��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopuptempGL()">������ǥ</A>|
											<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
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
									<TD CLASS=TD5 NOWRAP>������ȣ</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtNoteNoQry" NAME="txtNoteNoQry" SIZE=30 MAXLENGTH=30 tag="12XXXU"ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteQry" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenNoteInfo"></TD>
								<TR>		
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
								<TD CLASS=TD5 NOWRAP>ó������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDateTime1_txtStsDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptCD" NAME="txtDeptCD" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON"  ONCLICK="vbscript:Call OpenPopUpDept(frm1.txtDeptCD.Value, 1)">&nbsp;
													<INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 MAXLENGTH=40 STYLE="TEXT-ALIGN: left" tag="24X" ALT="�μ�"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="��������" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="��������" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDateTime1_txtIssueDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDateTime2_txtDueDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=10 MAXLENGTH=10   tag="24XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBpCd.Value, 4)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="24X" ALT="�ŷ�ó"> </TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10   tag="24XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 5)">&nbsp;
													 <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="����"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDoubleSingle1_txtNoteAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDoubleSingle2_txtSttlAmt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ڼ���</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f5121ma1_fpDoubleSingle1_txtIntRevAmt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>���ڰ���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIntAcctCd" ALT="���ڼ��Ͱ���" SIZE="10" MAXLENGTH="20"  tag="22X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIntAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtIntAcctCd.value, 7)">
													 <INPUT NAME="txtIntAcctNm" ALT="���ڼ��Ͱ�����" SIZE="20" tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=4><INPUT CLASS="clstxt" TYPE=TEXT ID="txtNoteDesc" NAME="txtNoteDesc" SIZE=70 MAXLENGTH=128  tag="2XX" ALT="���"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
									<script language =javascript src='./js/f5121ma1_OBJECT1_vspdData.js'></script>
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
	</TR>

	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hNoteNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hGlNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTempGlNo" tag="2" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="horgchangeid" tag="2" TABINDEX="-1">
<!-- ���������� ó���ݾ� sum -->
<script language =javascript src='./js/f5121ma1_hOBJECT1_htxtSumSttlAmt.js'></script>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

