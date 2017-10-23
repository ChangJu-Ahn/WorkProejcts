<%@ LANGUAGE="VBSCRIPT" %>

<!--===================================================================================================
'*  1. Module Name          : ACCOUNTING
'*  2. Function Name        : TREASURY - NOTE
'*  3. Program ID		    : f5107ma1
'*  4. Program Name         : ���������� 
'*  5. Program Desc         : ����������/����/����/��ȸ 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001.06.28
'*  8. Modified date(Last)  : 2003.03.23
'*  9. Modifier (First)     : Song,MunGil
'* 10. Modifier (Last)      : Oh, Soo Min
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
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

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
Const BIZ_PGM_ID  = "f5107mb1.asp"										'��: �����Ͻ� ���� ASP�� 
Const JUMP_PGM_ID_NOTE_INF = "f5101ma1"									'����������� 

Dim C_NOTE_STS_NM
Dim C_STS_DT
Dim C_GL_NO	
Dim C_TEMP_GL_NO	
Dim C_SEQ		
Dim C_NOTE_STS	
Dim C_DC_RATE	
Dim C_DC_INT_AMT
Dim C_CHARGE_AMT
Dim C_AMT	
Dim C_BP_CD	
Dim C_BP_NM	
Dim C_BANK_CD
Dim C_BANK_NM
Dim C_BANK_ACCT_NO
Dim C_RCPT_TYPE	
Dim C_RCPT_TYPE_NM
Dim C_CHG_NOTE_ACCT_CD
Dim C_CHG_NOTE_ACCT_NM  
Dim C_NOTE_ACCT_CD
Dim C_NOTE_ACCT_NM
Dim C_DC_INT_ACCT_CD
Dim C_DC_INT_ACCT_NM
Dim C_CHARGE_ACCT_CD
Dim C_CHARGE_ACCT_NM  
Dim C_NOTE_ITEM_DESC

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

<%
Dim dtToday
dtToday = GetSvrDate
%>

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

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
  	lgIntFlgMode = Parent.OPMD_CMODE							'��: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False									'��: Indicates that no value changed
	lgStrPrevKey = ""
	IsOpenPop = False											'��: ����� ���� �ʱ�ȭ 
	lgSortKey = 1
	lgPageNo  = ""
	
	lgBlnFlgChgValue = False
End Sub

Sub initSpreadPosVariables()
	C_NOTE_STS_NM		= 1			
	C_STS_DT			= 2
	C_GL_NO				= 3
	C_TEMP_GL_NO		= 4
	C_SEQ				= 5
	C_NOTE_STS			= 6
	C_DC_RATE			= 7
	C_DC_INT_AMT		= 8	
	C_CHARGE_AMT		= 9
	C_AMT				= 10
	C_BP_CD				= 11
	C_BP_NM				= 12
	C_BANK_CD			= 13
	C_BANK_NM			= 14
	C_BANK_ACCT_NO		= 15
	C_RCPT_TYPE			= 16
	C_RCPT_TYPE_NM		= 17
	C_CHG_NOTE_ACCT_CD	= 18
	C_CHG_NOTE_ACCT_NM	= 19
	C_NOTE_ACCT_CD		= 20
	C_NOTE_ACCT_NM		= 21
	C_DC_INT_ACCT_CD	= 22
	C_DC_INT_ACCT_NM	= 23	
	C_CHARGE_ACCT_CD	= 24
	C_CHARGE_ACCT_NM	= 25 
	C_NOTE_ITEM_DESC	= 26
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
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
	frm1.txtStsDt1.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)

	lgBlnFlgChgValue = False
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()

    With frm1
		.vspdData.Redraw = False
		.vspdData.MaxCols = C_NOTE_ITEM_DESC + 1
		.vspdData.Col = .vspdData.MaxCols	:	.vspdData.ColHidden = True		'��: ������Ʈ�� ��� Hidden Column
		.vspdData.MaxRows = 0

		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    
        Call GetSpreadColumnPos("A")
        
	    ggoSpread.SSSetEdit		C_SEQ,				"����", 8, , , 3
		ggoSpread.SSSetCombo	C_NOTE_STS,			"��������", 12
		ggoSpread.SSSetCombo  	C_NOTE_STS_NM,		"��������", 12
		ggoSpread.SSSetDate		C_STS_DT,			"����", 12, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit		C_GL_NO,			"��ǥ��ȣ", 15, , , 18
		ggoSpread.SSSetEdit		C_TEMP_GL_NO,		"������ǥ��ȣ", 15, , , 18
		ggoSpread.SSSetCombo 	C_RCPT_TYPE,		"", 12
		ggoSpread.SSSetCombo 	C_RCPT_TYPE_NM,		"", 12
		
		ggoSpread.SSSetEdit		C_DC_RATE,			"", 15, , , 18
		ggoSpread.SSSetEdit		C_DC_INT_AMT,		"", 15, , , 18
		ggoSpread.SSSetEdit		C_CHARGE_AMT,		"", 15, , , 18
		ggoSpread.SSSetEdit		C_AMT,				"", 15, , , 18
		ggoSpread.SSSetEdit		C_BP_CD,			"", 15, , , 18
		ggoSpread.SSSetEdit		C_BP_NM,			"", 15, , , 18
		ggoSpread.SSSetEdit		C_BANK_CD,			"", 15, , , 18
		ggoSpread.SSSetEdit		C_BANK_NM,			"", 15, , , 18
		ggoSpread.SSSetEdit		C_BANK_ACCT_NO,		"", 15, , , 18
		ggoSpread.SSSetEdit		C_RCPT_TYPE,		"", 15, , , 18
		ggoSpread.SSSetEdit		C_RCPT_TYPE_NM,		"", 15, , , 18
		
		ggoSpread.SSSetEdit		C_CHG_NOTE_ACCT_CD,	"", 15, , , 18
		ggoSpread.SSSetEdit		C_CHG_NOTE_ACCT_NM,	"", 15, , , 18
		ggoSpread.SSSetEdit		C_NOTE_ACCT_CD,	    "", 15, , , 18
		ggoSpread.SSSetEdit		C_NOTE_ACCT_NM,	    "", 15, , , 18
		ggoSpread.SSSetEdit		C_DC_INT_ACCT_CD,	"", 15, , , 18
		ggoSpread.SSSetEdit		C_DC_INT_ACCT_NM,	"", 15, , , 18		
		ggoSpread.SSSetEdit		C_CHARGE_ACCT_CD,	"", 15, , , 18
		ggoSpread.SSSetEdit		C_CHARGE_ACCT_NM,	"", 15, , , 18
		ggoSpread.SSSetEdit		C_NOTE_ITEM_DESC,	"", 15, , , 500
		
		Call ggoSpread.SSSetColHidden(C_SEQ,C_SEQ,True)
		Call ggoSpread.SSSetColHidden(C_NOTE_STS,C_NOTE_STS,True)
		Call ggoSpread.SSSetColHidden(C_DC_RATE,C_DC_RATE,True)
		Call ggoSpread.SSSetColHidden(C_DC_INT_AMT,C_DC_INT_AMT,True)
		Call ggoSpread.SSSetColHidden(C_CHARGE_AMT,C_CHARGE_AMT,True)
		Call ggoSpread.SSSetColHidden(C_AMT,C_AMT,True)
		Call ggoSpread.SSSetColHidden(C_BP_CD,C_BP_CD,True)
		Call ggoSpread.SSSetColHidden(C_BP_NM,C_BP_NM,True)
		Call ggoSpread.SSSetColHidden(C_BANK_CD,C_BANK_CD,True)
		Call ggoSpread.SSSetColHidden(C_BANK_NM,C_BANK_NM,True)
		Call ggoSpread.SSSetColHidden(C_BANK_ACCT_NO,C_BANK_ACCT_NO,True)
		Call ggoSpread.SSSetColHidden(C_RCPT_TYPE,C_RCPT_TYPE,True)
		Call ggoSpread.SSSetColHidden(C_RCPT_TYPE_NM,C_RCPT_TYPE_NM,True)
		Call ggoSpread.SSSetColHidden(C_CHG_NOTE_ACCT_CD,C_CHG_NOTE_ACCT_CD,True)
		Call ggoSpread.SSSetColHidden(C_CHG_NOTE_ACCT_NM,C_CHG_NOTE_ACCT_NM,True)
		Call ggoSpread.SSSetColHidden(C_NOTE_ACCT_CD,C_NOTE_ACCT_CD,True)
		Call ggoSpread.SSSetColHidden(C_NOTE_ACCT_NM,C_NOTE_ACCT_NM,True)
		Call ggoSpread.SSSetColHidden(C_DC_INT_ACCT_CD,C_DC_INT_ACCT_CD,True)
		Call ggoSpread.SSSetColHidden(C_DC_INT_ACCT_NM,C_DC_INT_ACCT_NM,True)		
		Call ggoSpread.SSSetColHidden(C_CHARGE_ACCT_CD,C_CHARGE_ACCT_CD,True)
		Call ggoSpread.SSSetColHidden(C_CHARGE_ACCT_NM,C_CHARGE_ACCT_NM,True)
		Call ggoSpread.SSSetColHidden(C_NOTE_ITEM_DESC,C_NOTE_ITEM_DESC,True)

    	.vspdData.Redraw = True
    End With
    
	Call SetSpreadLock
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		.ReDraw = False
		 ggoSpread.SpreadLockWithOddEvenRowColor()    
		.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		.vspdData.ReDraw = True
    End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Function InitComboBox()
	'�������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteFg ,lgF0  ,lgF1  ,Chr(11))
	
	'�������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))
	
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_NOTE_STS
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_NOTE_STS_NM
    
	'�������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_RCPT_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_RCPT_TYPE_NM
End Function

Function InitSpreadCombo()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))
	
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_NOTE_STS
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_NOTE_STS_NM
    
	'�������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_RCPT_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_RCPT_TYPE_NM
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
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("f5107ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f5107ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

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

	frm1.txtNoteNoQry.focus
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If UCase(frm1.txtBpCd1.className) = "PROTECTED" Then Exit Function

	arrParam(0) = strCode								'Code Condition
   	arrParam(1) = ""									'ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "T"									'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 

	IsOpenPop = True
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call EScPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If
End Function

'==================================================================================
'	Name : OpenPopUp()
'	Description : �����˾� ���� 
'==================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iTransType

	If IsOpenPop = True Then Exit Function

	'�������� �����ϴ� form�� �˾��� ���� 
	'������ ��ȸ���� �� ������ �⺻ ��������(txtNoteSts1)�� ������ 
	Select Case iWhere
		Case 0		' �������� 
			If frm1.txtNoteSts1.className = Parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "���������˾�"
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"								'popup�� sql�� 
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  " _
						& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD " _
						& " AND B.SEQ_NO = 4 "
			arrParam(5) = frm1.txtNoteSts1.Alt
	
			arrField(0) = "A.MINOR_CD"													' form1�� ������ minor_cd,nmǥ�� 
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtNoteSts1.Alt			
			arrHeader(1) = frm1.txtNoteStsNm1.Alt
		Case 1																			' �ŷ�ó 
			If frm1.txtBpCd1.className = Parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "�ŷ�ó�˾�"												' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER" 												' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = ""															' Where Condition
			arrParam(5) = "�ŷ�ó"													' �����ʵ��� �� ��Ī 

			arrField(0) = "BP_CD"														' Field��(0)
			arrField(1) = "BP_NM"														' Field��(1)

			arrHeader(0) = "�ŷ�ó�ڵ�"												' Header��(0)
			arrHeader(1) = "�ŷ�ó��"												' Header��(1)
		Case 2																			'��/������� 
			If frm1.txtRcptType1.className = Parent.UCN_PROTECTED Then Exit Function

 			arrParam(0) = frm1.txtRcptType1.Alt
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
						& " AND A.MAJOR_CD = B.MAJOR_CD AND A.MINOR_CD = B.MINOR_CD " _
						& " AND B.SEQ_NO = 1 AND B.REFERENCE = " & FilterVar("RP", "''", "S") & "  "
			arrParam(5) = frm1.txtRcptType1.Alt

			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"

			arrHeader(0) = frm1.txtRcptType1.Alt
			arrHeader(1) = frm1.txtRcptTypeNm1.Alt
		Case 3																			' ���� 
			If frm1.txtBankCd1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "���� �˾�"												' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"							' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "										' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(5) = "�����ڵ�"												' �����ʵ��� �� ��Ī 
			arrField(0) = "A.BANK_CD"													' Field��(0)
			arrField(1) = "A.BANK_NM"													' Field��(1)
			arrField(2) = "B.BANK_ACCT_NO"												' Field��(2)
			arrHeader(0) = "�����ڵ�"												' Header��(0)
			arrHeader(1) = "�����"													' Header��(1)
			arrHeader(2) = "���¹�ȣ"												' Header��(2)				
		Case 4																			' ���¹�ȣ 
			If frm1.txtBankAcct1.className = Parent.UCN_PROTECTED Then Exit Function
			arrParam(0) = "���¹�ȣ �˾�"											' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"							' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "										' Where Condition
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "
			arrParam(5) = "���¹�ȣ"												' �����ʵ��� �� ��Ī 
			arrField(0) = "B.BANK_ACCT_NO"												' Field��(0)
			arrField(1) = "A.BANK_CD"													' Field��(0)
			arrField(2) = "A.BANK_NM"													' Field��(0)
			arrHeader(0) = "���¹�ȣ"												' Header��(0)
			arrHeader(1) = "�����ڵ�"												' Header��(0)
			arrHeader(2) = "�����"													' Header��(0)
		Case 5																			' ��/������������ڵ� 
			If frm1.txtNoteAcctCd.className = "protected" Then Exit Function    

			arrParam(0) = "��/��ݰ����˾�"											' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM D	"				' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN002", "''", "S") & "  AND D.TRANS_TYPE = " & FilterVar("FN002", "''", "S") & " " 			' Where Condition
			arrParam(4) = arrParam(4) & " AND C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD  "
			arrParam(4) = arrParam(4) & " AND C.JNL_CD= D.JNL_CD AND D.SEQ = C.SEQ "
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  " 
			arrParam(4) = arrParam(4) & " AND C.JNL_CD =  " & FilterVar(frm1.txtRcptType1.Value, "''", "S") 	 	

			arrParam(5) = frm1.txtNoteAcctCd.Alt										' �����ʵ��� �� ��Ī 

			arrField(0) = "A.ACCT_CD"													' Field��(0)
			arrField(1) = "A.ACCT_NM"													' Field��(1)
			arrField(2) = "B.GP_CD"														' Field��(2)
			arrField(3) = "B.GP_NM"					 									' Field��(3)

			arrHeader(0) = frm1.txtNoteAcctCd.Alt										' Header��(0)
			arrHeader(1) = frm1.txtNoteAcctNm.Alt										' Header��(1)
			arrHeader(2) = "�׷��ڵ�"												' Header��(2)
			arrHeader(3) = "�׷��"													' Header��(3)	
		Case 6																			'���� ��������� 
			If frm1.txtChargeAcctCd.className = "protected" Then Exit Function    

			If UCase(frm1.txtNoteSts1.value) = "DC" Then
				iTransType = "FN002"
			ElseIf 	UCase(frm1.txtNoteSts1.value) = "DH" Then
				iTransType = "FN003"			 
			End If

			arrParam(0) = "����������˾�"											' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM 	D	"			' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar(iTransType, "''", "S") & "  " 										' Where Condition
			arrParam(4) = arrParam(4) & " AND D.TRANS_TYPE = " & FilterVar(iTransType, "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	C.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND 	C.TRANS_TYPE = D.TRANS_TYPE "
			arrParam(4) = arrParam(4) & " AND 	C.JNL_CD = D.JNL_CD "
			arrParam(4) = arrParam(4) & "	 AND 	C.DR_CR_FG = D.DR_CR_FG "
			arrParam(4) = arrParam(4) & "	 AND 	C.SEQ = D.SEQ "
			arrParam(4) = arrParam(4) & "	 AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "   "
			arrParam(4) = arrParam(4) & " AND  C.JNL_CD = " & FilterVar("FEE", "''", "S") & "   " 
			arrParam(5) = frm1.txtChargeAcctCd.Alt										' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"													' Field��(0)
			arrField(1) = "A.Acct_NM"													' Field��(1)
			arrField(2) = "B.GP_CD"														' Field��(2)
			arrField(3) = "B.GP_NM"														' Field��(3)

			arrHeader(0) = frm1.txtChargeAcctCd.Alt										' Header��(0)
			arrHeader(1) = frm1.txtChargeAcctNm.Alt										' Header��(1)
			arrHeader(2) = "�׷��ڵ�"												' Header��(2)
			arrHeader(3) = "�׷��"													' Header��(3)
		Case 7																			'����/�ε��������� 
			If frm1.txtChgNoteAcctCd.className = "protected" Then Exit Function    
			
			arrParam(0) = "���������˾�"											' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A"													' TABLE ��Ī 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition			
			If UCase(Trim(frm1.txtNoteSts1.Value)) = "DC" Then
				arrParam(4) = "A.ACCT_TYPE = " & FilterVar("D2", "''", "S") & "  " 									' Where Condition												
			ElseIf  UCase(Trim(frm1.txtNoteSts1.Value)) = "DH" Then
				arrParam(4) = "A.ACCT_TYPE = " & FilterVar("D4", "''", "S") & "  " 									' Where Condition
			End If
			arrParam(4) = arrParam(4) & " AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  "			
			arrParam(5) = frm1.txtChgNoteAcctCd.Alt										' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"													' Field��(0)
			arrField(1) = "A.Acct_NM"													' Field��(1)			
			
			arrHeader(0) = frm1.txtChgNoteAcctCd.Alt													' Header��(0)
			arrHeader(1) = frm1.txtChgNoteAcctNm.Alt													' Header��(1)
		Case 8																							' ���� ��������� 
			If frm1.txtDcIntAcctCd.className = "protected" Then Exit Function    

			arrParam(0) = "��������(���η�)�����˾�"												' �˾� ��Ī 
			arrParam(1) = "A_ACCT	A,A_ACCT_GP 	B, A_JNL_ACCT_ASSN	C,	A_JNL_FORM 	D	"			' TABLE ��Ī 
			arrParam(2) = strCode																		' Code Condition
			arrParam(3) = ""																			' Name Cindition
			arrParam(4) = "C.TRANS_TYPE = " & FilterVar("FN002", "''", "S") & "  " 						' Where Condition
			arrParam(4) = arrParam(4) & " AND D.TRANS_TYPE = " & FilterVar("FN002", "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND 	C.ACCT_CD = A.ACCT_CD AND A.GP_CD = B.GP_CD "
			arrParam(4) = arrParam(4) & " AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND 	C.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "
			arrParam(4) = arrParam(4) & " AND 	C.TRANS_TYPE = D.TRANS_TYPE "
			arrParam(4) = arrParam(4) & " AND 	C.JNL_CD = D.JNL_CD "
			arrParam(4) = arrParam(4) & "	 AND 	C.DR_CR_FG = D.DR_CR_FG "
			arrParam(4) = arrParam(4) & "	 AND 	C.SEQ = D.SEQ "
			arrParam(4) = arrParam(4) & "	 AND 	A.DEL_FG <> " & FilterVar("Y", "''", "S") & "   "
			arrParam(4) = arrParam(4) & " AND  C.JNL_CD = " & FilterVar("DCINT", "''", "S") & "   " 
			arrParam(5) = frm1.txtDcIntAcctCd.Alt										' �����ʵ��� �� ��Ī 

			arrField(0) = "A.Acct_CD"													' Field��(0)
			arrField(1) = "A.Acct_NM"													' Field��(1)
			arrField(2) = "B.GP_CD"														' Field��(2)
			arrField(3) = "B.GP_NM"														' Field��(3)

			arrHeader(0) = frm1.txtDcIntAcctCd.Alt										' Header��(0)
			arrHeader(1) = frm1.txtDcIntAcctNm.Alt										' Header��(1)
			arrHeader(2) = "�׷��ڵ�"												' Header��(2)
			arrHeader(3) = "�׷��"													' Header��(3)			
	End Select
  
	IsOpenPop = True
	If (iWhere = 3 or iWhere = 4) Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EScPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' �������� 
				.txtNoteSts1.value   = arrRet(0)
				.txtNoteStsNm1.value = arrRet(1)
				.txtNoteSts1.focus
				Call txtNoteSts1_OnChange
				lgBlnFlgChgValue = True
			Case 1		' �輭�ŷ�ó 
				.txtBpCd1.value = arrRet(0)
				.txtBpNM1.value = arrRet(1)
				.txtBpCd1.focus
				lgBlnFlgChgValue = True
			Case 2		' ��/������� 
				.txtRcptType1.value   = arrRet(0)
				.txtRcptTypeNm1.value = arrRet(1)
				.txtRcptType1.focus
				Call txtRcptType1_OnChange
				lgBlnFlgChgValue = True
			Case 3		' ���� 
				.txtBankCd1.value	= arrRet(0)
				.txtBankNm1.value	= arrRet(1)
				.txtBankAcct1.value =  arrRet(2)
				.txtBankCd1.focus
				lgBlnFlgChgValue = True
			Case 4		' ���¹�ȣ 
				.txtBankAcct1.value =  arrRet(0)
				.txtBankCd1.value	= arrRet(1)
				.txtBankNm1.value	= arrRet(2)					
				.txtBankAcct1.focus
				lgBlnFlgChgValue = True
			Case 5		' ��/��ݰ����ڵ� 
				.txtNoteAcctCd.value   = arrRet(0)
				.txtNoteAcctNm.value = arrRet(1)
				.txtNoteAcctCd.focus
				lgBlnFlgChgValue = True
			Case 6	'����������ڵ� 
				.txtChargeAcctCd.value = arrRet(0)
				.txtChargeAcctNm.value = arrRet(1)
				.txtChargeAcctCd.focus
				lgBlnFlgChgValue = True
			Case 7		' ���������ڵ� 
				.txtChgNoteAcctCd.value = arrRet(0)
				.txtChgNoteAcctNm.value = arrRet(1)
				.txtChgNoteAcctCd.focus
				lgBlnFlgChgValue = True
			Case 8		'��������(���η�) �����ڵ� 
				.txtDcIntAcctCd.value   = arrRet(0)
				.txtDcIntAcctNm.value   = arrRet(1)
				.txtDcIntAcctCd.focus
				lgBlnFlgChgValue = True
		End Select
	End With
End Function

Function EScPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' �������� 
				.txtNoteSts1.focus
			Case 1		' �輭�ŷ�ó 
				.txtBpCd1.focus
			Case 2		' ��/������� 
				.txtRcptType1.focus
			Case 3		' ���� 
				.txtBankCd1.focus
			Case 4		' ���¹�ȣ 
				.txtBankAcct1.focus
			Case 5		' ��/��ݰ����ڵ� 
				.txtNoteAcctCd.focus
			Case 6	'����������ڵ� 
				.txtChargeAcctCd.focus
			Case 7		' ���������ڵ� 
				.txtChgNoteAcctCd.focus
			Case 8		'��������(���η�) �����ڵ� 
				.txtDcIntAcctCd.focus				
		End Select
	End With
End Function

'============================================================
'ȸ����ǥ �˾� 
'============================================================
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

	arrParam(0) = Trim(frm1.txtGlNo1.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = ""			'Reference��ȣ 

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'============================================================
'������ǥ �˾� 
'============================================================
Function OpenPopupTempGL()
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo1.value)	'������ǥ��ȣ 
	arrParam(1) = ""							'Reference��ȣ 

	IsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/a5130ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

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
		Case JUMP_PGM_ID_NOTE_INF	'����������� 
			strTemp = frm1.txtNoteNo.value 
			Call WriteCookie("NOTE_NO", strTemp)
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
    Call LoadInfTB19029	    										'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.ClearField(Document, "1")							'��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData
    
    Call ggoOper.ClearField(Document, "3")
    Call ggoOper.LockField(Document, "N")							'��: Lock  Suitable  Field

	Call InitSpreadSheet											'Setup the Spread sheet
	Call InitComboBox
    Call InitVariables												'��: Initializes local global variables
    Call SetToolbar("1110100000001111")	
    Call CookiePage("FORM_LOAD")
	Call SetDefaultVal

    frm1.txtNoteNoQry.focus

    Set gActiveElement = document.activeElement
    
	' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing    
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
			C_NOTE_STS_NM		= iCurColumnPos(1)
			C_STS_DT			= iCurColumnPos(2)
			C_GL_NO				= iCurColumnPos(3) 
			C_TEMP_GL_NO		= iCurColumnPos(4)	
			C_SEQ				= iCurColumnPos(5)
			C_NOTE_STS			= iCurColumnPos(6)
			C_DC_RATE			= iCurColumnPos(7)
			C_DC_INT_AMT		= iCurColumnPos(8)
			C_CHARGE_AMT		= iCurColumnPos(9)
			C_AMT	            = iCurColumnPos(10)
			C_BP_CD	            = iCurColumnPos(11)
			C_BP_NM	            = iCurColumnPos(12)
			C_BANK_CD           = iCurColumnPos(13)
			C_BANK_NM           = iCurColumnPos(14)
			C_BANK_ACCT_NO      = iCurColumnPos(15)
			C_RCPT_TYPE 	    = iCurColumnPos(16)
			C_RCPT_TYPE_NM      = iCurColumnPos(17)
			C_CHG_NOTE_ACCT_CD	= iCurColumnPos(18)
			C_CHG_NOTE_ACCT_NM  = iCurColumnPos(19)
			C_NOTE_ACCT_CD	    = iCurColumnPos(20)
			C_NOTE_ACCT_NM      = iCurColumnPos(21)
			C_DC_INT_ACCT_CD    = iCurColumnPos(22)
			C_DC_INT_ACCT_NM    = iCurColumnPos(23)			
			C_CHARGE_ACCT_CD    = iCurColumnPos(24)
			C_CHARGE_ACCT_NM    = iCurColumnPos(25)
			C_NOTE_ITEM_DESC    = iCurColumnPos(26)
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

		frm1.vspdData.Col = C_NOTE_STS
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_NOTE_STS_NM
		frm1.vspdData.value = intindex

		frm1.vspdData.Col = C_RCPT_TYPE
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_RCPT_TYPE_NM
		frm1.vspdData.value = intindex
	Next
End Sub

'=======================================================================================================
'   Event Name : _DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStsDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtStsDt1.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtStsDt1.Focus           
    End If
End Sub

'=======================================================================================================
'   Event Name : _Change()
'   Event Desc : Changed Setting
'=======================================================================================================
Sub txtStsDt1_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtSttlAmt1_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtDcRate1_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtDcIntAmt1_Change()
    lgBlnFlgChgValue = True
    If (unicdbl(frm1.txtDcIntAmt1.Text) > 0 And Trim(UCase(frm1.txtNoteSts1.value)) <> "DH" )	Then
		If Trim(UCase(frm1.txtNoteSts1.value)) = "DS"  Then
			Call ggoOper.SetReqAttr(frm1.txtDcIntAcctCd, "Q")
		Else			
			Call ggoOper.SetReqAttr(frm1.txtDcIntAcctCd, "N")
		End If
	Else 
		frm1.txtDcIntAcctCd.value = ""
		frm1.txtDcIntAcctNm.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtDcIntAcctCd, "Q")			
	End If		
End Sub

Sub txtChargeAmt1_Change()
    lgBlnFlgChgValue = True
    If unicdbl(frm1.txtChargeAmt1.Text) > 0 Then
		If Trim(UCase(frm1.txtNoteSts1.value)) = "DS"  Then
			Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "Q")
		Else			
			Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "N")
		End If
	Else 
		frm1.txtChargeAcctCd.value = ""
		frm1.txtChargeAcctNm.value = ""		
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "Q")			
	End If		
End Sub

'=======================================================================================================
'   Event Desc : ��/��������� Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptType1_OnChange()
	Dim strval

	strval = UCase(frm1.txtRcptType1.value)

	If Trim(frm1.txtRcptType1.value) <> "" Then
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "Q")
	End If

	'�����ڵ�, ���¹�ȣ Protected Setting
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		Select Case UCase(lgF0)
			Case "DP" & Chr(11)			' ������ 
				Call ggoOper.SetReqAttr(frm1.txtBankCd1, "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "N")
				frm1.txtBankCd1.focus
			Case Else
				frm1.txtBankCd1.value = ""
				frm1.txtBankNm1.value = ""
				frm1.txtBankAcct1.value = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "Q")
		End Select
	Else
		frm1.txtBankCd1.value = ""
		frm1.txtBankNm1.value = ""
		frm1.txtBankAcct1.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "Q")
	End If

	frm1.txtNoteAcctCd.value = ""
	frm1.txtNoteAcctNm.value = ""
End Sub

'=======================================================================================================
'   Event Desc : ��/��������� Set Protected/Required Fields
'=======================================================================================================
Sub txtRcptType1_Change()
	Dim strval
	
	strval = UCase(frm1.txtRcptType1.value)
	
	If Trim(frm1.txtRcptType1.value) <> "" Then
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "Q")
	End If
	
	'�����ڵ�, ���¹�ȣ Protected Setting	
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strval , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then 
		Select Case UCase(lgF0)
			Case "DP" & Chr(11)			' ������ 
				Call ggoOper.SetReqAttr(frm1.txtBankCd1, "N")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "N")
				frm1.txtBankCd1.focus
			Case Else
				frm1.txtBankCd1.value = ""
				frm1.txtBankNm1.value = ""
				frm1.txtBankAcct1.value = ""
				Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "Q")
		End Select
	Else
		frm1.txtBankCd1.value = ""
		frm1.txtBankNm1.value = ""
		frm1.txtBankAcct1.value = ""
		Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "Q")
	End If

	frm1.txtNoteAcctCd.value = ""
	frm1.txtNoteAcctNm.value = ""
End Sub

'=======================================================================================================
'   Event Desc : �������º� Set Protected/Required Fields
'=======================================================================================================
Sub txtNoteSts1_OnChange()
	'�������º� Protected/Required
	frm1.txtRcptType1.value = ""
	frm1.txtRcptTypeNm1.value = ""
	frm1.txtChgNoteAcctCd.value = ""
	frm1.txtChgNoteAcctNm.value = ""
	frm1.txtNoteAcctCd.value = ""
	frm1.txtNoteAcctNm.value = ""

	Select Case UCase(frm1.txtNoteSts1.value)
		Case "DC"	'����		
			Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "N")
			Call ggoOper.SetReqAttr(frm1.txtRcptType1, "N")									
			Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtSttlAmt1, "Q")	'N:Required, Q:Protected, D:Default
			Call ggoOper.SetReqAttr(frm1.txtDcRate1, "N")
			Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "D")
			Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "D")
			Call ggoOper.SetReqAttr(frm1.txtDesc, "D")
			Call txtRcptType1_OnChange()			
		Case "DH"	'�ε�		
			Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "N")
			Call ggoOper.SetReqAttr(frm1.txtRcptType1, "N")			
			Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")			
			Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "D")
			Call ggoOper.SetReqAttr(frm1.txtDesc, "D")
			Call txtRcptType1_OnChange()
		Case "ED"	'�輭			
			Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "Q")			
			Call ggoOper.SetReqAttr(frm1.txtRcptType1, "Q")			
			Call ggoOper.SetReqAttr(frm1.txtBpCd1, "N")
			Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")			
			Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "Q")
			Call ggoOper.SetReqAttr(frm1.txtDesc, "D")
			Call txtRcptType1_OnChange()
		Case Else
	End Select

	frm1.txtDcRate1.text = 0
	frm1.txtDcIntAmt1.text = 0
	frm1.txtChargeAmt1.text = 0
End Sub

'=======================================================================================================
'   Event Desc : �������º� Set Protected/Required Fields
'=======================================================================================================
Sub txtNoteSts1_Change()
	'�������º� Protected/Required
	frm1.txtBpCd1.value = ""
	frm1.txtBpNM1.value = ""
	frm1.txtChgNoteAcctCd.value = ""
	frm1.txtChgNoteAcctNm.value = ""
	frm1.txtRcptType1.value = ""
	frm1.txtRcptTypeNm1.value = ""
	frm1.txtNoteAcctCd.value = ""
	frm1.txtNoteAcctNm.value = ""

	frm1.txtDcRate1.text = 0
	frm1.txtDcIntAmt1.text = 0	
	frm1.txtChargeAmt1.text = 0
End Sub

'==========================================================================================
'   Event Desc : Set Data from Spread
'==========================================================================================
Function SetNoteItemData(ByVal Row)
	With frm1.vspdData
		.Row = Row

		.Col = C_SEQ			:	frm1.txtSeq.value = .Text
		.Col = C_NOTE_STS		:	frm1.txtNoteSts1.value = .Text
		.Col = C_NOTE_STS_NM	:	frm1.txtNoteStsNm1.value = .Text
		Call txtNoteSts1_Change

		.Col = C_STS_DT			:	frm1.txtStsDt1.Text = .Text
		.Col = C_BP_CD			:	frm1.txtBpCd1.value = .Text
		.Col = C_BP_NM			:	frm1.txtBpNM1.value = .Text
		.Col = C_DC_RATE		:	frm1.txtDcRate1.Text = .Text
		.Col = C_DC_INT_AMT		:	frm1.txtDcIntAmt1.Text = .Text		
		.Col = C_CHARGE_AMT		:	frm1.txtChargeAmt1.Text = .Text
		.Col = C_RCPT_TYPE		:	frm1.txtRcptType1.value = .Text
		.Col = C_RCPT_TYPE_NM	:	frm1.txtRcptTypeNm1.value = .Text
		Call txtRcptType1_Change

		.Col = C_BANK_CD		:	frm1.txtBankCd1.value = .Text
		.Col = C_BANK_NM		:	frm1.txtBankNm1.value = .Text
		.Col = C_BANK_ACCT_NO	:	frm1.txtBankAcct1.value = .Text
		.Col = C_GL_NO			:	frm1.txtGlNo1.value = .Text
		.Col = C_TEMP_GL_NO		:	frm1.txtTempGlNo1.value = .Text

		.Col = C_CHG_NOTE_ACCT_CD	:	frm1.txtChgNoteAcctCd.value = .Text
		.Col = C_CHG_NOTE_ACCT_NM	:	frm1.txtChgNoteAcctNm.value = .Text
		.Col = C_NOTE_ACCT_CD		:	frm1.txtNoteAcctCd.value = .Text
		.Col = C_NOTE_ACCT_NM		:	frm1.txtNoteAcctNm.value = .Text
		.Col = C_DC_INT_ACCT_CD		:	frm1.txtDcIntAcctCd.value = .Text
		.Col = C_DC_INT_ACCT_NM		:	frm1.txtDCIntAcctNm.value = .Text
		.Col = C_CHARGE_ACCT_CD		:	frm1.txtChargeAcctCd.value = .Text
		.Col = C_CHARGE_ACCT_NM		:	frm1.txtChargeAcctNm.value = .Text
		.Col = C_NOTE_ITEM_DESC		:	frm1.txtDesc.value = .Text
	End With
	
'	Call txtChargeAmt1_Change()			'DH�� ���, ������ ���� Protect
	Call txtDcIntAmt1_Change()			'DH�� ���, ��������(���η�) ���� Protect 	 

	If frm1.vspdData.row < frm1.vspdData.MaxRows Then
		Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtStsDt1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtRcptType1, "Q")		
		Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "Q")		
		Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")		
		Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "Q")		
		Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")				
		Call ggoOper.SetReqAttr(frm1.txtDcIntAcctCd, "Q")				
		Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "Q")		
		Call ggoOper.SetReqAttr(frm1.txtChargeAcctCd, "Q")
		Call ggoOper.SetReqAttr(frm1.txtDesc, "Q")				
		Call SetToolbar("1110000000001111")
	Else
		Select Case UCase(frm1.txtNoteSts1.value)
			Case "DC"	'���� 
				Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtStsDt1, "N")
				Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "N")
				Call ggoOper.SetReqAttr(frm1.txtRcptType1, "N")	
				Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtSttlAmt1, "Q")	'N:Required, Q:Protected, D:Default
				Call ggoOper.SetReqAttr(frm1.txtDcRate1, "N")
				Call ggoOper.SetReqAttr(frm1.txtDCIntAmt1, "D")				
				Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "D")
				Call ggoOper.SetReqAttr(frm1.txtDesc, "D")
				Call SetToolbar("1111100000001111")
			Case "DH"	'�ε� 
				Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtStsDt1, "N")
				Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "N")
				Call ggoOper.SetReqAttr(frm1.txtRcptType1, "N")			
				Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")							
				Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "D")
				Call ggoOper.SetReqAttr(frm1.txtDesc, "D")
				Call SetToolbar("1111100000001111")
			Case "ED"	'�輭 
				Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtStsDt1, "N")
				Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptType1, "Q")		
				Call ggoOper.SetReqAttr(frm1.txtBpCd1, "N")
				Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDCIntAmt1, "Q")							
				Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDesc, "D")
				Call SetToolbar("1111100000001111")
			Case "SE"	'�輭�絵 
				Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtStsDt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptType1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")							
				Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDesc, "Q")
				Call SetToolbar("1110100000001111")
			Case "SM","MV" 	'�����Ϸ�, �̵��� 
				Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtStsDt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptType1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtNoteAcctCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtSttlAmt1, "Q")	'N:Required, Q:Protected, D:Default
				Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBankCd1, "Q")		
				Call ggoOper.SetReqAttr(frm1.txtBankAcct1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDesc, "Q")
				Call SetToolbar("1110000000001111")
			Case "DS"	'�ε�ó�� 
				Call ggoOper.SetReqAttr(frm1.txtNoteSts1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtStsDt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtChgNoteAcctCd, "Q")
				Call ggoOper.SetReqAttr(frm1.txtRcptType1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtBpCd1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcRate1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDcIntAmt1, "Q")															
				Call ggoOper.SetReqAttr(frm1.txtChargeAmt1, "Q")
				Call ggoOper.SetReqAttr(frm1.txtDesc, "D")				
				Call SetToolbar("1111100000001111")
			Case Else
		End Select
	End If

	lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False
End Function

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 
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
	Else
		Call SetNoteItemData(Row)
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

'============================================================================
'�� �̵��� Data Display
'============================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow And NewRow > 0 Then
		Call vspdData_Click(Col, NewRow)
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_NOTE_STS_NM Or NewCol <= C_NOTE_STS_NM Then
        Cancel = True
        Exit Sub
    End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
	'Erase contents area
	'----------------------- 
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.ClearField(Document, "3")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData		
    Call InitVariables														'��: Initializes local global variables
   
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
    
    FncNew = False															'��: Processing is NG
    
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True Then
	    IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")		'�� �ٲ�κ� 
	     If IntRetCD = vbNo Then
	         Exit Function
	     End If
	End If
    
	Call ggoOper.ClearField(Document, "3")									'��: Clear Condition Field        
	
    frm1.txtChgNoteAcctCd.value = ""
    frm1.txtChgNoteAcctNm.value = ""
    Call ggoOper.LockField(Document, "N")									'��: Lock  Suitable  Field

    Call InitVariables														'��: Initializes local global variables
    Call SetToolbar("1110100000001111")										'��: ��ư ���� ���� 

    FncNew = True															'��: Processing is OK
	
	frm1.txtStsDt1.text = UniConvDateAToB("<%=dtToday%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtNoteSts1.focus
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
        intRetCD = DisplayMsgBox("141433","x","x","x")						'��ȸ�� �����ϼ���.
        Exit Function
    End If
    
    If Trim(frm1.txtNoteNo.value) = "" Then
		intRetCD = DisplayMsgBox("970029","x",frm1.txtNoteNoQry.Alt,"x")	'~�� Ȯ���ϼ���.
		frm1.txtNoteNoQry.focus
		Exit Function
    End If
    
    If Trim(frm1.txtSeq.value) = "" Then
		intRetCD = DisplayMsgBox("900025","x","x","x")						'���õ� �׸��� �����ϴ�.
		Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")			'�����Ͻðڽ��ϱ�?
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
    If Not chkField(Document, "3") Then										'��: Check contents area
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

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
     On Error Resume Next																'��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    On Error Resume Next																'��: Protect system from crashing
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
    Call parent.FncExport(Parent.C_MULTI)												'��: ȭ�� ���� 
    Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)											'��:ȭ�� ����, Tab ���� 
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = Parent.SS_ACTION_ACTIVE_CELL
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
End Function

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
    Call InitSpreadCombo()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")				'�� �ٲ�κ� 
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
    Dim strVal
    
    Err.Clear																		'��: Protect system from crashing
    
    DbDelete = False																'��: Processing is NG
    
	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtNoteNo=" & UCase(Trim(frm1.txtNoteNo.value))				'��: ���� ���� ����Ÿ 
	strVal = strVal & "&txtSeq=" & UCase(Trim(frm1.txtSeq.value))	
	
	Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True																	'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()																'��: ���� ������ ���� ���� 
    Call InitVariables
    Call MainQuery()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    
    Err.Clear																		'��: Protect system from crashing
    
    DbQuery = False																	'��: Processing is NG
    
	Call LayerShowHide(1)
	
	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode		= " & Parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtNoteNoQry	= " & Trim(.txtNoteNoQry.value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey	= " & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode		= " & Parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtNoteNoQry	= " & Trim(.txtNoteNoQry.value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey	= " & "0"
		End If
		strVal = strVal & "&lgPageNo	=" & lgPageNo
		strVal = strVal & "&txtMaxRows	=" & .vspdData.MaxRows
	End With

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����											

	Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 

    DbQuery = True																	'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()																'��: ��ȸ ������ ������� 
	Call InitData
	Call SetToolbar("1111100000011111")
	Call vspdData_Click(1, 1)
     
	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = Parent.OPMD_UMODE											'��: Indicates that current mode is Update mode
	Else
		lgIntFlgMode = Parent.OPMD_CMODE
	End If
   
	frm1.txtSttlAmt1.text	= frm1.txtNoteAmt.Text
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 
	Dim strVal

    Err.Clear																		'��: Protect system from crashing

	DbSave = False																	'��: Processing is NG

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		
		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end		

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True																	'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()																	'��: ���� ������ ���� ���� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A> &nbsp;|
											<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
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
									<TD CLASS=TD5 NOWRAP>������ȣ</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="CLSTXT" TYPE="TEXT" ID="txtNoteNoQry" NAME="txtNoteNoQry" SIZE=30 MAXLENGTH=30  tag="12XXXU"ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteQry" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenNoteInfo"></TD>
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
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="��������" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="��������" STYLE="WIDTH: 100px" tag="24X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtBpCd" NAME="txtBpCd" SIZE=12 MAXLENGTH=10   tag="24XXXU" ALT="�ŷ�ó">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBpNM" NAME="txtBpNM" SIZE=25 MAXLENGTH=20  STYLE="TEXT-ALIGN: left" tag="24X" ALT="�ŷ�ó"> </TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=12 MAXLENGTH=10  tag="24XXXU" ALT="����">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=25 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="����"> </TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtIssueDt" NAME="txtIssueDt" SIZE=12 MAXLENGTH=10  STYLE="TEXT-ALIGN: center" tag="24X" ALT="������"></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtDueDt" NAME="txtDueDt" SIZE=12 MAXLENGTH=10  STYLE="TEXT-ALIGN: center" tag="24X" ALT="������"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
								<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNoteAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����ݾ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
								<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSttlAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����ݾ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							</TR>
							<TR HEIGHT=80%>
								<TD WIDTH=50% COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=OBJECT1 tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH=50% COLSPAN=2>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtNoteSts1" NAME="txtNoteSts1" SIZE=10 MAXLENGTH=2 tag="33XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteSts" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteSts1.Value, 0)">&nbsp;<INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtNoteStsNm1" NAME="txtNoteStsNm1" SIZE=20 STYLE="TEXT-ALIGN: left" tag="34X" ALT="�������¸�"></TD>
										</TR>										
										<TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>												
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChgNoteAcctCd" ALT="��������" SIZE="10" MAXLENGTH="20"  tag="34X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChgNoteAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtChgNoteAcctCd.value, 7)">
																   <INPUT NAME="txtChgNoteAcctNm" ALT="����������" SIZE="20" tag="34X"></TD>
										</TR>
									    <TR>
											<TD CLASS=TD5 NOWRAP>��������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID=fpStsDt1 NAME=txtStsDt1 CLASS=FPDTYYYYMMDD TITLE=FPDATETIME ALT="��������" tag="32X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�輭�ŷ�ó</TD>
											<TD CLASS=TD6 NOWRAP><INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtBpCd1" NAME="txtBpCd1" SIZE=10 MAXLENGTH=10 tag="31XXXU" ALT="�輭�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd1.Value, 1)">&nbsp;<INPUT CLASS="CLSTXT" TYPE=TEXT ID="txtBpNM1" NAME="txtBpNM1" SIZE=20 STYLE="TEXT-ALIGN: left" tag="34X" ALT="�輭�ŷ�ó��"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>�����ݾ�</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=fpSttlAmt1 NAME=txtSttlAmt1 CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="�����ݾ�" tag="34X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��/�������</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRcptType1" ALT="��/�������" SIZE="10" MAXLENGTH="2"  tag="31XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRcptType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtRcptType1.value, 2)">&nbsp;<INPUT NAME="txtRcptTypeNm1" ALT="��/���������" STYLE="TEXT-ALIGN: Left" tag="34X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��/��ݰ���</TD>												
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNoteAcctCd" ALT="��/��ݰ���" SIZE="10" MAXLENGTH="20"  tag="31XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteAcctCd.value, 5)">
																   <INPUT NAME="txtNoteAcctNm" ALT="��/��ݰ�����" SIZE="20" tag="34X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>����</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankCd1" ALT="����" SIZE="10" MAXLENGTH="10"  tag="31XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd1.value, 3)">&nbsp;<INPUT NAME="txtBankNm1" ALT="�����" STYLE="TEXT-ALIGN: Left" tag="34X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankAcct1" ALT="���¹�ȣ" SIZE="18" MAXLENGTH="30" tag="31XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcct" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcct1.value, 4)"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=fpDcRate1 NAME=txtDcRate1 CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="������" tag="31X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������(���η�)</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=fpDcIntAmt1 NAME=txtDcintAmt1 CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="��������(���η�)" tag="34X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��������(���η�)����</TD>												
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDcIntAcctCd" ALT="��������(���η�)����"   SIZE="10" MAXLENGTH="20"  tag="34X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDcIntAcctCd.value, 8)">
																   <INPUT NAME="txtDcIntAcctNm" ALT="��������(���η�)������" SIZE="20" tag="34X"></TD>
										</TR>										
										<TR>
											<TD CLASS=TD5 NOWRAP>������</TD>
											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=fpChargeAmt1 NAME=txtChargeAmt1 CLASS=FPDS140 TITLE=FPDOUBLESINGLE ALT="������" tag="34X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���������</TD>												
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChargeAcctCd" ALT="���������"   SIZE="10" MAXLENGTH="20"  tag="34X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtChargeAcctCd.value, 6)">
																   <INPUT NAME="txtChargeAcctNm" ALT="�����������" SIZE="20" tag="34X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>��ǥ��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo1" ALT="��ǥ��ȣ" SIZE="18" MAXLENGTH="30"  tag="34XXXU"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>������ǥ��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTempGlNo1" ALT="������ǥ��ȣ" SIZE="18" MAXLENGTH="30"  tag="34XXXU"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>���</TD>
											<TD CLASS=TD6 NOWRAP><INPUT ID="txtDesc" NAME=txtDesc ALT="���" MAXLENGTH=128 SIZE=36 tag="3X"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
											<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
										</TR>
									</TABLE>
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
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_NOTE_INF)">�����������</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtNoteNo"			tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtEndorseFg"		tag="2" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtSeq"			tag="34" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

 
