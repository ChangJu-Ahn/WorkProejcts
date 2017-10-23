
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3101ma1
'*  4. Program Name         : �����ݵ�� 
'*  5. Program Desc         : Register of Deposit Master
'*  6. Comproxy List        : FD0011, FD0019
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Kim, Jong Hwan
'* 10. Modifier (Last)      : Kim, Hee Jung
'* 11. Comment              : 2001.05.31 Song,MunGil ��������ݿ�/�ڱ��ʵ��߰�����ݿ� 
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->				
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' ������ ������ ���� Coding


Const BIZ_PGM_ID = "f3101mb1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "f3101mb2.asp"	

Const JUMP_PGM_ID_BANK_REP = "b1310ma1"										 '��: Jump Page to ����������� 

 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgBlnFlgConChg				'��: Condition ���� Flag

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgCurName()					'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim cboOldVal          
Dim IsOpenPop          

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
    lgIntFlgMode = parent.OPMD_CMODE   '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '��: Indicates that no value changed
    lgIntGrpCount = 0           '��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'��: ����� ���� �ʱ�ȭ 

End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
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

    if frm1.cboTransSts.length > 0 then
       frm1.cboTransSts.selectedindex = 0
    end if
	frm1.hTemp.value = ""
	frm1.txtStartDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 						
    frm1.hOrgChangeId.value = Parent.gChangeOrgId	
	frm1.txtDocCur.value	= Parent.gCurrency
	frm1.txtXchRate.text	= 1

	frm1.hTemp.value = ""

End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
		
	Dim arrData
	
	'�����ݱ��� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3011", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDpstFg ,lgF0  ,lgF1  ,Chr(11))
	
	'���������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDpstType ,lgF0  ,lgF1  ,Chr(11))
	
	'�ŷ����� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3014", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTransSts ,lgF0  ,lgF1  ,Chr(11))
	
	'�������� 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3013", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboBankAcctFg ,lgF0  ,lgF1  ,Chr(11))
End Sub


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
 '------------------------------------------  OpenRefBankAcctNo()  ----------------------------------------
'	Name : OpenRefBankAcctNo()
'	Description : ����������� 
'--------------------------------------------------------------------------------------------------------- 
'Function OpenRefBankAcctNo()
'	Dim arrRet
'	Dim arrParam(6), arrField(6), arrHeader(6)
	
'	If IsOpenPop = True Then Exit Function
	
'	arrParam(0) = "�����������"										' �˾� ��Ī 
'	arrParam(1) = "B_BANK A, B_BANK_ACCT B, B_MINOR C, B_MINOR D "			' TABLE ��Ī 
'	arrParam(2) = ""														' Code Condition
'	arrParam(3) = ""														' Name Cindition
'	arrParam(4) = "A.BANK_CD = B.BANK_CD AND (B.BP_CD IS NULL OR B.BP_CD = ' ') "
'	arrParam(4) = arrParam(4) & "AND C.MAJOR_CD = 'F3011' AND C.MINOR_CD = B.BANK_ACCT_TYPE "	
'	arrParam(4) = arrParam(4) & "AND D.MAJOR_CD = 'F3012' AND D.MINOR_CD = B.DPST_TYPE "		' Where Condition	
'	arrParam(5) = "�����ڵ�"											' �����ʵ��� �� ��Ī 

'	arrField(0) = "A.BANK_CD"								' Field��(0)
'	arrField(1) = "A.BANK_NM"								' Field��(1)
'	arrField(2) = "B.BANK_ACCT_NO"							' Field��(2)
'	arrField(3) = "C.MINOR_NM"								' Field��(3)
'   arrField(4) = "D.MINOR_NM"								' Field��(4)
'   arrField(5) = "HH" & Parent.gColSep & "C.MINOR_CD"		' Field��(5) - Hidden
'	arrField(6) = "HH" & Parent.gColSep & "D.MINOR_CD"		' Field��(6) - Hidden
    
'	arrHeader(0) = "�����ڵ�"							' Header��(0)
'	arrHeader(1) = "�����"								' Header��(1)
'	arrHeader(2) = "���¹�ȣ"							' Header��(2)
'	arrHeader(3) = "�����ݱ���"							' Header��(3)
'	arrHeader(4) = "����������"							' Header��(4)	
	
'	IsOpenPop = True
	
'	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
'			"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

'	IsOpenPop = False
	
'	If arrRet(0) = "" Then
'		Exit Function
'	Else
'		frm1.txtBankCd.value		= arrRet(0)
'		frm1.txtBankNM.value		= arrRet(1)
'		frm1.txtBankAcctNo.value	= arrRet(2)
'		frm1.cboDpstFg.Value		= arrRet(5)
'		frm1.cboDpstType.Value		= arrRet(6)
				
'	End If
	
'	Call cboDpstFg_Change()
	
'	frm1.txtBankAcctNo.focus
	
'End Function

Function OpenRefBankAcctNo(ByVal iOpt1, Byval iOpt2)
	Dim arrRet
	Dim arrParam(11)	                           '���Ѱ��� �߰� (3 -> 4)
	Dim IntRetCD	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("f3101ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f3101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
'	arrParam(4)	= lgAuthorityFlag              '���Ѱ��� �߰�	

   arrParam(5) = iOpt2
   arrParam(6) = iOpt1

	' ���Ѱ��� �߰� 
	arrParam(8) = lgAuthBizAreaCd
	arrParam(9) = lgInternalCd
	arrParam(10) = lgSubInternalCd
	arrParam(11) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) = ""  Then
		frm1.txtBankCd.focus			
		Exit Function
	Else
		frm1.txtBankCd.value		= arrRet(0)
		frm1.txtBankNM.value		= arrRet(1)
		frm1.txtBankAcctNo.value	= arrRet(2)
		frm1.cboDpstFg.Value		= arrRet(3)
		frm1.cboDpstType.Value		= arrRet(4)
	End If
	
	Call cboDpstFg_Change()

End Function

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere

		Case 5
			arrParam(0) = "�ŷ���ȭ �˾�"			' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY"		 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�ŷ���ȭ"					' �����ʵ��� �� ��Ī 

			arrField(0) = "CURRENCY"					' Field��(0)
			arrField(1) = "CURRENCY_DESC"				' Field��(0)
    
			arrHeader(0) = "�ŷ���ȭ"				' Header��(0)
			arrHeader(1) = "�ŷ���ȭ��"				' Header��(0)

		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True
	
	Select Case iWhere
	Case 0, 3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function OpenPopupDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("DeptPopupDt")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDt", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = strCode						'�μ��ڵ� 
	arrParam(1) = frm1.txtStartDt.Text			'��¥(Default:������)
	arrParam(2) = "1"							'�μ�����(lgUsrIntCd)
'	If lgIntFlgMode = parent.OPMD_UMODE then
'		arrParam(3) = "T"									' �������� ���� Condition  
'	Else
'		arrParam(3) = "F"									' �������� ���� Condition  
'	End If

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	End If
	
	frm1.txtDeptCd.value = arrRet(0)
	frm1.txtDeptNm.value = arrRet(1)
	Call txtDeptCD_Change()
	frm1.txtDeptCD.focus
	
	lgBlnFlgChgValue = True
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
			Case 5		' �ŷ���ȭ 
				.txtDocCur.value = arrRet(0)
				
				If Parent.gCurrency = UCase(Trim(frm1.txtDocCur.value)) Then
					frm1.txtXchRate.Text = "1"
				Else
					Call FncCalcRate
				End If
				
				call txtDocCur_OnChange()				
				lgBlnFlgChgValue = True	
				.txtDocCur.focus
		End Select

	End With
End Function

 '=========================================================================================================
'	Name : FncCalcRate()
'	Description : lookup exchange rate 
'========================================================================================================= 
Function FncCalcRate()
    Dim strXrate
    Dim strVal
    
    Err.Clear   
	
	FncCalcRate = False
	
	If Trim(frm1.txtDocCur.value) = "" then
		frm1.txtXchRate.Text = ""
	ElseIf Trim(frm1.txtDocCur.value) = Parent.gCurrency Then
		frm1.txtXchRate.Text  = "1"
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode=" & "XRate"	        
 		strVal = strVal & "&txtLocCurr=" & Parent.gCurrency
 		strVal = strVal & "&txtToCurr=" & Trim(frm1.txtDocCur.value)
 	 	 	
		If frm1.txtStartDt.Text = "" Then		   
			strVal = strVal & "&txtAppDt=" & Trim("1900-01-01")
		Else
			strVal = strVal & "&txtAppDt=" & Trim(frm1.txtStartDt.Text) '��: ��ȸ ���� ����Ÿ 
		End If	
			
 		Call RunMyBizASP(MyBizASP, strVal) 
 	End if
 	 	
 	FncCalcRate = True
 	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function CookiePage(ByVal Kubun)

	Dim strTemp

	Select Case Kubun		
	Case "FORM_LOAD"
		strTemp = ReadCookie("BANK_CD")
		Call WriteCookie("BANK_CD", "")
		
		If strTemp = "" then Exit Function
					
		frm1.txtBankCd.value = strTemp
				
		If Err.number <> 0 Then
			Err.Clear
			Call WriteCookie("BANK_CD", "")
			Exit Function 
		End If
				
		Call MainQuery()
	
	Case JUMP_PGM_ID_BANK_REP
		Call WriteCookie("BANK_CD", frm1.txtBankCd.value)

	Case Else
		Exit Function
	End Select
End Function	

Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	'-----------------------
	'Check previous data area
	'------------------------ 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"X","X")
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

    Call InitVariables							'��: Initializes local global variables
    Call LoadInfTB19029							'��: Load table , B_numeric_format
	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)    
	Call ggoOper.LockField(Document, "N")		'��: Lock  Suitable  Field
	  
    '----------  Coding part  -------------------------------------------------------------
	Call FncSetToolBar("New")  
	Call InitComboBox
    Call SetDefaultVal
	
	Call ggoOper.FormatNumber(frm1.txtPaymDt, "31", "0", False)					'���ݳ����� 
	Call ggoOper.FormatNumber(frm1.txtPaymPeriod, "99", "0", False)				'�����ֱ� 
	Call ggoOper.FormatNumber(frm1.txtPaymCnt, "9999", "0", True)				'����Ƚ�� 
	Call ggoOper.FormatNumber(frm1.txtTotPaymCnt, "9999", "0", True)			'�Ѻ���Ƚ�� 

	Call CookiePage("FORM_LOAD") 
	frm1.txtBankCd.focus
	
	lgBlnFlgChgValue = False


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

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '-----------------------------------------------------------------------------------------------------
'	Name : SetXchRate()
'	Description : lookup exchange rate 
'--------------------------------------------------------------------------------------------------------- 
Function SetXchRate()
    Dim strXrate
    Dim strVal
    
    Err.Clear   
	
	SetXchRate = False
	
	If Trim(frm1.txtDocCur.value) = "" Then
		frm1.txtXchRate.Text = ""
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode="   & "XchRate"	        
 		strVal = strVal & "&txtLocCur=" & Parent.gCurrency
 		strVal = strVal & "&txtDocCur=" & Trim(frm1.txtDocCur.value)
 	 	
		If Trim(frm1.txtStartDt.Text) = "" Then		   
			Call DisplayMsgBox("700110","X","X","X")
			'Msgbox "�ŷ��������� �Է��ϼ���."
			Exit Function
		Else
			strVal = strVal & "&txtAppDt=" & Trim(frm1.txtStartDt.Text) '��: ��ȸ ���� ����Ÿ 
		End If	
    
 		Call RunMyBizASP(MyBizASP, strVal) 	
 	End If

 	SetXchRate = True

End Function
 '-----------------------------------------------------------------------------------------------------
'	Name : SetCnclXchRate()
'	Description : lookup exchange rate 
'--------------------------------------------------------------------------------------------------------- 
Function SetCnclXchRate()
Dim strXrate
Dim strVal
    
    Err.Clear   
	
	SetCnclXchRate = False
	
	If Trim(frm1.txtDocCur.value) = "" Then
		frm1.txtCnclXchRate.Text = ""
	Else
		strVal = BIZ_PGM_ID2 & "?txtMode="   & "CnclXchRate"	        
 		strVal = strVal & "&txtLocCur=" & Parent.gCurrency
 		strVal = strVal & "&txtDocCur=" & Trim(frm1.txtDocCur.value)
 	 	
		If Trim(frm1.txtCnclDt.Text) = "" Then		   
			Call DisplayMsgBox("700111","X","X","X")
			'Msgbox "�ؾ����ڸ� �Է��ϼ���."
			Exit Function
		Else
			strVal = strVal & "&txtAppDt=" & Trim(frm1.txtCnclDt.Text) '��: ��ȸ ���� ����Ÿ 
		End If	
    
 		Call RunMyBizASP(MyBizASP, strVal) 	
 	End If

 	SetCnclXchRate = True
 	
End Function
 '-----------------------------------------------------------------------------------------------------
'	Name : Amt's fields'  event
'	Description : 
'--------------------------------------------------------------------------------------------------------- 

Sub txtStartDt_onblur()
 
End Sub

Sub txtCnclDt_onblur()
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
    End If
End Sub

Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
    End If
End Sub

Sub txtCnclDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtCnclDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_Change()
	
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii
	Dim arrVal1, arrVal2


	If Trim(frm1.txtDeptCd.value) <> "" and Trim(frm1.txtStartDt.Text <> "") Then
	
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStartDt.Text, gDateFormat,""), "''", "S") & "))"

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
	
	Call FncCalcRate()
    lgBlnFlgChgValue = True
End Sub

Sub txtEndDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtBankRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtXchRate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPaymDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymPeriod_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtTotPaymCnt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtContractAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtContractLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclXchRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclCapitalAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclCapLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub


Sub txtCnclIntRate_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclIntAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclIntLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtCnclLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPaymLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub



Sub txtDeptCD_Change()

    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii
	If Trim(frm1.txtDeptCd.value) = "" and Trim(frm1.txtStartDt.Text = "") Then		Exit Sub

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtStartDt.Text, gDateFormat,""), "''", "S") & "))"			

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
		'----------------------------------------------------------------------------------------

     lgBlnFlgChgValue = True
End Sub

Sub txtBankCd_OnChange()
End Sub

Sub txtBankAcctNo_OnChange()
End Sub

Sub txtDocCur_OnChange()
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	END IF	    
	
End Sub

Sub Type_itemChange()
	lgBlnFlgChgValue = True
End Sub

'=====================================================
'�����ݱ��� ����� 
'=======================================================
Sub cboDpstFg_Change()

	Select Case Trim(frm1.cboDpstFg.value)
	Case "SV", ""

    	Call ggoOper.SetReqAttr(frm1.txtEndDt, "Q")			'������ 
		Call ggoOper.SetReqAttr(frm1.txtPaymDt, "Q")		'������ 
		Call ggoOper.SetReqAttr(frm1.txtPaymPeriod, "Q")	'�����ֱ� 
		Call ggoOper.SetReqAttr(frm1.txtContractAmt, "Q")	'���ݾ� 
		Call ggoOper.SetReqAttr(frm1.txtContractLocAmt, "Q")'���ݾ�(�ڱ�)
		Call ggoOper.SetReqAttr(frm1.txtPaymAmt, "Q")		'�����Աݾ� 
		Call ggoOper.SetReqAttr(frm1.txtPaymLocAmt, "Q")	'�����Աݾ�(�ڱ�)

	Case Else
		Call ggoOper.SetReqAttr(frm1.txtEndDt, "D")			'������ 
		Call ggoOper.SetReqAttr(frm1.txtPaymDt, "D")		'������ 
		Call ggoOper.SetReqAttr(frm1.txtPaymPeriod, "D")	'�����ֱ� 
		Call ggoOper.SetReqAttr(frm1.txtContractAmt, "D")	'���ݾ� 
		Call ggoOper.SetReqAttr(frm1.txtContractLocAmt, "D")'���ݾ�(�ڱ�)
		Call ggoOper.SetReqAttr(frm1.txtPaymAmt, "D")		'�����Աݾ� 
		Call ggoOper.SetReqAttr(frm1.txtPaymLocAmt, "D")	'�����Աݾ�(�ڱ�)

	End Select

End Sub

'=====================================================
'�ŷ����� ����� 
'=======================================================
Sub cboTransSts_Change()
	Select Case Trim(frm1.cboTransSts.value)
	Case "TR", ""
		frm1.txtCnclDt.Text         = ""
		frm1.txtCnclXchRate.Text    = ""
		frm1.txtCnclCapitalAmt.Text = ""
		frm1.txtCnclCapLocAmt.Text  = ""
		frm1.txtCnclIntRate.Text    = ""
		frm1.txtCnclIntAmt.Text     = ""
		frm1.txtCnclIntLocAmt.Text  = ""
		frm1.txtCnclAmt.Text        = ""
		
		Call ggoOper.SetReqAttr(frm1.txtCnclDt, "Q")			'�ؾ��� 
		Call ggoOper.SetReqAttr(frm1.txtCnclXchRate, "Q")		'�ؾ�ȯ�� 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapitalAmt, "Q")	'�ؾ���� 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapLocAmt, "Q")		'�ؾ����(�ڱ�)
		Call ggoOper.SetReqAttr(frm1.txtCnclIntRate, "Q")		'�ؾ����� 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntAmt, "Q")		'�ؾ����� 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntLocAmt, "Q")		'�ؾ�����(�ڱ�)

	Case Else
		Call ggoOper.SetReqAttr(frm1.txtCnclDt, "D")			'�ؾ��� 
		Call ggoOper.SetReqAttr(frm1.txtCnclXchRate, "D")		'�ؾ�ȯ�� 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapitalAmt, "D")	'�ؾ���� 
		Call ggoOper.SetReqAttr(frm1.txtCnclCapLocAmt, "D")		'�ؾ����(�ڱ�)
		Call ggoOper.SetReqAttr(frm1.txtCnclIntRate, "D")		'�ؾ����� 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntAmt, "D")		'�ؾ����� 
		Call ggoOper.SetReqAttr(frm1.txtCnclIntLocAmt, "D")		'�ؾ�����(�ڱ�)

	End Select
	
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")				'�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
      '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field    
    Call SetDefaultVal
    Call InitVariables	
    
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not ChkField(Document, "1") Then		'��: This function check indispensable field
       Exit Function
    End If
    
    Call FncSetToolBar("New")
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery									'��: Query db data
       
    FncQuery = True									'��: Processing is OK
        
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False      '��: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X") '�� �ٲ�κ� 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    
    Call ggoOper.ClearField(Document, "A")  '��: Clear Condition/Contents(All) Field    
    
    Call InitVariables						'��: Initializes local global variables
    
    
    call txtDocCur_OnChange()
	Call cboTransSts_Change
	
	Call SetDefaultVal
		
	Call FncSetToolBar("New")

    lgBlnFlgChgValue = False
	frm1.txtBankCd.focus 
    
    FncNew = True							'��: Processing is OK

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
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002","X","X","X")  '�� �ٲ�κ� 
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")  '�� �ٲ�κ� 
    If IntRetCD = vbNo Then
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    
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
		IntRetCD = DisplayMsgBox("900001","X","X","X")  '�� �ٲ�κ� 
		Exit Function
	End If
	    
    '-----------------------
    'Check content area
    '-----------------------
    If Not ChkField(Document, "1")     then                              '��: Check contents area
       Exit Function
    End If

    If Not ChkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    
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
     On Error Resume Next                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '�� �ٲ�κ� 
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
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtBankCd=" & EnCoding(Trim(frm1.txtBankCd.value))
    strVal = strVal & "&txtBankAcctNo=" & EnCoding(Trim(frm1.txtBankAcctNo.value))	'��: ���� ���� ����Ÿ 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncNew()
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1

		ggoOper.FormatFieldByObjectOfCur .txtAmt,			.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,		.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtContractAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCnclCapitalAmt,.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCnclIntAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCnclAmt,       .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
Dim strVal
    
    Err.Clear																		'��: Protect system from crashing
    
    DbQuery = False																	 '��: Processing is NG
    
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtBankCd=" & Trim(frm1.txtBankCd.value)
	strVal = strVal & "&txtBankAcctNo=" & Trim(frm1.txtBankAcctNo.value)			'��: ��ȸ ���� ����Ÿ 

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

Function DbQueryOk()							'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")		'��: This function lock the suitable field

    Call cboTransSts_Change
   
	Call FncSetToolBar("Query")
    call txtDocCur_OnChange()
    	
    lgIntFlgMode = parent.OPMD_UMODE					'��: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
	frm1.txtBankCd.focus 
	Set gActiveElement = document.activeElement 
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
Dim strVal

    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG

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

    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()			'��: ���� ������ ���� ���� 
    Call InitVariables
	Call MainQuery()
End Function

'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1110100000001111")
	Case "QUERY"
		Call SetToolbar("1111100000011111")
	End Select
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenRefBankAcctNo(frm1.hTemp.value,1)">�����������</A>
					</TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10   tag="12XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefBankAcctNo(frm1.txtBankCd.Value, 2)">&nbsp;
														 <INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=20 MAXLENGTH=30  STYLE="TEXT-ALIGN: left" tag="24X" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>���¹�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtBankAcctNo" NAME="txtBankAcctNo" SIZE=18 MAXLENGTH=30  tag="12XXXU" ALT="���¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefBankAcctNo(frm1.txtBankAcctNo.Value, 3)"></TD>
								</TR>									
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>									
								<TR>
									<TD CLASS=TD5 NOWRAP>�����ݱ���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboDpstFg" NAME="cboDpstFg" ALT="�����ݱ���" STYLE="WIDTH: 132px" tag="14X" OnClick ="vbscript:Type_itemChange()" OnChange="vbscript:Call cboDpstFg_Change()" ><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>����������</TD>
									<TD CLASS=TD6 NOWRAP><SELECT ID="cboDpstType" NAME="cboDpstType" ALT="����������" STYLE="WIDTH: 132px" tag="14X" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
								<TD CLASS=TD5 NOWRAP>�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptCD" NAME="txtDeptCD" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" tag="23X" ONCLICK="vbscript:Call OpenPopupDept(frm1.txtDeptCD.Value, 1)">&nbsp;<INPUT CLASS="clstxt" TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 MAXLENGTH=40 STYLE="TEXT-ALIGN: left" tag="24X" ALT="�μ�"></TD>
								<TD CLASS=TD5 NOWRAP>�ŷ�������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpStartDt name=txtStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�ŷ�������" tag="22X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�����</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboTransSts" NAME="cboTransSts" ALT="�ŷ�����" STYLE="WIDTH: 132px" tag="22X" OnClick ="vbscript:Type_itemChange()" OnChange="vbscript:Call cboTransSts_Change()"><!--OPTION VALUE="" selected></OPTION--></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SELECT ID="cboBankAcctFg" NAME="cboBankAcctFg" ALT="��������" STYLE="WIDTH: 132px" tag="2XX" OnClick ="vbscript:Type_itemChange()"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpBankRate name=txtBankRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="����" tag="21X5" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;%</TD>							
								<TD CLASS=TD5 NOWRAP>���Ի���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDpstNm" SIZE="35" MAXLENGTH="40" tag="21X" ALT="���Ի���"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ȭ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT ID="txtDocCur" NAME="txtDocCur" SIZE=15 MAXLENGTH=3  tag="22XXXU" ALT="��ȭ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 5)"></TD>
								<TD CLASS=TD5 NOWRAP>ȯ��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpXchRate name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="ȯ��" tag="21X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ܾ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpAmt name=txtAmt title=FPDOUBLESINGLE ALT="�ܾ�" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�ܾ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpLocAmt name=txtLocAmt title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR></TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpEndDt name=txtEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="21X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPaymDt name=txtPaymDt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="������" tag="21X" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;��</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����ֱ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPaymPeriod name=txtPaymPeriod style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="�����ֱ�" tag="21X" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;����</TD>
								<TD CLASS=TD5 NOWRAP>����Ƚ��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpPaymCnt name=txtPaymCnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="����Ƚ��" tag="24X" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;/&nbsp;
											  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpTotPaymCnt name=txtTotPaymCnt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 50px" title=FPDOUBLESINGLE ALT="�Ѻ���Ƚ��" tag="24X" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����Ծ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpPaymAmt name=txtPaymAmt title=FPDOUBLESINGLE ALT="�����Ծ�" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�����Ծ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpPaymLocAmt name=txtPaymLocAmt title=FPDOUBLESINGLE ALT="�����Ծ�(�ڱ�)" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>���ݾ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpContractAmt name=txtContractAmt title=FPDOUBLESINGLE ALT="���ݾ�" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>���ݾ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpContractLocAmt name=txtContractLocAmt title=FPDOUBLESINGLE ALT="���ݾ�(�ڱ�)" tag="21X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR></TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ؾ�����</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpCnclDt name=txtCnclDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�ؾ�����" tag="21X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�ؾ��ȯ��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclXchRate name=txtCnclXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="�ؾ��ȯ��" tag="24X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ؾ��������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclIntRate name=txtCnclIntRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="�ؾ��������" tag="24X5Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ؾ�ÿ���</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclCapitalAmt name=txtCnclCapitalAmt title=FPDOUBLESINGLE ALT="�ؾ�ÿ���" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�ؾ�ÿ���(�ڱ�)</TD>	
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclCapLocAmt name=txtCnclCapLocAmt title=FPDOUBLESINGLE ALT="�ؾ�ÿ���(�ڱ�)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>�ؾ������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclIntAmt name=txtCnclIntAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="�ؾ������" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�ؾ������(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpCnclIntLocAmt name=txtCnclIntLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 132px" title=FPDOUBLESINGLE ALT="�ؾ������(�ڱ�)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ؾ�ݾ�</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclAmt name=txtCnclAmt title=FPDOUBLESINGLE ALT="�ؾ�ݾ�" tag="24X2Z"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�ؾ�ݾ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 id=fpCnclLocAmt name=txtCnclLocAmt title=FPDOUBLESINGLE ALT="�ؾ�ݾ�(�ڱ�)" tag="24X2Z" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP	COLSPAN=3><INPUT TYPE=TEXT NAME="txtDpstDesc" SIZE="70" MAXLENGTH="128" tag="21X" ALT="����������"></TD>
							</TR>
<!--							<% Call SubFillRemBodyTD5656(1) %> -->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_BANK_REP)">�����������</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="2"  Tabindex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"	tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="horgchangeid"		tag="2"  Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="hTemp"				tag="2"  Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
</BODY>
</HTML>

