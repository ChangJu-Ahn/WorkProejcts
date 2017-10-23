<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103ma1
'*  4. Program Name         : �����ڻ� MASTER ���� 
'*  5. Program Desc         : �����ڻ꺰 MASTER�� ����,��ȸ 
'*  6. Comproxy List        : +As0041ManageSvr
'                             +As0049LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2001/06/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : KIM HEE JUNG
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->						<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!--==========================================  1.1.1 Style Sheet  ===========================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--==========================================  1.1.2 ���� Include   =========================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                             '��: indicates that All variables must be declared in advance 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

Const BIZ_PGM_ID = "a7103mb1.asp"     											 '��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgBlnFlgConChg				'��: Condition ���� Flag
'@Dim lgBlnFlgChgValue				'��: Variable is for Dirty flag
'@Dim lgIntGrpCount				'��: Group View Size�� ������ ���� 
'@Dim lgIntFlgMode					'��: Variable is for Operation Status

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'-------------------  ���� Global ������ ����  ----------------------------------------------------------- 


'+++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        

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

    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '----------  Coding part  -----------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate = ""
    lgLlcGivenDt = ""
End Sub


'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=============================================================================================== 
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
    	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
End Sub

Sub InitComboBox()

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
	Dim intMaxRow, intLoopCnt
	Dim ArrTmpF0, ArrTmpF1
	
	On error resume next

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A2004", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ArrTmpF0 = split(lgF0,chr(11))
	ArrTmpF1 = split(lgF1,chr(11))
	
	intMaxRow = ubound(ArrTmpF0)
	
	If intRetCD1 <> False Then
		for intLoopCnt = 0 to intMaxRow - 1
			Call SetCombo(frm1.cboTaxDeprSts, ArrTmpF0(intLoopCnt), ArrTmpF1(intLoopCnt))
			Call SetCombo(frm1.cboCasDeprSts, ArrTmpF0(intLoopCnt), ArrTmpF1(intLoopCnt))
		next
	End If		

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A2005", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ArrTmpF0 = split(lgF0,chr(11))
	ArrTmpF1 = split(lgF1,chr(11))
	
	intMaxRow = ubound(ArrTmpF0)
	
	If intRetCD1 <> False Then
		for intLoopCnt = 0 to intMaxRow - 1
			Call SetCombo(frm1.cboAcqFg, ArrTmpF0(intLoopCnt), ArrTmpF1(intLoopCnt))
		next
	End If		
	'------ Developer Coding part (End )   --------------------------------------------------------------
end sub

'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'********************************************************************************************************* 


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
 '------------------------------------------  OpenMasterRef()  -------------------------------------------------
'	Name : OpenMasterRef()
'	Description : Asset Master Condition PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function	
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gStrRequestMenuID , Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPoRef(arrRet)
	End If	

	frm1.txtCondAsstNo.focus			
End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)
       
	frm1.txtCondAsstNo.value     = strRet(0)
	frm1.txtcondAsstNm.value	 = strRet(1)
		
End Sub


'----------------------------------------  OpenAcctCd()  -------------------------------------------------
'	Name : OpenAcctCd()
'	Description : Data Account Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True Or UCase(frm1.txtAcctCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "{{�����ڵ��˾�}}"			' �˾� ��Ī 
	arrParam(1) = "a_asset_acct, a_acct"		' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtAcctCd.value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "a_asset_acct.acct_cd = a_acct.acct_cd"		' Where Condition
	arrParam(5) = "{{�����ڵ�}}"				' �����ʵ��� �� ��Ī 
	
    arrField(0) = "a_asset_acct.acct_cd"		' Field��(0)
    arrField(1) = "a_acct.acct_sh_nm"			' Field��(1)
    
    arrHeader(0) = "{{�����ڵ�}}"				' Header��(0)
    arrHeader(1) = "{{������}}"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = 3
		Call SetReturnVal(arrRet,field_fg)
	End If	
End Function


'----------------------------------------  OpenMgmtId()  -------------------------------------------------
'	Name : OpenMgmtId()
'	Description : ������Id PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMgmtId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True Or UCase(txtMgmtUserId.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "{{����ڵ��˾�}}"			' �˾� ��Ī 
	arrParam(1) = ""							' TABLE ��Ī 
	arrParam(2) = ""							' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "{{����ڵ�}}"				' �����ʵ��� �� ��Ī 
	
    arrField(0) = ""							' Field��(0)
    arrField(1) = ""							' Field��(1)
    
    arrHeader(0) = "{{����ڵ�}}"				' Header��(0)
    arrHeader(1) = "{{�����}}"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = 4
		Call SetReturnVal(arrRet,field_fg)
	End If	
End Function


'------------------------------------------ OpenCurrency() -----------------------------------------------
'	Name : OpenCurrency()
'	Description : Data Currency Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg
    
	If IsOpenPop = True Or UCase(frm1.txtDocCur.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "{{�ŷ���ȭ �˾�}}"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "{{�ŷ���ȭ}}"
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "{{�ŷ���ȭ}}"		
    arrHeader(1) = "{{�ŷ���ȭ��}}"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = 5
		Call SetReturnVal(arrRet,field_fg)
	End If	
End Function


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(ByVal arrRet, ByVal field_fg)
	
	Select case field_fg
		case 3	'OpenAcctCd
			frm1.txtAcctCd.Value		= arrRet(0)
			frm1.txtAcctNm.Value		= arrRet(1)
		case 4	'OpenMgmtId
			frm1.txtMgmtUserId.Value	= arrRet(0)
			frm1.txtMgmtUserNm.Value	= arrRet(1)
			lgBlnFlgChgValue = True
		case 5	'OpenCurrency
			frm1.txtDocCur.Value		= arrRet(0)
	End select	

End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function funChkAmt()
    dim Ltaxbalamt, Lcasbalamt
    
    funChkAmt = False
    
	IF frm1.txtTaxBalAmt.value <> "" Then
		Ltaxbalamt = UNICDbl(frm1.txtTaxBalAmt.value)
	    if Ltaxbalamt < 0 then
			Call DisplayMsgBox("AS0049", "X", "X", "X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
	'       call MsgBox("�̻� �ܾ��� 0���� �۾Ƽ��� �ȵ˴ϴ�..",vbInformation)
	        frm1.txtTaxBalAmt.focus
	        Set gActiveElement = document.activeElement
	        Exit Function
	    End If
	End If    

	If frm1.txtCasBalAmt.value <> "" Then
	    Lcasbalamt = UNICDbl(frm1.txtCasBalAmt.value)
	    if Lcasbalamt < 0 then
			Call DisplayMsgBox("AS0049", "X", "X", "X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
	'       call MsgBox("�̻� �ܾ��� 0���� �۾Ƽ��� �ȵ˴ϴ�..",vbInformation)
	        frm1.txtCasBalAmt.focus
	        Set gActiveElement = document.activeElement
	        Exit Function
	    End if
	End If
    funChkAmt = True
     
End function


'##########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'#########################################################################################################

'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

'    Call GetGlobalVar
'    Call ClassLoad																	'��: Load Common DLL
    Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call AppendNumberPlace("7","3","0")
    Call AppendNumberPlace("6","2","0")
    Call AppendNumberRange("0","1","60")
    
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.FormatDate(frm1.txtDeprFrDt, gDateFormat, 2)

    Call ggoOper.FormatDate(frm1.txtTaxDeprEnd, gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtCasDeprEnd, gDateFormat, 2)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    Call InitVariables																'��: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("110000000000111")
    Call SetDefaultVal
    Call InitComboBox
	
	frm1.txtCondAsstNo.focus	
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


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'***************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 


'***********************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 


'-----------------------------  Coding part  ------------------------------------------------------------- 
Sub txtCondAsstNo_OnChange()
	If Trim(frm1.txtCondAsstNo.value) = "" Then
		frm1.txtCondAsstNm.value = ""
	End If
End Sub

'Sub txtAcqQty_OnChange()
'	frm1.txtInvQty.value = frm1.txtAcqQty.value
'End Sub

Sub txtCasDurYrs_OnChange()
	lgBlnFlgChgValue = True
End Sub


Sub txtDeprFrdt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDeprFrDt.Action = 7
	End If
End Sub

'=======================================================================================================
'   Event Name : txtTaxDeprEnd_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtTaxDeprEnd_DblClick(Button)
    If Button = 1 Then
        frm1.txtTaxDeprEnd.Action = 7
    End If
End Sub


'=======================================================================================================
'   Event Name : txtCasDeprEnd_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtCasDeprEnd_DblClick(Button)
    If Button = 1 Then
        frm1.txtCasDeprEnd.Action = 7
    End If
End Sub

Sub cboTaxDeprSts_OnChange()

'	frm1.txtTaxDeprEnd.value = ""
	lgBlnFlgChgValue = True
'	If frm1.cboTaxDeprSts.value = "02" Then
'		ReleaseTag(frm1.txtTaxDeprEnd)
'	Else
'		ProtectTag(frm1.txtTaxDeprEnd)
'	End If
End Sub

Sub cboCasDeprSts_OnChange()
	lgBlnFlgChgValue = True
'	frm1.txtCasDeprEnd.value = ""
'
'	If frm1.cboCasDeprSts.value = "02" Then
'		ReleaseTag(frm1.txtCasDeprEnd)
'	Else
'		ProtectTag(frm1.txtCasDeprEnd)
'	End If
End Sub

Sub txtTaxDeprEnd_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCasDeprEnd_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxDeprTotAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCasDeprTotAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxCptTotAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCasCptTotAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxBalAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCasBalAmt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxDurYrs_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCasDurYrs_Change()
	lgBlnFlgChgValue = True
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

'********************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
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
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
'		IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field

'    ggoSpread.Source = frm1.vspdData
'	ggospread.ClearSpreadData		'Buffer Clear

    Call InitVariables															'��: Initializes local global variables
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
     
    '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'��: Query db data

    FncQuery = True																'��: Processing is OK
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
'		IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��Է��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                      '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call InitVariables															'��: Initializes local global variables
	Call SetToolBar("110000000000111")
    Call SetDefaultVal
    
    FncNew = True																'��: Processing is OK

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
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
'        Call MsgBox("��ȸ���Ŀ� ������ �� �ֽ��ϴ�.", vbInformation)
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function

'----------------------------------------------------------
'  Functions before fncSave
'----------------------------------------------------------
Function FncChkAmt()
	'-----------------------------------------------------------	
	' ���ݾ�,�ں������⴩��ݾ�,�󰢴���ݾ�,�̻��ܾ� Check
	'-----------------------------------------------------------
	Dim varAcqAmt, varDeprTotamt,varCptTotAmt,varBalAmt
	Dim strRegDt,strFiscDt
	
	FncChkAmt = False
	
	strRegDt	= UniConvDateToYYYYMMDD(frm1.RegDateTime1.Text, gDateFormat, "")
    strFiscDt   = UniConvDateToYYYYMMDD(parent.gFiscStart, parent.gAPDateFormat,"")  ' ��� ���ۿ� 
    
	varAcqAmt	  = UNICDbl(frm1.txtAcqLocAmt.value) 
	varDeprTotAmt = UNICDbl(frm1.txtTaxDeprTotAmt.value) 
	varCptTotAmt  = UNICDbl(frm1.txtTaxCptTotAmt.value) 
	varBalAmt	  = UNICDbl(frm1.txtTaxBalAmt.value) 

	'-------------------------------------------------------------
	' ���������� ���Ŀ� ����� ��� ���⸻ ����Ÿ�� ����� �Ѵ�.
	'-------------------------------------------------------------

	if strRegDt >= strFiscDt then   
		if varDeprTotAmt > 0 or varCptTotAmt > 0 or varBalAmt >0  then
			Call DisplayMsgBox("117428", "X", "X", "X")  '''������Ŀ� ����� �ڻ��� ���⸻�󰢳����� �Է��� �� �����ϴ�.
			exit function
		end if	
	else		
		If (varAcqAmt + varCptTotAmt - varDeprTotAmt) <> varBalAmt then
			Call DisplayMsgBox("117424", "X", "X", "X")                               
			Exit Function
		End if				
	end if

	
	varAcqAmt	  = UNICDbl(frm1.txtAcqLocAmt.value) 
	varDeprTotAmt = UNICDbl(frm1.txtCasDeprTotAmt.value) 
	varCptTotAmt  = UNICDbl(frm1.txtCasCptTotAmt.value) 
	varBalAmt	  = UNICDbl(frm1.txtCasBalAmt.value) 	

	'-------------------------------------------------------------
	' ���������� ���Ŀ� ����� ��� ���⸻ ����Ÿ�� ����� �Ѵ�.
	'-------------------------------------------------------------
	if strRegDt >= strFiscDt then   
		if varDeprTotAmt > 0 or varCptTotAmt > 0 or varBalAmt >0 then
			Call DisplayMsgBox("117428", "X", "X", "X")  '''������Ŀ� ����� �ڻ��� ���⸻�󰢳����� �Է��� �� �����ϴ�.
			exit function
		end if	
	else
		If (varAcqAmt + varCptTotAmt -varDeprTotAmt) <> varBalAmt then
			Call DisplayMsgBox("117424", "X", "X", "X")                              
			Exit Function
		End if			
	end if
	
	FncChkAmt = True
	
End Function

Function fncChkDeprSts()
	fncChkDeprSts = False

	if frm1.hTaxDeprSts.value <> "03" then '�󰢴���� �ڻ꿡 ���� ����� ���� �� 
		if frm1.cboTaxDeprSts.value = "03" then
			Call DisplayMsgBox("117423", "X", "X", "X")
			frm1.cboTaxDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if		
	end if
	
	if frm1.hCasDeprSts.value <> "03" then '�󰢴���� �ڻ꿡 ���� ����� ���� �� 
		if frm1.cboCasDeprSts.value = "03" then
			Call DisplayMsgBox("117423", "X", "X", "X")
			frm1.cboCasDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if
	end if	
	
	if frm1.cboTaxDeprSts.value = "02" then  '�󰢿Ϸ��� �� 
		if frm1.fpDateTime1.text = "" then ' �󰢿Ϸ����� �Է����� ���� ��� 
			Call DisplayMsgBox("117422", "X", "X", "X")
			Exit Function
		end if
	end if
	if frm1.cboCasDeprSts.value = "02" then  '�󰢿Ϸ��� �� 
		if frm1.toDateTime1.text = "" then ' �󰢿Ϸ����� �Է����� ���� ��� 
			Call DisplayMsgBox("117422", "X", "X", "X")
			Exit Function
		end if
	end if	

	fncChkDeprSts = True
	
End Function

Function FncChkBalAmt()
	Dim strRemRate
	Dim varRemRate   ''������ 
	Dim varAcqAmt,varCptTotAmtTax,varBalAmtTax
	Dim varCptTotAmtCas,varBalAmtCas
	Dim varRemAmtTax,varRemAmtCas
	Dim varInvQty
	Dim strRegDt,strFiscDt
			
	FncChkBalAmt = False
	strRegDt	= UniConvDateToYYYYMMDD(frm1.RegDateTime1.Text, gDateFormat, "")
    strFiscDt   = UniConvDateToYYYYMMDD(parent.gFiscStart, parent.gAPDateFormat,"")  ' ��� ���� �� 
	
	
	'-------------------------------------------------------------
	' ���������� ���Ŀ� ����� ��� ���⸻ ����Ÿ�� ����� �Ѵ�.
	'-------------------------------------------------------------	
	if strRegDt >= strFiscDt then
		if frm1.cboTaxDeprSts.value = "02" then    '��������: �󰢿Ϸ� 
			Call DisplayMsgBox("117423", "X", "X", "X")  '''�󰢻��¸� Ȯ���Ͻʽÿ�.
			frm1.cboTaxDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if
		if frm1.cboCasDeprSts.value = "02" then    '���ȸ�����: �󰢿Ϸ� 
			Call DisplayMsgBox("117423", "X", "X", "X")  '''�󰢻��¸� Ȯ���Ͻʽÿ�.
			frm1.cboCasDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if			
	
	else
		strRemRate = Trim(frm1.htxtRemRate.value)   '������(����:0%,����: 5%)
		varInvQty  = CInt(frm1.txtInvQty.value) 
	
		if isNull(strRemRate) then
			varRemRate = 0
		else
		    If isnumeric(strRemRate) Then
    	       varRemRate = CDbl(strRemRate)
    	    Else   
    	       varRemRate = 0
    	    End If    
		end if
	
		varAcqAmt		 = UNICDbl(frm1.txtAcqLocAmt.value) 
			
		varCptTotAmtTax  = UNICDbl(frm1.txtTaxCptTotAmt.value) 
		varBalAmtTax     = UNICDbl(frm1.txtTaxBalAmt.value)  
		''''varRemAmtTax     = ((varAcqAmt + varCptTotAmtTax) * varRemRate * 0.01 )
		varRemAmtTax     = ((varAcqAmt + varCptTotAmtTax) * 5 * 0.01 )
		
		if varInvQty * 1000 < varRemAmtTax then
			varRemAmtTax = varInvQty * 1000
		end if
	
		varCptTotAmtCas  = UNICDbl(frm1.txtCasCptTotAmt.value) 
		varBalAmtCas	 = UNICDbl(frm1.txtCasBalAmt.value) 	
		'''''varRemAmtCas	 = ((varAcqAmt + varCptTotAmtcas) * varRemRate * 0.01 )
		varRemAmtCas	 = ((varAcqAmt + varCptTotAmtcas) * 5 * 0.01 )
	
		if varInvQty * 1000 < varRemAmtCas then
			varRemAmtCas = varInvQty * 1000
		end if
	
		'1. ��������				
		if frm1.cboTaxDeprSts.value = "02" then 
		'**************************************************************
		'  �󰢻���-�󰢿Ϸ� ���� ��, �̻󰢱ݾ׵� �󰢿Ϸ�Ǵ� �ݾ����� üũ: 
		'**************************************************************	
		'''(��氡+�ں�������ݾ�) * ������ * 0.01 < �̻��ܾ�?
			IF varRemAmtTax < varBalAmtTax then
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if	
		else
		'************************************************************
		' �󰢿Ϸ� �ƴϸ鼭 �̻󰢱ݾ��� �󰢿Ϸ�Ǵ� �ݾ��̸� Error
		'************************************************************
			IF varRemAmtTax >= varBalAmtTax then   '�̸��ݾ� >= �̻󰢱ݾ�? ��,�󰢿Ϸ�Ǵ� �ݾ� 
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if			
		end if

		'2. ���ȸ����� 
		IF frm1.cboCasDeprSts.value = "02" THEN 		
		'**************************************************************
		'  �󰢿Ϸ� ���� ��, �̻󰢱ݾ׵� �󰢿Ϸ�Ǵ� �ݾ����� üũ: 
		'**************************************************************		
			IF varRemAmtCas < varBalAmtCas then
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if					
		ELSE
		'************************************************************
		' �󰢿Ϸ� �ƴϸ鼭 �̻󰢱ݾ��� �󰢿Ϸ�Ǵ� �ݾ��̸� Error
		'***********************************************************	
			IF varRemAmtCas >= varBalAmtCas then   '�̸��ݾ� >= �̻󰢱ݾ�? ��,�󰢿Ϸ�Ǵ� �ݾ� 
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if	
		END IF
	end if
	
	FncChkBalAmt = True
	
End Function

Function fncChkDeprStarYymm
	fncChkDeprStarYymm = False
	Dim strRegDt,strDeprFrDt,strEndDt
	Dim strYear
	Dim strMonth
	Dim strDay	
	
	strRegDt	= UniConvDateToYYYYMM(frm1.RegDateTime1.Text, gDateFormat, "")
	Call ExtractDateFrom(frm1.DeprDateTime1.Text,frm1.DeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    strDeprFrDt = strYear & strMonth
	if strRegDt > strDeprFrDt then
		Call DisplayMsgBox("117426", "X", "X", "X")   ''�󰢽��۳���� ��������� ũ�ų� ���ƾ� �մϴ�.
		Exit Function
	end if
	strEndDt = UniConvDateToYYYYMM(frm1.fpDateTime1.Text, gDateFormat, "")
		
	if strEnddt <> "" then
		if strRegDt > strEndDt then
			Call DisplayMsgBox("117427", "X", "X", "X")   ''�󰢿Ϸ����� ��������� Ŀ�� �մϴ�.
			Exit Function
		end if	
	end if
		
	strEndDt = ""
	strEndDt = UniConvDateToYYYYMM(frm1.toDateTime1.Text, gDateFormat, "")
	if strEnddt <> "" then	
		if strRegDt > strEndDt then
			Call DisplayMsgBox("117427", "X", "X", "X")   ''�󰢿Ϸ����� ��������� Ŀ�� �մϴ�.
			Exit Function
		end if
	end if
		
	fncChkDeprStarYymm = True
	
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD    
    	
    if Not funChkAmt then
       exit function
    end if
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                              '��: Protect system from crashing    
	'-----------------------
    ' Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                   '��: No data changed!!
        Exit Function
    End If    
	'-----------------------
    ' Check content area
    '-----------------------
    If Not chkField(Document, "2") Then										'��: Check contents area
       Exit Function
    End If

	if IsNull(frm1.txtTaxDeprEnd.text) then
		frm1.txtTaxDeprEnd.text = ""
	end if
	
	if IsNull(frm1.txtCasDeprEnd.text) then
		frm1.txtCasDeprEnd.text = ""
	end if	
	
	if IsNull(frm1.txtDeprFrdt.text) then
		frm1.txtDeprFrdt.text = ""
	end if		

	'********************************************************************
	' FncChkAmt(): ��氡+�ں�������-�󰢴���ݾ� = �̻󰢱ݾ� ?
	'********************************************************************
	If FncChkAmt = False Then
		exit function
	End if
	
	'********************************************************************
	' FncChkBalAmt(): �̻󰢱ݾװ� �󰢿ϷῩ�ο� �󰢿Ϸ��� Check
	'********************************************************************
	If FncChkBalAmt = False Then
		exit function
	End if
			
	'********************************************************************
	' FncChkDeprSts(): �󰢻��¿� �󰢿Ϸ��� Check
	'********************************************************************
	If fncChkDeprSts = False Then
		Exit function
	End if
	
	'********************************************************************
	' FncChkDeprSts(): �󰢿ϷῩ�ο� �󰢿Ϸ��� Check
	'***************************************************
	if fncChkDeprStarYymm = False then
		Exit Function
	end if
		
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
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
'		IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												'��: Indicates that current mode is Crate mode
    
     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                                      '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'��: This function lock the suitable field
    
    frm1.txtAssetCd2.value = ""
    frm1.txtAssetNm2.value = ""
    frm1.txtAssetCd2.focus
    Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                 '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        'Call MsgBox("��ȸ���Ŀ� �����˴ϴ�.", vbInformation)
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")                                 '��: 
		'Call MsgBox("���� ����Ÿ�� �����ϴ�..", vbInformation)
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
    strVal = strVal & "&txtAssetCd1=" & lgPrevNo							'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")                                '��: 
    End If
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���°� 
    strVal = strVal & "&txtAssetCd1=" & lgNextNo							'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
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
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtAssetCd1=" & Trim(frm1.txtAssetCd1.value)		'��: ���� ���� ����Ÿ 

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


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG

	Call LayerShowHide(1)
	
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCondAsstNo=" & Trim(frm1.txtCondAsstNo.value)	'��: ��ȸ ���� ����Ÿ 

	' ���Ѱ��� �߰�
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' �����
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ�
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

	Call SetToolBar("110010000001111")	'111010000001111

	if frm1.cboTaxDeprSts.value = "03" Then'����� ��� 

		ggoOper.SetReqAttr frm1.OBJECT1, "Q"
		ggoOper.SetReqAttr frm1.fpDoubleSingle1, "Q"
		ggoOper.SetReqAttr frm1.fpDoubleSingle2, "Q"
		ggoOper.SetReqAttr frm1.fpDoubleSingle3, "Q"
				
		ggoOper.SetReqAttr frm1.cboTaxDeprSts, "Q"
		ggoOper.SetReqAttr frm1.fpDateTime1,   "Q"
		'ggoOper.SetReqAttr frm1.DeprDateTime1, "Q"
	else
		ggoOper.SetReqAttr frm1.OBJECT1, "N"
		ggoOper.SetReqAttr frm1.fpDoubleSingle1, "N"
		ggoOper.SetReqAttr frm1.fpDoubleSingle2, "N"
		ggoOper.SetReqAttr frm1.fpDoubleSingle3, "N"
				
		ggoOper.SetReqAttr frm1.cboTaxDeprSts, "N"
		ggoOper.SetReqAttr frm1.fpDateTime1,   "Q"
		
		'ggoOper.SetReqAttr frm1.DeprDateTime1, "D"	
	end if	

	if frm1.cboCasDeprSts.value = "03" Then'����� ��� 
		ggoOper.SetReqAttr frm1.OBJECT2, "Q"
		ggoOper.SetReqAttr frm1.fpDoubleSingle4, "Q"
		ggoOper.SetReqAttr frm1.fpDoubleSingle5, "Q"
		ggoOper.SetReqAttr frm1.fpDoubleSingle6, "Q"
		
		ggoOper.SetReqAttr frm1.cboCasDeprSts,   "Q"
		ggoOper.SetReqAttr frm1.toDateTime1,     "Q"
		'ggoOper.SetReqAttr frm1.DeprDateTime1,   "Q"
	else
		ggoOper.SetReqAttr frm1.OBJECT2, "N"
		ggoOper.SetReqAttr frm1.fpDoubleSingle4, "N"
		ggoOper.SetReqAttr frm1.fpDoubleSingle5, "N"
		ggoOper.SetReqAttr frm1.fpDoubleSingle6, "N"
		
		ggoOper.SetReqAttr frm1.cboCasDeprSts,   "N"
		ggoOper.SetReqAttr frm1.toDateTime1,     "Q"
		'ggoOper.SetReqAttr frm1.DeprDateTime1,   "D"	
	end if	

	lgBlnFlgChgValue = False
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
    Dim strVal
	Dim varDeprdt,varTaxDt,varCasDt
	Dim strYear
	Dim strMonth
	Dim strDay	
	
    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG
  
	Call LayerShowHide(1)
	With frm1
		.txtMode.value = parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
	
	
		Call ExtractDateFrom(frm1.DeprDateTime1.Text,frm1.DeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
		varDeprdt = strYear & strMonth
		
		varTaxDt  = UniConvDateToYYYYMM(frm1.fpDateTime1.Text, gDateFormat, "")
		varCasDt  = UniConvDateToYYYYMM(frm1.toDateTime1.Text, gDateFormat, "")
		
		frm1.htxtDeprYymm.value   = varDeprDt
		frm1.htxtTaxDeprEnd.value = varTaxDt	
		frm1.htxtCasDeprEnd.value = varCasDt

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

Function DbSaveOk()															'��: ���� ������ ���� ���� 

    'frm1.txtAssetCd1.value = frm1.txtAssetCd2.value    'Conditon�� �ڻ��ڵ� 
    'frm1.txtAssetNm1.value = frm1.txtAssetNm2.value 
     
    Call InitVariables
    
    Call dbQuery()

End Function

Sub txtDeprFrdt_Change()
       lgBlnFlgChgValue = true
End Sub
'***************************************************************************************************************

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>

	<!-- �Ǳ���  -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
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

	<!-- ��������  -->
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
									<TD CLASS="TD5" NOWRAP>{{�ڻ��ȣ}}</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE="Text" NAME="txtCondAsstNo" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="{{�ڻ��ȣ}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef()"> <INPUT TYPE="Text" NAME="txtCondAsstNm" SIZE=30 MAXLENGTH=30 tag="14" ALT="{{�ڻ��}}"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD WIDTH=50%>
									<FIELDSET STYLE="HEIGHT: 100%"><LEGEND>{{�⺻����}}</LEGEND>
									<TABLE CLASS="TB2" CELLSPACING=0 STYLE="HEIGHT: 96%">
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�ڻ��}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNm" SIZE=44 MAXLENGTH=40 TAG="2x" ALT="{{�ڻ��}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{������ȣ}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtRefNo" SIZE=30 MAXLENGTH=30 TAG="2x" ALT="{{������ȣ}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�����μ�}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCd" SIZE=15 MAXLENGTH=10 tag="24" ALT="{{�����μ��ڵ�}}"> <INPUT TYPE="Text" NAME="txtDeptNm" SIZE=27 MAXLENGTH=40 tag="24" ALT="{{�����μ���}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�������}}</TD>
											<TD CLASS="TD6" NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=RegDateTime1 name=txtRegDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="{{�������}}"></TD> </OBJECT>');</SCRIPT>											    
											</TD>										
										</TR>																	
<%	If gIsShowLocal <> "N" Then	%>										
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�ŷ���ȭ}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDocCur" SIZE=10 MAXLENGTH=3 STYLE="TEXT-ALIGN: left" TAG="24" ALT="{{�ŷ���ȭ}}"> <INPUT TYPE="Text" NAME="txtXchRate" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: right" TAG="24X5" ALT="{{ȯ��}}">
										</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur"><INPUT TYPE=HIDDEN NAME="txtXchRate">
<%	End If %>																				
										<TR>
											<TD CLASS="TD5" NOWRAP>{{���ݾ�}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcqAmt" SIZE=22 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{���ݾ�}}"></TD>
										</TR>
<%	If gIsShowLocal <> "N" Then	%>										
										<TR>
											<TD CLASS="TD5" NOWRAP>{{���ݾ�(�ڱ�)}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcqLocAmt" SIZE=22 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{���ݾ�(�ڱ�)}}">
											</TD>
										</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtAcqLocAmt">
<%	End If %>										
										<TR>
											<TD CLASS="TD5" NOWRAP>{{������}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcqQty" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{������}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{������}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtInvQty" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{������}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�����ڵ�}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCd" SIZE=15 MAXLENGTH=15 tag="24" ALT="{{�����ڵ�}}"> <INPUT TYPE="Text" NAME="txtAcctNm" SIZE=27 MAXLENGTH=30 tag="24" ALT="{{������}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�ŷ�ó}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=15 MAXLENGTH=15 tag="24" ALT="{{�ŷ�ó}}"> <INPUT TYPE="Text" NAME="txtBpNm" SIZE=27 MAXLENGTH=30 tag="24" ALT="{{�ŷ�ó��}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{��汸��}}</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboAcqFg" STYLE="WIDTH:120px;" tag="24" ALT="{{��汸��}}"><OPTION VALUE=""></OPTION></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{����/�뵵/ũ��}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtSpec" SIZE=25 MAXLENGTH=30 TAG="2x" ALT="{{����/�뵵/ũ��}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{����}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDesc" SIZE=40 MAXLENGTH=30 TAG="2x	" ALT="{{����}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�����󰢽��۳��}}</TD>
											<TD CLASS="TD6" NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDeprFrdt" CLASS=FPDTYYYYMM tag="24" Title="FPDATETIME" ALT={{�����󰢽��۳��}} id=DeprDateTime1> </OBJECT>');</SCRIPT>
											</TD>							
										</TR>											
									</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET STYLE="HEIGHT: 41%"><LEGEND>{{���⸻ �󰢳���: ��������(�ڱ�)}}</LEGEND>
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{���뿬��}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE style="LEFT: 0px; WIDTH: 80px; TOP: 0px; HEIGHT: 20px" name=txtTaxDurYrs CLASSID=<%=gCLSIDFPDS%> tag="22X60" ALT="{{���뿬��}}" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{����}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTaxDeprRate" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: right" TAG="24X5" ALT="{{����}}"> %</TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�󰢴���}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxDeprTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{�󰢴���}}" tag="22X2" id=fpDoubleSingle1></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�ں������⴩��}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxCptTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{�ں������⴩��}}" tag="22X2" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�̻��ܾ�}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxBalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{�̻��ܾ�}}" tag="22X2" id=fpDoubleSingle3></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�󰢻���}}</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboTaxDeprSts" STYLE="WIDTH:150px;" tag="23" ALT="{{�󰢻���}}"><OPTION VALUE=""></OPTION></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�󰢿Ϸ���}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtTaxDeprEnd" style="HEIGHT: 20px; WIDTH: 90px" tag="24" Title="FPDATETIME" ALT={{�󰢿Ϸ���}} id=fpDateTime1></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
									</FIELDSET><BR>
									<FIELDSET STYLE="HEIGHT: 41%"><LEGEND>{{���⸻ �󰢳���: ���ȸ�����(�ڱ�)}}</LEGEND>
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{���뿬��}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE style="LEFT: 0px; WIDTH: 80px; TOP: 0px; HEIGHT: 20px" name=txtCasDurYrs CLASSID=<%=gCLSIDFPDS%> tag="22X60" ALT="{{���뿬��}}" VIEWASTEXT id=OBJECT2></OBJECT>');</SCRIPT>										
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{����}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCasDeprRate" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: right" TAG="24X5" ALT="{{����}}"> %</TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�󰢴���}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCasDeprTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{�󰢴���}}" tag="22X2" id=fpDoubleSingle4></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�ں������⴩��}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCasCptTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{�ں������⴩��}}" tag="22X2" id=fpDoubleSingle5></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�̻��ܾ�}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCasBalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{�̻��ܾ�}}" tag="22X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�󰢻���}}</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboCasDeprSts" STYLE="WIDTH:150px;" tag="23" ALT="{{�󰢻���}}"><OPTION VALUE=""></OPTION></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{�󰢿Ϸ���}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtCasDeprEnd" CLASS=FPDTYYYYMM style="HEIGHT: 20px; WIDTH: 90px" tag="24" Title="FPDATETIME" ALT={{�󰢿Ϸ���}} id=toDateTime1></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
									</FIELDSET>
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
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=20><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode"        tag="24"><INPUT TYPE=hidden NAME="txtFlgMode" tag="24">
<INPUT TYPE=hidden NAME="htxtDeprYymm"   tag="24">
<INPUT TYPE=hidden NAME="htxtTaxDeprEnd" tag="24">
<INPUT TYPE=hidden NAME="htxtCasDeprEnd" tag="24">
<INPUT TYPE=hidden NAME="htxtRemRate"    tag="24">
<INPUT TYPE=hidden NAME="hTaxDeprSts"    tag="24">
<INPUT TYPE=hidden NAME="hCasDeprSts"    tag="24">
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


