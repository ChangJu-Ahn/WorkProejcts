<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7103ma1
'*  4. Program Name         : 고정자산 MASTER 수정 
'*  5. Program Desc         : 고정자산별 MASTER를 수정,조회 
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
'												1. 선 언 부 
'##########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->						<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!--==========================================  1.1.1 Style Sheet  ===========================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<!--==========================================  1.1.2 공통 Include   =========================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                             '☜: indicates that All variables must be declared in advance 
<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

Const BIZ_PGM_ID = "a7103mb1.asp"     											 '☆: 비지니스 로직 ASP명 

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgBlnFlgConChg				'☜: Condition 변경 Flag
'@Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
'@Dim lgIntGrpCount				'☜: Group View Size를 조사할 변수 
'@Dim lgIntFlgMode					'☜: Variable is for Operation Status

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""


'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 

'-------------------  공통 Global 변수값 정의  ----------------------------------------------------------- 


'+++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '----------  Coding part  -----------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
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


'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
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

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'********************************************************************************************************* 


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 


'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
 '------------------------------------------  OpenMasterRef()  -------------------------------------------------
'	Name : OpenMasterRef()
'	Description : Asset Master Condition PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' 권한관리 추가
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

	arrParam(0) = "{{계정코드팝업}}"			' 팝업 명칭 
	arrParam(1) = "a_asset_acct, a_acct"		' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtAcctCd.value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "a_asset_acct.acct_cd = a_acct.acct_cd"		' Where Condition
	arrParam(5) = "{{계정코드}}"				' 조건필드의 라벨 명칭 
	
    arrField(0) = "a_asset_acct.acct_cd"		' Field명(0)
    arrField(1) = "a_acct.acct_sh_nm"			' Field명(1)
    
    arrHeader(0) = "{{계정코드}}"				' Header명(0)
    arrHeader(1) = "{{계정명}}"					' Header명(1)
    
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
'	Description : 관리자Id PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMgmtId()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg

	If IsOpenPop = True Or UCase(txtMgmtUserId.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "{{사원코드팝업}}"			' 팝업 명칭 
	arrParam(1) = ""							' TABLE 명칭 
	arrParam(2) = ""							' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "{{사원코드}}"				' 조건필드의 라벨 명칭 
	
    arrField(0) = ""							' Field명(0)
    arrField(1) = ""							' Field명(1)
    
    arrHeader(0) = "{{사원코드}}"				' Header명(0)
    arrHeader(1) = "{{사원명}}"					' Header명(1)
    
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

	arrParam(0) = "{{거래통화 팝업}}"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "{{거래통화}}"
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "{{거래통화}}"		
    arrHeader(1) = "{{거래통화명}}"		
    
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
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'-------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
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


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function funChkAmt()
    dim Ltaxbalamt, Lcasbalamt
    
    funChkAmt = False
    
	IF frm1.txtTaxBalAmt.value <> "" Then
		Ltaxbalamt = UNICDbl(frm1.txtTaxBalAmt.value)
	    if Ltaxbalamt < 0 then
			Call DisplayMsgBox("AS0049", "X", "X", "X")                                '☆: 밑에 메세지를 ID로 처리해야 함 
	'       call MsgBox("미상각 잔액이 0보다 작아서는 안됩니다..",vbInformation)
	        frm1.txtTaxBalAmt.focus
	        Set gActiveElement = document.activeElement
	        Exit Function
	    End If
	End If    

	If frm1.txtCasBalAmt.value <> "" Then
	    Lcasbalamt = UNICDbl(frm1.txtCasBalAmt.value)
	    if Lcasbalamt < 0 then
			Call DisplayMsgBox("AS0049", "X", "X", "X")                                '☆: 밑에 메세지를 ID로 처리해야 함 
	'       call MsgBox("미상각 잔액이 0보다 작아서는 안됩니다..",vbInformation)
	        frm1.txtCasBalAmt.focus
	        Set gActiveElement = document.activeElement
	        Exit Function
	    End if
	End If
    funChkAmt = True
     
End function


'##########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'#########################################################################################################

'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

'    Call GetGlobalVar
'    Call ClassLoad																	'⊙: Load Common DLL
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call AppendNumberPlace("7","3","0")
    Call AppendNumberPlace("6","2","0")
    Call AppendNumberRange("0","1","60")
    
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.FormatDate(frm1.txtDeprFrDt, gDateFormat, 2)

    Call ggoOper.FormatDate(frm1.txtTaxDeprEnd, gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtCasDeprEnd, gDateFormat, 2)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    Call InitVariables																'⊙: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("110000000000111")
    Call SetDefaultVal
    Call InitComboBox
	
	frm1.txtCondAsstNo.focus	
	Set gActiveElement = document.activeElement	

	' 권한관리 추가
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인
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


'***************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 


'***********************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
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
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtTaxDeprEnd_DblClick(Button)
    If Button = 1 Then
        frm1.txtTaxDeprEnd.Action = 7
    End If
End Sub


'=======================================================================================================
'   Event Name : txtCasDeprEnd_DblClick(Button)
'   Event Desc : 달력을 호출한다.
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
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 

'********************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD     
    
    FncQuery = False                                                        '⊙: Processing is NG    

    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
'		IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field

'    ggoSpread.Source = frm1.vspdData
'	ggospread.ClearSpreadData		'Buffer Clear

    Call InitVariables															'⊙: Initializes local global variables
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
     
    '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
'		IntRetCD = MsgBox("데이타가 변경되었습니다. 신규입력을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    Call InitVariables															'⊙: Initializes local global variables
	Call SetToolBar("110000000000111")
    Call SetDefaultVal
    
    FncNew = True																'⊙: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False														'⊙: Processing is NG
    
  '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
'        Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If
    
  '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
    
End Function

'----------------------------------------------------------
'  Functions before fncSave
'----------------------------------------------------------
Function FncChkAmt()
	'-----------------------------------------------------------	
	' 취득금액,자본적지출누계금액,상각누계금액,미상각잔액 Check
	'-----------------------------------------------------------
	Dim varAcqAmt, varDeprTotamt,varCptTotAmt,varBalAmt
	Dim strRegDt,strFiscDt
	
	FncChkAmt = False
	
	strRegDt	= UniConvDateToYYYYMMDD(frm1.RegDateTime1.Text, gDateFormat, "")
    strFiscDt   = UniConvDateToYYYYMMDD(parent.gFiscStart, parent.gAPDateFormat,"")  ' 당기 시작월 
    
	varAcqAmt	  = UNICDbl(frm1.txtAcqLocAmt.value) 
	varDeprTotAmt = UNICDbl(frm1.txtTaxDeprTotAmt.value) 
	varCptTotAmt  = UNICDbl(frm1.txtTaxCptTotAmt.value) 
	varBalAmt	  = UNICDbl(frm1.txtTaxBalAmt.value) 

	'-------------------------------------------------------------
	' 당기시작일자 이후에 취득한 경우 전기말 데이타는 없어야 한다.
	'-------------------------------------------------------------

	if strRegDt >= strFiscDt then   
		if varDeprTotAmt > 0 or varCptTotAmt > 0 or varBalAmt >0  then
			Call DisplayMsgBox("117428", "X", "X", "X")  '''당기이후에 취득한 자산은 전기말상각내역을 입력할 수 없습니다.
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
	' 당기시작일자 이후에 취득한 경우 전기말 데이타는 없어야 한다.
	'-------------------------------------------------------------
	if strRegDt >= strFiscDt then   
		if varDeprTotAmt > 0 or varCptTotAmt > 0 or varBalAmt >0 then
			Call DisplayMsgBox("117428", "X", "X", "X")  '''당기이후에 취득한 자산은 전기말상각내역을 입력할 수 없습니다.
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

	if frm1.hTaxDeprSts.value <> "03" then '상각대상의 자산에 대해 비상각을 선택 시 
		if frm1.cboTaxDeprSts.value = "03" then
			Call DisplayMsgBox("117423", "X", "X", "X")
			frm1.cboTaxDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if		
	end if
	
	if frm1.hCasDeprSts.value <> "03" then '상각대상의 자산에 대해 비상각을 선택 시 
		if frm1.cboCasDeprSts.value = "03" then
			Call DisplayMsgBox("117423", "X", "X", "X")
			frm1.cboCasDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if
	end if	
	
	if frm1.cboTaxDeprSts.value = "02" then  '상각완료일 때 
		if frm1.fpDateTime1.text = "" then ' 상각완료년월을 입력하지 않은 경우 
			Call DisplayMsgBox("117422", "X", "X", "X")
			Exit Function
		end if
	end if
	if frm1.cboCasDeprSts.value = "02" then  '상각완료일 때 
		if frm1.toDateTime1.text = "" then ' 상각완료년월을 입력하지 않은 경우 
			Call DisplayMsgBox("117422", "X", "X", "X")
			Exit Function
		end if
	end if	

	fncChkDeprSts = True
	
End Function

Function FncChkBalAmt()
	Dim strRemRate
	Dim varRemRate   ''잔존율 
	Dim varAcqAmt,varCptTotAmtTax,varBalAmtTax
	Dim varCptTotAmtCas,varBalAmtCas
	Dim varRemAmtTax,varRemAmtCas
	Dim varInvQty
	Dim strRegDt,strFiscDt
			
	FncChkBalAmt = False
	strRegDt	= UniConvDateToYYYYMMDD(frm1.RegDateTime1.Text, gDateFormat, "")
    strFiscDt   = UniConvDateToYYYYMMDD(parent.gFiscStart, parent.gAPDateFormat,"")  ' 당기 시작 월 
	
	
	'-------------------------------------------------------------
	' 당기시작일자 이후에 취득한 경우 전기말 데이타는 없어야 한다.
	'-------------------------------------------------------------	
	if strRegDt >= strFiscDt then
		if frm1.cboTaxDeprSts.value = "02" then    '세법기준: 상각완료 
			Call DisplayMsgBox("117423", "X", "X", "X")  '''상각상태를 확인하십시오.
			frm1.cboTaxDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if
		if frm1.cboCasDeprSts.value = "02" then    '기업회계기준: 상각완료 
			Call DisplayMsgBox("117423", "X", "X", "X")  '''상각상태를 확인하십시오.
			frm1.cboCasDeprSts.focus
			Set gActiveElement = document.activeElement
			exit function
		end if			
	
	else
		strRemRate = Trim(frm1.htxtRemRate.value)   '잔존율(정액:0%,정률: 5%)
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
	
		'1. 세법기준				
		if frm1.cboTaxDeprSts.value = "02" then 
		'**************************************************************
		'  상각상태-상각완료 선택 시, 미상각금액도 상각완료되는 금액인지 체크: 
		'**************************************************************	
		'''(취득가+자본적지출금액) * 잔존율 * 0.01 < 미상각잔액?
			IF varRemAmtTax < varBalAmtTax then
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if	
		else
		'************************************************************
		' 상각완료 아니면서 미상각금액이 상각완료되는 금액이면 Error
		'************************************************************
			IF varRemAmtTax >= varBalAmtTax then   '미망금액 >= 미상각금액? 즉,상각완료되는 금액 
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if			
		end if

		'2. 기업회계기준 
		IF frm1.cboCasDeprSts.value = "02" THEN 		
		'**************************************************************
		'  상각완료 선택 시, 미상각금액도 상각완료되는 금액인지 체크: 
		'**************************************************************		
			IF varRemAmtCas < varBalAmtCas then
				Call DisplayMsgBox("117425", "X", "X", "X")
				Exit Function
			End if					
		ELSE
		'************************************************************
		' 상각완료 아니면서 미상각금액이 상각완료되는 금액이면 Error
		'***********************************************************	
			IF varRemAmtCas >= varBalAmtCas then   '미망금액 >= 미상각금액? 즉,상각완료되는 금액 
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
		Call DisplayMsgBox("117426", "X", "X", "X")   ''상각시작년월은 취득년월보다 크거나 같아야 합니다.
		Exit Function
	end if
	strEndDt = UniConvDateToYYYYMM(frm1.fpDateTime1.Text, gDateFormat, "")
		
	if strEnddt <> "" then
		if strRegDt > strEndDt then
			Call DisplayMsgBox("117427", "X", "X", "X")   ''상각완료년월은 취득년월보다 커야 합니다.
			Exit Function
		end if	
	end if
		
	strEndDt = ""
	strEndDt = UniConvDateToYYYYMM(frm1.toDateTime1.Text, gDateFormat, "")
	if strEnddt <> "" then	
		if strRegDt > strEndDt then
			Call DisplayMsgBox("117427", "X", "X", "X")   ''상각완료년월은 취득년월보다 커야 합니다.
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
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                              '☜: Protect system from crashing    
	'-----------------------
    ' Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                   '⊙: No data changed!!
        Exit Function
    End If    
	'-----------------------
    ' Check content area
    '-----------------------
    If Not chkField(Document, "2") Then										'⊙: Check contents area
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
	' FncChkAmt(): 취득가+자본적지출-상각누계금액 = 미상각금액 ?
	'********************************************************************
	If FncChkAmt = False Then
		exit function
	End if
	
	'********************************************************************
	' FncChkBalAmt(): 미상각금액과 상각완료여부와 상각완료년월 Check
	'********************************************************************
	If FncChkBalAmt = False Then
		exit function
	End if
			
	'********************************************************************
	' FncChkDeprSts(): 상각상태와 상각완료년월 Check
	'********************************************************************
	If fncChkDeprSts = False Then
		Exit function
	End if
	
	'********************************************************************
	' FncChkDeprSts(): 상각완료여부와 상각완료년월 Check
	'***************************************************
	if fncChkDeprStarYymm = False then
		Exit Function
	end if
		
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
'		IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												'⊙: Indicates that current mode is Crate mode
    
     ' 조건부 필드를 삭제한다. 
    Call ggoOper.ClearField(Document, "1")                                      '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")									'⊙: This function lock the suitable field
    
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
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
        Call DisplayMsgBox("900002", "X", "X", "X")                                 '☆: 밑에 메세지를 ID로 처리해야 함 
        'Call MsgBox("조회한후에 지원됩니다.", vbInformation)
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")                                 '☆: 
		'Call MsgBox("이전 데이타가 없습니다..", vbInformation)
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
    strVal = strVal & "&txtAssetCd1=" & lgPrevNo							'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")                                '☆: 
    End If
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태값 
    strVal = strVal & "&txtAssetCd1=" & lgNextNo							'☆: 조회 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
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
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
    Err.Clear                                                               '☜: Protect system from crashing
    
    DbDelete = False														'⊙: Processing is NG
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtAssetCd1=" & Trim(frm1.txtAssetCd1.value)		'☜: 삭제 조건 데이타 

	' 권한관리 추가
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인

     
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	Call FncNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

    Err.Clear                                                               '☜: Protect system from crashing
    
    DbQuery = False                                                         '⊙: Processing is NG

	Call LayerShowHide(1)
	
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCondAsstNo=" & Trim(frm1.txtCondAsstNo.value)	'☆: 조회 조건 데이타 

	' 권한관리 추가
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                          '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

	Call SetToolBar("110010000001111")	'111010000001111

	if frm1.cboTaxDeprSts.value = "03" Then'비상각인 경우 

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

	if frm1.cboCasDeprSts.value = "03" Then'비상각인 경우 
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
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
    Dim strVal
	Dim varDeprdt,varTaxDt,varCasDt
	Dim strYear
	Dim strMonth
	Dim strDay	
	
    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG
  
	Call LayerShowHide(1)
	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
	
	
		Call ExtractDateFrom(frm1.DeprDateTime1.Text,frm1.DeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
		varDeprdt = strYear & strMonth
		
		varTaxDt  = UniConvDateToYYYYMM(frm1.fpDateTime1.Text, gDateFormat, "")
		varCasDt  = UniConvDateToYYYYMM(frm1.toDateTime1.Text, gDateFormat, "")
		
		frm1.htxtDeprYymm.value   = varDeprDt
		frm1.htxtTaxDeprEnd.value = varTaxDt	
		frm1.htxtCasDeprEnd.value = varCasDt

		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()															'☆: 저장 성공후 실행 로직 

    'frm1.txtAssetCd1.value = frm1.txtAssetCd2.value    'Conditon의 자산코드 
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

	<!-- 탭구분  -->
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

	<!-- 본문내용  -->
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
									<TD CLASS="TD5" NOWRAP>{{자산번호}}</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE="Text" NAME="txtCondAsstNo" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="{{자산번호}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef()"> <INPUT TYPE="Text" NAME="txtCondAsstNm" SIZE=30 MAXLENGTH=30 tag="14" ALT="{{자산명}}"></TD>
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
									<FIELDSET STYLE="HEIGHT: 100%"><LEGEND>{{기본정보}}</LEGEND>
									<TABLE CLASS="TB2" CELLSPACING=0 STYLE="HEIGHT: 96%">
										<TR>
											<TD CLASS="TD5" NOWRAP>{{자산명}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNm" SIZE=44 MAXLENGTH=40 TAG="2x" ALT="{{자산명}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{참조번호}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtRefNo" SIZE=30 MAXLENGTH=30 TAG="2x" ALT="{{참조번호}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{관리부서}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCd" SIZE=15 MAXLENGTH=10 tag="24" ALT="{{관리부서코드}}"> <INPUT TYPE="Text" NAME="txtDeptNm" SIZE=27 MAXLENGTH=40 tag="24" ALT="{{관리부서명}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{취득일자}}</TD>
											<TD CLASS="TD6" NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=RegDateTime1 name=txtRegDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="{{등록일자}}"></TD> </OBJECT>');</SCRIPT>											    
											</TD>										
										</TR>																	
<%	If gIsShowLocal <> "N" Then	%>										
										<TR>
											<TD CLASS="TD5" NOWRAP>{{거래통화}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDocCur" SIZE=10 MAXLENGTH=3 STYLE="TEXT-ALIGN: left" TAG="24" ALT="{{거래통화}}"> <INPUT TYPE="Text" NAME="txtXchRate" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: right" TAG="24X5" ALT="{{환율}}">
										</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur"><INPUT TYPE=HIDDEN NAME="txtXchRate">
<%	End If %>																				
										<TR>
											<TD CLASS="TD5" NOWRAP>{{취득금액}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcqAmt" SIZE=22 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{취득금액}}"></TD>
										</TR>
<%	If gIsShowLocal <> "N" Then	%>										
										<TR>
											<TD CLASS="TD5" NOWRAP>{{취득금액(자국)}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcqLocAmt" SIZE=22 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{취득금액(자국)}}">
											</TD>
										</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtAcqLocAmt">
<%	End If %>										
										<TR>
											<TD CLASS="TD5" NOWRAP>{{취득수량}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcqQty" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{취득수량}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{재고수량}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtInvQty" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: right" TAG="24" ALT="{{재고수량}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{계정코드}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCd" SIZE=15 MAXLENGTH=15 tag="24" ALT="{{계정코드}}"> <INPUT TYPE="Text" NAME="txtAcctNm" SIZE=27 MAXLENGTH=30 tag="24" ALT="{{계정명}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{거래처}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=15 MAXLENGTH=15 tag="24" ALT="{{거래처}}"> <INPUT TYPE="Text" NAME="txtBpNm" SIZE=27 MAXLENGTH=30 tag="24" ALT="{{거래처명}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{취득구분}}</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboAcqFg" STYLE="WIDTH:120px;" tag="24" ALT="{{취득구분}}"><OPTION VALUE=""></OPTION></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{구조/용도/크기}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtSpec" SIZE=25 MAXLENGTH=30 TAG="2x" ALT="{{구조/용도/크기}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{적요}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDesc" SIZE=40 MAXLENGTH=30 TAG="2x	" ALT="{{적요}}"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{감가상각시작년월}}</TD>
											<TD CLASS="TD6" NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDeprFrdt" CLASS=FPDTYYYYMM tag="24" Title="FPDATETIME" ALT={{감가상각시작년월}} id=DeprDateTime1> </OBJECT>');</SCRIPT>
											</TD>							
										</TR>											
									</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET STYLE="HEIGHT: 41%"><LEGEND>{{전기말 상각내역: 세법기준(자국)}}</LEGEND>
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{내용연수}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE style="LEFT: 0px; WIDTH: 80px; TOP: 0px; HEIGHT: 20px" name=txtTaxDurYrs CLASSID=<%=gCLSIDFPDS%> tag="22X60" ALT="{{내용연수}}" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각율}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTaxDeprRate" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: right" TAG="24X5" ALT="{{상각율}}"> %</TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각누계}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxDeprTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{상각누계}}" tag="22X2" id=fpDoubleSingle1></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{자본적지출누계}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxCptTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{자본적지출누계}}" tag="22X2" id=fpDoubleSingle2></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{미상각잔액}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTaxBalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{미상각잔액}}" tag="22X2" id=fpDoubleSingle3></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각상태}}</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboTaxDeprSts" STYLE="WIDTH:150px;" tag="23" ALT="{{상각상태}}"><OPTION VALUE=""></OPTION></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각완료년월}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtTaxDeprEnd" style="HEIGHT: 20px; WIDTH: 90px" tag="24" Title="FPDATETIME" ALT={{상각완료년월}} id=fpDateTime1></OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
									</FIELDSET><BR>
									<FIELDSET STYLE="HEIGHT: 41%"><LEGEND>{{전기말 상각내역: 기업회계기준(자국)}}</LEGEND>
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{내용연수}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE style="LEFT: 0px; WIDTH: 80px; TOP: 0px; HEIGHT: 20px" name=txtCasDurYrs CLASSID=<%=gCLSIDFPDS%> tag="22X60" ALT="{{내용연수}}" VIEWASTEXT id=OBJECT2></OBJECT>');</SCRIPT>										
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각율}}</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCasDeprRate" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: right" TAG="24X5" ALT="{{상각율}}"> %</TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각누계}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCasDeprTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{상각누계}}" tag="22X2" id=fpDoubleSingle4></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{자본적지출누계}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCasCptTotAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{자본적지출누계}}" tag="22X2" id=fpDoubleSingle5></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{미상각잔액}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCasBalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE ALT="{{미상각잔액}}" tag="22X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각상태}}</TD>
											<TD CLASS="TD6" NOWRAP><SELECT NAME="cboCasDeprSts" STYLE="WIDTH:150px;" tag="23" ALT="{{상각상태}}"><OPTION VALUE=""></OPTION></SELECT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>{{상각완료년월}}</TD>
											<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtCasDeprEnd" CLASS=FPDTYYYYMM style="HEIGHT: 20px; WIDTH: 90px" tag="24" Title="FPDATETIME" ALT={{상각완료년월}} id=toDateTime1></OBJECT>');</SCRIPT></TD>
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


