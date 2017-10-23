
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  p1501ma1.asp
'*  4. Program Name         :  Resource Management
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p1501mb1.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "p1501mb2.asp"											'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "p1501mb3.asp"	
Const BIZ_PGM_LOOKUP_ID = "p1501mb4.asp"
Const BIZ_PGM_LOOKUP_CUR_ID = "p1502mb4.asp"	'추가 2003-04-17

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo
Dim iDBSYSDate
Dim StartDate, EndDate
Dim IsOpenPop          
Dim lgRdoOldVal
Dim lgCurCd
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, gDateFormat)
	frm1.txtValidFromDt.text  = startdate
	frm1.txtValidToDt.text	= UniConvYYYYMMDDToDate(gDateFormat, "2999","12","31")
	frm1.cboResourceType.value = "L"
	frm1.txtNoOfResource.Value = "1"
	frm1.txtCostType.value = "E"
	frm1.txtEfficiency.text = "100"
	frm1.txtUtilization.text = "100"
	frm1.rdoRunCrp2.checked = true
	frm1.rdoRunRccp2.checked = true
End Sub

Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1502", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboResourceType, lgF0, lgF1, Chr(11))
    Call CommonQueryRs(" RULE_TYPE,DESCRIPTION "," P_APS_RULE_DETAIL "," RULE_TYPE_CD = " & FilterVar("RSSLRL", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    '<!--	RCCP 관련 삭제 Start
    'lgF0 = "0" & Chr(11) & lgF0
    'lgF1 = "" & Chr(11) & lgF1
    'Call SetCombo2(frm1.txtSelectionRule, lgF0, lgF1, Chr(11))
    'RCCP 관련 삭제 End	-->
    Call CommonQueryRs(" RULE_TYPE,DESCRIPTION "," P_APS_RULE_DETAIL "," RULE_TYPE_CD = " & FilterVar("RSSQRL", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = "0" & Chr(11) & lgF0
    lgF1 = "" & Chr(11) & lgF1
    '<!--	RCCP 관련 삭제 Start
    'Call SetCombo2(frm1.txtSequenceRule, lgF0, lgF1, Chr(11))
    'RCCP 관련 삭제 End	-->
End Sub

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    arrField(2) = "CUR_CD"
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    arrHeader(2) = "통화코드"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()
	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "자원팝업"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtResourceCd1.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
	arrParam(5) = "자원"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "자원"		
    arrHeader(1) = "자원명"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceCd1.focus
	
End Function

'------------------------------------------  OpenResourceGroup()  -------------------------------------------------
'	Name : OpenResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtResourceGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "자원그룹팝업"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtResourceGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
	arrParam(5) = "자원그룹"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "자원그룹"		
    arrHeader(1) = "자원그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetResourceGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtResourceGroupCd.focus
	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "자원기준단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtResourceUnitCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION = " & FilterVar("TM", "''", "S") & ""			
	arrParam(5) = "자원기준단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "자원기준단위"		
    arrHeader(1) = "자원기준단위명"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceUnitCd.focus
		
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtCurCd.value      = UCase(arrRet(2))
	lgCurCd = UCase(arrRet(2))		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.txtResourceCd1.Value    = arrRet(0)		
	frm1.txtResourceNm1.Value    = arrRet(1)		
End Function

'------------------------------------------  SetResourceGroup()  --------------------------------------------------
'	Name : SetResourceGroup()
'	Description : ResourceGroup Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResourceGroup(byval arrRet)
	frm1.txtResourceGroupCd.Value    = arrRet(0)		
	frm1.txtResourceGroupNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Resource Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(byval arrRet)
	frm1.txtResourceUnitCd.Value    = arrRet(0)	
	frm1.txtResourceUnitCd1.value   = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function ChkValidData()
	ChkValidData = False
	
	With frm1
		If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function       
		'<!--	RCCP 관련 삭제 Start
		'If frm1.txtSelectionRule.value <> "" Then
		'	If CInt(frm1.txtSelectionRule.value) > 39 Then
		'		Call DisplayMsgBox("970025","X", "선택규칙","20")
		'		.txtSelectionRule.focus 
		'		Set gActiveElement = document.activeElement 
		'		Exit Function
		'	End IF	
		'End If
		
		'If frm1.txtSequenceRule.value <> "" Then
		'	If CInt(frm1.txtSequenceRule.value) > 39 Then
		'		Call DisplayMsgBox("970025","X", "순번규칙","20")
		'		.txtSequenceRule.focus 
		'		Set gActiveElement = document.activeElement 
		'		Exit Function
		'	End IF	
		'End If
		'RCCP 관련 삭제 End	-->
	End With
	
	ChkValidData = True
End Function

Sub ChkNumKeyPress()
	Dim KeyCode
	KeyCode = window.event.keyCode 
	
	If KeyCode < 48 Or KeyCode > 57 Then
		window.event.keyCode = 8
		Exit Sub
	End If
End Sub

'==========================================  2.5.6 LookUpRuleType() =======================================
'=	Event Name : LookUpRuleType																				=
'=	Event Desc :																						=
'========================================================================================================
Sub LookUpRuleType()
	LayerShowHide(1) 
		
	Dim strVal

	strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001				'☜: 비지니스 처리 ASP의 상태 	
	'<!--	RCCP 관련 삭제 Start
	'strVal = strVal & "&txtSelectionRule=" & Trim(frm1.txtSelectionRule.value)		'☆: 조회 조건 데이타 
	'RCCP 관련 삭제 End	-->
	strVal = strVal & "&PrevNextFlg=" & ""	
	strVal = strVal & "&lgCurDate=" & startdate
	
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
End Sub

'==========================================  2.5.7 LookUpItemOk() =======================================
'=	Event Name : LookUpItemOk																				=
'=	Event Desc :																						=
'========================================================================================================
Sub LookUpRuleTypeOk()
	IsOpenPop = False
End Sub

Sub LookUpRuleTypeNotOk()
	IsOpenPop = False
End Sub

Function CurCdLookUp()
		Dim strVal
		lgCurCd = ""
		frm1.txtCurCd.value = ""
		
		strVal = BIZ_PGM_LOOKUP_CUR_ID & "?txtMode=" & parent.UID_M0001					'☜: 비지니스 처리 ASP의 상태 	
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'☆: 조회 조건 데이타 
		strVal = strVal & "&PrevNextFlg=" & ""	
	
		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
End Function

Function CurCdLooKUpOk()
		lgCurCd = frm1.txtCurCd.value 
		IsOpenPop = False
End Function

Function CurCdLooKUpNotOk()
		
		IsOpenPop = False
End Function
'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call InitComboBox
    Call SetDefaultVal
    Call InitVariables
    Call SetToolbar("11101000000011")
    If parent.gPlant <> "" then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		Call CurCdLooKUp()
		frm1.txtResourceCd1.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

Sub txtResourceEa_Change()
    lgBlnFlgChgValue = True    
	frm1.txtResourceEa1.value = frm1.txtResourceEa.value	
End Sub

Sub txtMfgCost_Change()
    lgBlnFlgChgValue = True    	
End Sub

Sub txtResourceUnitCd_OnChange()
	frm1.txtResourceUnitCd1.value = frm1.txtResourceUnitCd.value 
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtNoOfResource_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtEfficiency_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtUtilization_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtOverloadTol_Change()
    lgBlnFlgChgValue = True
End Sub

Sub rdoRunRccp1_onChange()
    lgBlnFlgChgValue = True
End Sub
Sub rdoRunRccp2_onChange()
    lgBlnFlgChgValue = True
End Sub
Sub rdoRunCrp1_onChange()
    lgBlnFlgChgValue = True
End Sub
Sub rdoRunCrp2_onChange()
    lgBlnFlgChgValue = True
End Sub

Sub rdoInfiniteResourceFlg1_OnClick()
	If lgRdoOldVal = 1 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal = 1
End Sub

Sub cboResourceType_onchange()
	If frm1.cboResourceType.value = "M" Then
		frm1.txtCostType.value = "E"
	Else
		frm1.txtCostType.value = "L"
	End IF
	lgBlnFlgChgValue = True
End Sub

'<!--	RCCP 관련 삭제 Start
'Sub txtSelectionRule_onKeyPress()
'	Call ChkNumKeyPress
'	lgBlnFlgChgValue = True
'End Sub
'
'Sub txtSequenceRule_onKeyPress()
'	Call ChkNumKeyPress
'	lgBlnFlgChgValue = True
'End Sub
'RCCP 관련 삭제 End	-->

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'----------  Coding part  ------------------------------------------------------------- 


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
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False															'⊙: Processing is NG
    
    Err.Clear																	'☜: Protect system from crashing

   '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
 '-----------------------
    'Erase contents area
    '----------------------- 
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtResourceCd1.value = "" Then
		frm1.txtResourceNm1.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call SetDefaultVal
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
    If DbQuery = False Then   
		Exit Function           
    End If     										'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False																'⊙: Processing is NG
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")					'⊙: "데이타가 변경되었습니다. 신규입력을 하시겠습니까?"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    frm1.txtResourceCd1.value = ""
    frm1.txtResourceNm1.value = "" 
    Call ggoOper.ClearField(Document, "2")                                      '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
	Call SetToolbar("11101000000011")    
	frm1.txtResourceCd2.focus 
	Set gActiveElement = document.activeElement 
	frm1.txtCurCd.value = lgCurCd	
    FncNew = True																'⊙: Processing is OK
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim intRetCD
    
    FncDelete = False														'⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then   
		Exit Function           
    End If     														'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '⊙: No data changed!!
        Exit Function
    End If
	
    If Not chkField(Document, "2") Then                             '⊙: Check contents area
       Exit Function
    End If
	
	If UniCDbl(frm1.txtResourceEa.Text) <= CDbl(0) Then
		Call DisplayMsgBox("970022","X",frm1.txtResourceEa.alt,"0")
		frm1.txtResourceEa.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
    If DbSave = False Then
		Exit Function           
    End If     				                                                '☜: Save db data

    FncSave = True                                                          '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												'⊙: Indicates that current mode is Crate mode
    
    Call ggoOper.LockField(Document, "N")									'⊙: This function lock the suitable field
	Call SetToolbar("11101000000011")
	
    frm1.txtResourceCd1.value = ""
    frm1.txtResourceNm1.value = ""
    
    frm1.txtResourceCd2.value = ""
	
	frm1.txtValidFromDt.text  = startdate
	frm1.txtValidToDt.text	= UniConvYYYYMMDDToDate(gDateFormat, "2999","12","31")
    
    frm1.txtResourceCd2.focus
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
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                               '☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables

    Err.Clear                                                               '☜: Protect system from crashing
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
	strVal = strVal & "&txtResourceCd1=" & Trim(frm1.txtResourceCd1.value)				'☜: 조회 조건 데이타 
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & "P"
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                               '☆: 밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 
	strVal = strVal & "&txtResourceCd1=" & Trim(frm1.txtResourceCd1.value)				'☜: 조회 조건 데이타 
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & "N"
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)											'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
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
    
    LayerShowHide(1) 

    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
	strVal = strVal & "&txtResourceCd=" & Trim(frm1.hResourceCd.value)
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                         '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Err.Clear                                                               '☜: Protect system from crashing
    DbQuery = False                                                         '⊙: Processing is NG
    
    Dim strVal
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtResourceCd1=" & Trim(frm1.txtResourceCd1.value)
    strVal = strVal & "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)
	
    DbQuery = True
End Function

'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
Function DbQueryOk()
    frm1.hPlantCd.value = frm1.txtPlantCd.value 
    lgCurCd = frm1.txtCurCd.value
	
    lgIntFlgMode = parent.OPMD_UMODE
    lgBlnFlgChgValue = false
        
    Call ggoOper.LockField(Document, "Q")

	Call SetToolbar("11111000111111")
	
	frm1.txtResourceNm2.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 
	Dim BlnRetCd

    Err.Clear																'☜: Protect system from crashing

	DbSave = False															'⊙: Processing is NG

    Dim strVal

	BlnRetCd = ChkValidData

	If BlnRetCd = False Then
		Exit Function
	End if

	LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With

    DbSave = True                                                           '⊙: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
    frm1.txtResourceCd1.value = frm1.txtResourceCd2.value 
    frm1.txtResourceNm1.value = frm1.txtResourceNm2.value 

    Call InitVariables
    
    Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자원등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()"> <INPUT TYPE=TEXT ID="txtPlantNm" NAME="arrCond" SIZE=50 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>자원</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd1" SIZE=20 MAXLENGTH=10 tag="12XXXU" ALT="자원"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm1" SIZE=50 tag="14"></TD>
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
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100% valign=top>
									<FIELDSET>
										<LEGEND>일반정보</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd2" SIZE=20 MAXLENGTH=10 tag="23XXXU" ALT="자원">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm2" SIZE=50 MAXLENGTH=40 tag="22XXXX" ALT="자원명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원그룹</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroupCd" SIZE=20 MAXLENGTH=10 tag="23XXXU" ALT="자원그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResourceGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroupNm" SIZE=50 tag="24"></TD>
											</TR>	
											<TR>
												<TD CLASS=TD5 NOWRAP>자원구분</TD>
												<TD CLASS=TD656 NOWRAP><SELECT NAME="cboResourceType" ALT="자원구분" STYLE="Width: 98px;" tag="22"></SELECT></TD>																								
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원수</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I434898900_txtNoOfResource.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>효율</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I599973791_txtEfficiency.js'></script>&nbsp;%
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>가동율</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I784308445_txtUtilization.js'></script>&nbsp;%
												</TD>
											</TR>		
											<TR ID=Q1>
												<TD CLASS=TD5 NOWRAP>RCCP부하계산대상</TD>
												<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="2X" ID="rdoRunRccp1" VALUE="Y"><LABEL FOR="rdoRunRccp1">예</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="2X" ID="rdoRunRccp2" VALUE="N"><LABEL FOR="rdoRunRccp2">아니오</LABEL></TD>
											</TR>
											<TR ID=Q2>
												<TD CLASS=TD5 NOWRAP>CRP부하계산대상</TD>
												<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="24" ID="rdoRunCrp1" VALUE="Y"><LABEL FOR="rdoRunCrp1">예</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="24" ID="rdoRunCrp2" VALUE="N"><LABEL FOR="rdoRunCrp2">아니오</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>과부하허용율</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I161376506_txtOverloadTol.js'></script>&nbsp;%
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자원기준수량</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I833706259_txtResourceEa.js'></script>												
												</TD>
											</TR>																																	
											<TR>
												<TD CLASS=TD5 NOWRAP>자원기준단위</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceUnitCd" SIZE=5 MAXLENGTH=3 tag="22XXXU" ALT="자원기준단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>기준단위당 단위제조경비</TD>
												<TD CLASS=TD656 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>															
																<script language =javascript src='./js/p1501ma1_I324865192_txtMfgCost.js'></script>
															</TD>
															<TD>											
																&nbsp;<INPUT TYPE=TEXT NAME="txtCurCd" tag=24 SIZE=5 MAXLENGTH=3 ALT="통화코드">&nbsp;/&nbsp;
															</TD>
															<TD>
																<script language =javascript src='./js/p1501ma1_I870693814_txtResourceEa1.js'></script>												
															</TD>
															<TD>
																&nbsp;<INPUT TYPE=TEXT NAME="txtResourceUnitCd1" SIZE=5 MAXLENGTH=3 tag="24XXXU" ALT="자원기준단위">
															</TD>
														</TR>
													</TABLE>												
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I130084407_txtValidFromDt.js'></script>
													&nbsp;~&nbsp;
													<script language =javascript src='./js/p1501ma1_I571680456_txtValidToDt.js'></script>										
												</TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hResourceCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCostType" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
