<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4317oa1.asp
'*  4. Program Name         : 자재출고요청서출력(품목별)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/01/10
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 						<% '⊙: 화면처리ASP에서 서버작업이 필요한 경우 %>
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--☜:Print Program needs this vbs file-->
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = "<%=GetSvrDate%>"	
StartDate = UniConvDateAToB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormat)     	'☆: 초기화면에 뿌려지는 시작 날짜 
EndDate = UNIDateAdd("M",1,StartDate,parent.gDateFormat)							    '☆: 초기화면에 뿌려지는 마지막 날짜 

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
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
	frm1.txtReqDt1.Text	= StartDate
	frm1.txtReqDt2.Text	= EndDate
End Sub
'=======================================================================================================
'   Event Name : txtReqDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReqDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqDt1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReqDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqDt2.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqDt2.Focus
    End If
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
End Sub

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

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call ggoOper.FormatField(Document, "x",parent.ggStrIntegeralPart, parent.ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call InitVariables                                                      '⊙: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("10000000000011")										'⊙: 버튼 툴바 제어 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd1.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 	 
		Set gActiveElement = document.activeElement
	End If    
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
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
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint() 
	
	Dim strEbrFile
    Dim objName
	
    Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then							'⊙: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If parent.ValidDateCheck(frm1.txtReqDt1, frm1.txtReqDt2) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	
	if frm1.txtWorkCtCd1.value = "" then 
		var2 = "0"	
		frm1.txtWorkCtNm1.value = ""
	else
		var2 = UCase(Trim(frm1.txtWorkCtCd1.value))
	End If
	
	if frm1.txtWorkCtCd2.value = "" then 
		var3 = "zzzzzzzzzzzzzzzzzz"	
		frm1.txtWorkCtNm2.value = ""
	else
		var3 = UCase(Trim(frm1.txtWorkCtCd2.value))
	End If
	
	if frm1.txtItemCd1.value = "" then 
		frm1.txtItemNm1.value = ""
		var4 = "0"	
	else
		var4 = UCase(Trim(frm1.txtItemCd1.value))
	End If
	
	if frm1.txtItemCd2.value = "" then 
		frm1.txtItemNm2.value = ""
		var5 = "zzzzzzzzzzzzzzzzzz"	
	else
		var5 = UCase(Trim(frm1.txtItemCd2.value))
	End If
	
	var6 = UniConvDateAToB(frm1.txtReqDt1.text,parent.gDateFormat,parent.gServerDateFormat)
	var7 = UniConvDateAToB(frm1.txtReqDt2.text,parent.gDateFormat,parent.gServerDateFormat)

	strUrl = strUrl & "plant_cd|" & var1
	strUrl = strUrl & "|fr_work_ct|" & var2
	strUrl = strUrl & "|to_work_ct|" & var3
	strUrl = strUrl & "|fr_item_cd|" & var4
	strUrl = strUrl & "|to_item_cd|" & var5
	strUrl = strUrl & "|fr_req_dt|" & var6
	strUrl = strUrl & "|to_req_dt|" & var7 
	
	strEbrFile = "p4317oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
'----------------------------------------------------------------
' Print 함수에서 추가되는 부분 
'----------------------------------------------------------------
	call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	

	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement
	
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview() 
	
	Dim strEbrFile
    Dim objName
	
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	
	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then							'⊙: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If parent.ValidDateCheck(frm1.txtReqDt1, frm1.txtReqDt2) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))

	
	if frm1.txtWorkCtCd1.value = "" then 
		var2 = "0"	
		frm1.txtWorkCtNm1.value = ""
	else
		var2 = UCase(Trim(frm1.txtWorkCtCd1.value))
	End If
	
	if frm1.txtWorkCtCd2.value = "" then 
		var3 = "zzzzzzzzzzzzzzzzzz"	
		frm1.txtWorkCtNm2.value = ""
	else
		var3 = UCase(Trim(frm1.txtWorkCtCd2.value))
	End If
	
	if frm1.txtItemCd1.value = "" then 
		frm1.txtItemNm1.value = ""
		var4 = "0"	
	else
		var4 = UCase(Trim(frm1.txtItemCd1.value))
	End If
	
	if frm1.txtItemCd2.value = "" then 
		frm1.txtItemNm2.value = ""
		var5 = "zzzzzzzzzzzzzzzzzz"	
	else
		var5 = UCase(Trim(frm1.txtItemCd2.value))
	End If
	
	var6 = UniConvDateAToB(frm1.txtReqDt1.text,parent.gDateFormat,parent.gServerDateFormat) 
	var7 = UniConvDateAToB(frm1.txtReqDt2.text,parent.gDateFormat,parent.gServerDateFormat) 
	
	strUrl = strUrl & "plant_cd|" & var1
	strUrl = strUrl & "|fr_work_ct|" & var2
	strUrl = strUrl & "|to_work_ct|" & var3
	strUrl = strUrl & "|fr_item_cd|" & var4
	strUrl = strUrl & "|to_item_cd|" & var5
	strUrl = strUrl & "|fr_req_dt|" & var6
	strUrl = strUrl & "|to_req_dt|" & var7
	
	strEbrFile = "p4317oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr") 

	call FncEBRPreview("p4317oa1.ebr", strUrl)
	
	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement

End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================
Function FncQuery()
	On Error Resume Next                                                    '☜: Protect system from crashing	
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)									<%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                         <%'☜: Protect system from crashing%>
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
    	arrField(0) = "PLANT_CD"				' Field명(0)
    	arrField(1) = "PLANT_NM"				' Field명(1)
    
    	arrHeader(0) = "공장"				' Header명(0)
    	arrHeader(1) = "공장명"				' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConWC(ByVal strCode, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call parent.DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	IsOpenPop = True

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & parent.FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")		' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    arrField(2) = "INSIDE_FLG"
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    arrHeader(2) = "작업장타입"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet,iPos)
	End If
	
	If iPos = "0" Then
		frm1.txtWorkCtCd1.focus
	Else
		frm1.txtWorkCtCd2.focus
	End If
		
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item1 PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(ByVal strCode, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
		
	If frm1.txtPlantCd.value = "" Then
		Call parent.DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	iCalledAspName = AskPRAspName("B1B11PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.Value)						   ' Plant Code
	arrParam(1) = strCode	' Item Code
	arrParam(2) = ""		' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"	
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"	
    arrField(2) = 3								' Field명(1) : "ITEM_ACCT"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent ,arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If	
	
	Call SetFocusToDocument("M")
	
	If iPos = "0" Then
		frm1.txtItemCd1.focus
	Else
		frm1.txtItemCd2.focus
	End If
	
End Function

'------------------------------------------  SetPlantCd()  --------------------------------------------------
'	Name : SetPlantCd()
'	Description : Plant  Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function

Function SetConWC(ByVal arrRet, ByVal iPos)	
	If iPos = 0 Then
		frm1.txtWorkCtCd1.value = arrRet(0) 
		frm1.txtWorkCtNm1.value = arrRet(1)
	ElseIf iPos = 1 Then
		frm1.txtWorkCtCd2.value = arrRet(0) 
		frm1.txtWorkCtNm2.value = arrRet(1)
	End If
End Function


'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : SetItemCd Popup에서 return된 값 
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(ByVal arrRet, ByVal iPos)	
	If iPos = 0 Then
		frm1.txtItemCd1.value = arrRet(0) 
		frm1.txtItemNm1.value = arrRet(1)
	ElseIf iPos = 1 Then
		frm1.txtItemCd2.value = arrRet(0) 
		frm1.txtItemNm2.value = arrRet(1)
	End If
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출고요청서(품목별)</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>		
								<TR>	
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="x2xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="공장명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>필요일</TD>
									<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4317oa1_txtReqDt1_txtReqDt1.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p4317oa1_txtReqDt2_txtReqDt2.js'></script>	
								</TR>					
								<TR>
									<TD CLASS="TD5" NOWRAP>품목코드</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd1.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40 MAXLENGTH=40 tag="x4" ALT="품목코드">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd2.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=40 MAXLENGTH=40 tag="x4" ALT="품목코드"></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtWorkCtCd1" SIZE=10 MAXLENGTH=7 tag="x1xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWorkCtCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC frm1.txtWorkCtCd1.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtWorkCtNm1" SIZE=40 MAXLENGTH=40 tag="x4" ALT="작업장">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtWorkCtCd2" SIZE=10 MAXLENGTH=7 tag="x1xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWorkCtCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC frm1.txtWorkCtCd2.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtWorkCtNm2" SIZE=40 MAXLENGTH=40 tag="x4" ALT="작업장"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
	    <TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				  <TD WIDTH = 10></TD>
		          <TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON></TD>		
	            </TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
