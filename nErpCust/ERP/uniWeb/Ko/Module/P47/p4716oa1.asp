
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : production
'*  2. Function Name        : 
'*  3. Program ID           : p4716oa1
'*  4. Program Name         : (p)자원소비실적출력 
'*  5. Program Desc         :  P4716OA1.EBR (오더별) 사용 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001.12.12
'*  9. Modifier (First)     : Jaehyun Chen
'* 10. Modifier (Last)      : Jaehyun Chen
'* 11. Comment              :
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT> 
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************


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

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim startDate
Dim EndDate
Dim strYear
Dim strMonth
Dim strDay
Dim BaseDate

BaseDate = "<%=GetsvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)	    	'☆: 초기화면에 뿌려지는 시작 날짜 
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")	   '☆: 초기화면에 뿌려지는 마지막 날짜 

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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
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
	frm1.txtStartDt.Text = StartDate
	frm1.txtEndDt.Text = EndDate
	
End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()
<%

%>
        
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "P","NOCOOKIE","OA") %>
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
    
	Call ggoOper.FormatField(Document, "x",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitComboBox	
	Call SetDefaultVal
    Call InitVariables		'⊙: Initializes local global variables
   	
    Call SetToolbar("10000000000011")										'⊙: 버튼 툴바 제어 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtOrderNo.focus 
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


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)									'☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()                                         '☜: Protect system from crashing
    Call parent.FncPrint()
End Function

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
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	Dim var8
	Dim var9
	Dim var10
	Dim var11
	Dim var12
	
	dim strUrl
	dim arrParam, arrField, arrHeader
	Dim strEbrFile
	Dim objName
	
	If frm1.txtFromWcCd.value = "" Then
		frm1.txtFromWcNm.value = "" 
	End If
	
	If frm1.txtToWcCd.value = "" Then
		frm1.txtToWcNm.value = "" 
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	
	
    Call BtnDisabled(1)	
	
	If Not chkfield(Document, "x") Then									'⊙: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	
	If Trim(frm1.txtPlantCd.value)= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))	
	
	var2 = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	var3 = UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If Trim(frm1.txtOrderNo.value) = "" Then
		var4 = "0"
	Else
		var4 = Trim(frm1.txtOrderNo.value)  
	End If
	
	If Trim(frm1.txtOrderNo1.value) = "" Then
		var5 = "zzzzzzzzzzzzzzzz"
	Else
		var5 = Trim(frm1.txtOrderNo1.value)
	End If
	
	If Trim(frm1.txtFromItemCd.value) = "" Then
		var6 = "0"
	Else
		var6 = Trim(frm1.txtFromItemCd.value)  
	End If
	
	If Trim(frm1.txtToItemCd.value) = "" Then
		var7 = "zzzzzzzzzzzzzzzzzz"
	Else
		var7 = Trim(frm1.txtToItemCd.value)
	End If
	
	
	If Trim(frm1.txtFromWcCd.value) = "" Then
		var8 = "0"		
	Else
		var8 = Trim(frm1.txtFromWcCd.value)
	End If
	
	If Trim(frm1.txtToWcCd.value) = "" Then
		var9 = "zzzzzzz"	
	Else
		var9 = Trim(frm1.txtToWcCd.value)
	End If
	
	If Trim(frm1.txtTrackingNo1.value) = "" Then
		var10 = "!"		
	Else
		var10 = Trim(frm1.txtTrackingNo1.value)
	End If
	
	If Trim(frm1.txtTrackingNo2.value) = "" Then
		var11 = "zzzzzzzzz"	
	Else
		var11 = Trim(frm1.txtTrackingNo2.value)
	End If

	strUrl = strUrl & "plant_cd|" & var1 
	strUrl = strUrl & "|from_consumed_dt|" & var2 
	strUrl = strUrl & "|to_consumed_dt|" & var3
	strUrl = strUrl & "|fr_prod_order_no|" & var4 
	strUrl = strUrl & "|to_prod_order_no|" & var5 
	strUrl = strUrl & "|from_item_cd|" & var6 
	strUrl = strUrl & "|to_item_cd|" & var7 
	strUrl = strUrl & "|from_wc_cd|" & var8 
	strUrl = strUrl & "|to_wc_cd|" & var9 
	strUrl = strUrl & "|fr_tracking_no|" & var10
	strUrl = strUrl & "|to_tracking_no|" & var11 


'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------
	strEbrFile = "p4716oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
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
'On Error Resume Next                                                    '☜: Protect system from crashing
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	Dim var8
	Dim var9
	Dim var10
	Dim var11
	
	dim strUrl
	dim arrParam, arrField, arrHeader
	Dim strEbrFile
	Dim objName
	
	If frm1.txtFromWcCd.value = "" Then
		frm1.txtFromWcNm.value = "" 
	End If
	
	If frm1.txtToWcCd.value = "" Then
		frm1.txtToWcNm.value = "" 
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then									'⊙: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	
	If Trim(frm1.txtPlantCd.value)= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))	
	
	var2 = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	var3 = UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If Trim(frm1.txtOrderNo.value) = "" Then
		var4 = "0"
	Else
		var4 = Trim(frm1.txtOrderNo.value  )
	End If
	
	If Trim(frm1.txtOrderNo1.value) = "" Then
		var5 = "zzzzzzzzzzzzzzzz"
	Else
		var5 = frm1.txtOrderNo1.value
	End If
	
	If Trim(frm1.txtFromItemCd.value) = "" Then
		var6 = "0"
	Else
		var6 = Trim(frm1.txtFromItemCd.value  )
	End If
	
	If Trim(frm1.txtToItemCd.value) = "" Then
		var7 = "zzzzzzzzzzzzzzzzzz"
	Else
		var7 = Trim(frm1.txtToItemCd.value)
	End If
	
	
	If Trim(frm1.txtFromWcCd.value) = "" Then
		var8 = "0"		
	Else
		var8 = Trim(frm1.txtFromWcCd.value)
	End If
	
	If Trim(frm1.txtToWcCd.value) = "" Then
		var9 = "zzzzzzz"	
	Else
		var9 = Trim(frm1.txtToWcCd.value)
	End If
	
	If Trim(frm1.txtTrackingNo1.value) = "" Then
		var10 = "!"		
	Else
		var10 = Trim(frm1.txtTrackingNo1.value)
	End If
	
	If Trim(frm1.txtTrackingNo2.value) = "" Then
		var11 = "zzzzzzzzz"	
	Else
		var11 = Trim(frm1.txtTrackingNo2.value)
	End If
	
	strUrl = strUrl & "plant_cd|" & var1 
	strUrl = strUrl & "|from_consumed_dt|" & var2 
	strUrl = strUrl & "|to_consumed_dt|" & var3
	strUrl = strUrl & "|fr_prod_order_no|" & var4 
	strUrl = strUrl & "|to_prod_order_no|" & var5 
	strUrl = strUrl & "|from_item_cd|" & var6 
	strUrl = strUrl & "|to_item_cd|" & var7 
	strUrl = strUrl & "|from_wc_cd|" & var8 
	strUrl = strUrl & "|to_wc_cd|" & var9 
	strUrl = strUrl & "|fr_tracking_no|" & var10
	strUrl = strUrl & "|to_tracking_no|" & var11 
	
	strEbrFile = "p4716oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
	On Error Resume Next        
End Function



'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
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

	arrParam(0) = "공장팝업"		' 팝업 명칭 
	arrParam(1) = "B_PLANT"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "공장"			' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"			' Field명(0)
    arrField(1) = "PLANT_NM"			' Field명(1)
    
    arrHeader(0) = "공장"			' Header명(0)
    arrHeader(1) = "공장명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenFrProdOrderNo()  -------------------------------------------------
'	Name : OpenFrProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFrProdOrderNo()

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtOrderNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""

	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetFrProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOrderNo.focus
	
End Function

'------------------------------------------  OpenToProdOrderNo()  -------------------------------------------------
'	Name : OpenToProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenToProdOrderNo()

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtOrderNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtOrderNo1.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetToProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOrderNo1.focus
	
End Function

'------------------------------------------  OpenFromWcCd()  -------------------------------------------------
'	Name : OpenFromWcCd()
'	Description : WcPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromWcCd()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = frm1.txtFromWcCd.value  
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 
				 
	arrParam(5) = "작업장"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"		
    
    arrHeader(0) = "작업장"		
    arrHeader(1) = "작업장명"		
    arrHeader(2) = "작업장타입"		
    arrHeader(3) = "작업장담당자"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetFromWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtFromWcCd.focus
	
End Function
'------------------------------------------  OpenToWcCd()  -------------------------------------------------
'	Name : OpenToWcCd()
'	Description : WcPopup
'---------------------------------------------------------------------------------------------------------
Function OpenToWcCd()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = frm1.txtToWcCd.value  
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 
				  
	arrParam(5) = "작업장"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"	
       
    arrHeader(0) = "작업장"		
    arrHeader(1) = "작업장명"		
    arrHeader(2) = "작업장타입"		
    arrHeader(3) = "작업장담당자"
       
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetToWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtToWcCd.focus
	
End Function
'------------------------------------------  OpenFromItemCd()  -------------------------------------------------
'	Name : OpenFromItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenFromItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtfromItemCd.Value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
    iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetFromItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtfromItemCd.focus

End Function
'------------------------------------------  OpenToItemCd()  -------------------------------------------------
'	Name : OpenToItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenToItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtToItemCd.value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"   
	
	iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetToItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtToItemCd.focus

End Function

'--------------------------------------  OpenTrackingInfo1()  ------------------------------------------
'	Name : OpenTrackingInfo1()
'	Description : From OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo1()
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo1.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtStartDt.Text
	arrParam(4) = frm1.txtEndDt.Text	
	
	iCalledAspName = AskPRAspName("p4600pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4600pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtTrackingNo1.value = arrRet(0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo1.focus

End Function

'--------------------------------------  OpenTrackingInfo2()  ------------------------------------------
'	Name : OpenTrackingInfo2()
'	Description : To OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo2()
    If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo2.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtStartDt.Text
	arrParam(4) = frm1.txtEndDt.Text	
	
	iCalledAspName = AskPRAspName("p4600pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4600pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtTrackingNo2.value = arrRet(0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo2.focus
	
End Function

Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function
Function SetFrProdOrderNo(ByVal arrRet)
	frm1.txtOrderNo.value = arrRet(0) 
End Function

Function SetToProdOrderNo(ByVal arrRet)
	frm1.txtOrderNo1.value = arrRet(0) 
End Function
Function SetFromWcCd(ByVal arrRet)
	frm1.txtFromWcCd.value = arrRet(0)
	frm1.txtFromWcNm.value = arrRet(1)  
End Function
Function SetToWcCd(ByVal arrRet)
	frm1.txtToWcCd.value = arrRet(0)
	frm1.txtToWcNm.value = arrRet(1)  
End Function
Function SetFromItemCd(ByVal arrRet)
	frm1.txtFromItemCd.value = arrRet(0)
	frm1.txtFromItemNm.value = arrRet(1)  
End Function
Function SetToItemCd(ByVal arrRet)
	frm1.txtToItemCd.value = arrRet(0)
	frm1.txtToItemNm.value = arrRet(1)  
End Function

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtEndDt.Focus
    End If
End Sub

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
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>자원소비출력(오더별)</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" colspan="2">
			<TABLE CLASS="BasicTB" CELLSPACING=0 >	
	    		<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="x2xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="공장명">
									</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>제조오더번호</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtOrderNo" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFrProdOrderNo()">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtOrderNo1" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToProdOrderNo()">
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="품목명">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="품목명">&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromWcCd" SIZE=7 MAXLENGTH=7 tag="x1xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromWcNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="작업장명">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToWcCd" SIZE=7 MAXLENGTH=7 tag="x1xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToWcCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToWcNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="작업장명">&nbsp;
									</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtTrackingNo1" SIZE=25 MAXLENGTH=25 tag="x1xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo1">&nbsp;~&nbsp;
									<INPUT TYPE=TEXT NAME="txtTrackingNo2" SIZE=25 MAXLENGTH=25 tag="x1xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo2">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>자원소비일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4716oa1_I243214565_txtStartDt.js'></script>
										&nbsp;~&nbsp; 
										<script language =javascript src='./js/p4716oa1_I707943058_txtEndDt.js'></script>								
									</TD>
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
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
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
