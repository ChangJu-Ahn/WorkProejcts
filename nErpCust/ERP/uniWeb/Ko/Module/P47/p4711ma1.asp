<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711ma1
'*  4. Program Name         : Resource Consumption (Batch)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->							<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우 -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'======================================================================================================select * from b_message====-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_SHIFT		= "p4711mb1.asp"										 '☆: 비지니스 로직 ASP명 
Const BIZ_EXECUTE_ID	= "p4711mb2.asp"
Const BIZ_CANCEL_ID		= "p4711mb3.asp"
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim strDate 
Dim StartDate 
Dim strYear
Dim strMonth
Dim strDay
Dim BaseDate

BaseDate = "<%=GetsvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

Dim IsOpenPop
Dim lgShiftCnt
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
    IsOpenPop = False														'☆: 사용자 변수 초기화 
End Sub

'=============================== 2.1.2 \fTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029() 
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
End Sub 

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtReportDtFrom.text = StartDate
	frm1.txtReportDtTo.text   = StrDate
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Dim i

	For i = lgShiftCnt To 1 Step -1
		frm1.cboShiftCdFrom.remove(i)
		frm1.cboShiftCdTo.remove(i)  
	Next

    Dim strVal
	
	strVal = BIZ_PGM_SHIFT & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'☜: 조회 조건 데이타 
	
    Call RunMyBizASP(MyBizASP, strVal)
	
End Sub

'==========================================  2.2.6 InitStatusCombo()  =======================================
'	Name : InitStatusCombo()
'	Description : Combo Display
'========================================================================================================= 
Sub InitStatusCombo()
	Call SetCombo(frm1.cboStatus, "R", "실행됨")
	Call SetCombo(frm1.cboStatus, "C", "취소됨")		'⊙: InitCombo 에서 해야 되는데 임시로 넣은 것임 
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

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
   	arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
    
    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBatchRunNo()  -------------------------------------------------
'	Name : OpenBatchRunNo()
'	Description : Batch Run No. PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBatchRunNo()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtBatchRunNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))
	arrParam(1) = Trim(frm1.txtPlantNm.value)
	arrParam(2) = UCase(Trim(frm1.txtBatchRunNo.value))

	iCalledAspName = AskPRAspName("p4711pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4711pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBatchRunNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBatchRunNo.focus
	
End Function

'------------------------------------------  OpenProdOrderNoFrom()  -------------------------------------------------
'	Name : OpenProdOrderNoFrom()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNoFrom()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	If IsOpenPop = True Or UCase(frm1.txtProdtOrderNoFrom.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrderNoFrom.value)
	arrParam(6) = ""
	arrParam(7) = Trim(frm1.txtItemCdFrom.value)
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
		Call SetProdOrderNoFrom(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrderNoFrom.focus
		
End Function

'------------------------------------------  OpenProdOrderNoTo()  -------------------------------------------------
'	Name : OpenProdOrderNoTo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNoTo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	If IsOpenPop = True Or UCase(frm1.txtProdtOrderNoTo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrderNoTo.value)
	arrParam(6) = ""
	arrParam(7) = Trim(frm1.txtItemCdTo.value)
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
		Call SetProdOrderNoTo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrderNoTo.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCdFrom()
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
	arrParam(1) = frm1.txtItemCdFrom.Value		' Item Code
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
		Call SetItemCdFrom(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCdFrom.focus

End Function
'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemCdTo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCdTo()
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
	arrParam(1) = frm1.txtItemCdTo.value		' Item Code
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
		Call SetItemCdTo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCdTo.focus

End Function

'------------------------------------------  OpenWcCdFrom()  ------------------------------------------------
'	Name : OpenWcCdFrom()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCdFrom()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWCCdFrom.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNmFrom.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")		' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCdFrom(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCdFrom.focus
	
End Function

'------------------------------------------  OpenWcCdTo()  ------------------------------------------------
'	Name : OpenWcCdTo()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCdTo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWCCdTo.Value)								' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNmTo.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCdTo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCdTo.focus
	
End Function

'------------------------------------------  OpenErrorRef()  -------------------------------------------------
'	Name : OpenErrorRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenErrorRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtBatchRunNo.value= "" Then
		Call DisplayMsgBox("971012","X", "이력번호","X")
		frm1.txtBatchRunNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	arrParam(0) = Trim(UCase(frm1.txtPlantCd.value))	'☆: 조회 조건 데이타 
	arrParam(1) = Trim(frm1.txtPlantNm.value)			'☆: 조회 조건 데이타 
	arrParam(2) = Trim(frm1.txtBatchRunNo.value)		'☆: 조회 조건 데이타 
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4711ra2")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4711ra2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	Call InitComboBox
End Function

'------------------------------------------  SetBatchRunNo()  --------------------------------------------------
'	Name : SetBatchRunNo()
'	Description : ResourceGroup Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBatchRunNo(byval arrRet)
	frm1.txtBatchRunNo.Value = arrRet(0)
	frm1.cboStatus.Value	 = arrRet(1)
	frm1.txtSuccessCnt.Value = arrRet(2)
	frm1.txtErrorCnt.Value	 = arrRet(3)
End Function

'------------------------------------------  SetFrProdOrderNo()  -------------------------------------------------
'	Name : SetFrProdOrderNo()
'	Description : ProdOrderNo Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNoFrom(ByVal arrRet)
	frm1.txtProdtOrderNoFrom.value = arrRet(0) 
End Function

'------------------------------------------  SetProdOrderNoTo()  -------------------------------------------------
'	Name : SetProdOrderNoTo()
'	Description : ProdOrderNo Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNoTo(ByVal arrRet)
	frm1.txtProdtOrderNoTo.value = arrRet(0) 
End Function

'------------------------------------------  SetItemCdFrom()  -------------------------------------------------
'	Name : SetItemCdFrom()
'	Description : ProdOrderNo Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCdFrom(ByVal arrRet)
	frm1.txtItemCdFrom.value = arrRet(0)
	frm1.txtItemNmFrom.value = arrRet(1)  
End Function

'------------------------------------------  SetItemCdTo()  -------------------------------------------------
'	Name : SetItemCdTo()
'	Description : ProdOrderNo Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCdTo(ByVal arrRet)
	frm1.txtItemCdTo.value = arrRet(0)
	frm1.txtItemNmTo.value = arrRet(1)  
End Function

'------------------------------------------  SetWcCdFrom()  -------------------------------------------------
'	Name : SetWcCdFrom()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCdFrom(byval arrRet)
	frm1.txtWCCdFrom.Value    = arrRet(0)		
	frm1.txtWCNmFrom.Value    = arrRet(1)		
End Function

'------------------------------------------  SetWcCdTo()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCdTo(byval arrRet)
	frm1.txtWCCdTo.Value    = arrRet(0)		
	frm1.txtWCNmTo.Value    = arrRet(1)		
End Function

Sub txtPlantCd_OnChange()
    If frm1.txtPlantCd.value <> "" Then
		Call InitComboBox	
	End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'++++++++++++++++++++++++++++++++++++++++++  2.5.2 Execute  +++++++++++++++++++++++++++++++++++++++
'        Name : Execute()
'        Description : MRP 전개 Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function Execute()

	Dim strVal
		
    Err.Clear															'☜: Protect system from crashing
    Execute = False														'⊙: Processing is NG

    If Not chkField(Document, "1") Then									'⊙: Check contents area
       Exit Function
    End If
	
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If ValidDateCheck(frm1.txtReportDtFrom, frm1.txtReportDtTo) = False Then Exit Function
        
    Call LayerShowHide(1)
    
	strVal = BIZ_EXECUTE_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtProdtOrderNoFrom=" & Trim(frm1.txtProdtOrderNoFrom.value)
	strVal = strVal & "&txtProdtOrderNoTo=" & Trim(frm1.txtProdtOrderNoTo.value)
	strVal = strVal & "&txtItemCdFrom=" & Trim(frm1.txtItemCdFrom.value)
	strVal = strVal & "&txtItemCdTo=" & Trim(frm1.txtItemCdTo.value)
	strVal = strVal & "&txtWcCdFrom=" & Trim(frm1.txtWcCdFrom.value)
	strVal = strVal & "&txtWcCdTo=" & Trim(frm1.txtWcCdTo.value)
	strVal = strVal & "&cboShiftCdFrom=" & Trim(frm1.cboShiftCdFrom.value)
	strVal = strVal & "&cboShiftCdTo=" & Trim(frm1.cboShiftCdTo.value)
	strVal = strVal & "&txtReportDtFrom=" & Trim(frm1.txtReportDtFrom.text)
	strVal = strVal & "&txtReportDtTo=" & Trim(frm1.txtReportDtTo.text)
	
	Call RunMyBizASP(MyBizASP, strVal)
	
    Execute = True 
            
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5.2 Cancel()  +++++++++++++++++++++++++++++++++++++++
'        Name : Cancel()
'        Description : MRP 전개 Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function Cancel()

	Dim strVal
		
    Err.Clear															'☜: Protect system from crashing
    Cancel = False														'⊙: Processing is NG

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtBatchRunNo.value= "" Then
		Call DisplayMsgBox("971012","X", "이력번호","X")
		frm1.txtBatchRunNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call LayerShowHide(1)
    
	strVal = BIZ_CANCEL_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtBatchRunNo=" & Trim(frm1.txtBatchRunNo.value)
    Call RunMyBizASP(MyBizASP, strVal)
	
    Cancel = True 
            
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
    Call SetDefaultVal
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call SetDefaultVal																	'⊙: Initializes local global variables
    Call InitVariables
	Call InitStatusCombo
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm	
		Call InitComboBox()
		frm1.txtProdtOrderNoFrom.focus 
	ELSE
		frm1.txtPlantCd.focus 
	End If   
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtReportDtFrom_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDtFrom_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportDtFrom.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtReportDtFrom.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtReportDtTo_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDtTo_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportDtTo.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtReportDtTo.Focus
	End If 
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
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
	On Error Resume Next       
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
	On Error Resume Next                                                    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
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
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 

End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()															'☆: 저장 성공후 실행 로직 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>자원소비등록(Batch)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>이력번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBatchRunNo" SIZE=18 MAXLENGTH=18 tag="11XXXU"  ALT="이력번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBatchRunNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBatchRunNo()" >&nbsp;<SELECT NAME="cboStatus" ALT="Status" STYLE="Width: 98px;" tag="14"></SELECT></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>감안된실적수</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtSuccessCnt" SIZE=16 MAXLENGTH=16 tag="14xxxU" ALT="감안된실적수"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>ERROR수</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtErrorCnt" SIZE=16 MAXLENGTH=16 tag="14xxxU" ALT="ERROR수"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>제조오더번호</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtProdtOrderNoFrom" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNoFrom()">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtProdtOrderNoTo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNoTo()">
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCdFrom" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCdFrom" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCdFrom()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmFrom" SIZE=40 MAXLENGTH=40 tag="14" ALT="품목명">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCdTo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCdTo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmTo" SIZE=40 MAXLENGTH=40 tag="14" ALT="품목명">&nbsp;
								</TD>
							</TR>
							<TR>								
								<TD CLASS=TD5 NOWRAP>작업장</TD>
								<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtWCCdFrom" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCdFrom()"> <INPUT TYPE=TEXT  NAME="txtWCNmFrom" SIZE=40 MAXLENGTH=40 tag="14">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtWCCdTo" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCdTo()"> <INPUT TYPE=TEXT  NAME="txtWCNmTo" SIZE=40 MAXLENGTH=40 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Shift</TD>
								<TD CLASS=TD6 NOWRAP>
								<SELECT NAME="cboShiftCdFrom" ALT="시작 Shift" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
								&nbsp;~&nbsp;
								<SELECT NAME="cboShiftCdTo" ALT="종료 Shift" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>생산실적일</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4711ma1_I498841530_txtReportDtFrom.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p4711ma1_I689011383_txtReportDtTo.js'></script>
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>					
					<TD><BUTTON NAME="btnExecute" CLASS="CLSMBTN" Flag=1 onclick="Execute()">실행</BUTTON>&nbsp;<BUTTON NAME="btnCancel" CLASS="CLSMBTN" Flag=1 onclick="Cancel()">취소</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
