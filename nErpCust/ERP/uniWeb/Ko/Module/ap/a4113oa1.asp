
<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : A/P
'*  3. Program ID           : a4113ma1
'*  4. Program Name         : 채무상세출력 
'*  5. Program Desc         : Report of A/P Detail
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001.01.29
'*  8. Modified date(Last)  : 2004.01.12
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Kim Chang Jin
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance


'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'⊙: 비지니스 로직 ASP명 
'Const BIZ_PGM_ID = "f5111mb1.asp"			'☆: 비지니스 로직 ASP명 


'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
Dim lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

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

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    '---- Coding part--------------------------------------------------------------------    
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
	Dim strYear, strMonth, strDay, EndDate, StartDate
<%  
	Dim dtToday
	dtToday = GetSvrDate
%>	
	Call ExtractDateFrom("<%=dtToday%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)

	frm1.txtDateMid.Text = EndDate
	frm1.txtDateFr.Text  = StartDate
	frm1.txtDateTo.Text  = EndDate

	frm1.hOrgChangeId.value = parent.gChangeOrgId
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","OA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","OA") %>
End Sub
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

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 0, 5
			arrParam(0) = "사업장코드 팝업"								' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"										' Field명(0)
			arrField(1) = "BIZ_AREA_NM"										' Field명(1)
    
			arrHeader(0) = "사업장코드"									' Header명(0)
			arrHeader(1) = "사업장명"									' Header명(1)
			
		Case 3, 4
			
			If frm1.rdoPayBp.checked = False then
				arrParam(0) = "공급처팝업"
				arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AP B " 
				arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.DEAL_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AP_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
				IF frm1.txtDateFr.Text<>"" THEN 	arrParam(1) = arrParam(1) & " AND AP_DT >= " & FilterVar(UNIConvDate(frm1.txtDateFr.Text), "''", "S") & ""
				IF frm1.txtDateTo.Text<>"" THEN		arrParam(1) = arrParam(1) & " AND AP_DT <= " & FilterVar(UNIConvDate(frm1.txtDateTo.Text), "''", "S") & ""
				arrParam(1) = arrParam(1) & ") TMP"
			
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "공급처"			
	
				arrField(0) = "TMP.BP_CD"	
				arrField(1) = "TMP.BP_NM"	

				arrHeader(0) = "공급처"                                     ' Header명(0)
				arrHeader(1) = "공급처명"                                   ' Header명(1)
			
			ELse
				arrParam(0) = "지급처팝업"
				arrParam(1) = "(SELECT DISTINCT A.BP_CD,A.BP_NM FROM B_BIZ_PARTNER A, A_OPEN_AP B " 
				arrParam(1) = arrParam(1) & "WHERE  A.BP_CD=B.PAY_BP_CD AND B.CONF_FG = " & FilterVar("C", "''", "S") & "  AND B.AP_STS=" & FilterVar("O", "''", "S") & "  AND B.BAL_AMT <> 0" 
				IF frm1.txtDateFr.Text<>"" THEN 	arrParam(1) = arrParam(1) & " AND AP_DT >= " & FilterVar(UNIConvDate(frm1.txtDateFr.Text), "''", "S") & ""
				IF frm1.txtDateTo.Text<>"" THEN arrParam(1) = arrParam(1) & " AND AP_DT <= " & FilterVar(UNIConvDate(frm1.txtDateTo.Text), "''", "S") & ""
				arrParam(1) = arrParam(1) & ") TMP"
			
				arrParam(2) = strCode
				arrParam(3) = ""
				arrParam(4) = ""
				arrParam(5) = "지급처"			
	
				arrField(0) = "TMP.BP_CD"	
				arrField(1) = "TMP.BP_NM"	

				arrHeader(0) = "지급처"                                     ' Header명(0)
				arrHeader(1) = "지급처명"                                   ' Header명(1)
			End IF

		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
			Select Case iWhere
			Case 0
				frm1.txtBizAreaCd.focus
			Case 3
				frm1.txtDealBpCdFr.focus
			Case 4
				frm1.txtDealBpCdTo.focus
			Case 5
				frm1.txtBizAreaCd1.focus
			Case Else
		End Select
		Exit Function
	Else
		Select Case iWhere
			Case 0
				frm1.txtBizAreaCd.value = arrRet(0)
				frm1.txtBizAreaNm.value = arrRet(1)
				frm1.txtBizAreaCd.focus
			Case 3
				frm1.txtDealBpCdFr.value = arrRet(0)
				frm1.txtDealBpNmFr.value = arrRet(1)
				frm1.txtDealBpCdFr.focus
			Case 4
				frm1.txtDealBpCdTo.value = arrRet(0)
				frm1.txtDealBpNmTo.value = arrRet(1)
				frm1.txtDealBpCdTo.focus
			Case 5
			    frm1.txtBizAreaCd1.value = arrRet(0)
			    frm1.txtBizAreaNm1.value = arrRet(1)
				frm1.txtBizAreaCd1.focus
			Case Else
		End Select
	End If	

End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 


'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Account Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 



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

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029()                           '⊙: Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")			'⊙: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")			'⊙: 조건에 맞는 Field locking
    
    Call InitVariables()                            '⊙: Initializes local global Variables
    Call SetDefaultVal()
    
    Call SetToolbar("1000000000000111")				'⊙: 버튼 툴바 제어 
	frm1.txtBizAreaCd.focus

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

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 



'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Event 처리	
'********************************************************************************************************* 
Function rdoDelBp_OnClick() 
	if frm1.rdoDelBp.checked = True then
		BP_Cd.innerHTML = "공급처"
	end if
End Function
Function rdoPayBp_OnClick() 
	if frm1.rdoPayBp.checked = True then
		BP_Cd.innerHTML = "지급처"
	end if
End Function
'======================================================================================================
'   Event Name : txtDateMid_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtDateMid_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateMid.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtDateMid.Focus
    End If
End Sub

Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateFr.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtDateFr.Focus
    End If
End Sub

Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateTo.Action = 7
    	Call SetFocusToDocument("M")
		Frm1.txtDateTo.Focus
    End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Function SetPrintCond(StrEbrFile, StrUrl)

	Dim	VarBizAreaCd, VarBizAreaCd1, VarDateMid, VarDateFr, VarDateTo, VarDealBpCdFr, VarDealBpCdTo, VarBpLabel
	Dim	strAuthCond	
		
	SetPrintCond = False

	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		Exit Function
	End If

	If CompareDateByFormat(frm1.txtDateFr.text,frm1.txtDateTo.text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
	    	               "970025",frm1.txtDateFr.UserDefinedFormat,parent.gComDateType, true) = False Then
		frm1.txtDateFr.focus
		Exit Function
	End If
		
	frm1.txtDealBpCdFr.value = UCase(Trim(frm1.txtDealBpCdFr.value))
	frm1.txtDealBpCdTo.value = UCase(Trim(frm1.txtDealBpCdTo.value))
		
	If frm1.txtDealBpCdFr.value <> "" And frm1.txtDealBpCdTo.value <> "" Then
		If frm1.txtDealBpCdFr.value > frm1.txtDealBpCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtDealBpCdFr.Alt, frm1.txtDealBpCdTo.Alt)
			frm1.txtDealBpCdFr.focus
			Exit Function
		End If
	End If

	StrEbrFile = "a4113ma1"
	VarBpLabel= "DEAL_BP_CD"
		
	If frm1.txtBizAreaCd.value = "" then
		frm1.txtBizAreaNm.value = ""
		VarBizAreaCd = ""
	else 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
	end if
		
	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		VarBizAreaCd1 = "ZZZZZZZZZZ"
	else 
		VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
	end if
		
	VarDateMid		= UNIConvDate(frm1.txtDateMid.Text)
	VarDateFr		= UNIConvDate(frm1.txtDateFr.Text)
	VarDateTo		= UNIConvDate(frm1.txtDateTo.Text)
	VarDealBpCdFr	= UCase(Trim(frm1.txtDealBpCdFr.value))
	VarDealBpCdTo	= UCase(Trim(frm1.txtDealBpCdTo.value))
		
	If VarDealBpCdFr = "" Then VarDealBpCdFr = ""
	If VarDealBpCdTo = "" Then VarDealBpCdTo = "ZZZZZZZZZZ"
	If frm1.rdoPayBp.checked = True then VarBpLabel = "PAY_BP_CD"


	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_OPEN_AP.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_OPEN_AP.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_OPEN_AP.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_OPEN_AP.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	StrUrl = StrUrl & "BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	StrUrl = StrUrl & "|DateMid|"		& VarDateMid
	StrUrl = StrUrl & "|DateFr|"		& VarDateFr
	StrUrl = StrUrl & "|DateTo|"		& VarDateTo
	StrUrl = StrUrl & "|DealBpCdFr|"	& VarDealBpCdFr
	StrUrl = StrUrl & "|DealBpCdTo|"	& VarDealBpCdTo
	StrUrl = StrUrl & "|BpLabel|"		& VarBpLabel

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond
	
	SetPrintCond = True

End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrEbrFile, StrUrl,objName

	On Error Resume Next                                                    '☜: Protect system from crashing
	Err.Clear 


	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
		
	objName = AskEBDocumentName(StrEbrFile,"ebr")
		
	Call FncEBRPrint(EBAction,objName,StrUrl)
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	Dim StrEbrFile, StrUrl,objName

	On Error Resume Next                                                    '☜: Protect system from crashing
	Err.Clear 


	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
		
	objName = AskEBDocumentName(StrEbrFile,"ebr")
		
	Call FncEBRPreview(objName,StrUrl)
End Function

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

'Function FncQuery() 
'End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
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
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                        '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                        '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
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
'********************************************************************************************************** 

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
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete()
	On Error Resume Next
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>채무일자</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=시작채무일자 id=fpDateFr></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=종료채무일자 id=fpDateTo></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>기준일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateMid" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=기준일자 id=fpDateMid></OBJECT>');</SCRIPT>&nbsp;&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래처기준</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=RADIO CLASS="RADIO" NAME="rdoBpLabel" ID="rdoDealBp" VALUE="S" TAG="11" Checked><LABEL FOR="rdoReport1">공급처</LABEL>&nbsp;&nbsp
													 <INPUT TYPE=RADIO CLASS="RADIO" NAME="rdoBpLabel" ID="rdoPayBp" VALUE="D" TAG="11"><LABEL FOR="rdoReport2">지급처</LABEL>
								</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" ID= BP_Cd NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCdFr" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="시작공급처코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDealBpCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDealBpCdFr.Value, 3)">
													   <INPUT TYPE="Text" NAME="txtDealBpNmFr" SIZE=25 tag="14X" ALT="시작공급처명">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDealBpCdTo" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="종료공급처코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDealBpCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDealBpCdTo.Value, 4)">
													   <INPUT TYPE="Text" NAME="txtDealBpNmTo" SIZE=25 tag="14X" ALT="종료공급처명">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)">
													   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명">
								</TD>
							</TR>
							<TR>
	                            <TD CLASS="TD5" NOWRAP></TD>
                                <TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,5)">
                                                       <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명">
                                </TD>
                            </TR>
						</TABLE>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>	
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>	
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>

