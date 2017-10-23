

<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2111ma1
'*  4. Program Name         : 예산실적출력 
'*  5. Program Desc         : Report of Budget Result
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001.01.06
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
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


'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim strSvrDate

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
	strSvrDate = "<%=GetSvrDate%>"
	
	frm1.hOrgChangeId.value = parent.gChangeOrgId
'	frm1.fpDateTime1.Text = UNIDateClientFormat(strSvrDate)
'	frm1.fpDateTime2.Text = UNIDateClientFormat(strSvrDate)
	frm1.fpDateTime1.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.fpDateTime2.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.fpDateTime1, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.fpDateTime2, parent.gDateFormat, 2)

	frm1.Rb_Fg1.checked = True
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================

Sub InitComboBox()

End Sub


'==========================================  2.2.7 SetCheckBox()  =======================================
'	Name : SetCheckBox()
'	Description : 체크박스 선택 처리(1개만 선택되도록 함)
'========================================================================================================= 
Function SetCheckBox(objCheckBox)
	Dim idx
	
	For idx = 0 To Document.All.Length - 1
		Select Case Document.All(idx).TagName
		Case "INPUT"
			If UCase(Document.All(idx).Type) = "CHECKBOX" Then
				Document.All(idx).Checked = False
			End If
		End Select
	Next
	
	objCheckBox.Checked = True
End Function


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
'		Case 0, 1
'			arrParam(0) = "부서코드 팝업"								' 팝업 명칭 
'			arrParam(1) = "B_ACCT_DEPT" 									' TABLE 명칭 
'			arrParam(2) = strCode											' Code Condition
'			arrParam(3) = ""												' Name Cindition
'			arrParam(4) = "ORG_CHANGE_ID = '" & parent.gChangeOrgId & "'"			' Where Condition
'			arrParam(5) = "부서코드"									' 조건필드의 라벨 명칭 
'
'			arrField(0) = "DEPT_CD"											' Field명(0)
'			arrField(1) = "DEPT_NM"											' Field명(1)
 '   
'			arrHeader(0) = "부서코드"									' Header명(0)
'			arrHeader(1) = "부서명"										' Header명(1)

		Case 2, 3
			arrParam(0) = "예산코드 팝업"			' 팝업 명칭 
			arrParam(1) = "F_BDG_ACCT"		 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "예산코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "BDG_CD"						' Field명(0)
			arrField(1) = "GP_ACCT_NM"					' Field명(1)
    
			arrHeader(0) = "예산코드"					' Header명(0)
			arrHeader(1) = "예산명"						' Header명(1)

		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
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
	
	Select Case iWhere
'		Case 0		'시작부서코드 
'			frm1.txtDeptCdFr.value = arrRet(0)
'			frm1.txtDeptNmFr.value = arrRet(1)
'		Case 1		'종료부서코드 
'			frm1.txtDeptCdTo.value = arrRet(0)
'			frm1.txtDeptNmTo.value = arrRet(1)
		Case 2		'시작예산코드 
			frm1.txtBdgCdFr.value = arrRet(0)
			frm1.txtBdgNmFr.value = arrRet(1)
		Case 3		'종료예산코드 
			frm1.txtBdgCdTo.value = arrRet(0)
			frm1.txtBdgNmTo.value = arrRet(1)
		Case Else
	End select	

End Function
'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup( ByVal iWhere)
	Dim arrRet
	Dim arrParam(8)
	Dim strYear,strMonth,strDay

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0)	= UniConvDateAToB(frm1.txtDymFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	arrParam(1)	= UniConvDateAToB(frm1.txtDymTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	arrParam(1)	= UNIDateAdd("M", +1, arrParam(1),parent.gServerDateFormat)
	arrParam(1)	= UNIDateAdd("D", -1, arrParam(1),parent.gServerDateFormat)	    

	arrParam(0)	=  UniConvDateAToB(arrParam(0),parent.gServerDateFormat,gDateFormat)
	arrParam(1)	=  UniConvDateAToB(arrParam(1),parent.gServerDateFormat,gDateFormat)

'	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCdFr.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(iWhere,arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept( ByVal iWhere,Byval arrRet)
	Select Case iWhere
		Case 0		'시작부서코드 

			'frm1.hOrgChangeId.value=arrRet(2)
			
			frm1.txtDeptCdFr.value = arrRet(0)
			frm1.txtDeptNmFr.value = arrRet(1)		
		Case 1		'시작부서코드 
			'frm1.hOrgChangeId.value=arrRet(2)
			
			frm1.txtDeptCdTo.value = arrRet(0)
			frm1.txtDeptNmTo.value = arrRet(1)		
	End Select 

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

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
	' 현재 Page의 Form Element들을 Clear한다. 
		
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call SetDefaultVal
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    
    'Call InitSpreadSheet                          '⊙: Setup the Spread Sheet
    Call InitVariables                            '⊙: Initializes local global Variables
    
    '----------  Coding part  -------------------------------------------------------------
	'Call InitComboBox
	
	' [Main Menu ToolBar]의 각 버튼을 [Enable/Disable] 처리하는 부분 
    Call SetToolbar("1000000000001111")				'⊙: 버튼 툴바 제어 
    
	frm1.txtDymFr.focus

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

Sub Radio_Fg_Click()
	If frm1.Rb_Fg1.checked = True Then	'부서별 
		lblTitle1.innerHTML = "예산년월"
		lblHyphen.innerHTML = "~"
		Call ElementVisible(frm1.fpDateTime2, 1)	'Visible
		Call ggoOper.FormatDate(frm1.txtDymFr, parent.gDateFormat, 2)	'년월 
	ElseIf frm1.Rb_Fg2.checked = True Then	'예산코드별 
		lblTitle1.innerHTML = "예산년월"
		lblHyphen.innerHTML = "~"
		Call ElementVisible(frm1.fpDateTime2, 1)	'Visible
		Call ggoOper.FormatDate(frm1.txtDymFr, parent.gDateFormat, 2)	'년월 
	ElseIf frm1.Rb_Fg3.checked = True Then	'년간 
		lblTitle1.innerHTML = "예산년도"
		lblHyphen.innerHTML = ""
		Call ElementVisible(frm1.fpDateTime2, 0)	'InVisible
		Call ggoOper.FormatDate(frm1.txtDymFr, parent.gDateFormat, 3)	'년도 
	End If
End Sub

'======================================================================================================
'   Event Name : 
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtDymFr_DblClick(Button)
    If Button = 1 Then
		frm1.fpDateTime1.Action = 7
		Call SetFocusToDocument("M")	
		frm1.fpDateTime1.Focus       
        
    End If
End Sub

Sub txtDymTo_DblClick(Button)
    If Button = 1 Then
		frm1.fpDateTime2.Action = 7
		Call SetFocusToDocument("M")	
		frm1.fpDateTime2.Focus               
    End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, StrUrl)
	Dim StrFg, VarDeptCdFr, VarDeptCdTo, VarBdgCdFr, VarBdgCdTo, VarDymFr, VarDymTo
	Dim strYear, strMonth, strDay
	Dim strYear1, strMonth1, strDay1

	Dim strAuthCond

	If frm1.Rb_Fg1.checked = True Then	 '부서별 
		StrEbrFile = "f2111ma1a"
	ElseIf frm1.Rb_Fg2.checked = True Then	 '예산코드별 
		StrEbrFile = "f2111ma1b"
	ElseIf frm1.Rb_Fg3.checked = True Then	 '년간 
		StrEbrFile = "f2111ma1c"
	End If

	VarDeptCdFr	= " "
	VarDeptCdTo	= "ZZZZZZZZZZ"
	VarBdgCdFr	= " "
	VarBdgCdTo	= "ZZZZZZZZZZZZZZZZZZ"
	
	Call ExtractDateFrom(frm1.fpDateTime1.Text,frm1.fpDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	VarDymFr = strYear & strMonth
	
	Call ExtractDateFrom(frm1.fpDateTime2.Text,frm1.fpDateTime2.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	VarDymTo = strYear1 & strMonth1

	If frm1.Rb_Fg3.checked = True Then	 '년간인 경우, FromDate의 년도만 사용 
		Call ExtractDateFrom(frm1.fpDateTime2.Text,frm1.fpDateTime2.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	    VarDymTo = strYear1 & strMonth1
				
		VarDymFr = Trim(frm1.fpDateTime1.Text)
		
	End If
	
	If Trim(frm1.txtDeptCdFr.value) <> ""	Then VarDeptCdFr = FilterVar(Trim(frm1.txtDeptCdFr.value),"","SNM")
	If Trim(frm1.txtDeptCdTo.value) <> ""	Then VarDeptCdTo = FilterVar(Trim(frm1.txtDeptCdTo.value),"","SNM")
	If Trim(frm1.txtBdgCdFr.value) <> ""	Then VarBdgCdFr = FilterVar(Trim(frm1.txtBdgCdFr.value),"","SNM")
	If Trim(frm1.txtBdgCdTo.value) <> ""	Then VarBdgCdTo = FilterVar(Trim(frm1.txtBdgCdTo.value),"","SNM")
	
	'-----------------------------------------------------------------------------------
	
	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_BDG.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_BDG.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_BDG.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_BDG.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	StrUrl = StrUrl & "DeptCdFr|"	& VarDeptCdFr
	StrUrl = StrUrl & "|DeptCdTo|"	& VarDeptCdTo
	StrUrl = StrUrl & "|BdgCdFr|"	& VarBdgCdFr
	StrUrl = StrUrl & "|BdgCdTo|"	& VarBdgCdTo
	StrUrl = StrUrl & "|DymFr|"		& VarDymFr
	StrUrl = StrUrl & "|DymTo|"		& VarDymTo
	StrUrl = StrUrl & "|DYear|"		& Trim(VarDymFr) & "__"
	
	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond
	

	'-----------------------------------------------------------------------------------
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
    On Error Resume Next
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile	
	Dim ObjName
    	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtDymFr.Text, frm1.txtDymTo.Text, frm1.txtDymFr.Alt, frm1.txtDymTo.Alt, _
				"970025", frm1.txtDymFr.UserDefinedFormat, parent.gComDateType, true) = False Then
		frm1.txtDymFr.focus											'⊙: GL Date Compare Common Function
		Exit Function
	End if	
	
	
	frm1.txtDeptCdFr.value = Trim(frm1.txtDeptCdFr.value)
	frm1.txtDeptCdTo.value = Trim(frm1.txtDeptCdTo.value)
	If frm1.txtDeptCdFr.value <> "" And frm1.txtDeptCdTo.value <> "" Then
		If frm1.txtDeptCdFr.value > frm1.txtDeptCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtDeptCdFr.Alt, frm1.txtDeptCdTo.Alt)
			frm1.txtDeptCdFr.focus 
			Exit Function
		End If
	End If
	
		
	frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
	frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
	If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus 
			Exit Function
		End If
	End If
		
	Call SetPrintCond(StrEbrFile, StrUrl)
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	

End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	Dim StrFg
    Dim StrUrl, StrUrl2
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile	
	Dim ObjName
        
	If frm1.Rb_Fg1.checked = True Then	 '부서별 
		StrFg = "a"
	ElseIf frm1.Rb_Fg2.checked = True Then	 '예산코드별 
		StrFg = "b"
	ElseIf frm1.Rb_Fg3.checked = True Then	 '년간 
		StrFg = "c"
	End If
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	If StrFg <> "c" Then
		If CompareDateByFormat(frm1.txtDymFr.Text, frm1.txtDymTo.Text, frm1.txtDymFr.Alt, frm1.txtDymTo.Alt, _
					"970025", frm1.txtDymFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtDymFr.focus											'⊙: GL Date Compare Common Function
			Exit Function	
		End if	
	End If
	
	frm1.txtDeptCdFr.value = Trim(frm1.txtDeptCdFr.value)
	frm1.txtDeptCdTo.value = Trim(frm1.txtDeptCdTo.value)
	If frm1.txtDeptCdFr.value <> "" And frm1.txtDeptCdTo.value <> "" Then
		If frm1.txtDeptCdFr.value > frm1.txtDeptCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtDeptCdFr.Alt, frm1.txtDeptCdTo.Alt)
			frm1.txtDeptCdFr.focus 
			Exit Function
		End If
	End If
	
	frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
	frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
	If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus 
			Exit Function
		End If
	End If
	
	Call SetPrintCond(StrEbrFile, StrUrl)
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPreview(ObjName,StrUrl)
		
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
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
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
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
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>출력구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg1 ONCLICK="vbscript:Call Radio_Fg_Click()"><LABEL FOR=Rb_Fg1>부서별</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg2 ONCLICK="vbscript:Call Radio_Fg_Click()"><LABEL FOR=Rb_Fg2>예산코드별</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg3 ONCLICK="vbscript:Call Radio_Fg_Click()"><LABEL FOR=Rb_Fg3>년간</LABEL>&nbsp;
									</TD>
								</TR>
								<TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle1">예산년월</SPAN></TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDymFr" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=시작예산년월 id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen">~</SPAN>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDymTo" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=종료예산년월 id=fpDateTime2></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCdFr" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="시작부서코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup(0)">&nbsp;<INPUT TYPE="Text" NAME="txtDeptNmFr" SIZE=25 tag="14X" ALT="시작부서명">&nbsp;~
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCdTo" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="종료부서코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup(1)">&nbsp;<INPUT TYPE="Text" NAME="txtDeptNmTo" SIZE=25 tag="14X" ALT="종료부서명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>예산코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBdgCdFr" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT ="시작예산코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBdgCdFr.Value, 2)">&nbsp;<INPUT TYPE="Text" NAME="txtBdgNmFr" SIZE=25 tag="14X" ALT="시작예산코드명">&nbsp;~
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBdgCdTo" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT ="종료예산코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBdgCdTo.Value, 3)">&nbsp;<INPUT TYPE="Text" NAME="txtBdgNmTo" SIZE=25 tag="14X" ALT="종료예산코드명">
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TabIndex="-1">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TabIndex="-1">	
</FORM>
</BODY>
</HTML>

