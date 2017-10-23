
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module명          : 회계관리 
'*  2. Function명        : 자산관리 
'*  3. Program ID        : a7117ma1.asp
'*  4. Program 이름      : 고정자산명세표출력 
'*  5. Program 설명      : 고정자산명세표출력 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2000/12/19
'*  8. 최종 수정년월일   : 2004/01/30
'*  9. 최초 작성자       : KIM HEE JUNG
'* 10. 최종 작성자       : U & I (Kim Chang Jin)
'* 11. 전체 comment      : 사업장 조건 추가 
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'##########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance


'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'⊙: 비지니스 로직 ASP명 
'Const BIZ_PGM_ID = "a7120mb1.asp"			'☆: 비지니스 로직 ASP명 


'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns


'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
'Dim lgIntFlgMode               ' Variable is for Operation Status 


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

<!-- #Include file="../../inc/lgvariables.inc" --> 
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

	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	dtToday = "<%=GetSvrDate%>"
	Call ExtractDateFrom(dtToday, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, gDateFormat)

	frm1.fpDateTime1.Text = StartDate
	frm1.fpDateTime2.Text = EndDate

	frm1.Rb_acct.checked = True
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>
End Sub


'==========================================  2.2.7 SetCheckBox()  =======================================
'	Name : SetCheckBox()
'	Description : 감가상각집계표 출력물 체크박스 선택 처리(1개만 선택되도록 함)
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

Function OpenPopUp(Byval strCode, Byval Cond)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case Cond
	Case "FrAcct"
		arrParam(0) = "계정코드팝업"			' 팝업 명칭 
		arrParam(1) = "a_acct"						' TABLE 명칭 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = "acct_type = " & FilterVar("K0", "''", "S") & " "			' Where Condition
		arrParam(5) = "계정코드"				' 조건필드의 라벨 명칭 
	
		arrField(0) = "acct_cd"						' Field명(0)
		arrField(1) = "acct_nm"						' Field명(1)
    
		arrHeader(0) = "계정코드"				' Header명(0)
		arrHeader(1) = "계정명"					' Header명(1)
    
	Case "ToAcct"
		arrParam(0) = "계정코드팝업"			' 팝업 명칭 
		arrParam(1) = "a_acct"						' TABLE 명칭 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = "acct_type = " & FilterVar("K0", "''", "S") & " "			' Where Condition
		arrParam(5) = "계정코드"				' 조건필드의 라벨 명칭 
	
		arrField(0) = "acct_cd"						' Field명(0)
		arrField(1) = "acct_nm"						' Field명(1)
    
		arrHeader(0) = "계정코드"				' Header명(0)
		arrHeader(1) = "계정명"					' Header명(1)    

	Case 0, 1
		arrParam(0) = "사업장코드 팝업"			' 팝업 명칭 
		arrParam(1) = "B_BIZ_AREA" 					' TABLE 명칭 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition

		' 권한관리 추가 
		If lgAuthBizAreaCd <> "" Then
			arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
		Else
			arrParam(4) = ""
		End If

		arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 

		arrField(0) = "BIZ_AREA_CD"					' Field명(0)
		arrField(1) = "BIZ_AREA_NM"					' Field명(1)

		arrHeader(0) = "사업장코드"				' Header명(0)
		arrHeader(1) = "사업장명"				' Header명(1)

	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, Cond)
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

Function SetReturnVal(ByVal arrRet, ByVal field_fg)	
	Select case field_fg
		case "FrAcct"
			frm1.txtFrAcctCd.Value		= arrRet(0)
			frm1.txtFrAcctNm.Value		= arrRet(1)

		case "ToAcct"
			frm1.txtToAcctCd.Value		= arrRet(0)
			frm1.txtToAcctNm.Value		= arrRet(1)

		Case 0	'사업장코드 
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
			
		Case 1	'사업장코드 
			frm1.txtBizAreaCd1.focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
			
	End select	

End Function
'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup(Byval iwhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	Select Case iwhere
	Case "1"
		arrParam(0) = frm1.txtFrYymm.text								'  Code Condition
		arrParam(1) = frm1.txtToYymm.Text
		'arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
		arrParam(3) = frm1.txtFrDeptCd.value
		arrParam(4) = "F"									' 결의일자 상태 Condition  
	Case "2"
		arrParam(0) = frm1.txtFrYymm.text								'  Code Condition
		arrParam(1) = frm1.txtToYymm.Text
		'arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
		arrParam(3) = frm1.txtToDeptCd.value
		arrParam(4) = "F"									' 결의일자 상태 Condition  
	End Select
	
	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(iwhere, arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept(Byval iwhere,Byval arrRet)
	Select Case iwhere
	Case "1"
		frm1.txtFrDeptCd.focus
		frm1.txtFrDeptCd.value	= Trim(arrRet(0))
		frm1.txtFrDeptNm.value	= arrRet(1)
		frm1.txtFrYymm.text		= arrRet(4)
		frm1.txtToYymm.text		= arrRet(5)
	Case "2"
		frm1.txtToDeptCd.focus
		frm1.txtToDeptCd.value	= Trim(arrRet(0))
		frm1.txtToDeptNm.value	= arrRet(1)
		frm1.txtFrYymm.text		= arrRet(4)
		frm1.txtToYymm.text		= arrRet(5)
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
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    frm1.fpDateTime1.focus    
    Call InitVariables                            '⊙: Initializes local global Variables
    Call SetDefaultVal
    
    Call SetToolbar("10000000000011")

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
'======================================================================================================
'   Event Name : txtDeprYYYYMM_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================

Sub txtFrYymm_DblClick(Button)
    If Button = 1 Then
       frm1.txtFrYymm.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtFrYymm.Focus       
    End If
End Sub

Sub txtToYymm_DblClick(Button)
    If Button = 1 Then
       frm1.txtToYymm.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtToYymm.Focus       
    End If
End Sub


'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
'========================================================================================================
sub txtBizAreaCd_Onchange()
	Dim strCd

	strCd = frm1.txtBizAreaCd.value
	Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtBizAreaNm.value = ""
	else
		frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub


'========================================================================================================
'   Event Name : txtBizAreaCd1_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
'========================================================================================================
sub txtBizAreaCd1_Onchange()
	Dim strCd

	strCd = frm1.txtBizAreaCd1.value
	Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtBizAreaNm1.value = ""
	else
		frm1.txtBizAreaNm1.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile,StrUrl)

	Dim	VarFrAcct,VarToAcct,VarFromDt,VarTodt,VarBizAreaCd,VarBizAreaCd1,VarFrDept,VarToDept
	Dim	strAuthCond
	
	
	If frm1.Rb_acct.checked = true then
		StrEbrFile = "a7117ma1"
	elseif frm1.Rb_dur.checked = true then 
		StrEbrFile = "a7117ma2"
	else
		StrEbrFile = "a7117ma3"
	end if 
	

	If Len(frm1.txtFrAcctCd.value ) < 1 Then
		VarFrAcct = " "
	Else
		VarFrAcct = FilterVar(frm1.txtFrAcctCd.value,"","SNM")
	End If
	
	If Len(frm1.txtToAcctCd.value) < 1 Then
		VarToAcct = "ZZZZZZZZZZZZZZZZZZZZ"
	Else
		VarToAcct = FilterVar(frm1.txtToAcctCd.value,"","SNM")
	End If
	
	If Len(frm1.txtFrDeptCd.value) < 1 Then
		VarFrDept = " "
	Else
		VarFrDept = FilterVar(frm1.txtFrDeptCd.value,"","SNM")
	End If
	
	If Len(frm1.txtToDeptCd.value) < 1 Then
		VarToDept = "ZZZZZZZZZZ"
	Else
		VarToDept  = FilterVar(frm1.txtToDeptCd.value,"","SNM")
	End If
	
	VarFromDt = UniConvDateToYYYYMMDD(frm1.fpDateTime1.Text, gDateFormat,"")
	VarToDt   = UniConvDateToYYYYMMDD(frm1.fpDateTime2.Text, gDateFormat,"")
	
	If frm1.txtBizAreaCd.value = "" then 
		frm1.txtBizAreaNm.value = ""
		VarBizAreaCd = " "
	else 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
	end if
	
	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		VarBizAreaCd1 = "ZZZZZZZZZZ"
	else 
		VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
	end if


	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_ASSET_MASTER.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_ASSET_MASTER.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_ASSET_MASTER.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_ASSET_MASTER.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	



	StrUrl = StrUrl & "VarFrAcct|"		& VarFrAcct
	StrUrl = StrUrl & "|VarToAcct|"		& VarToAcct
	
	StrUrl = StrUrl & "|VarFrDept|"		& VarFrDept
	StrUrl = StrUrl & "|VarToDept|"		& VarToDept
	
	StrUrl = StrUrl & "|VarFromDt|"		& VarFromDt
	StrUrl = StrUrl & "|VarToDt|"		& VarToDt
	
	StrUrl = StrUrl & "|BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrEbrFile
	Dim ObjName
	Dim StrUrl
	Dim IntRetCd
		
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		Exit Function
	End If

	If CompareDateByFormat(frm1.txtFrYymm.text,frm1.txtToYymm.text,frm1.txtFrYymm.Alt,frm1.txtToYymm.Alt, _
	    	               "970025",frm1.txtFrYYMM.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFrYymm.focus
	   Exit Function
	End If

	If Trim(frm1.txtFrAcctCd.value) <> "" and   Trim(frm1.txtToAcctCd.value) <> "" Then		
		If Trim(frm1.txtFrAcctCd.value) > Trim(frm1.txtToAcctCd.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrAcctCd.Alt, frm1.txtToAcctCd.Alt)
			frm1.txtFrAcctCd.focus
			Exit Function
		End If
	End If

	If Trim(frm1.txtFrDeptCd.value) <> "" and   Trim(frm1.txtToDeptCd.value) <> "" Then    
		If Trim(frm1.txtFrDeptCd.value) > Trim(frm1.txtToDeptCd.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrDeptCd.Alt, frm1.txtToDeptCd.Alt)
			frm1.txtFrDeptCd.focus
			Exit Function
		End If
	End If
		
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If

	Call SetPrintCond(StrEbrFile,StrUrl)
		
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
		
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
	Dim StrEbrFile
	Dim ObjName
	Dim StrUrl
	Dim IntRetCd
	    
	If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If

	If CompareDateByFormat(frm1.txtFrYymm.text,frm1.txtToYymm.text,frm1.txtFrYymm.Alt,frm1.txtToYymm.Alt, _
	    	               "970025",frm1.txtFrYYMM.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFrYymm.focus
	   Exit Function
	End If

	If Trim(frm1.txtFrAcctCd.value) <> "" and   Trim(frm1.txtToAcctCd.value) <> "" Then				
		If Trim(frm1.txtFrAcctCd.value) > Trim(frm1.txtToAcctCd.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrAcctCd.Alt, frm1.txtToAcctCd.Alt)
			frm1.txtFrAcctCd.focus
			Exit Function
		End If
	End If

	If Trim(frm1.txtFrDeptCd.value) <> "" and   Trim(frm1.txtToDeptCd.value) <> "" Then    
		If Trim(frm1.txtFrDeptCd.value) > Trim(frm1.txtToDeptCd.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrDeptCd.Alt, frm1.txtToDeptCd.Alt)
			frm1.txtFrDeptCd.focus
			Exit Function
		End If
	End If
		
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If

	Call SetPrintCond(StrEbrFile,StrUrl)
	   
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
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
End Function


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
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
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

Sub txtFrAcctCd_onBlur()
	if frm1.txtFrAcctCd.value = "" then
		frm1.txtFrAcctNm.value = ""
	end if
End Sub
Sub txtToAcctCd_onBlur()
	if frm1.txtToAcctCd.value = "" then
		frm1.txtToAcctNm.value = ""
	end if
End Sub
Sub txtFrDeptCd_onBlur()
	if frm1.txtFrDeptCd.value = "" then
		frm1.txtFrDeptNm.value = ""
	end if
End Sub
Sub txtToDeptCd_onBlur()
	if frm1.txtToDeptCd.value = "" then
		frm1.txtToDeptNm.value = ""
	end if
End Sub
Sub txtBizAreaCd_onBlur()
	if frm1.txtBizAreaCd.value = "" then
		frm1.txtBizAreaNm.value = ""
	end if
End Sub
Sub txtBizAreaCd1_onBlur()
	if frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
	end if
End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0 >
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
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
			
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_acct Checked><LABEL FOR=Rb_WK1>계정별</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_dur><LABEL FOR=Rb_WK2>내용년수별</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_dept><LABEL FOR=Rb_WK2>부서별</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>취득일</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrYymm" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=시작취득일 id=fpDateTime1> </OBJECT>');</SCRIPT>~&nbsp
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToYymm" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=종료취득일 id=fpDateTime2> </OBJECT>');</SCRIPT>									
								</TD>					
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="시작사업장" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)"> 
													   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="종료사업장" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,1)"> 
													   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP>계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFrAcctCd" NAME="txtFrAcctCd" SIZE=15 MAXLENGTH=20 tag="11XXXU" ALT="시작계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtFrAcctCd.value, 'FrAcct')"> <INPUT TYPE="Text" NAME="txtFrAcctNm" SIZE=25 MAXLENGTH=30 tag="14X" ALT="계정코드명">&nbsp;~&nbsp;</TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtToAcctCd" NAME="txtToAcctCd" SIZE=15 MAXLENGTH=20 tag="11XXXU" ALT="종료계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtToAcctCd.value, 'ToAcct')"> <INPUT TYPE="Text" NAME="txtToAcctNm" SIZE=25 MAXLENGTH=30 tag="14X" ALT="계정코드명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>부서코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFrDeptCd" NAME="txtFrDeptCd" SIZE=12 MAXLENGTH=10 tag="11XXXU" ALT="시작부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup('1')"> <INPUT TYPE="Text" NAME="txtFrDeptNm" SIZE=25 MAXLENGTH=30 tag="14X" ALT="부서코드명">&nbsp;~&nbsp;</TD>
							</TR>			
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtToDeptCd" NAME="txtToDeptCd" SIZE=12 MAXLENGTH=10 tag="11XXXU" ALT="종료부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup('2')"> <INPUT TYPE="Text" NAME="txtToDeptNm" SIZE=25 MAXLENGTH=30 tag="14X" ALT="부서코드명"></TD>
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
					<TD><BUTTON NAME="btn출력" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp<BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX = "-1" >	
</FORM>
</BODY>
</HTML>

