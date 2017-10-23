<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset
'*  3. Program ID           : a7120ma1
'*  4. Program Name         : 감가상각집계표출력 
'*  5. Program Desc         : Report of Depreciation Total
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000.11.29
'*  8. Modified date(Last)  : 2004/01/30
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : U & I (Kim Chang Jin)
'* 11. Comment              : 사업장 조건 추가 
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
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
<SCRIPT LANGUAGE="VBScript">

'Option Explicit																	'☜: indicates that All variables must be declared in advance


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
<!-- #Include file="../../inc/lgvariables.inc" --> 

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
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
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

	frm1.Rb_Total.checked = True
	frm1.fpDateTime1.Text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)

    Call ggoOper.FormatDate(frm1.txtDeprYYYYMM, Parent.gDateFormat, 2)   
' 	frm1.txtFromIssueDt.TEXT  = svrDate      	
    frm1.fpDateTime2.Text =  UNIMonthClientFormat(Parent.gFiscStart)
'    frm1.fpDateTime2.Text =UNIMonthClientFormat("<%=GetSvrDate%>")
    
	frm1.Rb_WK1.checked = True
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>

End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================

Sub SetSpreadLock()
End Sub


'================================== 2.2.5 SetSpreadColor() ================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'==========================================================================================================

Sub SetSpreadColor(ByVal lRow)
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
End Sub


'==========================================  2.2.7 SetCheckBox()  =======================================
'	Name : SetCheckBox()
'	Description : 감가상각집계표 출력물 체크박스 선택 처리(1개만 선택되도록 함)
'========================================================================================================= 
Function SetCheckBox(objCheckBox)
	Exit Function
	
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

Function OpenPopUp(strCode, iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
	Case 0, 1
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
		
    Case 2, 3
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
		Select Case iWhere
		Case 0
			frm1.txtAcctCdFR.Value = arrRet(0)
			frm1.txtAcctNmFR.Value = arrRet(1)
		Case 1
			frm1.txtAcctCdTO.Value = arrRet(0)
			frm1.txtAcctNmTO.Value = arrRet(1)
			
		Case 2	'사업장코드 
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
			
		Case 3	'사업장코드 
			frm1.txtBizAreaCd1.focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)		
		Case Else
		
		End Select
	End If	
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 


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
	Dim IntRetCd
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)                        
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
    Call ggoOper.FormatDate(frm1.txtFiscStYYYYMM, Parent.gDateFormat, 2)
    Call InitVariables                            '⊙: Initializes local global Variables

    Call SetDefaultVal
    
    Call SetToolbar("10000000000111")				'⊙: 버튼 툴바 제어 
	
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


'======================================================================================================
'   Event Name : txtDeprYYYYMM_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtDeprYYYYMM_DblClick(Button)
    If Button = 1 Then
        'frm1.txtFromReqrdDt.Action = 7
        frm1.fpDateTime1.Action = 7
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
Function SetPrintCond(StrEbrFile,StrUrl)

	Dim strYear, strMonth, strDay
	Dim strYear1, strMonth1, strDay1
	Dim	VarBizAreaCd,VarBizAreaCd1, VarAcctCdFR, VarAcctCdTO, VarDeprYYYYMM,VarFiscStYYYYMM, VarDurYrsFg

	Dim IntRetCd

	Dim	strAuthCond
		

	SetPrintCond = False


    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	If Trim(frm1.txtAcctCdFR.value) <> "" and Trim(frm1.txtAcctCdTO.value) <> "" Then
		If Trim(frm1.txtAcctCdFR.value) > Trim(frm1.txtAcctCdTO.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtAcctCdFR.Alt, frm1.txtAcctCdTO.Alt)
			frm1.txtAcctCdFR.focus
			Exit Function
		End If
	End If

	If CompareDateByFormat(frm1.txtFiscStYYYYMM.text,frm1.txtDeprYYYYMM.text,frm1.txtFiscStYYYYMM.Alt,frm1.txtDeprYYYYMM.Alt, _
        	               "970025",frm1.txtFiscStYYYYMM.UserDefinedFormat,gComDateType, true) = False Then
	   frm1.txtDeprYYYYMM.focus
	   Exit Function
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If



	If frm1.Rb_Total.checked = True Then
		StrEbrFile = "a7120ma1a"
	ElseIf frm1.Rb_Year_DeptCD.checked = True Then
		StrEbrFile = "a7120ma1b"
	ElseIf frm1.Rb_Month_DeptCD.checked = True Then
		StrEbrFile = "a7120ma1c"
	ElseIf frm1.Rb_Year_AccdCD.checked = True Then
		StrEbrFile = "a7120ma1d"
	ElseIf frm1.Rb_Month_AccdCD.checked = True Then
		StrEbrFile = "a7120ma1e"
	Else
		StrEbrFile = ""
	End If

	If Len(frm1.txtAcctCdFR.value) < 1 Then
		VarAcctCdFR = " "
	Else
		VarAcctCdFR = FilterVar( frm1.txtAcctCdFR.value,"","SNM")
	End If
	
	If Len(frm1.txtAcctCdTO.value) < 1 Then
		VarAcctCdTO = "ZZZZZZZZZZZZZZZZZZZZ"
	Else
		VarAcctCdTO  = FilterVar( frm1.txtAcctCdTO.value,"","SNM")
	End If
	
 	Call ExtractDateFrom(frm1.txtDeprYYYYMM.Text,frm1.txtDeprYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	VarDeprYYYYMM = strYear & strMonth
 	Call ExtractDateFrom(frm1.txtFiscStYYYYMM.Text,frm1.txtFiscStYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear1,strMonth1,strDay1)
	VarFiscStYYYYMM = strYear1 & strMonth1
		
	If frm1.Rb_WK1.checked = True Then
		VarDurYrsFg = "C"
	Else
		VarDurYrsFg = "T"
	End If
	
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
		strAuthCond		= strAuthCond	& " AND ISNULL(A_ASSET_CHG.TO_BIZ_AREA_CD,A_ASSET_CHG.FROM_BIZ_AREA_CD) = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND (case when ISNULL(A_ASSET_CHG.TO_INTERNAL_CD,'') <> '' then A_ASSET_CHG.TO_INTERNAL_CD else A_ASSET_CHG.FROM_INTERNAL_CD end) = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND (case when ISNULL(A_ASSET_CHG.TO_INTERNAL_CD,'') <> '' then A_ASSET_CHG.TO_INTERNAL_CD else A_ASSET_CHG.FROM_INTERNAL_CD end) LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_ASSET_MASTER.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	StrUrl = StrUrl & "AcctCDFR|"		& VarAcctCdFR
	StrUrl = StrUrl & "|AcctCDTO|"		& VarAcctCdTO
	StrUrl = StrUrl & "|DeprYYYYMM|"	& VarDeprYYYYMM
	StrUrl = StrUrl & "|FiscStrYYYYMM|"	& VarFiscStYYYYMM
	StrUrl = StrUrl & "|DurYrsFg|"		& VarDurYrsFg
	StrUrl = StrUrl & "|BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	
	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond


	SetPrintCond = True

End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
    Dim StrEbrFile, StrUrl, ObjName
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
    Dim StrEbrFile, StrUrl, ObjName
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
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
	Call parent.FncFind(C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
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

Sub txtAcctCdFR_onBlur()
	if frm1.txtAcctCdFR.value = "" then
		frm1.txtAcctNmFR.value = ""
	end if
End Sub

Sub txtAcctCdTo_onBlur()
	if frm1.txtAcctCdTo.value = "" then
		frm1.txtAcctNmTo.value = ""
	end if
End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	

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
					<TD>
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
					<TD WIDTH=100% HEIGHT=50% valign=top>
						<FIELDSET STYLE="HEIGHT:80%"><LEGEND>출력구분</LEGEND>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RadioFg ID=Rb_Total><LABEL FOR=Rb_Total>감가상각집계표(총괄)</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RadioFg ID=Rb_Year_DeptCD><LABEL FOR=Rb_Year_DeptCD>감가상각집계표(년간/부서별)</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RadioFg ID=Rb_Month_DeptCD><LABEL FOR=Rb_Month_DeptCD>감가상각집계표(월간/부서별)</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RadioFg ID=Rb_Year_AccdCD><LABEL FOR=Rb_Year_AccdCD>감가상각집계표(년간/계정별)</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RadioFg ID=Rb_Month_AccdCD><LABEL FOR=Rb_Month_AccdCD>감가상각집계표(월간/계정별)</LABEL></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<FIELDSET><LEGEND>출력조건</LEGEND>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="시작사업장" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,2)"> 
														   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="종료사업장" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,3)"> 
														   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계정코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCdFR" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="시작계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCDFR.Value, 0)"> <INPUT TYPE="Text" NAME="txtAcctNmFR" SIZE=25 MAXLENGTH=30 tag="14" ALT="시작계정명">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCdTO" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="종료계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdTO" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCDTO.Value, 1)"> <INPUT TYPE="Text" NAME="txtAcctNmTO" SIZE=25 MAXLENGTH=30 tag="14" ALT="종료계정명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>상각년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDeprYYYYMM" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=상각년월 id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>당기시작년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFiscStYYYYMM" CLASS=FPDTYYYYMM tag="14X1" Title="FPDATETIME" ALT=당기시작년월 id=fpDateTime2></OBJECT>');</SCRIPT>&nbsp;</TD>
								</TR>
<!--									<TD CLASS="HIDDEN"><INPUT TYPE="RADIO" CLASS="Radio" NAME="Radio1" TAG="12" ID=Rb_WK1 Checked><INPUT TYPE="RADIO" CLASS="Radio" NAME="Radio1" TAG="12" ID=Rb_WK2></TD>				포삼 -->
								<TR>
									<TD CLASS="TD5" NOWRAP>내용년수 구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK1 tag="12X1"  Checked><LABEL FOR=Rb_WK1>기업회계기준</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2 tag="12X1"><LABEL FOR=Rb_WK2>세법기준</LABEL></TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
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
