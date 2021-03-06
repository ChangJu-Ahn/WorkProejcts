
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4242oa1
'*  4. Program Name         : 미지급이자월결산내역출력 
'*  5. Program Desc         : Report of Loan Repay Plan
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003.05.19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ahn, do hyun
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">				</SCRIPT>

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
<!-- #Include file="../../inc/lgvariables.inc" -->	


'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim varDateFrist, varDateToto



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
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	                   'Select                 From        Where                Return value list  
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	Dim FirstDt, NextDtTo
	DIm strYear, strMonth, strDay

	Dim FirstDate, LastDate
 
    FirstDate	= UNIGetFirstDay("<%=GetSvrDate%>",	Parent.gServerDateFormat)				'Get Server DB Date
	LastDate	= UNIDateAdd("M",6,FirstDate,Parent.gServerDateFormat)
	LastDate	= UNIDateAdd("D",-1,LastDate,Parent.gServerDateFormat)


	Call ggoOper.FormatDate(frm1.txtDateFr, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtDateTo, Parent.gDateFormat, 2)
	
	Call ExtractDateFrom(FirstDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)

	Call InitComboBox

	frm1.txtDateFr.Year = strYear
	frm1.txtDateFr.Month = strMonth
	frm1.txtDateFr.Day = strDay
	
	Call ExtractDateFrom(LastDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtDateTo.Year = strYear
	frm1.txtDateTo.Month = strMonth
	frm1.txtDateTo.Day = strDay

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","OA") %>
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

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
        Case 1	
			arrParam(0) = "차입용도팝업"			' 팝업 명칭 
			arrParam(1) = "b_minor" 				    ' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "major_cd=" & FilterVar("f1000", "''", "S") & " "	        ' Where Condition
			arrParam(5) = "차입용도"				' 조건필드의 라벨 명칭 

			arrField(0) = "minor_cd"						' Field명(0)
			arrField(1) = "minor_nm"						' Field명(1)
    
			arrHeader(0) = frm1.txtLoanType.Alt				' Header명(0)
			arrHeader(1) = frm1.txtLoanTypeNm.Alt				    ' Header명(1)
		Case 2
			arrParam(0) = "거래통화팝업"								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "CURRENCY"										' Field명(0)
		    arrField(1) = "CURRENCY_DESC"									' Field명(1)

		    arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드명"									' Header명(1)
			
		Case 3, 4
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
		Select Case iWhere
			
			Case 1		'차입용도 
				frm1.txtLoanType.focus
				
			Case 2
				frm1.txtDocCur.focus
				
			Case 3
				frm1.txtBizAreaCd.focus
				
			Case 4
				frm1.txtBizAreaCd1.focus
		End Select
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetReturnPopUp()  --------------------------------------------------
'	Name : SetReturnPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 1		'차입용도 
				frm1.txtLoanType.value = arrRet(0)
				frm1.txtLoanTypeNm.value = arrRet(1)
				frm1.txtLoanType.focus
				
			Case 2
				frm1.txtDocCur.value = arrRet(0)
				frm1.txtDocCur.focus
				
			Case 3	'사업장코드 
				frm1.txtBizAreaCd.value = arrRet(0)
				frm1.txtBizAreaNm.value = arrRet(1)
				frm1.txtBizAreaCd.focus
					
			Case 4	'사업장코드 
				frm1.txtBizAreaCd1.value = arrRet(0)
				frm1.txtBizAreaNm1.value = arrRet(1)
				frm1.txtBizAreaCd1.focus
		End Select

	End With
	
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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
	
	Call InitVariables                            '⊙: Initializes local global Variables
	Call SetDefaultVal
    '----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("1000000000001111")				'⊙: 버튼 툴바 제어 
    
	frm1.txtDateFr.focus

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
'   Event Name : txtDateFr_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================

Sub txtDateFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDateFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateFr.Focus
	End If
End Sub


'========================================================================================================
'   Event Name : txtDateTo_DblClick
'   Event Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateTo.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : txtLoanType_Onchange()
'   Event Desc : 월차코드를 직접입력할경우에 월차코드명을 설정해준다.
'========================================================================================================
sub txtLoanType_Onchange()
	Dim strCd

	strCd = frm1.txtLoanType.value
	Call CommonQueryRs("A.minor_NM","B_MINOR A","A.major_cd=" & FilterVar("f1000", "''", "S") & "  AND A.minor_cd = " & FilterVar(strCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		frm1.txtLoanTypeNm.value = ""
	else
		frm1.txtLoanTypeNm.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub


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
Sub SetPrintCond(condvar, StrEbrFile)

	Dim i
	Dim VarLoanType, VarDocCur
	Dim varDateFr(6), varDateTo(5)
	Dim strYear, strMonth, strDay
	Dim VarBizAreaCd, VarBizAreaCd1
	
	Dim	strAuthCond
	
	StrEbrFile = "f4242oa1"

	VarLoanType = "%"
	VarDocCur = "%"

	If Trim(frm1.txtLoanType.value) <> "" Then VarLoanType = FilterVar(Trim(frm1.txtLoanType.value)	, "", "SNM")
	If Trim(frm1.txtDocCur.value) <> "" Then VarDocCur = FilterVar(Trim(frm1.txtDocCur.value)	, "", "SNM")

	varDateFr(0) = varDateFrist
	For i = 1 to 6
		varDateFr(i) = UNIDateAdd("M",i,varDateFr(0),Parent.gServerDateFormat)
		varDateTo(i-1) = UNIDateAdd("D",-1,varDateFr(i),Parent.gServerDateFormat)
		If varDateToto < varDateTo(i-1) Then
			varDateFr(i) = varDateFr(i-1)
			varDateTo(i-1) = varDateToto
		End If
	Next
	
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


	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_LN_INFO.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	
	condvar = "LoanType|" & VarLoanType
	condvar = condvar & "|DocCur|"			& VarDocCur
	condvar = condvar & "|DateFr1|"			& varDateFr(0)
	condvar = condvar & "|DateTo1|"			& varDateTo(0)
	condvar = condvar & "|DateFr2|"			& varDateFr(1)
	condvar = condvar & "|DateTo2|"			& varDateTo(1)
	condvar = condvar & "|DateFr3|"			& varDateFr(2)
	condvar = condvar & "|DateTo3|"			& varDateTo(2)
	condvar = condvar & "|DateFr4|"			& varDateFr(3)
	condvar = condvar & "|DateTo4|"			& varDateTo(3)
	condvar = condvar & "|DateFr5|"			& varDateFr(4)
	condvar = condvar & "|DateTo5|"			& varDateTo(4)
	condvar = condvar & "|DateFr6|"			& varDateFr(5)
	condvar = condvar & "|DateTo6|"			& varDateTo(5)
	condvar = condvar & "|BizAreaCd|"		& VarBizAreaCd
	condvar = condvar & "|BizAreaCd1|"		& VarBizAreaCd1

	condvar = condvar & "|strAuthCond|"		& strAuthCond
	
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim condvar
    Dim StrEbrFile
    
	varDateFrist = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateFr.Year, frm1.txtDateFr.Month, "01")
	varDateToto = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateTo.Year, frm1.txtDateTo.Month, "01")
	varDateToto = UNIGetLastDay (varDateToto,Parent.gServerDateFormat)

	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'⊙: GL Date Compare Common Function
			Exit Function
	End if

	
	'6개월차이 
	If UNICDbl(DateDiff("M",varDateFrist,varDateToto)) > 5 Then 
		Call DisplayMsgBox("141158","x", frm1.txtDateFr.Alt,frm1.txtDateTo.Alt)					 '☜: Data is changed.  Do you want to display it? 
		Exit Function
		frm1.txtDateFr.focus
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
    
    'On Error Resume Next       	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
  
	Call SetPrintCond(condvar, StrEbrFile)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	Call FncEBRPrint(EBAction,ObjName,condvar)	
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
    
    Dim condvar
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile    
    Dim ConDateFr, ConDateTo
	'On Error Resume Next                                                    '☜: Protect system from crashing
   
	varDateFrist = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateFr.Year, frm1.txtDateFr.Month, "01")
	varDateToto = UniConvYYYYMMDDToDate(Parent.gServerDateFormat, frm1.txtDateTo.Year, frm1.txtDateTo.Month, "01")
	varDateToto = UNIGetLastDay (varDateToto,Parent.gServerDateFormat)

	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'⊙: GL Date Compare Common Function
			Exit Function
	End if

	'6개월차이 
	If UNICDbl(DateDiff("M",varDateFrist,varDateToto)) > 5 Then 
		Call DisplayMsgBox("141158","x", frm1.txtDateFr.Alt,frm1.txtDateTo.Alt)					 '☜: Data is changed.  Do you want to display it? 
		Exit Function
		frm1.txtDateFr.focus
	End If
    
    If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
	Call SetPrintCond(condvar, StrEbrFile)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	Call FncEBRPreview(ObjName,condvar)
			
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

Function FncQuery()
	Exit Function
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD HEIGHT=* WIDTH=100%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>기준일자</TD>
									<TD CLASS=TD6 NOWRAP colspan=2 >
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작기준일자월" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작종료일자월" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="시작사업장" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,3)"> 
														   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명">&nbsp;~&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="종료사업장" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,4)"> 
														   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>차입용도</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtLoanType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="차입용도코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtLoanType.Value,1)">
														   <INPUT TYPE="Text" NAME="txtLoanTypeNm" SIZE=25 tag="14X" ALT="차입용도명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래통화</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 2)">
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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

