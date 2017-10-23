
<%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module명          : 회계-자금관리-어음 
'*  2. Function명        : 
'*  3. Program ID        : f5111ma1.asp
'*  4. Program 이름      : 지급어음명세서출력 
'*  5. Program 설명      : 지급어음명세서출력 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2003/01/08
'*  8. 최종 수정년월일   : 
'*  9. 최초 작성자       : Kim Chang Jin
'* 10. 최종 작성자       : 
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
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
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
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

'Dim  lgBlnFlgChgValue           ' Variable is for Dirty flag 
'Dim  lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim  IsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================
'=                       1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       1.4 User-defind Variables
'========================================================================================================
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
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtDateFr.Text = frDt
	frm1.txtDateTo.Text = toDt
		
	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	frm1.Rb_Dt1.checked = True	 '만기일 
	Call Radio_Dt_Click
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","Q") %>
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================

Sub InitComboBox()
	'어음상태 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboNoteSts ,lgF0  ,lgF1  ,Chr(11))
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
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	
End Function


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

	IsOpenPop = True
	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 5
			arrParam(0) = "사업장코드 팝업"								' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition

			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"										' Field명(0)
			arrField(1) = "BIZ_AREA_NM"										' Field명(1)
    
			arrHeader(0) = "사업장코드"									' Header명(0)
			arrHeader(1) = "사업장명"									' Header명(1)
			
		Case 2
			arrParam(0) = "은행 팝업"	' 팝업 명칭 
			arrParam(1) = "B_BANK"			 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "은행코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "BANK_CD"						' Field명(0)
			arrField(1) = "BANK_NM"						' Field명(1)
    
			arrHeader(0) = "은행코드"					' Header명(0)
			arrHeader(1) = "은행명"						' Header명(1)
		Case 3,4
			arrParam(0) = "어음번호"	' 팝업 명칭 
			arrParam(1) = "f_note"			 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = " 1=1 "							' Where Condition


			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			End If

			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			End If

			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			End If

			arrParam(5) = "어음번호"					' 조건필드의 라벨 명칭 

			arrField(0) = "note_no"						' Field명(0)
    
			arrHeader(0) = "어음번호"					' Header명(0)
			
		Case Else
			Exit Function
	End Select
    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
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
		Case 0		'사업장코드 
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
			frm1.txtBizAreaCd.focus
		Case 1		'거래처코드 
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNM.value = arrRet(1)
			frm1.txtBpCd.focus
		Case 2		'은행코드 
			frm1.txtBankCd.value = arrRet(0)
			frm1.txtBankNM.value = arrRet(1)
			frm1.txtBankCd.focus
		Case 3		'어음번호 
			frm1.txtNoteNoFr.value = arrRet(0)
			frm1.txtNoteNoFr.focus
		Case 4		'어음번호 
			frm1.txtNoteNoTo.value = arrRet(0)
			frm1.txtNoteNoTo.focus
		Case 5		'사업장코드 
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
			frm1.txtBizAreaCd1.focus
		Case Else
	End select	

End Function

'------------------------------------------  EscPopUp()  ---------------------------------------------
'	Name : EscPopUp()
'	Description : Account Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function EscPopUp(ByVal iWhere)
	
	Select Case iWhere
		Case 0		'사업장코드 
			frm1.txtBizAreaCd.focus
		Case 1		'거래처코드 
			frm1.txtBpCd.focus
		Case 2		'은행코드 
			frm1.txtBankCd.focus
		Case 3		'어음번호 
			frm1.txtNoteNoFr.focus
		Case 4		'어음번호 
			frm1.txtNoteNoTo.focus
		Case 5		'사업장코드 
			frm1.txtBizAreaCd1.focus
		Case Else
	End select	

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
	Call InitComboBox
	
    Call SetToolbar("1000000000001111")				'⊙: 버튼 툴바 제어 
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
'   Event Name : Radio_Dt
'   Event Desc : 만기일자 조건 변경 
'=======================================================================================================
Sub Radio_Dt_Click()
	With frm1
		If .Rb_Dt1.checked = True Then	 '만기일 
			lblTitle1.innerHTML = "만기일자"
			lblHyphen.innerHTML = "~"
			Call ElementVisible(frm1.fpDateTime2, 1)
		ElseIf .Rb_Dt2.checked = True Then	 '발행일 
			lblTitle1.innerHTML = "발행일자"
			lblHyphen.innerHTML = "~"
			Call ElementVisible(frm1.fpDateTime2, 1)
		Else	 '기준일 
			lblTitle1.innerHTML = "기준일자"
			lblHyphen.innerHTML = ""
			Call ElementVisible(frm1.fpDateTime2, 0)
		End If
	End With
End Sub

'======================================================================================================
'   Event Name : txtDateFr_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.fpDateTime1.Focus
    End If
End Sub

Sub txtDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.fpDateTime2.Action = 7
		Call SetFocusToDocument("M")
		Frm1.fpDateTime2.Focus
	End If
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrUrl, StrEbrFile)
	Dim StrDt, StrFg, VarBizAreaCd, VarBizAreaCd1, VarDateFr, VarDateTo, VarBpCd, VarBankCd, VarNoteSts
	Dim varNoteNoFr, varNoteNoTo
	Dim	strAuthCond
	
	If frm1.Rb_Dt1.checked = True Then
		StrDt = "a"
	ElseIf frm1.Rb_Dt2.checked = True Then
		StrDt = "b"
	Else
		StrDt = "c"
	End If

	If frm1.Rb_Fg1.checked = True Then
		StrFg = "1"
	ElseIf frm1.Rb_Fg2.checked = True Then
		StrFg = "2"
	ElseIf frm1.Rb_Fg3.checked = True Then
		StrFg = "3"
	Else
		StrFg = "4"
	End If
	
	StrEbrFile = "f5111ma1" & StrDt & StrFg

	VarBizAreaCd = "%"
	VarBpCd      = "%"
	VarBankCd    = "%"
	VarNoteSts   = "%"
	varNoteNoFr  = "0"
	varNoteNoTo  = "ZZZZZZZZZZZZZZZZZZ"
	
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, Parent.gServerDateType)
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, Parent.gServerDateType)
	
	If Trim(frm1.txtBizAreaCd.value) <> "" Then 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCd.value,"","SNM")
	else
		VarBizAreaCd = ""
	end if
	
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
	
	If Trim(frm1.txtBpCd.value)		<> "" Then VarBpCd = FilterVar(Trim(frm1.txtBpCd.value), "", "SNM")
	If Trim(frm1.txtBankCd.value)	<> "" Then VarBankCd = FilterVar(Trim(frm1.txtBankCd.value), "", "SNM")
	If Trim(frm1.cboNoteSts.value)	<> "" Then VarNoteSts = Trim(frm1.cboNoteSts.value)
	If Trim(frm1.txtNoteNoFr.value)	<> "" Then varNoteNoFr = Trim(frm1.txtNoteNoFr.value)
	If Trim(frm1.txtNoteNoTo.value)	<> "" Then varNoteNoTo = Trim(frm1.txtNoteNoTo.value)
	
	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_NOTE.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_NOTE.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_NOTE.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_NOTE.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	StrUrl = StrUrl & "BizAreaCd|"		& VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|"	& VarBizAreaCd1
	StrUrl = StrUrl & "|DateFr|"		& VarDateFr
	StrUrl = StrUrl & "|DateTo|"		& VarDateTo
	StrUrl = StrUrl & "|BpCd|"			& VarBpCd
	StrUrl = StrUrl & "|BankCd|"		& VarBankCd
	StrUrl = StrUrl & "|NoteSts|"		& VarNoteSts
	StrUrl = StrUrl & "|NoteNoFr|"		& varNoteNoFr
	StrUrl = StrUrl & "|NoteNoTo|"		& varNoteNoTo

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile	
    Dim ObjName
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	If (frm1.Rb_Dt1.checked = True) Or (frm1.Rb_Dt2.checked = True) Then
		If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'⊙: GL Date Compare Common Function
			Exit Function
		End if
	End If

	Call SetPrintCond(StrUrl, StrEbrFile)
	
'    On Error Resume Next                                                    '☜: Protect system from crashing
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim StrUrl, StrUrl2
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile	    
    Dim ObjName
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
	If (frm1.Rb_Dt1.checked = True) Or (frm1.Rb_Dt2.checked = True) Then
		If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
					"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'⊙: GL Date Compare Common Function
			Exit Function
		End if
	End If
	
	Call SetPrintCond(StrUrl, StrEbrFile)
	
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

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    FncQuery = True
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
					<TD WIDTH=100% HEIGHT=20%>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>출력구분</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg1 checked><LABEL FOR=Rb_Fg1>만기일</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg2 ><LABEL FOR=Rb_Fg2>발행일</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg3 ><LABEL FOR=Rb_Fg3>은행</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Fg ID=Rb_Fg4 ><LABEL FOR=Rb_Fg4>어음번호</LABEL>&nbsp;
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>날짜선택</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Dt ID=Rb_Dt1 checked ONCLICK="vbscript:Call Radio_Dt_Click()"><LABEL FOR=Rb_WK1>만기일</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Dt ID=Rb_Dt2 ONCLICK="vbscript:Call Radio_Dt_Click()"><LABEL FOR=Rb_WK2>발행일</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio_Dt ID=Rb_Dt3 ONCLICK="vbscript:Call Radio_Dt_Click()"><LABEL FOR=Rb_WK3>기준일</LABEL>&nbsp;
									</TD>
								</TR>
							</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=*>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)">
														   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명">&nbsp;~
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,5)">
									<INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblTitle1">만기일자</SPAN></TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=시작일자 id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;<SPAN CLASS="normal" ID="lblHyphen">~</SPAN>&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=종료일자 id=fpDateTime2></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBpCd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="거래처코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value, 1)">
														   <INPUT TYPE="Text" NAME="txtBpNM" SIZE=25 MAXLENGTH=40  tag="14X" ALT="거래처명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>은행</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="Text" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 2)">
														   <INPUT CLASS="Text" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=25 MAXLENGTH=30  tag="14X" ALT="은행명">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>어음번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="Text" TYPE=TEXT ID="txtNoteNoFr" NAME="txtNoteNoFr" SIZE=15 MAXLENGTH=18   tag="11XXXU" ALT="어음번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnNoteNoFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNoFr.Value, 3)">
														   ~ <INPUT CLASS="Text" TYPE=TEXT ID="txtNoteNoTo" NAME="txtNoteNoTo" SIZE=15 MAXLENGTH=18   tag="11XXXU" ALT="어음번호"><IMG SRC="../../image/btnPopup.gif" NAME="btnNoteNoTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtNoteNoTo.Value, 4)"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>어음상태</TD>
									<TD CLASS="TD6" NOWRAP><SELECT ID="cboNoteSts" NAME="cboNoteSts" ALT="어음상태" STYLE="WIDTH: 132px" tag="11X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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
 
