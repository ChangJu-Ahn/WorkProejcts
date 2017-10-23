<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : B2210MA1
'*  4. Program Name         : Company Register(법인정보등록)
'*  5. Program Desc         : 법인정보등록 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/3/20
'*  8. Modified date(Last)  : 2000/8/29
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Kwon Yong Gyoun / Cho Ig Sung/kang eun kyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'***********************************************************************k*********************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '☜: indicates that All variables must be declared in advance 


'********************************************  1.2 Global 변수/상수 선언  *********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global 상수 선언  ====================================
'==========================================================================================================

Const BIZ_PGM_ID = "b2210mb1.asp"											 '☆: 비지니스 로직 ASP명 

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'☆: 사용자 변수 초기화 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""

	frm1.txtCO_CD.value = parent.gCompany
	frm1.txtco_cd.focus  
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


 
'------------------------------------------  InitComboBox()  ----------------------------------------------
'	Name :InitComboBox()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox_One()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTaxPolicy ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Two()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = " & FilterVar("B0004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCurPolicy ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Three()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = " & FilterVar("A1004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboxch_rate_fg ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Four()
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("A1020", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboOpenAcctFg ,lgF0  ,lgF1  ,Chr(11))  '미결관리여부(계정코드용)
	Call SetCombo2(frm1.cboXchErrorUseFg ,lgF0  ,lgF1  ,Chr(11))  '사용자환율계산 
End Sub

Sub InitComboBox_Five()
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("Z0015", "''", "S") & " ) ORDER BY MINOR_CD DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboQmdpalignopt ,lgF0  ,lgF1  ,Chr(11))  '멀티화폐소수점정렬(조회용)
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("Z0016", "''", "S") & " ) ORDER BY MINOR_CD DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboImdpalignopt ,lgF0  ,lgF1  ,Chr(11))  '멀티화폐소수점정렬(입력용)
End Sub

Sub InitComboBox_Six()
	Dim IntRetCD1
	Call CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("B9040", "''", "S") & " ) ORDER BY MINOR_CD DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboInvPostingFg ,lgF0  ,lgF1  ,Chr(11))  '재고포스팅방법 
End Sub

'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenCompanyInfo()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 

Function OpenCompanyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "법인 팝업"						' 팝업 명칭 
	arrParam(1) = "B_COMPANY"						' TABLE 명칭 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "법인"

    arrField(0) = "Upper(CO_CD)"					' Field명(0)
    arrField(1) = "CO_FULL_NM"						' Field명(1)

    arrHeader(0) = "법인코드"						' Header명(0)
    arrHeader(1) = "법인명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	Else
		Call SetCompanyInfo(arrRet,iWhere)
	End If	

End Function



'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------------
Function SetCompanyInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCO_CD.focus
			.txtCO_CD.value     = arrRet(0)
			.txtCO_FULLNM.value = arrRet(1)
		End If
'		lgBlnFlgChgValue = False
	End With

End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : OpenCountryInfo()
'	Description : 국가코드 popup
'========================================================================================================= 

Function OpenCountryInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "국가 팝업"							' 팝업 명칭 
	arrParam(1) = "B_COUNTRY"							' TABLE 명칭 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "국가"

    arrField(0) = "COUNTRY_CD"							' Field명(0)
    arrField(1) = "COUNTRY_NM"							' Field명(1)

    arrHeader(0) = "국가코드"							' Header명(0)
    arrHeader(1) = "국가명"							' Header명(1)
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCountryCd.focus
	    Exit Function
	Else
		Call SetCountryInfo(arrRet,iWhere)
	End If
End Function


'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetCountryInfo()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCountryInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCountryCd.focus
			.txtCountryCd.value = arrRet(0)
		End If
		lgBlnFlgChgValue = True
	End With

End Function


'========================================================================================================= 
Function OpenCurrencyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "자국통화 팝업"						' 팝업 명칭 
	arrParam(1) = "B_CURRENCY"							' TABLE 명칭 
	arrParam(2) = strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "자국통화"

    arrField(0) = "CURRENCY"							' Field명(0)
    arrField(1) = "CURRENCY_DESC"						' Field명(1)

    arrHeader(0) = "자국통화코드"						' Header명(0)
    arrHeader(1) = "자국통화명"						' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtLOC_CUR.focus
	    Exit Function
	Else
		Call SetCurrencyInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function OpenIndclassInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "업태 팝업"							' 팝업 명칭 
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) =  strCode								' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("B9003", "''", "S") & "  "					' Where Condition
	arrParam(5) = "업태"

    arrField(0) = "MINOR_CD"							' Field명(0)
    arrField(1) = "MINOR_NM"							' Field명(1)

    arrHeader(0) = "업태코드"							' Header명(0)
    arrHeader(1) = "업태명"							' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtInd_class.focus
	    Exit Function
	Else
		Call SetOpenIndclassInfo(arrRet,iWhere)
	End If
End Function

'========================================================================================================= 
Function OpenIndTypeInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	'B9003
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "업종 팝업"							' 팝업 명칭 
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) =  strCode							 	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("B9002", "''", "S") & "  "					' Where Condition
	arrParam(5) = "업종"

    arrField(0) = "MINOR_CD"							' Field명(0)
    arrField(1) = "MINOR_NM"							' Field명(1)

    arrHeader(0) = "업종코드"							' Header명(0)
    arrHeader(1) = "업종명"							' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtInd_Type.focus
	    Exit Function
	Else
		Call SetOpenIndTypeInfo(arrRet,iWhere)
	End If
End Function


'========================================================================================================= 
Function OpenZipCode(ByVal strCode, ByVal iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If Trim(frm1.txtCountryCd.value) = "" Then
		MsgBox "국가를 먼저 입력하세요", vbInformation, "uniERP(Information)"
		frm1.txtCountryCd.focus
		IsOpenPop = False
		Exit Function
	End IF
	iCalledAspName = AskPRAspName("ZipPopup")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.parent.VB_INFORMATION, "ZipPopup", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode
	arrParam(1) = ""
	arrParam(2) = Trim(frm1.txtCountryCd.value)

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
			frm1.txtzip_code.focus
	    Exit Function
	Else
		Call SetCurrencyInfo(arrRet,iWhere)
	End If
End Function


'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetCurrency()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrencyInfo(Byval arrRet,byval iWhere)'
	With frm1
		If iWhere = 0 Then
			.txtLOC_CUR.focus
			.txtLOC_CUR.value = arrRet(0)
		ElseIf iWhere = 1 Then
			.txtzip_code.focus
			.txtzip_code.value = arrRet(0)
			.txtaddr.value     = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function SetOpenIndclassInfo(Byval arrRet,byval iWhere)'
	With frm1
		If iWhere = 0 Then

			.txtInd_class.focus
			.txtInd_class.value = arrRet(0)
			.txtInd_class_Nm.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================= 
Function SetOpenIndTypeInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtInd_Type.focus
			.txtInd_Type.value = arrRet(0)
			.txtInd_Type_Nm.value = arrRet(1)
		End If
		lgBlnFlgChgValue = True
	End With

End Function



'========================================================================================================= 
Sub Form_Load()
    Call InitVariables																'⊙: Initializes local global variables
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("1100100000001111")
    Call InitComboBox_One
    Call InitComboBox_Two
    Call InitComboBox_Three
	Call InitComboBox_Four
	Call InitComboBox_Five
	Call InitComboBox_Six

	Call ggoOper.FormatDate(frm1.txtFirstDeprYyyymm, parent.gDateFormat, 2)
    'Call ggoOper.FormatDate(frm1.txtLastDeprYyyymm, parent.gDateFormat, 2)

	frm1.txtco_cd.focus 

    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed

	FncQuery

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
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
    Call DbQuery
    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables

    Call SetToolbar("1100100000001111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strYear,strMonth,strDay
    Dim strYear1,strMonth1,strDay1

	FncSave = False
	Err.Clear

	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '⊙: No data changed!!
	    Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '⊙: Check contents area
	   Exit Function
	End If

	If CompareDateByFormat(frm1.txtFISC_Start_DT.text,frm1.txtFISC_End_DT.text,frm1.txtFISC_Start_DT.Alt,frm1.txtFISC_End_DT.Alt, _
        	               "970024",frm1.txtFISC_Start_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFISC_Start_DT.focus
	   Exit Function
	End If
   
 	Call ExtractDateFrom(frm1.FDeprDateTime1.Text,frm1.FDeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    frm1.hFirstDeprYyyymm.value = strYear & strMonth

 	'Call ExtractDateFrom(frm1.LDeprDateTime1.Text,frm1.LDeprDateTime1.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    'frm1.hLastDeprYyyymm.value = strYear1 & strMonth1


	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then
		Exit Function
	End If

	FncSave = True
End Function


'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode

     ' 조건부 필드를 삭제한다. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
	lgBlnFlgChgValue = True

    frm1.txtCO_CD_Body.value = ""

    frm1.txtCO_CD_Body.focus
    
End Function


'========================================================================================
Function FncCancel()
     On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF

    response.write lgPrevNo

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtco_cd = " & lgPrevNo

	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						  '☜: 비지니스 처리 ASP의 상태값 
    strVal = strVal & "&txtco_cd=" & lgNextNo

	Call RunMyBizASP(MyBizASP, strVal)
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtco_cd=" & Trim(frm1.txtco_cd.value)				'☜: 삭제 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function


'=======================================================================================================
'   Event Name : txtFISC_START_DT_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFISC_START_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_START_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_START_DT.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFOUNDATION_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFOUNDATION_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFOUNDATION_DT.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFISC_END_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_END_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_END_DT.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFirstDeprYyyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtFirstDeprYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFirstDeprYyyymm.Focus
    End If
End Sub

'=======================================================================================================
Sub txtLastDeprYyyymm_DblClick(Button)
    If Button = 1 Then
        frm1.txtLastDeprYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLastDeprYyyymm.Focus
    End If
End Sub

'=======================================================================================================
Sub txtFISC_START_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtFOUNDATION_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtFISC_END_DT_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub txtTransStartDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub txtFirstDeprYyyymm_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub txtLastDeprYyyymm_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : cboQmdpalignopt_OnChange()
' Function Desc : 
'========================================================================================
Sub cboQmdpalignopt_OnChange()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================================= 
Sub cboImdpalignopt_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub cboTaxPolicy_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub cboCurPolicy_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboXCH_RATE_FG_OnChange()
	lgBlnFlgChgValue = True
End Sub
'========================================================================================
Sub txtFISC_CNT_Change() 
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboOpenAcctFg_OnChange() 
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub cboXchErrorUseFg_OnChange() 
	lgBlnFlgChgValue = True
End Sub

Sub cboInvPostingFg_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()
    Call SetToolbar("1100100000011111")
    lgBlnFlgChgValue = False
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	'20030916 jsk 최초상각년월 
	Call settxtFirstDeprYyyymmMode '기초자산이있으면 수정할 수 없다 
    lgIntFlgMode = parent.OPMD_UMODE
End Function

Function SettxtFirstDeprYyyymmMode()

	call CommonQueryRs("TOP 1 ACQ_NO"," A_ASSET_ACQ "," ACQ_FG = " & FilterVar("03", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If lgF0 <> "" Then
		Call ggoOper.SetReqAttr(frm1.txtFirstDeprYyyymm, "Q")
	Else
		Call ggoOper.SetReqAttr(frm1.txtFirstDeprYyyymm, "N")
	End If	
End Function
'========================================================================================
Function DbSave() 

    Err.Clear
	DbSave = False

    Dim strVal

    Call LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
    frm1.txtCO_CD.value = frm1.txtCO_CD_Body.value 
    lgBlnFlgChgValue = False
    FncQuery
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>법인</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtCO_CD" MAXLENGTH="10" SIZE=10 ALT ="법인코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCompanyInfo(frm1.txtco_cd.value,0)"> <INPUT NAME="txtCO_FULLNM" MAXLENGTH="30" SIZE=30 ALT ="법인명" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>법인코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_CD_Body" ALT="법인코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN:Left" tag = "23"></TD>
								<TD CLASS=TD5 NOWRAP>법인약명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_NM" ALT="법인약명" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>법인명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCO_FULL_NM_Body" ALT="법인명" MAXLENGTH="50" SIZE=45 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
								<TD CLASS=TD5 NOWRAP>법인영문명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtENG_NM" ALT="법인영문명" MAXLENGTH="50" SIZE=30 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>법인등록번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOwn_Rgst_No" ALT="법인등록번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="22"></TD>
								<TD CLASS=TD5 NOWRAP>대표자명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtREPRE_NM" ALT="대표자명" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대표자주민등록번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_Rgst_No" ALT="대표자주민등록번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="22" ></TD>
								<TD CLASS=TD5 NOWRAP>FAX번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFAX_NO" ALT="FAX번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag ="2" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>업태</TD>								
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_class" ALT="업태" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenIndclassInfo(frm1.txtInd_class.value,0)">
								<INPUT NAME="txtInd_class_Nm" ALT="업태" SIZE="20" tag = "24" ></TD>
								<TD CLASS=TD5 NOWRAP>전화번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTEL_NO" ALT="전화번호" MAXLENGTH="20" STYLE="TEXT-ALIGN:left" tag  ="2"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>업종</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="업종" MAXLENGTH="10" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenIndTypeInfo(frm1.txtInd_Type.value,0)">
								<INPUT NAME="txtInd_Type_Nm" ALT="업종" SIZE="20" tag = "24" ></TD>
								<TD CLASS=TD5 NOWRAP>국가코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCountryCd" ALT="국가코드" MAXLENGTH="2" SIZE="4" STYLE="TEXT-ALIGN:left" tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCountryInfo(frm1.txtCountryCd.value,0)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>회기</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE style="LEFT: 0px; WIDTH: 40px; TOP: 0px; HEIGHT: 20px" name=txtFISC_CNT CLASSID=<%=gCLSIDFPDS%> tag="22X6Z" ALT="회기" VIEWASTEXT id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>자국통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLOC_CUR" ALT="자국통화" MAXLENGTH="3" SIZE="4" STYLE="TEXT-ALIGN:left" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCurrencyInfo(frm1.txtLOC_CUR.value,0)"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>당기시작일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_START_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="당기시작일자" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>법인설립일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFOUNDATION_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="법인설립일자" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>당기종료일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_END_DT CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="당기종료일자" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>미결관리여부</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOpenAcctFg" ALT="미결관리여부" STYLE="WIDTH: 100px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최초감가상각년월</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFirstDeprYyyymm CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="최초감가상각년월" tag="21X1" id=FDeprDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
<!--								<TD CLASS=TD5 NOWRAP>최종감가상각년월</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtLastDeprYyyymm CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="최종감가상각년월" tag="21X1" id=LDeprDateTime1></OBJECT>');</SCRIPT></TD>
-->
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최종조직변경일</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtTransStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="최종조직변경일" tag="24" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>조직변경ID</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCurOrgChangeID" ALT="조직변경ID" MAXLENGTH="5" Size = "5" STYLE="TEXT-ALIGN:Center" tag = "24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세금계산정책</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboTaxPolicy" ALT="세금계산정책" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>환율구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboXCH_RATE_FG" ALT="환율구분" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>조회용소수점자리수</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboQmdpalignopt" ALT="조회용소수점자리수" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>입력용소수점자리수</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboImdpalignopt" ALT="입력용소수점자리수" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>환율재계산불가</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboXchErrorUseFg" ALT="환율재계산불가" STYLE="WIDTH: 100px" tag="22"></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>재고포스팅방법</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboInvPostingFg" ALT="재고포스팅방법" STYLE="WIDTH: 100px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>우편번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtzip_code" ALT="우편번호" MAXLENGTH="12" SIZE="11" STYLE="TEXT-ALIGN:left" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZipCode(frm1.txtZip_Code.value, 1)"></TD>
								<TD CLASS=TD5 NOWRAP>외환환율정책</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCurPolicy" ALT="외환환율정책" STYLE="WIDTH: 170px" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주소</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtaddr" ALT="주소" MAXLENGTH="128" SIZE="95" STYLE="TEXT-ALIGN:left"  tag="22" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>영문주소</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txteng_addr" ALT="영문주소" MAXLENGTH="128" SIZE="95" STYLE="TEXT-ALIGN:left"  tag="2" ></TD>
							</TR>
<!--							<% Call SubFillRemBodyTd5656(2) %> -->
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hFirstDeprYyyymm" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hLastDeprYyyymm" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

