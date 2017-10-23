
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3102ma1
'*  4. Program Name         : 예적금조회 
'*  5. Program Desc         : Query of Deposit
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.11
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  은행코드, 계좌번호 오류 Check 로직 추가 
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->				
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!--#Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim IsOpenPop                                               '☜: Popup status                           
Dim lgIsOpenPop

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

<%

'--------------- 개발자 coding part(실행로직,Start)-----------------------------------------------------------
'  Call GetAdoFiledInf("F3102MA1","S", "A")						'☆: spread sheet 필드정보 query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- 개발자 coding part(실행로직,End)-------------------------------------------------------------
%>

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f3102mb1.asp"						'☆: 비지니스 로직 ASP명 

'Dim lsPoNo                                                '☆: Jump시 Cookie로 보낼 Grid value
Const C_MaxKey          = 3                                    '☆☆☆☆: Max key value
Const C_GL_NO			= 9
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function부 
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgPageNo         = 0
    lgSortKey        = 1

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'                 lgSort...로 시작하는 변수 영역에 sort대상 목록을 저장 
'                 IsPopUpR 변수영역에 sort 정보의 기본이 되는 값 저장 
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
		

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "F","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","QA") %>

End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
        Call SetZAdoSpreadSheet("f3102ma1","S","A","V20030410",Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
        Call SetSpreadLock("A") 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval iOpt)
	  If iOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		End With
      End If
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()
<%	
	Dim arrData
	
'	arrData = InitCombo("F3011", "frm1.cboDpstFg")
'	arrData = InitCombo("F3014", "frm1.cboTransSts")
%>
End Sub
 
<%
Function InitCombo(ByVal strMajorCd, ByVal objCombo)

    Dim pB1a028
    Dim intMaxRow
    Dim intLoopCnt
    Dim strCodeList
    Dim strNameList
        
    Err.Clear                                                               '☜: Clear error no
	On Error Resume Next

	Set pB1a028 = Server.CreateObject("B1a028.B1a028ListMinorCode")
	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
		Set pB1a028 = Nothing												'☜: ComProxy Unload
		Call MessageBox(Err.description, I_INSCRIPT)						'⊙:
		Response.End														'☜: 비지니스 로직 처리를 종료함 
	End If

	pB1a028.ImportBMajorMajorCd = strMajorCd									'⊙: Major Code
    pB1a028.ServerLocation = ggServerIP
    pB1a028.ComCfg = gConnectionString
    pB1a028.Execute															'☜:
    
    '-----------------------
    'Com action result check area(DB,internal)
    '-----------------------
    If Not (pB1a028.OperationStatusMessage = Parent.MSG_OK_STR) Then
		Call MessageBox(pB1a028.OperationStatusMessage, I_INSCRIPT)         '☆: you must release this line if you change msg into code
		Set pB1a028 = Nothing												'☜: ComProxy Unload
		Response.End														'☜: 비지니스 로직 처리를 종료함 
    End If

	intMaxRow = pB1a028.ExportGroupCount
	
	For intLoopCnt = 1 To intMaxRow
%>
		Call SetCombo(<%=objCombo%>, "<%=pB1a028.ExportItemBMinorMinorCd(intLoopCnt)%>", "<%=pB1a028.ExportItemBMinorMinorNm(intLoopCnt)%>")		'⊙: InitCombo 에서 해야 되는데 임시로 넣은 것임 
<%
		strCodeList = strCodeList & vbtab & pB1a028.ExportItemBMinorMinorCd(intLoopCnt)
		strNameList = strNameList & vbtab & pB1a028.ExportItemBMinorMinorNm(intLoopCnt)
	Next
	
	InitCombo = Array(strCodeList, strNameList)
		
	Set pB1a028 = Nothing

End Function
%>

 '**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0
			arrParam(0) = "은행 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BANK A(NOLOCK), F_DPST B(NOLOCK)" 			' TABLE 명칭 
			arrParam(2) = Trim(strCode)					' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD"		' Where Condition

'			' 권한관리 추가 
'			If lgAuthBizAreaCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
'			End If
'			If lgInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
'			End If
'			If lgSubInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")
'			End If
'			If lgAuthUsrID <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
'			End If

			arrParam(5) = "은행코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "A.BANK_CD"					' Field명(0)
			arrField(1) = "A.BANK_NM"					' Field명(1)
			arrField(2) = "B.BANK_ACCT_NO"				' Field명(2)
    
			arrHeader(0) = "은행코드"				' Header명(0)
			arrHeader(1) = "은행명"					' Header명(1)
			arrHeader(2) = "계좌번호"				' Header명(1)
			
		Case 1
			arrParam(0) = "계좌번호 팝업"			' 팝업 명칭 
			arrParam(1) = "B_BANK A(NOLOCK), F_DPST B(NOLOCK)" 			' TABLE 명칭 
			arrParam(2) = Trim(strCode)					' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD"		' Where Condition

'			' 권한관리 추가 
'			If lgAuthBizAreaCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
'			End If
'			If lgInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
'			End If
'			If lgSubInternalCd <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")
'			End If
'			If lgAuthUsrID <> "" Then
'				arrParam(4) = arrParam(4) & " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
'			End If

			arrParam(5) = "계좌번호"				' 조건필드의 라벨 명칭 
	
			arrField(0) = "B.BANK_ACCT_NO"				' Field명(0)
			arrField(1) = "B.BANK_CD"					' Field명(1)
			arrField(2) = "A.BANK_NM"					' Field명(2)
    
			arrHeader(0) = "계좌번호"				' Header명(0)
			arrHeader(1) = "은행코드"				' Header명(1)
			arrHeader(2) = "은행명"					' Header명(1)

		Case 2
			arrParam(0) = "통화코드 팝업"				' 팝업 명칭 
			arrParam(1) = " B_CURRENCY A"					' TABLE 명칭 
			arrParam(2) = Trim(strCode)						' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "통화코드"					' 조건필드의 라벨 명칭 
	
			arrField(0) = "A.CURRENCY"						' Field명(0)
			arrField(1) = "A.CURRENCY_DESC"					' Field명(1)
    
			arrHeader(0) = "통화코드"					' Header명(0)
			arrHeader(1) = "통화명"						' Header명(1)
 		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True
	
	If iWhere = 2 Then 
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=450px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If 
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Select Case iWhere
		Case 0
			frm1.txtBankCd.value = arrRet(0)
			frm1.txtBankNm.value = arrRet(1)
			frm1.txtBankAcctNo.value = arrRet(2)
			frm1.txtBankAcctNo.focus
		Case 1
			frm1.txtBankAcctNo.value = arrRet(0)
			frm1.txtBankCd.value = arrRet(1)
			frm1.txtBankNm.value = arrRet(2)
			frm1.txtBankAcctNo.focus
		Case 2
			frm1.txtDocCur.value = arrRet(0)
			frm1.txtDocCur.focus
		End Select
	End If	
	
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(strCode)					' Code Condition
	arrParam(3) = ""							' Name Cindition

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function


'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet, iWhere)
	Select Case iWhere
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value	= arrRet(0)
			frm1.txtBizAreaNm1.Value	= arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	
	Dim arrField
	Dim ii

	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .ActiveRow > 0 Then
			.Row = .ActiveRow
			.Col = C_GL_NO
			arrParam(0) = Trim(.Text)	'회계전표번호 
			arrParam(1) = ""			'Reference번호 
		Else
			Call DisplayMsgBox("900025","X","X","X")
			Exit Function
		End If
	End With
	
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
End Function

'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
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
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	'Call InitComboBox()
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Call FncSetToolBar("New")
	frm1.txtBankCd.focus 
	
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
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
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
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 


'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'==========================================================================================

Sub txtDateFr_DblClick(Button)
	if Button = 1 then
		frm1.txtDateFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateFr.Focus
	End if
End Sub

Sub txtDateTo_DblClick(Button)
	if Button = 1 then
		frm1.txtDateTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateTo.Focus
	End if
End Sub

Sub txtDateFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtDateTo.Focus
	   Call MainQuery
	End If   
End Sub

Sub txtDateTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtDateFr.Focus
	   Call MainQuery
	End If   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
'	lsPoNo=frm1.vspdData.Text
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DbQuery
		End If
   End if
    
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
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData										'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'⊙: GL Date Compare Common Function
			Exit Function
	End if
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	If Trim(frm1.txtBizAreaCd.value) = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If Trim(frm1.txtBizAreaCd1.value) = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
	Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------
    
    If DbQuery	= False Then
       Exit Function
    End If

    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

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
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtBankCd="		& Trim(.txtBankCd.value)
		strVal = strVal & "&txtBankAcctNo="		& Trim(.txtBankAcctNo.value)
		strVal = strVal & "&txtDateFr="			& Trim(.txtDateFr.Text)
		strVal = strVal & "&txtDateTo="			& Trim(.txtDateTo.Text)
		strVal = strVal & "&txtDocCur="			& Trim(.txtDocCur.value)
		strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
		strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
		strVal = strVal & "&txtBankCd_Alt="		& Trim(.txtBankCd.Alt)
		strVal = strVal & "&txtBankAcctNo_Alt="	& Trim(.txtBankAcctNo.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt1="	& Trim(.txtBizAreaCd1.Alt)
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------

        strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey                      '☜: Next key tag
        strVal = strVal & "&lgPageNo="			& lgPageNo                      '☜: Next key tag		
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
    
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call FncSetToolBar("Query")
	Call CurFormatNumericOCX()
	
	Set gActiveElement = document.activeElement 
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	Dim IntRetCD
	Dim strGBCurrency
	Dim strBankCd
	Dim strBankAcctNo
	
	With frm1
        strBankCd =  frm1.txtBankCd.value
        strBankAcctNo =  frm1.txtBankAcctNo.value
         
		If Trim(.txtDocCur.value) = "" Then

		  intRetCD = CommonQueryRs("top 1 doc_cur"," f_dpst_item "," bank_cd =  " & FilterVar(strBankCd , "''", "S") & " and bank_acct_no =  " & FilterVar(strBankAcctNo , "''", "S") & " " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
           
          If intRetCD = True Then	
		  	strGBCurrency = Trim(Replace(lgF0,Chr(11),""))
		  Else
		    strGBCurrency = 	parent.gCurrency
		  End If					 

			ggoOper.FormatFieldByObjectOfCur .txtPreAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtRcptAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		Else
			ggoOper.FormatFieldByObjectOfCur .txtPreAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtRcptAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec		
		End If
	End With

End Sub

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################


'==========================================================
'툴바버튼 세팅 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>
						<TABLE CELLSPACING=0 CELLPADDING=0 align=right>
							<TR>
								<td>
									<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</a>
								</td>
						    </TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING="0" CELLPADDING="0">
								<TR>
									<TD CLASS=TD5 NOWRAP>은행</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankCd" NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="12XXXU" ALT="은행코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankCd.Value, 0)">
										<INPUT CLASS="clstxt" TYPE=TEXT ID="txtBankNM" NAME="txtBankNM" SIZE=25 MAXLENGTH=30  tag="24X" ALT="은행명">
									</TD>
									<TD CLASS=TD5 NOWRAP>계좌번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE="Text" ID="txtBankAcctNo" NAME="txtBankAcctNo" SIZE=20 MAXLENGTH=30 tag="12XXXU" ALT="계좌번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankAcctNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankAcctNo.Value, 1)">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>조회기간</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="조회시작일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="조회종료일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>통화</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="통화" SIZE = "10" MAXLENGTH="3"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 2)"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenBizAreaCd(frm1.txtBizAreaCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenBizAreaCd(frm1.txtBizAreaCd1.Value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100%>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS="TD5" NOWRAP>이월금액</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPreAmt" title=FPDOUBLESINGLE ALT="이월금액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
											
											<TD CLASS="TD5" NOWRAP>이월금액(자국)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPreLocAmt" title=FPDOUBLESINGLE ALT="이월금액(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>입금합계</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtRcptAmt" title=FPDOUBLESINGLE ALT="입금합계" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>입금합계(자국)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtRcptLocAmt" title=FPDOUBLESINGLE ALT="입금합계(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>출금합계</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPaymAmt" title=FPDOUBLESINGLE ALT="출금합계" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>출금합계(자국)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPaymLocAmt" title=FPDOUBLESINGLE ALT="출금합계(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>잔액</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtBalAmt" title=FPDOUBLESINGLE ALT="잔액" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>잔액(자국)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtBalLocAmt" title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HGIEHT_TYPE_01%>></td>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

