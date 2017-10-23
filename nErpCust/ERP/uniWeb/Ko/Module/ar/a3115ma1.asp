
<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : PrePayment management
'*  3. Program ID           : a3115ma1.asp
'*  4. Program Name         : 채권상세조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002.12.18
'*  8. Modified date(Last)  : 2004/02/04
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : U&I(Kim Chang Jin)
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2001.01.13
							  2004/02/04	사업장조건 추가 
'**********************************************************************************************
 -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'############################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<Script Language="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID        = "a3115MB1.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a3115MB2.asp"                         '☆: Biz logic spread sheet for #2
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================


Const C_MaxKey            = 5                                    '☆☆☆☆: Max key value
Const C_MaxKey_B            = 3                                   '☆☆☆☆: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop                                            '☜: Popup status                           
Dim  lgKeyPosVal
Dim  IsOpenPop												'☜: Popup status   
Dim  lgPageNo_A                                              '☜: Next Key tag                          
Dim  lgSortKey_A                                             '☜: Sort상태 저장변수                     
Dim  lgPageNo_B                                              '☜: Next Key tag                          
Dim  lgSortKey_B                                             '☜: Sort상태 저장변수                     

ReDim  lgKeyPosVal(C_MaxKey)
Dim strYear, strMonth, strDay,  EndDate, StartDate


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


<%
	Dim dtToday
	dtToday = GetSvrDate
%>	
	Call ExtractDateFrom("<%=dtToday%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, parent.gDateFormat)


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
Sub  InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B		 = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'                 lgSort...로 시작하는 변수 영역에 sort대상 목록을 저장 
'                 IsPopUpR 변수영역에 sort 정보의 기본이 되는 값 저장 
'========================================================================================================= 
Sub  SetDefaultVal()
	frm1.txtFromDt.text	= StartDate
	frm1.txtToDt.text	= EndDate
End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub  LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>
End Sub
'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A3115MA01","S","A","V20021211",Parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetZAdoSpreadSheet("A3115MA02","S","B","V20021211",Parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey_B, "X","X")
	Call SetSpreadLock("A")
	Call SetSpreadLock("B")																		
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock(Byval iOpt )
	If iOpt = "A" Then                                   ' 초기화 Spreadsheet #1 
		With frm1.vspdData
			.ReDraw = False
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadLockWithOddEvenRowColor()	
			.ReDraw = True
		End With 
    Else                                                ' 초기화 Spreadsheet #2 
		With frm1.vspdData2
			.ReDraw = False       
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadLockWithOddEvenRowColor()	
			.ReDraw = True
		End With 
    End If   
End Sub

 '**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 

'------------------------------------------  OpenDept()  ---------------------------------------
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
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""								'ToDt 
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	Select Case iWhere
		Case 1
			arrParam(5) = "SOL"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 2
			arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	End Select
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtDealBpCd.focus
			Case 2
				frm1.txtPayBpCd.focus
		End Select
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)

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
	arrParam(2) = strCode						' Code Condition
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


'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenConRouting()
'	Description : Routing PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim Field_fg

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	Field_fg = 3
	
	arrParam(0) = "계정코드팝업"								' 팝업 명칭 
	arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE 명칭 
	arrParam(2) = Trim(strCode)											' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
	arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

	arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
	arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
	arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
	arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
	arrHeader(0) = "계정코드"									' Header명(0)
	arrHeader(1) = "계정코드명"									' Header명(1)
	arrHeader(2) = "그룹코드"									' Header명(2)
	arrHeader(3) = "그룹명"										' Header명(3)

	lgIsOpenPop = True
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,Field_fg)
	End If
End Function

Function OpenARPopUp()
	Dim arrRet
	Dim Field_fg
	Dim arrParam
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a3101ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	ReDim arrParam(8)

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	Field_fg = 4
			
	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			

	lgIsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtArNo.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,Field_fg)
	End If	 
End Function			
			
'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		Case 1
			frm1.txtDealBpCd.value = arrRet(0)
			frm1.txtDealBpNm.value = arrRet(1)
			frm1.txtDealBpCd.focus				
		Case 2
			frm1.txtPayBpCd.value = arrRet(0)
			frm1.txtPayBpNm.value = arrRet(1)				
			frm1.txtPayBpCd.focus
		Case 3
			frm1.txtAcctCd.value = arrRet(0)
			frm1.txtAcctNm.value = arrRet(1)
			frm1.txtArNo.focus
		Case 4
			frm1.txtArNo.value = arrRet(0)
			frm1.txtArNo.focus
		case 5
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 6
			frm1.txtBizAreaCd1.Value	= arrRet(0)
			frm1.txtBizAreaNm1.Value	= arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : 
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim iGridPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			iGridPos = "A"
		Case "VSPDDATA2"			
			iGridPos = "B"
	End Select			
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(iGridPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(iGridPos,arrRet(0),arrRet(1))
       Call InitVariables()
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
Sub  Form_Load()
	Call LoadInfTB19029()			
	
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec) 
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field
    
	Call InitVariables()														'⊙: Initializes local global variables
	Call SetDefaultVal()
	Call InitSpreadSheet()
    Call SetToolbar("1100000000000111")											'⊙: 버튼 툴바 제어 

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
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

 '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub txtFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then
		frm1.txtToDt.focus
		Call FncQuery
	ENd if
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromDt.focus
		Call FncQuery
	End if		
End Sub

'========================================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'=========================================================================================================
Sub  txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus 	
	End If
End Sub

'========================================================================================================
'   Event Name : txtPoToDt
'   Event Desc :
'========================================================================================================
Sub  txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus 		
	End If
End Sub



'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
	
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData        
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
	End If
	
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	    

	Call DbQuery("2")
    
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
    lgPageNo_B       = ""                                  'initializes Previous Key
    lgSortKey_B      = 1
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    ggoSpread.Source = frm1.vspdData
			If lgSortKey_A = 1 Then
				ggoSpread.SSSort, lgSortKey_A
	            lgSortKey_A = 2
		    Else
			    ggoSpread.SSSort, lgSortKey_A
				lgSortKey_A = 1
	        End If    
		    Exit Sub
	    End If
	    
		Call SetSpreadColumnValue("A",frm1.vspdData,Col,NewRow)	        
    
		Call DbQuery("2")
     
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	
		lgPageNo_B       = ""                                  'initializes Previous Key
		lgSortKey_B      = 1
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
	
    gMouseClickStatus = "SP2C"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData2            
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If

	Call SetSpreadColumnValue("B",frm1.vspdData2,Col,Row)	
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("1")
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("2")
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
Function  FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    Err.Clear     
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
		Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtFromDt.text,frm1.txtToDt.text,frm1.txtFromDt.Alt,frm1.txtToDt.Alt, _
        	               "970025",frm1.txtFromDt.UserDefinedFormat,parent.gComDateType, true) = False Then
		frm1.txtFromDt.focus
		Exit Function
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	if frm1.txtBizAreaCd.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd.alt,"X")            '☜ : No data is found. 
	  	frm1.txtBizAreaNm.value = ""
	  	frm1.txtBizAreaCd.focus
 	  	Exit Function
	  End If
	End If
	  
	if frm1.txtBizAreaCd1.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd1.alt,"X")            '☜ : No data is found.
	  	frm1.txtBizAreaNm1.value = ""
	  	frm1.txtBizAreaCd1.focus
 	  	Exit Function
	  End If
	End If
	
	If Trim(frm1.txtDealBpCd.value) = "" Then
		frm1.txtDealBpNm.value = ""
	End If	
	
	If Trim(frm1.txtPayBpCd.value) = "" Then
		frm1.txtPayBpNm.value = ""
	End If	
	
	If Trim(frm1.txtAcctCd.value) = "" Then
		frm1.txtAcctnm.value = ""
	End If	
	
	If Trim(frm1.txtBizAreaCd.value) = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If Trim(frm1.txtBizAreaCd1.value) = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData    
	    
    Call InitVariables() 														'⊙: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery("1")															'☜: Query db data

    FncQuery = True		
	
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    Call parent.FncPrint()
    	
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
    	
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    Dim iColumnLimit2
    
    If gMouseClickStatus = "SPCRP" Then
		iColumnLimit = 3
       
		ACol = Frm1.vspdData.ActiveCol
		ARow = Frm1.vspdData.ActiveRow

		If ACol > iColumnLimit Then
		   iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
		   Exit Function  
		End If   
    
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
    
		ggoSpread.Source = Frm1.vspdData
    
		ggoSpread.SSSetSplit(ACol)    
    
		Frm1.vspdData.Col = ACol
		Frm1.vspdData.Row = ARow
    
		Frm1.vspdData.Action = 0    
    
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
	'----------------------------------------
	' Spread가 두개일 경우 2번째 Spread
	'----------------------------------------
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = 4
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function  FncExit()
    FncExit = True
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function  DbQuery(ByVal iOpt) 
	Dim strVal
	
    Err.Clear																						'☜: Protect system from crashing
	On Error Resume Next
	
    DbQuery = False
    Call DisableToolBar(parent.TBC_QUERY)															'☜: Disable Query Button Of ToolBar
	Call LayerShowHide(1)
    
    With frm1
		Select Case iOpt 
			Case "1" 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
				strVal = BIZ_PGM_ID & "?txtFromDt="		& Trim(.txtFromDt.Text)
				strVal = strVal & "&txtToDt="			& Trim(.txtToDt.Text)
				strVal = strVal & "&txtDealBpCd="		& Trim(.txtDealBpCd.value)
				strVal = strVal & "&txtPayBpCd="		& Trim(.txtPayBpCd.value)
				strVal = strVal & "&txtAcctCd="			& Trim(.txtAcctCd.value)
				strVal = strVal & "&txtDesc="			& Trim(.txtDesc.value)
				strVal = strVal & "&txtArNo="			& Trim(.txtArNo.value)
				strVal = strVal & "&txtRefNo="			& Trim(.txtRefNo.value)
				strVal = strVal & "&txtInvDocNo="		& Trim(.txtInvDocNo.value)
				strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
				strVal = strVal & "&txtDealBpCd_ALT="	& .txtDealBpCd.alt
				strVal = strVal & "&txtPayBpCd_ALT="	& .txtPayBpCd.alt
				strVal = strVal & "&txtAcctCd_ALT="		& .txtAcctCd.alt
				strVal = strVal & "&txtDesc_ALT="		& .txtDesc.alt
				strVal = strVal & "&txtArNo_ALT="		& .txtArNo.alt
				strVal = strVal & "&txtRefNo_ALT="		& .txtRefNo.alt
				strVal = strVal & "&txtInvDocNo_ALT="	& .txtInvDocNo.alt
				strVal = strVal & "&txtBizAreaCd_ALT="	& .txtBizAreaCd.alt
				strVal = strVal & "&txtBizAreaCd_ALT1="	& .txtBizAreaCd1.alt
				strVal = strVal & "&txtProject="		& Trim(.txtProject.value)
				
    '--------- Developer Coding Part (End) ----------------------------------------------------------
				strVal = strVal & "&lgPageNo="			& lgPageNo_A									'☜: Next key tag
				strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

				' 권한관리 추가 
				strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
				strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
				strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
				strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

			Case "2"
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------				
				strVal = BIZ_PGM_ID1 & "?txtArNo="		& GetKeyPosVal("A",1)

				strVal = strVal & "&lgPageNo="			& lgPageNo_B									'☜: Next key tag
				strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("B")
				strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("B")
				strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("B"))
		End Select 
      
		Call RunMyBizASP(MyBizASP, strVal)															'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(byval iOpt)																		'☆: 조회 성공후 실행로직 
    lgIntFlgMode = parent.OPMD_UMODE																'⊙: Indicates that current mode is Update mode
    
	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If																							'⊙: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")															'⊙: This function lock the suitable field 
End Function

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
					<TD WIDTH="*" align=right>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
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
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
									<TD CLASS="TD5" NOWRAP>발생기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="12" VIEWASTEXT > </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12" VIEWASTEXT > </OBJECT>');</SCRIPT>					
									</TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBizAreaCd(frm1.txtBizAreaCd.Value, 5)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDealBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="1XXXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtDealBpNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBizAreaCd(frm1.txtBizAreaCd1.Value, 6)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
								</TR>			 					
			 					</TR>
									<TD CLASS="TD5" NOWRAP>수금처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPayBpCd" SIZE=10 MAXLENGTH=10  STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="수금처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value, 2)">&nbsp;<INPUT TYPE=TEXT NAME="txtPayBpNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ><IMG align=top name=btnCalType onclick="vbscript:CALL OpenAcctPopUp(frm1.txtAcctCd.value)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
														<INPUT NAME="txtAcctnm" ALT="계정코드명" MAXLENGTH="20" SIZE=25 tag  ="14"></TD>										
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>채권번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtArNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="채권번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenARPopUp()"></TD>
									<TD CLASS=TD5 NOWRAP>송장번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvDocNo" ALT="송장번호" MAXLENGTH="50" STYLE="TEXT-ALIGN: Left" tag="11XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="30" tag="11XXXU" ></TD>
									<TD CLASS=TD5 NOWRAP>참조번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호" MAXLENGTH="30" STYLE="TEXT-ALIGN: Left" tag="11XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>프로젝트</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="프로젝트" MAXLENGTH=25 SIZE=25 tag="1X"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									
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
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=6>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TDT NOWRAP>채권액(자국)</TD>
								<TD CLASS=TDT NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="채권액(자국)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TDT NOWRAP>반제금액(자국)</TD>
								<TD CLASS=TDT NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="반제금액(자국)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TDT NOWRAP>잔액(자국)</TD>
								<TD CLASS=TDT NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotBalLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="잔액(자국)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan=6>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

