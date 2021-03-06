
<%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module Name          : Finance
'*  2. Function Name        : PrePayment management
'*  3. Program ID           : f6103ma1.asp
'*  4. Program Name         : 선급금내역조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001.01.16
'*  9. Modifier (First)     : Hersheys
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*                            2001.01.13
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   ***************************************************
'	기능: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
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
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance


'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID        = "F6103MB1.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "F6103MB2.asp"                         '☆: Biz logic spread sheet for #2
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey				= 6                                    '☆☆☆☆: Max key value
Const C_MaxKey_B			= 2 
Const C_SHEETMAXROWS_A		= 30
Const C_SHEETMAXROWS_B		= 10
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim lgIsOpenPop                                             '☜: Popup status                           

'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   

Dim lgPageNo_A                                              '☜: Next Key tag                          
Dim lgSortKey_A                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   

Dim lgPageNo_B                                              '☜: Next Key tag                          
Dim lgSortKey_B                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet temp---------------------------------------------------------------------------   
Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 



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
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode
    
    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

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
	
	frm1.hOrgChangeId.value		= parent.gChangeOrgId	
	frm1.txtFromDt.Text			= frDt
	frm1.txtToDt.Text			= toDt
End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>

End Sub
'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("f6103ma1","S","A","V20021215",Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
    Call SetZAdoSpreadSheet("f6103ma2","S","B","V20021215",Parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey_B,"X","X")
    Call SetSpreadLock("A")
    Call SetSpreadLock("B")
    
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
    Else
       With frm1
            .vspdData2.ReDraw = False
    		ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()	
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub

 '**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
		lgBlnFlgChgValue = True
	End If

End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	Select Case iWhere
		Case 2
			arrParam(0) = "선급금유형"									' 팝업 명칭 
			arrParam(1) = "a_jnl_item"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "jnl_type = " & FilterVar("pp", "''", "S") & "  "								' Where Condition
			arrParam(5) = "선급금유형"									' 조건필드의 라벨 명칭 

			arrField(0) = "JNL_CD"											' Field명(0)
			arrField(1) = "JNL_NM"											' Field명(1)
    
			arrHeader(0) = "선급금유형"									' Header명(0)
			arrHeader(1) = "선급금유형명"			
    End Select 
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopUp(Byval iWhere)

	With frm1
		Select Case iWhere
			Case 1
				.txtBpCd.focus
			Case 2
				.txtPrpaymType.focus
		End Select
	End With
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 1
				.txtBpCd.value = arrRet(0)
				.txtBpNm.value = arrRet(1)				
				.txtBpCd.focus
			Case 2
				.txtPrpaymType.value = arrREt(0)
				.txtPrpaymTypeNm.value = arrRet(1)
				.txtPrpaymType.focus
		End Select
	End With
End Function

'======================================================================================================
'   Function Name : OpenDeptOrgPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = frm1.txtFromDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  
	
	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFromDt.text = arrRet(4)
		frm1.txtToDt.text = arrRet(5)
		frm1.txtDeptCd.focus
End Function

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "VSPDDATA2"                  
	            gPos = "B"
	       End Select     
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
    
    arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
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
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
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
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value = arrRet(0)
			frm1.txtBizAreaNm1.Value = arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
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
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
															    '  Arg 2. S is for Sort , G is for group 
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
  	
	frm1.txtFromDt.focus
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Set gActiveElement = document.activeElement 

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
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

'========================================================================================================
'   Event Name : DblClick
'   Event Desc :
'=========================================================================================================
Sub txtFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus 
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus 	
	End if
End Sub

'========================================================================================================
'   Event Name : KeyPress
'   Event Desc :
'========================================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToDt.focus
		Call FncQuery
	end if
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFromDt.focus
		Call FncQuery
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row = 0 Then
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
     
     ggoSpread.Source = frm1.vspdData2
     ggospread.ClearSpreadData

     lgPageNo_B       = ""                                  'initializes Previous Key
     lgSortKey_B      = 1
     
     Call DbQuery("2") 
    
End Sub


'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : 컬럼을 방향키로 이동할때 
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    Dim ii
    
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
        ggospread.ClearSpreadData
		lgPageNo_B       = ""                                  'initializes Previous Key
		lgSortKey_B      = 1
    End if 
     
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    call SetPopupMenuItemInf("00000000001") 
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
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	gMouseClickStatus = "SP2C"	'Split 상태코드 

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button <> "1"And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
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
        
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)
            If DbQuery("1") = False Then
              Call RestoreToolBar()
              Exit Sub
			End IF
		End If
	End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery("2") = False Then
              Call RestoreToolBar()
              Exit Sub
          End if
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
	Dim IntRetCD 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear     

    '-----------------------
    'Erase contents area
    '-----------------------
	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables
    
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
	    
    If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = ""
    End If
    
    If frm1.txtDeptCd.value = "" Then
		frm1.txtDeptNm.value = ""
    End If
	
	If Trim(frm1.txtPrpaymType.value)="" then
		frm1.txtPrpaymTypeNm.value="" 
	End if
	
	If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If frm1.txtBizAreaCd1.value = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("124600","X","X","X")           '⊙: Display Message(There is no changed data.)
		Exit Function
	End if
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
	  If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
	  	IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
	  	frm1.txtBizAreaCd.focus
	  	Exit Function
	  End If
	End If
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

Function FncPrint() 
    Call parent.FncPrint()
	    		
	Set gActiveElement = document.activeElement    

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	    		
	Set gActiveElement = document.activeElement    

End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
	    		
	Set gActiveElement = document.activeElement    

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
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbQuery(ByVal iOpt) 
	Dim strVal,strDeptCd,strBpCd

	DbQuery = False
	    
	Err.Clear                                                               '☜: Protect system from crashing
	Call LayerShowHide(1)
		
	  
	With frm1
		If iOpt = "1" Then
			'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
			strDeptCd	= Trim(.txtDeptCd.value)
			strBpCd		= Trim(.txtBpCd.value)
					   
			strVal = BIZ_PGM_ID & "?txtFromDt="		& Trim(.txtFromDt.Text)
			strVal = strVal & "&txtToDt="			& Trim(.txtToDt.Text)
			strVal = strVal & "&txtDeptCd="			& Trim(.txtDeptCd.value) 'EnCoding(strDeptCd)
			strVal = strVal & "&txtBpCd="			& Trim(.txtBpCd.value)   'EnCoding(strBpCd)
			strVal = strVal & "&txtPrpaymType="		& Trim(.txtPrpaymType.value)           
			strVal = strVal & "&txtDeptCd_Alt="		& .txtDeptCd.Alt
			strVal = strVal & "&txtBpCd_Alt="		& .txtBpCd.Alt
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt=" & Trim(frm1.txtBizAreaCd1.alt)

			strVal = strVal & "&OrgChangeId="		& Trim(.hOrgChangeId.Value)
			strVal = strVal & "&lgPageNo="			& lgPageNo_A                          '☜: Next key tag
			strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
			strVal = strVal & "&lgMaxCount="		& C_SHEETMAXROWS_A
			strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))
		           
			' 권한관리 추가 
			strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
			strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
			strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
			strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

		Else   
			strVal = BIZ_PGM_ID1 & "?txtPrrcptNo="	& GetKeyPosVal("A",1)

			strVal = strVal & "&lgPageNo="			& lgPageNo_B                          '☜: Next key tag
			strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("B")
			strVal = strVal & "&lgMaxCount="		& C_SHEETMAXROWS_B
			strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("B")
			strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("B"))
		   
		End If   

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	End With
	    
	DbQuery = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk(byval iOpt)														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 
	
	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If							                                     '⊙: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 
 	
End Function
'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFromDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

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
					<TD WIDTH=*>&nbsp;</TD>
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
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="12" VIEWASTEXT id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12" VIEWASTEXT id=fpDateTime2></OBJECT>');</SCRIPT>					
									</TD>			 					
			 						<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="부서"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="시작사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>선급금유형</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPrpaymType" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="선급금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtPrpaymType.value, 2)">&nbsp;
										<INPUT TYPE=TEXT NAME="txtPrpaymTypeNm" SIZE=30 tag="14XXXU" ALT="선급금유형명"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="종료사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd1.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
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
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24" TabIndex="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

