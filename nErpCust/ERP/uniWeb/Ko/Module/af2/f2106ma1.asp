

<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2106ma1
'*  4. Program Name         : 예산실적조회 
'*  5. Program Desc         : Query of Budget Result
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.02.12
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  부서코드, 예산코드 오류 Check
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'Dim lgBlnFlgChgValue                                        '☜: Variable is for Dirty flag            
Dim lgIsOpenPop                                             '☜: Popup status                          

'☜:--------Spreadsheet #1----------------------------------------------------------------------------- 



Dim lgStrPrevKey_A                                          '☜: Next Key tag                          
Dim lgSortKey_A                                             '☜: Sort상태 저장변수                     

'☜:--------Spreadsheet #2----------------------------------------------------------------------------- 

Dim lgStrPrevKey_B                                          '☜: Next Key tag                          
Dim lgSortKey_B                                             '☜: Sort상태 저장변수                     

'☜:--------Spreadsheet temp---------------------------------------------------------------------------  
                                                            '☜:--------Buffer for Spreadsheet -----   
Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         
Dim strFrDt
Dim strToDt


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f2106mb1.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "f2106mb2.asp"                         '☆: Biz logic spread sheet for #2

Const C_MaxKey            = 5                                    '☆☆☆☆: Max key value

<%
Dim lsSvrDate
lsSvrDate = GetSvrDate
%>
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

    lgStrPrevKey_A   = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgStrPrevKey_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'                 lgSort...로 시작하는 변수 영역에 sort대상 목록을 저장 
'                 IsPopUpR 변수영역에 sort 정보의 기본이 되는 값 저장 
'=========================================================================================================
Sub SetDefaultVal()
'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	Dim strSvrDate
	strSvrDate = "<%=GetSvrDate%>"
    frm1.fpBdgYymmFr.focus 
	
	frm1.txtBdgYymmFr.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.txtBdgYymmTo.Text = UniConvDateAToB(strSvrDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtBdgYymmFr, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtBdgYymmTo, parent.gDateFormat, 2)
	frm1.hOrgChangeId.value = parent.gChangeOrgId
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("F2106MA1","S","A","V20021211",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")    
    Call SetZAdoSpreadSheet("F2106MA1","S","B","V20021211",parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
    
    Call SetSpreadLock("A")
    Call SetSpreadLock("B")
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval iOpt )
    If iOpt = "A" Then
       With frm1
			.vspdData.ReDraw = False
			ggoSpread.Source = .vspdData 
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			.vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2 
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub
 '**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	Select Case iWhere

		Case "BdgCdFr", "BdgCdTo"
			arrParam(0) = "예산코드 팝업"								' 팝업 명칭 
			arrParam(1) = "F_BDG_ACCT A "									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""
			arrParam(5) = "예산코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A.BDG_CD"	     								' Field명(0)
			arrField(1) = "A.GP_ACCT_NM"			    					' Field명(1)
			
			arrHeader(0) = "예산코드"									' Header명(0)
			arrHeader(1) = "예산명"										' Header명(1)
			
	   Case "DeptCd"
'			arrParam(0) = "부서코드 팝업"								' 팝업 명칭 
'			arrParam(1) = "B_ACCT_DEPT A "									' TABLE 명칭 
'			arrParam(2) = strCode											' Code Condition
'			arrParam(3) = ""												' Name Cindition
'			arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(parent.gChangeOrgId , "''", "S") & ""
'			arrParam(5) = "부서코드"									' 조건필드의 라벨 명칭 
'
'			arrField(0) = "A.DEPT_CD"
'			arrField(1) = "A.DEPT_NM"
'			
'			arrHeader(0) = "부서코드"									' Header명(0)
'			arrHeader(1) = "부서명"										' Header명(1)
'		
		Case Else
			Exit Function
	End Select	
	
	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		    
		    Case "BdgCdFr"
				.txtBdgCdFr.value = arrRet(0)
				.txtBdgNmFr.value = arrRet(1)
				.txtBdgCdFr.focus
				
		    Case "BdgCdTo"
				.txtBdgCdTo.value = arrRet(0)
				.txtBdgNmTo.value = arrRet(1)
				.txtBdgCdTo.focus
				
			Case "DeptCd"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtDeptCd.focus
			
		End Select
    
	End With

End Function
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	'iCalledAspName = AskPRAspName("DeptPopupOrg")
	'If Trim(iCalledAspName) = "" Then
	'	IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupOrg", "X")
	'	lgIsOpenPop = False
	'	Exit Function
	'End If

	arrParam(0) = UniConvDateAToB(frm1.txtBdgYymmFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	arrParam(1) = UniConvDateAToB(frm1.txtBdgYymmTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	arrParam(1) = UNIDateAdd("M", +1, arrParam(1),parent.gServerDateFormat)
	arrParam(1) = UNIDateAdd("D", -1, arrParam(1),parent.gServerDateFormat)	    

	arrParam(0)  =  UniConvDateAToB(arrParam(0),parent.gServerDateFormat,gDateFormat)
	arrParam(1) =  UniConvDateAToB(arrParam(1),parent.gServerDateFormat,gDateFormat)

	'arrParam(0)	= frm1.txtBdgYymmFr.text								'  Code Condition
   	'arrParam(1)	= frm1.txtBdgYymmTo.Text
	arrParam(2)		= lgUsrIntCd                            ' 자료권한 Condition  
	'arrParam(3)	= frm1.txtDeptCd.value
	arrParam(4)		= "F"									' 결의일자 상태 Condition  
	
	' 권한관리 추가 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
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
	Dim strStartDt, strEndDt
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.hOrgChangeId.value=arrRet(2)
		strStartDt = UniConvDateAToB(arrRet(4),parent.gDateFormat,parent.gServerDateFormat) 
		strEndDt = UniConvDateAToB(arrRet(5),parent.gDateFormat,parent.gServerDateFormat) 
		frm1.txtBdgYymmFr.Text = UNIMonthClientFormat(strStartDt) 
		frm1.txtBdgYymmTo.Text = UNIMonthClientFormat(strEndDt)

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

 '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Sub InitComboBox()
		
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2000", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCtrlFg ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F2010", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboCtrlUnit ,lgF0  ,lgF1  ,Chr(11))

End Sub


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

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Call FncSetToolBar("New")
	Call InitComboBox
	
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

Sub txtBdgYymmFr_DblClick(Button)
    If Button = 1 Then
       frm1.txtBdgYymmFr.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtBdgYymmFr.Focus       
    End If
End Sub

Sub txtBdgYymmTo_DblClick(Button)
    If Button = 1 Then
       frm1.txtBdgYymmTo.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtBdgYymmTo.Focus       
    End If
End Sub

Sub txtBdgYymmFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBdgYymmTo.focus
	   Call MainQuery
	End If   
End Sub


Sub txtBdgYymmTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtBdgYymmFr.focus
	   Call MainQuery
	End If   
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
	Dim lgF2By2
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
    tmpBdgYymmddFr = UniConvDateAToB(frm1.txtBdgYymmFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(frm1.txtBdgYymmTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("M", +1, tmpBdgYymmddTo,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("D", -1, tmpBdgYymmddTo,parent.gServerDateFormat)			

	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddFr , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddTo , "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
'		msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere
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

Sub txtBdgCdFr_OnChange()
	frm1.txtBdgNmFr.value = ""
End Sub

Sub txtBdgCdTo_OnChange()
	frm1.txtBdgNmTo.value = ""
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
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
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	
    gMouseClickStatus = "SP2C"	'Split 상태코드 
        
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

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow And NewRow > 0 Then
        Call SetSpreadColumnValue("A",frm1.vspdData,NewCol,NewRow)    
        Call SetSpread2(NewRow)
	End If
End Sub

'==========================================================================================
' Row 이동시 vspdData2에 데이터 Query 실행 
'==========================================================================================
Sub SetSpread2(Row)

'on Error Resume Next
'Err.Clear

    Dim ii
    Dim TempGetSqlSelectListA   

	' For ii = 1 to UBound(lgKeyPos)
    '    frm1.vspdData.Col = lgKeyPos(ii)
     '   frm1.vspdData.Row = Row
     '   lgKeyPosVal(ii) = frm1.vspdData.text
	 'Next
	      
     Call DbQuery("B")

    'frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

       lgStrPrevKey_B   = ""                                  'initializes Previous Key
     lgSortKey_B      = 1
     
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	frm1.vspdData.Row = Row
  
    TempGetSqlSelectListA    = Split(EnCoding(GetSQLSelectList("A")),",")
    For ii = LBound(TempGetSqlSelectListA) To UBound(TempGetSqlSelectListA)
		
		With frm1
			.vspdData.Col = ii + 1
			Select Case TempGetSqlSelectListA(ii) 
			Case "B.BDG_CTRL_UNIT"
				.cboCtrlUnit.value = Trim(.vspdData.Text)
			Case "B.ACCT_CTRL_FG"
				.cboCtrlFg.value   = Trim(.vspdData.Text)
			Case "B.ADD_FG"
				If .vspdData.Text = "1" Then
					.txtAddFg.value = "추가가능"
				Else
					.txtAddFg.value = ""
				End If
			Case "B.DIVERT_FG"
				If .vspdData.Text = "1" Then
					.txtDivertFg.value = "이월가능"
				Else
					.txtDivertFg.value = ""
				End If
			Case "B.TRANS_FG"
				If .vspdData.Text = "1" Then
					.txtTransFg.value = "전용가능"
				Else
					.txtTransFg.value = ""
				End If
			End Select
		End With
	Next
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("00000000001")

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

End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If    
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("A")
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgStrPrevKey_B <> "" Then                        '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("B")
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
	Dim strFrYear, strFrMonth, strFrDay
	Dim strToYear, strToMonth, strToDay
	
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear     

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

	'---------------------------------------------
	' 예산년월 조회기간 Check
	'---------------------------------------------
	
    If CompareDateByFormat(frm1.txtBdgYymmFr.Text, frm1.txtBdgYymmTo.Text, frm1.txtBdgYymmFr.Alt, frm1.txtBdgYymmTo.Alt, _
						"970025", frm1.txtBdgYymmFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtBdgYymmFr.focus														'⊙: GL Date Compare Common Function
			Exit Function
	End if
	
	Call ExtractDateFrom(frm1.txtBdgYymmFr.Text,frm1.txtBdgYymmFr.UserDefinedFormat,parent.gComDateType,strFrYear,strFrMonth,strFrDay)
    strFrDt = strFrYear & strFrMonth
    
    Call ExtractDateFrom(frm1.txtBdgYymmTo.Text,frm1.txtBdgYymmTo.UserDefinedFormat,parent.gComDateType,strToYear,strToMonth,strToDay)
    strToDt = strToYear & strToMonth
	'---------------------------------------------
	' 예산코드 조회조건 Check
	'---------------------------------------------
	frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
	frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
	
	If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus
			Exit Function
		End If
	End If
	
	IF NOT CheckOrgChangeId Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
		Exit Function
	End if
	Call FncSetToolBar("New")	
	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery("A")															'☜: Query db data

    FncQuery = True		
End Function


Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo
 
	CheckOrgChangeId = True
 
	With frm1
    tmpBdgYymmddFr = UniConvDateAToB(frm1.txtBdgYymmFr,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(frm1.txtBdgYymmTo,parent.gDateFormatYYYYMM,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("M", +1, tmpBdgYymmddTo,parent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("D", -1, tmpBdgYymmddTo,parent.gServerDateFormat)			
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddFr , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddTo , "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

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
	Call parent.FncExport(parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
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
	Dim strVal
	Dim ColYymm1, ColYymm2, ii
	Dim strDt
	Dim strYear
	Dim	strMonth
	Dim	strDay
    Dim TempGetSqlSelectListA
    Dim TempGetSqlSelectListB
    DbQuery = False
    
    'Err.Clear                                                               '☜: Protect system from crashing

    Call LayerShowHide(1)

    With frm1
        If iOpt = "A" Then
            TempGetSqlSelectListA    = Split(EnCoding(GetSQLSelectList("A")),",")
            For ii = LBound(TempGetSqlSelectListA) To UBound(TempGetSqlSelectListA)
                If TempGetSqlSelectListA(ii) = "A.BDG_YYYYMM" Then 
                    ColYymm1 = ii	'예산년월 컬럼 
                    Exit For
                End If
            Next
        '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
            strVal = BIZ_PGM_ID & "?txtBdgYymmFr=" & strFrDt
            strVal = strVal & "&txtBdgYymmTo=" & strToDt
            strVal = strVal & "&txtDeptCd=" & UCase(Trim(.txtDeptCd.value))
            strVal = strVal & "&txtBdgCdFr=" & UCase(Trim(.txtBdgCdFr.value))
            strVal = strVal & "&txtBdgCdTo=" & UCase(Trim(.txtBdgCdTo.value))
            strVal = strVal & "&txtDeptCd_Alt=" & .txtDeptCd.Alt
            strVal = strVal & "&txtBdgCdFr_Alt=" & .txtBdgCdFr.Alt
            strVal = strVal & "&txtBdgCdTo_Alt=" & .txtBdgCdTo.Alt
            strVal = strVal & "&txtColYymm=" & ColYymm1
            strVal = strVal & "&txtDateType=" & parent.gComDateType
			strVal = strVal & "&OrgChangeId="       & Trim(frm1.hOrgChangeId.Value)
        Else   
            TempGetSqlSelectListB    = Split(EnCoding(GetSQLSelectList("B")),",")
            For ii = LBound(TempGetSqlSelectListB) To UBound(TempGetSqlSelectListB)
                If TempGetSqlSelectListB(ii) = "A.CUR_BDG_YYYYMM" Then 
                   ColYymm2 = ii	'사용년월 컬럼 
                    Exit For
                End If
            Next

           Call ExtractDateFrom(GetKeyPosVal("A",2),parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
            strDt = strYear & strMonth
            strVal = BIZ_PGM_ID1 & "?txtBdgCd=" & GetKeyPosVal("A",1)'lgKeyPosVal(1)
            strVal = strVal & "&txtBdgYymm=" & strDt
            strVal = strVal & "&txtDeptCd=" & GetKeyPosVal("A",3)'lgKeyPosVal(3)
            strVal = strVal & "&txtOrgChangeId=" & GetKeyPosVal("A",4)'lgKeyPosVal(4)
            strVal = strVal & "&txtInternalCd=" & GetKeyPosVal("A",5)'lgKeyPosVal(5)
            strVal = strVal & "&txtColYymm=" & ColYymm2
            strVal = strVal & "&txtDateType=" & parent.gComDateType

        End If   
       
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        If iOpt = "A" Then
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_A                      '☜: Next key tag
            strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
            strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
            strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A")) 
        Else   
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_B                      '☜: Next key tag
            strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
            strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
            strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
        End If   


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
	
    frm1.vspdData.focus
    Call SetSpreadColumnValue("A",frm1.vspdData, ,1)    

    If frm1.vspdData.MaxRows > 0 Then Call SetSpread2(1)	'Sub_Query 실행 

    '-----------------------
    'Reset variables area
    '-----------------------
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	Call FncSetToolBar("Query")
	'Call SetGridFocus("A")	
	Set gActiveElement = document.activeElement 
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################


'=========================================================================================================
' Function Name : SetPopUpInitialInf
' Function Desc : set popup initial information according to iOpt
'===========================================================================================================
Sub SetPopUpInitialInf(Byval iOpt)
	Dim ii,kk	
	Dim iCast

    lgSortFieldNm_T  = ""
    lgSortFieldCD_T  = ""

    ReDim lgPopUpR_T(parent.C_MaxSelList - 1,1)

    For ii = 0 To UBound(lgFieldNM_T) - 1                                            'Sort 대상 list  저장 
        iCast = lgDefaultT_T(ii)
        If  IsNumeric(iCast) Or Trim(lgDefaultT_T(ii)) = "V" Then
            If IsNumeric(iCast) Then 
               If IsBetween(1,parent.C_MaxSelList,CInt(iCast)) Then                         'Sort정보 default값 저장 
                  lgPopUpR_T(CInt(lgDefaultT_T(ii)) - 1,0) = Trim(lgFieldCD_T(ii))
                  lgPopUpR_T(CInt(lgDefaultT_T(ii)) - 1,1) = "ASC"
               End If
            End If
            lgSortFieldNm_T = lgSortFieldNm_T & Trim(lgFieldNM_T(ii)) & Chr(11)
            lgSortFieldCD_T = lgSortFieldCD_T & Trim(lgFieldCD_T(ii)) & Chr(11)
        End If
    Next
    
    If iOpt = "1" Then          
       lgSortFieldCD_A       = Split (lgSortFieldCD_T ,Chr(11))                      '배열화 
       lgSortFieldNM_A       = Split (lgSortFieldNm_T ,Chr(11))

    Else
       lgSortFieldCD_B       = Split (lgSortFieldCD_T ,Chr(11))
       lgSortFieldNM_B       = Split (lgSortFieldNm_T ,Chr(11))          
    End If       
    
End Sub 


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

        Frm1.vspdData2.Row = 1
        Frm1.vspdData2.Col = 1
        Frm1.vspdData2.Action = 1
		
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
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>예산년월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYymmFr" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT=시작예산년월 id=fpBdgYymmFr></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYymmTo" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT=종료예산년월 id=fpBdgYymmTo></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" MAXLENGTH="10" SIZE=10 ALT ="부서코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
														   <INPUT NAME="txtDeptNm" MAXLENGTH="40" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="부서명" tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>시작예산</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCdFr" MAXLENGTH="18" SIZE=10 ALT ="시작예산코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBdgCdFr.Value, 'BdgCdFr')">
														   <INPUT NAME="txtBdgNmFr" MAXLENGTH="30" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="시작예산명" tag="24">
									</TD>
									<TD CLASS="TD5" NOWRAP>종료예산</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCdTo" MAXLENGTH="18" SIZE=10 ALT ="종료예산코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBdgCdTo.Value, 'BdgCdTo')">
														   <INPUT NAME="txtBdgNmTo" MAXLENGTH="30" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="종료예산명" tag="24">
									</TD>
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
								<TD HEIGHT="60%" COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=vaSpread1 tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="60%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=vaSpread2 tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH="40%">
									<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT:100%">
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS=TD5 NOWRAP>예산통제단위</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCtrlUnit" ALT="예산통제단위" STYLE="WIDTH: 100px" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>구분</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCtrlFg" ALT="구분" STYLE="WIDTH: 100px" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>추가여부</TD>
												<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAddFg" SIZE=12 STYLE="TEXT-ALIGN:left" ALT ="추가가능" tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>이월여부</TD>
												<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDivertFg" SIZE=12 STYLE="TEXT-ALIGN:left" ALT ="이월가능" tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>전용여부</TD>
												<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransFg" SIZE=12 STYLE="TEXT-ALIGN:left" ALT ="전용가능" tag="24"></TD>
											</TR>
										</TABLE>
									</FIELDSET>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId" tag="">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

