<%@ LANGUAGE="VBSCRIPT" %>


<!--
'**********************************************************************************************
'*  1. Module Name          : FINANCE
'*  2. Function Name        : 
'*  3. Program ID           : a3220ra1
'*  4. Program Name         : 채권반제채권번호Pop
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/04/17
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : U&I (KIM HEA RAN)
'* 11. Comment              :
'* 12. Common Coding Guide  : 
'* 13. History              :
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<!--'=====================================  1.1.2 공통 Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit                              '☜: indicates that All variables must be declared in advance



<!-- #Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim lgIsOpenPop                                          
Dim IsOpenPop    

Const BIZ_PGM_ID        = "a3125rb1.ASP"
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_SHEETMAXROWS_D  = 30  
Const C_MaxKey          = 2	

Dim  arrReturn
Dim  arrParent
Dim  arrParam					
	
	'------ Set Parameters from Parent ASP ------ 
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)
top.document.title = "입금반제팝업"


Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

dtToday = "<%=GetSvrDate%>"
Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

EndDate = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
StartDate = UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)

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
   
   Redim arrReturn(0)
   
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
           
	Self.Returnvalue = arrReturn
	
End Sub


'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'========================================================================================================= 
Sub  SetDefaultVal()

	frm1.txtFrRcptDt.Text	= StartDate
	frm1.txtToRcptDt.Text	= EndDate
	frm1.hOrgChangeId.value = PopupParent.gChangeOrgId    
End Sub


'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub  LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '☆: 
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","RA") %>
End Sub



'===========================================  2.3.1 OkClick()  ==========================================
'	Name : OkClick()
'	Description : Return Array to Opener Window when OK button click
'				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.
'========================================================================================================	
Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
		
		if frm1.vspdData.ActiveRow > 0 Then 			
		
			intInsRow = 0
			
			Redim arrReturn(1)
			
			For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
				frm1.vspdData.Row = intRowCnt + 1
			
				If frm1.vspdData.SelModeSelected Then
				   frm1.vspdData.Col = 1
				   arrReturn(intColCnt) = frm1.vspdData.Text										
				   intInsRow = intInsRow + 1					
				End IF
			Next
			
		End if			
		
		
	Self.Returnvalue = arrReturn
	Self.Close()
					
End Function



'=========================================  2.3.2 CancelClick()  ========================================
'	Name : CancelClick()
'	Description : Return Array to Opener Window for Cancel button click
'========================================================================================================
Function CancelClick()

		Self.Close()			

End Function




'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub  InitSpreadSheet()
    
    frm1.vspdData.operationmode = 5
    Call SetZAdoSpreadSheet("a3125RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock()

    
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()

    .vspdData.ReDraw = True

    End With
End Sub

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFrRcptDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToRcptDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  
	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(popupparent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
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
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.hOrgChangeId.value=arrRet(2)
		frm1.txtFrRcptDt.text = arrRet(4)
		frm1.txtToRcptDt.text = arrRet(5)
		frm1.txtDeptCd.focus		
End Function
  
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
	arrParam(2) = ""							'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet,iWhere)
	End If	
End Function
'==========================================  2.2 OpenPopUp()  ========================================
'	Name : OpenPopUp()
'	Description : 
'========================================================================================================= 
Function OpenPopUp(Byval strCode, iMode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iMode
	
	    Case 0
			
		Case 1
			arrParam(0) = "수금처 팝업"
			arrParam(1) = "B_Biz_Partner"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "수금처코드"

			arrField(0) = "bp_CD"
			arrField(1) = "bp_NM"
    
			arrHeader(0) = "수금처코드"
			arrHeader(1) = "수금처명"
			
		
	End Select		
	
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iMode)
		Exit Function
	Else
		Call SetPopUp(arrRet, iMode)
	End If	

End Function

'==========================================  2.2 SetPopUp()  ========================================
'	Name : EscPopUp()
'	Description : 
'========================================================================================================= 
Function EscPopUp(iMode)

	With frm1
	  Select Case iMode
	    CASE 0
		 CASE 1
		    .txtBPCd.focus
	  End Select    
	End With
End Function

'==========================================  2.2 SetPopUp()  ========================================
'	Name : SetPopUp()
'	Description : 
'========================================================================================================= 
Function SetPopUp(Byval arrRet, iMode)

	With frm1
	  Select Case iMode
		CASE 0
		 CASE 1
		    .txtBPCd.value = arrRet(0)
		    .txtBPNm.value = arrRet(1)
		    .txtBPCd.focus
	  End Select    
	End With
End Function


 
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

	Call LoadInfTB19029
   Call ggoOper.FormatField(Document, "1", PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
   Call ggoOper.FormatField(Document, "2", PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
    
   Call ggoOper.LockField(Document, "N")
   
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

 

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************

'==========================================================================================
'   Event Name : txtFrRcptDt
'   Event Desc :
'==========================================================================================

Sub  txtFrRcptDt_DblClick(Button)

	if Button = 1 then
		frm1.txtFrRcptDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtFrRcptDt.Focus
	End if

End Sub


Sub txtFrRcptDt_KeyPress(KeyAscii)
	On Error Resume Next
	
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF

End Sub

'==========================================================================================
'   Event Name : txtToRcptDt
'   Event Desc :
'==========================================================================================
Sub  txtToRcptDt_DblClick(Button)

	if Button = 1 then
		frm1.txtToRcptDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtToRcptDt.Focus
	End if

End Sub

Sub txtToRcptDt_KeyPress(KeyAscii)

	On Error Resume Next

	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF

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
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    
End Sub



'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
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
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
End Sub



Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function



Sub  vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub


'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function  FncQuery() 
    Dim IntRetCD
    
    FncQuery = False
    
    Err.Clear
	 
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("800600","X",frm1.txtFrRcptDt.alt,"X")            '⊙: Display Message(There is no changed data.)
		Exit Function
	End if
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    
    Call DbQuery

    FncQuery = True		
End Function


'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function  FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function  DbQuery() 
	Dim strVal

   DbQuery = False
    
   Err.Clear            
    
	Call LayerShowHide(1)

    
    With frm1

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
			strVal = BIZ_PGM_ID & "?txtFrRcptDt=" & Trim(.txtFrRcptDt.Text)
			strVal = strVal & "&txtToRcptDt="     & Trim(.txtToRcptDt.Text)
			strVal = strVal & "&txtFrRcptNo="     & Trim(.txtFrRcptNo.value)
			strVal = strVal & "&txtToRcptNo="     & Trim(.txtToRcptNo.value)
			strVal = strVal & "&txtdeptcd="       & Trim(.txtdeptcd.value)
			strVal = strVal & "&txtBPcd="         & Trim(.txtBPcd.value)	
			strVal = strVal & "&txtDeptCd_Alt=" & Trim(.txtdeptcd.alt)		
    '--------- Developer Coding Part (End) ------------------------------------------------------------									
			strVal = strVal & "&txtOrgChangeId=" & Trim(.hOrgChangeId.Value)
			strVal = strVal & "&lgPageNo="       & lgPageNo         
			strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
         
		  Call RunMyBizASP(MyBizASP, strVal)
		  
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
    lgBlnFlgChgValue = True                                                 'Indicates that no value changed

End Function




'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
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
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFrRcptDt.Text, popupparent.gDateFormat,""), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToRcptDt.Text, popupparent.gDateFormat,""), "''", "S") & ") "
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


Sub txtBpCd_OnChange()
    
   
    Call CommonQueryRs("bp_cd, bp_nm", "b_biz_partner", "bp_cd =  " & FilterVar(frm1.txtBpCD.value , "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
    if lgF0 <> "" and lgF0 <> "X" then
       frm1.txtBPCD.value = Left(lgF0, Len(lgF0)-1)
       frm1.txtBpNm.value = Left(lgF1, Len(lgF1)-1)
    else
       frm1.txtBPCD.value = ""
       frm1.txtBpNm.value = ""
    End if

End Sub

Sub txtDeptCd_OnChange()
    
   
    Call CommonQueryRs("dept_cd, dept_nm", "b_acct_dept", "dept_cd =  " & FilterVar(frm1.txtdeptCD.value , "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
    if lgF0 <> "" and lgF0 <> "X" then
       frm1.txtDeptCD.value = Left(lgF0, Len(lgF0)-1)
       frm1.txtDeptNm.value = Left(lgF1, Len(lgF1)-1)
    else
       frm1.txtDeptCD.value = ""
       frm1.txtDeptNm.value = ""
    End if

End Sub

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenOrderBy()
	Dim arrRet
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables()
       Call InitSpreadSheet()       
   End If
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>				
						<TD CLASS=TD5 NOWRAP>입금일자</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/a3125ra1_I146329796_txtFrRcptDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/a3125ra1_I849238234_txtToRcptDt.js'></script>
						</TD>												
						<TD CLASS=TD5 NOWRAP>입금번호</TD>				
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE="Text" NAME="txtFrRcptNo" SIZE=15 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="1XXXXU" ALT="입금번호">&nbsp;~&nbsp;
						<INPUT TYPE="Text" NAME="txtToRcptNo" SIZE=15 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="1XXXXU" ALT="입금번호">
						</TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>부서코드</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag  ="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup">&nbsp;
						<INPUT NAME="txtDeptNm" ALT="부서명" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
						<TD CLASS=TD5 NOWRAP>수금처코드</TD>
						<TD CLASS=TD6 NOWRAP>
						      <INPUT NAME="txtBPCd" ALT="수금처코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag  ="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnBPCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBPCd.Value, 1)">&nbsp;
						      <INPUT NAME="txtBPNm" ALT="수금처명" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
					</TR>
					
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<script language =javascript src='./js/a3125ra1_vspdData_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenOrderBy()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;
					</TD>

				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=1><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=1 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT		TYPE=hidden	 NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
