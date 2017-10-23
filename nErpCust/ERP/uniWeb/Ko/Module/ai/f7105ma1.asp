<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : f7105ma1
'*  4. Program Name         : 선수금기초치 등록 
'*  5. Program Desc         : 선수금기초치 등록 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/09/25
'*  8. Modified date(Last)  : 2002/11/19
'*  9. Modifier (First)     : Hee Jung, Kim
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'=======================================================================================================
'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
'@PGM_ID
Const BIZ_PGM_QRY_ID	= "f7105mb1.asp"											'비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID	= "f7105mb2.asp"											'비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID	= "f7105mb3.asp"											'비지니스 로직 ASP명 

Const PreReceiptJnlType = "PR"

Const gIsShowLocal = "Y"

'@Global_Var
Dim IsOpenPop						                        'Popup
Dim	lgFormLoad
Dim	lgQueryOk
Dim lgstartfnc

'2002.01.10 추가된 사항 
<%
Dim dtToday 
dtToday = GetSvrDate 
%>

'======================================================================================================
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                    'Indicates that no value changed
    '---- Coding part--------------------------------------------------------------------
    lgstartfnc=False
	lgFormLoad=True	    
    lgStrPrevKey = ""                                           'initializes Previous Key
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtDocCur.value = parent.gCurrency
	frm1.txtPrrcptDt.text = UniConvDateAToB("<%=dtToday%>",parent.gServerDateFormat,gDateFormat)
	frm1.txtgldt.text= frm1.txtPrrcptDt.text
<%	If gIsShowLocal <> "N" Then	%>
	frm1.txtXchRate.Text	= 1
<%  Else %>
	frm1.txtXchRate.Value	= 1
<%  End If %>
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	lgBlnFlgChgValue = False
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

Function OpenPopupTempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'결의전표번호 
	arrParam(1) = ""							'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
'   Function Name : OpenPopupPR()
'   Function Desc : 
'=======================================================================================================
Function OpenPopupPR()
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("f7105ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f7105ra1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrRet = window.ShowModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		frm1.txtPrrcptNo.focus
		Exit Function
	Else
		frm1.txtPrrcptNo.value = arrRet(0)
	End If	

	frm1.txtPrrcptNo.focus
End Function

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("DeptPopupDtA2")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtPrrcptDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		     Case "0"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtPrrcptDt.text = arrRet(3)
				           
				Call txtDeptCd_OnChange()
				frm1.txtDeptCd.focus 
        End Select
	End With
End Function  
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function
	
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
       frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
		frm1.txtBpCd.focus
		lgBlnFlgChgValue = True
	End If

End Function
'=======================================================================================================
'	Name : OpenPopup()
'	Description : 공통코드팝업 
'=======================================================================================================
Function OpenPopup(strCode, strWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case strWhere
		Case "BP"
			If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "거래처 팝업"									' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER A" 								' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "거래처"										' 조건필드의 라벨 명칭 

		    arrField(0) = "A.BP_CD"											' Field명(0)
		    arrField(1) = "A.BP_NM"											' Field명(1)
    
		    arrHeader(0) = "거래처코드"									' Header명(0)
			arrHeader(1) = "거래처명"									' Header명(1)

		Case "CURR"
			If frm1.txtDocCur.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = "통화 팝업"									' 팝업 명칭 
			arrParam(1) = "B_CURRENCY A"									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "통화"										' 조건필드의 라벨 명칭 

		    arrField(0) = "A.CURRENCY"										' Field명(0)
		    arrField(1) = "A.CURRENCY_DESC"									' Field명(1)
    
		    arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화명"										' Header명(1)

		Case "PRRCPTTYPE"
			If frm1.txtPrrcptType.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtPrrcptType.Alt								' 팝업 명칭 
			arrParam(1) = "a_jnl_item"	 									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPrrcptType.Value)						' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "jnl_type =  " & FilterVar(PreReceiptJnlType, "''", "S") & " "			' Where Condition
			arrParam(5) = frm1.txtPrrcptType.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "JNL_CD"											' Field명(0)
		    arrField(1) = "JNL_NM"											' Field명(1)
    
		    arrHeader(0) = frm1.txtPrrcptType.Alt								' Header명(0)
			arrHeader(1) = frm1.txtPrrcptTypeNm.Alt								' Header명(1)

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case strWhere
			Case "BP"
				frm1.txtBpCd.focus
			Case "CURR"
				frm1.txtDocCur.focus
			Case "PRRCPTTYPE"
				frm1.txtPrrcptType.focus	
		End Select	
			Exit Function
	End If

	Select Case strWhere

		Case "BP"
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNm.value = arrRet(1)
			lgBlnFlgChgValue = True
			frm1.txtBpCd.focus
		Case "CURR"
			frm1.txtDocCur.value = arrRet(0)
			Call txtDocCur_OnChange()
			Call XchLocRate()
			lgBlnFlgChgValue = True
			frm1.txtDocCur.focus
		Case "PRRCPTTYPE"
			frm1.txtPrrcptType.value = arrRet(0)
			frm1.txtPrrcptTypeNm.value = arrRet(1)
			frm1.txtPrrcptType.focus	
		Case Else
			Exit Function
	End Select
End Function
 
'======================================================================================================
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     'Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitVariables                                                      'Initializes local global variables

	Call FncNew																'add.여기서 call할때 SetDefaultVal()도 함깨 call한다.
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPrrcptDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPrrcptDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtPrrcptDt.Focus        
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGlDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtGlDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtGlDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtGlDt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtPrrcptAmt_Change()
	If UCase(Trim(frm1.txtDocCur.value)) <> UCase(parent.gCurrency) Then
		frm1.txtPrrcptLocAmt.Text = "0"
	End If
    lgBlnFlgChgValue = True
End Sub

Sub txtPrrcptLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtPrrcptDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtPrrcptDt.Text, gDateFormat,""), "''", "S") & "))"			
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hOrgChangeId.value = ""
	Else 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.value = Trim(arrVal2(2))
		Next	
	End If
	'----------------------------------------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : txtPrrcptDt_Change
'   Event Desc : 
'==========================================================================================
Sub txtPrrcptDt_Change()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	
	
	<% If gIsShowLocal <> "N" Then %>
	frm1.txtXchRate.Text = 0
	frm1.txtPrrcptLocAmt.text = 0
    <% Else %>	
	frm1.txtXchRate.Value = 0
	frm1.txtPrrcptLocAmt.value = 0
    <% End If %>
    
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
				If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtPrrcptDt.Text <> "") Then
		'----------------------------------------------------------------------------------------
					strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
					strFrom		=			 " b_acct_dept(NOLOCK) "		
					strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
					strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
					strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
					strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtPrrcptDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124600","X","X","X")
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
					Else
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						jj = Ubound(arrVal1,1)
						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))			
							frm1.hOrgChangeId.value = Trim(arrVal2(2))
						Next	
					End If 
				End If
			End With
		'----------------------------------------------------------------------------------------
		End If
	End IF
	Call XchLocRate()
End Sub

'==========================================================================================
'   Event Name : txtXchRate_Change
'   Event Desc : 
'==========================================================================================
Sub txtXchRate_Change()
    lgBlnFlgChgValue = True
    
 If lgQueryOk <> true Then    
		With Frm1    
			.txtPrrcptLocAmt.text="0"
		End with 
	End if    


End Sub

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'======================================================================================================

'======================================================================================================
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'======================================================================================================
Function FncQuery() 
    Dim IntRetCD
    
    FncQuery = False                                                        
	lgstartfnc = True           
    Err.Clear                                                               'Protect system from crashing

	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call InitVariables                                                      'Initializes local global variables
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
	'-----------------------
    'Query function call area
    '----------------------- 
    frm1.hCommand.value = "LOOKUP"
    Call DbQuery															'Query db data
       
    FncQuery = True	
    lgstartfnc = False       														
	
	Set gActiveElement = document.activeElement   
End Function


'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False            
	lgstartfnc = True    	                                              
	
	'-----------------------
	'Check previous data area
	'-----------------------
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
	
	Call txtDocCur_OnChange()
	
	frm1.txtPrrcptAmt.text = 0
	
	Call InitVariables                                                      'Initializes local global variables
	Call SetDefaultVal
	Call SetToolbar("111010000000111")

    frm1.txtPrrcptNo.focus 
	Set gActiveElement = document.activeElement
	
	FncNew = True                  
	lgFormLoad = True							' tempgldt read
    lgQueryOk = False
    lgstartfnc = False    
	Set gActiveElement = document.activeElement   
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete() 
    Dim IntRetCD
	FncDelete = False
		
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")   '삭제하시겠습니까?  
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Precheck area
	'-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")                                
    	Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete                                                          '☜: Delete db data
    
    FncDelete = True
	Set gActiveElement = document.activeElement       
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False
	
	Err.Clear                                                               
    
    If lgBlnFlgChgValue = False  Then  '⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '⊙: Display Message(There is no changed data.)
        Exit Function
    End If

    If len(frm1.txtPrrcptAmt.Text)= 0 then
    	Call DisplayMsgBox("970021","X",frm1.txtPrrcptAmt.alt,"X")  
		Exit Function
    ElseIf UNICDbl(frm1.txtPrrcptAmt.Text) = 0 then
		Call DisplayMsgBox("141704","X",frm1.txtPrrcptAmt.alt,"X")  
		Exit Function
    End if 
    
    If Not chkField(Document, "2") Then               '⊙: Check required field(Single area)
       Exit Function
    End If
	
	'-----------------------
	'Save function call area
	'-----------------------
	Call DbSave				                                                '☜: Save db data
	
	FncSave = True
	Set gActiveElement = document.activeElement   
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
	
	Set gActiveElement = document.activeElement   
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev()
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                     '밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If

	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If

	frm1.hCommand.value = "PREV"
	Call DbQuery
	
	Set gActiveElement = document.activeElement   
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                     '밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If

	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If

	frm1.hCommand.value = "NEXT"
	Call DbQuery
	Set gActiveElement = document.activeElement   
	
End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)										
	Set gActiveElement = document.activeElement   
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
    Set gActiveElement = document.activeElement   
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
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
	Set gActiveElement = document.activeElement   
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete() 
    Dim strVal
    
    DbDelete = False														'⊙: Processing is NG 
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtPrrcptNo=" & Trim(frm1.txtPrrcptNo.value)			'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         '⊙: Processing is NG
	Set gActiveElement = document.activeElement   
End Function


'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()												        '삭제 성공후 실행 로직 
	'Call FncNew()
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Contents  Field
	Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
	
	Call txtDocCur_OnChange()
	
	frm1.txtPrrcptAmt.text = 0
	
	Call InitVariables                                                      'Initializes local global variables
	Call SetDefaultVal
	Call SetToolbar("111010000000111")

    frm1.txtPrrcptNo.focus 
	Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	DbQuery = False                                                         
	
	Call LayerShowHide(1)
	
	Dim strVal
	
	With frm1
       	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'☜: 
       	strVal = strVal & "&txtPrrcptNo=" & Trim(.txtPrrcptNo.value)	'조회 조건 데이타 
       	strVal = strVal & "&txtCommand=" & Trim(.hCommand.value)
       	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										'비지니스 ASP 를 가동 
	
	DbQuery = True                                                          
	lgQueryOk = True 	   
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================================
Function DbQueryOk()
	Dim strTemp													'조회 성공후 실행로직 
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	lgQueryOK = True 
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field	
	Call SetToolbar("111110001101111")									'버튼 툴바 제어 
	
	If gIsShowLocal <> "N" Then
		strTemp = frm1.txtXchRate.Text
	Else
	    strTemp = frm1.txtXchRate.Value 
	End if
			
	Call txtDocCur_OnChange()
	If gIsShowLocal <> "N" Then
		frm1.txtXchRate.Text = strTemp
    Else
        frm1.txtXchRate.value  = strTemp
    End if		
	lgBlnFlgChgValue = False
	lgQueryOK = false 
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	DbSave = False                                                          
	
	On Error Resume Next                                                   
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = parent.UID_M0002											'☜: 저장 상태 
		.txtFlgMode.value = lgIntFlgMode									'☜: 신규입력/수정 상태 
	End With
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data 연결 규칙 
	' 0: Flag , 1: Row위치, 2~N: 각 데이타 

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'☜: 저장 비지니스 ASP 를 가동 
	
	DbSave = True                                                           
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
   	lgBlnFlgChgValue = False	

	Call FncQuery
End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If gIsShowLocal <> "N" Then
		frm1.txtXchRate.Text = 0
    Else
		frm1.txtXchRate.value = 0
    End if
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	End If
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 선수금액 
		ggoOper.FormatFieldByObjectOfCur .txtPrrcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 청산금액 
		ggoOper.FormatFieldByObjectOfCur .txtSttlAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

'===================================== XchLocRate()  ======================================
'	Name : XchLocRate()
'	Description : 통화가 변경될경우 통화에 따른 자국금액 
'====================================================================================================
Sub XchLocRate()

	If gIsShowLocal <> "N" Then
		frm1.txtPrrcptLocAmt.text = "0"
		frm1.txtXchRate.text = "0"
	else
		frm1.txtPrrcptLocAmt.value = "0"
		frm1.txtXchRate.value = "0"
	end if

End Sub

Sub chkLimitFg_onchange()
	If frm1.chkLimitFg.checked = True Then
		frm1.txtLimitFg.value = "Y"
	Else
		frm1.txtLimitFg.value = "N"	
	End If
	lgBlnFlgChgValue = True	
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>선수금기초치등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
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
									<TD CLASS="TD5" NOWRAP>선수금번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtPrrcptNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="선수금번호" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrrcptNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopupPR"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>선수금유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrrcptType" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="선수금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrrcptType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('','PRRCPTTYPE')">&nbsp;<INPUT TYPE=TEXT NAME="txtPrrcptTypeNm" SIZE=25 tag="24" ALT="선수금유형명"></TD>
								<TD CLASS="TD5" NOWRAP><LABEL FOR=chkConfFg>여신관리</LABEL></TD>
								<TD CLASS="TD6" NOWRAP><INPUT type="checkbox" CLASS="STYLE CHECK"  NAME=chkLimitFg ID=chkLimitFg tag="1" onclick=chkLimitFg_onchange()></TD>
							</TR>						
							<TR>
								<TD CLASS="TD5" NOWRAP>부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="부서" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 tag="24" ALT="회계부서명"></TD>
								<TD CLASS="TD5" NOWRAP>입금일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDateTime1_txtPrrcptDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="거래처코드" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 'BP')">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24" ALT="거래처명"></TD>
								<TD CLASS="TD5" NOWRAP>참조번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=30 MAXLENGTH=30 tag="21XXXU" ALT="참조번호" ></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>								
							<TR>
								<TD CLASS="TD5" NOWRAP>거래통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 MAXLENGTH=3 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.value, 'CURR')"></TD>
								
								<TD CLASS="TD5" NOWRAP>환율</TD>
   	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle1_txtXchRate.js'></script></TD>

							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtXchRate" TABINDEX="-1">
<%	End If %>								
							<TR>
								<TD CLASS="TD5" NOWRAP>선수금액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle2_txtPrrcptAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>	      								
								<TD CLASS="TD5" NOWRAP>선수금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle3_txtPrrcptLocAmt.js'></script></TD>
							</TR>
							<TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtPrrcptLocAmt" TABINDEX="-1">
<%	End If %>								
								<TD CLASS="TD5" NOWRAP>반제금액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle4_txtClsAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>	 								
								<TD CLASS="TD5" NOWRAP>반제금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle5_txtClsLocAmt.js'></script></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtClsLocAmt" TABINDEX="-1">
<%	End If %>								
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>청산금액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle6_txtSttlAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>								
								<TD CLASS="TD5" NOWRAP>청산금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle7_txtSttlLocAmt.js'></script></TD>
							</TR>
							<TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtSttlLocAmt" TABINDEX="-1">
<%	End If %>										
								<TD CLASS="TD5" NOWRAP>잔액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle8_txtBalAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>								
								<TD CLASS="TD5" NOWRAP>잔액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDoubleSingle9_txtBalLocAmt.js'></script></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtBalLocAmt" TABINDEX="-1">
<%	End If %>							
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="24" ALT="회계전표번호"></TD>
								<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=18 tag="24" ALT="G/L번호"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>전표일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f7105ma1_fpDateTime1_txtGlDt.js'></script></TD>

								<TD CLASS="TD5" NOWRAP>프로젝트</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProjectNo"  SIZE=14 MAXLENGTH=25 TAG="21xxxU" ALT="프로젝트"></TD>	                     
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP ><INPUT TYPE=TEXT NAME="txtPrrcptDesc" SIZE=50 MAXLENGTH=100 tag="2X" ALT="비고"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
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
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     tag="24" TABINDEX="-1">
<INPUT TYPE=TEXT   NAME="hDocumentNo1"   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCommand"       tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtLimitFg"     tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

