
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : Prepayment
'*  3. Program ID           : f6105ma1
'*  4. Program Name         : 선급금 기초등록 
'*  5. Program Desc         : 선급금 기초등록 
'*  6. Modified date(First) : 2000/09/22
'*  7. Modified date(Last)  : 2002/11/18
'*  8. Modifier (First)     : 조익성 
'*  9. Modifier (Last)      : Jeong Yong Kyun
'* 10. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================

'=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit                                                             '☜: indicates that All variables must be declared in advance 

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
'@PGM_ID
Const BIZ_PGM_ID = "f6105mb1.asp"											'비지니스 로직 ASP명 
Const PrePaymentJnlType = "PP"

Const gIsShowLocal = "Y"


'@Global_Var
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgNextNo						                                        '☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						                                        

Dim IsOpenPop          
Dim	lgFormLoad
Dim	lgQueryOk
Dim lgstartfnc

<%
If gIsShowLocal <> "N" Then
	Dim dtToday
	dtToday = GetSvrDate
End If
%>
'=======================================================================================================
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================
'=======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                                               'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                'Indicates that no value changed
                                                           
    IsOpenPop = False		
    lgstartfnc=False
	lgFormLoad=True			
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'=======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()
	If gIsShowLocal = "N" Then
		frm1.txtPrpaymDt.Text = UniConvDateAToB("<%=GetSvrDate%>") 'UNIFormatDate("<%=GetSvrDate%>")
		frm1.txtXchRate.Value	= 1
	Else
		frm1.txtPrpaymDt.Text = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,gDateFormat)
		frm1.txtXchRate.Text   = 1
	End If
	frm1.txtglDt.text=frm1.txtPrpaymDt.Text
	frm1.txtDocCur.value = parent.gCurrency
	frm1.hOrgChangeId.value = parent.gChangeOrgId
		
	lgBlnFlgChgValue = False
End Sub

'=======================================================================================================
'Description : 회계전표 생성내역 팝업 
'=======================================================================================================
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

'=======================================================================================================
'Description : 선급금번호 팝업 
'=======================================================================================================
Function OpenPopupPP()
	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName
	
	
	iCalledAspName = AskPRAspName("f6105ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f6105ra1", "X")
		IsOpenPop = False
		Exit Function
	End If	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam), _
		     "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPrpaymNo.focus
		Exit Function
	Else
		frm1.txtPrpaymNo.Value = arrRet(0)
	End If	
	
	frm1.txtPrpaymNo.focus
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

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = frm1.txtPrpaymDt.Text
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
  				If lgQueryOk <> True Then
				     .txtPrpaymDt.text = arrRet(3)
				End If
				           
				call txtDeptCd_OnChange()  
				.txtDeptCd.focus
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
'	Description : 공통코드 팝업 
'=======================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case UCase(iWhere)
		Case "BP"
			If frm1.txtBpCd.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtBpCd.Alt									' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER" 									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtBpCd.value)							' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtBpCd.Alt									' 조건필드의 라벨 명칭 

		    arrField(0) = "BP_CD"											' Field명(0)
		    arrField(1) = "BP_NM"											' Field명(1)
    
		    arrHeader(0) = frm1.txtBpCd.Alt									' Header명(0)
			arrHeader(1) = frm1.txtBpNm.Alt									' Header명(1)
		Case "CURR"
			If frm1.txtDocCur.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtDocCur.Alt								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDocCur.value)						' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "CURRENCY"										' Field명(0)
		    arrField(1) = "CURRENCY_DESC"									' Field명(1)
    
		    arrHeader(0) = frm1.txtDocCur.Alt								' Header명(0)
			arrHeader(1) = "거래통화명"									' Header명(1)
		Case "PRPAYMTYPE"
			If frm1.txtPrpaymType.className = parent.UCN_PROTECTED Then Exit Function
			
			arrParam(0) = frm1.txtPrpaymType.Alt								' 팝업 명칭 
			arrParam(1) = "a_jnl_item"	 									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPrpaymType.Value)						' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "jnl_type =  " & FilterVar(PrePaymentJnlType, "''", "S") & " "			' Where Condition
			arrParam(5) = frm1.txtPrpaymType.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "JNL_CD"											' Field명(0)
		    arrField(1) = "JNL_NM"											' Field명(1)
    
		    arrHeader(0) = frm1.txtPrpaymType.Alt								' Header명(0)
			arrHeader(1) = frm1.txtPrpaymTypeNm.Alt								' Header명(1)
		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
		
	If arrRet(0) = "" Then
		Select Case UCase(iWhere)
			Case "BP"
				frm1.txtBpCd.focus
			Case "CURR"
				frm1.txtDocCur.focus
			Case "PRPAYMTYPE"
				frm1.txtPrpaymType.focus
		End Select	
		Exit Function
	End If
	
	Select Case UCase(iWhere)
		Case "BP"
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNm.value = arrRet(1)
			frm1.txtBpCd.focus
		Case "CURR"
			frm1.txtDocCur.value = arrRet(0)
			Call txtDocCur_OnChange()
			Call XchLocRate()
			frm1.txtDocCur.focus
		Case "PRPAYMTYPE"
			frm1.txtPrpaymType.value = arrRet(0)
			frm1.txtPrpaymTypeNm.value = arrRet(1)
			frm1.txtPrpaymType.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'   개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================

'=======================================================================================================
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'=======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029																'Load table , B_numeric_format

    Call ggoOper.LockField(Document, "N")											 'Lock  Suitable  Field    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitVariables																'Initializes local global variables
    Call SetDefaultVal
	Call FncNew	
End Sub

'=======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc : load했던 'uni2kcm.dll"의 클래스들을 unload한다.
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPrpaymDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPrpaymDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPrpaymDt.Focus
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
'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtPrpaymDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtPrpaymDt.Text, gDateFormat,""), "''", "S") & "))"			
		
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
'	call XchLocRate()
End Sub
'==========================================================================================
'   Event Name : txtPrpaymDt_Change
'   Event Desc : 
'==========================================================================================
Sub txtPrpaymDt_Change()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	
	If gIsShowLocal = "N" Then
		frm1.txtXchRate.Value	= 0
		frm1.txtPrpaymLocAmt.value = 0
	Else
		frm1.txtXchRate.Text	= 0
    	frm1.txtPrpaymLocAmt.text = 0

	End If
	
	If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
				If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtPrpaymDt.Text <> "") Then
		'----------------------------------------------------------------------------------------
						strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtPrpaymDt.Text, gDateFormat,""), "''", "S") & "))"			
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
End Sub

Sub txtPrpaymAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtPrpaymLocAmt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtXchRate_Change()
    lgBlnFlgChgValue = True
	If lgQueryOk <> true Then    
		With Frm1    
			.txtPrpaymLocAmt.text="0"
		End with 
	End if    

End Sub

'=======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================

'=======================================================================================================
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'=======================================================================================================

'=======================================================================================================
'   Function Name : FncQuery
'   Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False     
	lgstartfnc = True                                                      
    Err.Clear                                                           

	'-----------------------
	'Check previous data area
	'----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase contents area
    '----------------------- 
    If Not chkField(Document, "1") Then									'This function check indispensable field
       Exit Function
    End If

'    Call ggoOper.ClearField(Document, "2")								'Clear Contents  Field
    Call InitVariables													'Initializes local global variables
	'-----------------------
	'Check condition area
	'----------------------- 
    If Not chkField(Document, "1") Then							'This function check indispensable field
       Exit Function
    End If
	'-----------------------
    'Query function call area
    '-----------------------
    frm1.hCommand.value = "QUERY" 
    Call DbQuery														'☜: Query db data
       
    FncQuery = True		
	Set gActiveElement = document.activeElement       												
End Function

'=======================================================================================================
'   Function Name : FncNew
'   Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                      
	lgstartfnc = True       
    
	'-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                              'Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                               'Lock  Suitable  Field
    
    Call txtDocCur_OnChange()
    
    frm1.txtPrpaymAmt.text = 0
    
    Call SetToolbar("111010000000111")
    Call SetDefaultVal
    Call InitVariables													'Initializes local global variables
    
    frm1.txtPrpaymNo.focus 
    Set gActiveElement = document.activeElement
    
    FncNew = True		   												'Processing is OK
    lgFormLoad = True													' tempgldt read
    lgQueryOk = False
    lgstartfnc = False        
	
	Set gActiveElement = document.activeElement   
	
End Function

'=======================================================================================================
'   Function Name : FncDelete
'   Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete()
	Dim IntRetCd
    
    FncDelete = False													
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
	'-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
	If IntRetCD = vbNo Then
		Exit Function
	End If
        
    Call DbDelete														'☜: Delete db data
    
    FncDelete = True                                                    
	
	Set gActiveElement = document.activeElement   
End Function

'=======================================================================================================
'   Function Name : FncSave
'   Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                     
    
    Err.Clear                                                           '☜: Protect system from crashing
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                      
        Exit Function
    End If
	'-----------------------
    'Check content area
    '-----------------------
    If len(frm1.txtPrpaymAmt.Text)= 0 then
		Call DisplayMsgBox("970021","X",frm1.txtPrpaymAmt.alt,"X")  
		Exit Function
    ElseIf UNICDbl(frm1.txtPrpaymAmt.Text) = 0 then
		Call DisplayMsgBox("141704","X",frm1.txtPrpaymAmt.alt,"X")  
		Exit Function
    End if
    
    If Not chkField(Document, "2") Then                         'Check contents area
       Exit Function
    End If
	'-----------------------
    'Save function call area
    '-----------------------
    CAll DbSave				                                            '☜: Save db data
    
    FncSave = True   
    
    Set gActiveElement = document.activeElement                                                      
End Function

'=======================================================================================================
'   Function Name : FncCopy
'   Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	Dim IntRetCD
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")	'"Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
     ' 조건부 필드를 삭제한다. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
    frm1.txtPrpaymNo.value = ""
    frm1.txtPrpaymNo.focus
    
    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode
	
	Set gActiveElement = document.activeElement   
End Function

'=======================================================================================================
'   Function Name : FncCancel
'   Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
    On Error Resume Next                                                '☜: Protect system from crashing
End Function

'=======================================================================================================
'   Function Name : FncInsertRow
'   Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow() 
     On Error Resume Next                                               '☜: Protect system from crashing
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    On Error Resume Next                                                '☜: Protect system from crashing
End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call parent.FncPrint()   
    
    Set gActiveElement = document.activeElement                                              '☜: Protect system from crashing
End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                            '밑에 메세지를 ID로 처리해야 함 
        Exit Function
    End If
	'-----------------------
	'Check previous data area
	'----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
	'Check condition area
	'----------------------- 
    If Not chkField(Document, "1") Then									'This function check indispensable field
       Exit Function
    End If
    
    Call InitVariables													'Initializes local global variables

	frm1.hCommand.value = "PREV"
	Call DbQuery
	
	Set gActiveElement = document.activeElement   
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    Dim IntRetCD

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                  'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                            
        Exit Function
    End If
	'-----------------------
	'Check previous data area
	'----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
	'Check condition area
	'----------------------- 
    If Not chkField(Document, "1") Then									'This function check indispensable field
       Exit Function
    End If
    
    Call InitVariables													'Initializes local global variables

	frm1.hCommand.value = "NEXT"
	Call DbQuery
	
	Set gActiveElement = document.activeElement   	
End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)										
	
	Set gActiveElement = document.activeElement   
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                               
	
	Set gActiveElement = document.activeElement   
End Function

'=======================================================================================================
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
    Err.Clear                                                           '☜: Protect system from crashing
    
    DbDelete = False													
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003						'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPrpaymNo=" & Trim(frm1.txtPrpaymNo.value)	'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtGlDt=" & Trim(frm1.txtGlDt.text)	'☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	
    DbDelete = True                                                     
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'=======================================================================================================
Function DbDeleteOk()													'삭제 성공후 실행 로직 
	Call FncNew()
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery() 
    Err.Clear                                                           '☜: Protect system from crashing
    
    DbQuery = False                                                     
    
    Call LayerShowHide(1)
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPrpaymNo=" & Trim(frm1.txtPrpaymNo.value)
    strVal = strVal & "&SelectChar=" & Trim(frm1.hCommand.value)
    
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

    DbQuery = True   
    lgQueryOk = True 	                                                
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQueryOk()													'☆: 조회 성공후 실행로직 
	 Dim strTemp
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE											'Indicates that current mode is Update mode
    lgQueryOk= true 

    Call ggoOper.LockField(Document, "Q")								'This function lock the suitable field
    Call SetToolbar("111110001101111")
    
    If gIsShowLocal = "N" Then
		strTemp = frm1.txtXchRate.Value
		Call txtDocCur_OnChange()	
		frm1.txtXchRate.Value = strTemp
	Else
		strTemp = frm1.txtXchRate.text
		Call txtDocCur_OnChange()	
		frm1.txtXchRate.text = strTemp
	End If
	lgBlnFlgChgValue = False
	lgQueryOk= False
End Function

'=======================================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'=======================================================================================================
Function DbSave() 
    Dim strVal
    Err.Clear															'☜: Protect system from crashing

	DbSave = False														
    
    Call LayerShowHide(1)

	With frm1
		.txtMode.value = parent.UID_M0002										'☜: 비지니스 처리 ASP 의 상태 
		.txtFlgMode.value = lgIntFlgMode
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										

    DbSave = True   
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function DbSaveOk()														'저장 성공후 실행 로직 
    Call InitVariables
    Call FncQuery()
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
	End If
    IF  CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
	End If	    
End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 선급금액 
		ggoOper.FormatFieldByObjectOfCur .txtPrpaymAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 반제금액 
		ggoOper.FormatFieldByObjectOfCur .txtClsAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 청산금액 
		ggoOper.FormatFieldByObjectOfCur .txtSttlAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec

	End With
End Sub

'===================================== XchLocRate()  ======================================
'	Name : XchLocRate()
'	Description : 통화가 변경될경우 통화에 따른 자국금액 
'====================================================================================================
Sub XchLocRate()
	If gIsShowLocal <> "N" Then
	
	frm1.txtPrpaymLocAmt.text = "0"
	frm1.txtXchRate.text = "0"
    else
	frm1.txtPrpaymLocAmt.value = "0"
	frm1.txtXchRate.value = "0"
    end if
End Sub

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
	
'======================================================================================================= -->
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>선급금기초치등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>선급금번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtPrpaymNo" SIZE=20 MAXLENGTH=18 tag="12XXXU"  ALT="선급금 번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopupPP()"></TD>
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
								<TD CLASS="TD5" NOWRAP>선급금유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrpaymType" SIZE=10 MAXLENGTH=10  tag="22XXXU" ALT="선급금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('','PrpaymType')">&nbsp;<INPUT TYPE=TEXT NAME="txtPrpaymTypeNm" SIZE=25 tag="24" ALT="선급금유형명"></TD>
								<TD CLASS="TD5" NOWRAP>부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="부서코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ALT="부서명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>거래처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value,'BP')">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="24" ALT="거래처명"></TD>
								<TD CLASS="TD5" NOWRAP>출금일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDateTime1_txtPrpaymDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>결의전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="24" ALT="회계전표번호"></TD>
								<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="24" ALT="회계전표번호"></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>								
							<TR>
								<TD CLASS="TD5" NOWRAP>거래통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" TYPE="Text" SIZE=10 MAXLENGTH=3 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup('','CURR')"></TD>
								<TD CLASS="TD5" NOWRAP>환율</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle1_txtXchRate.js'></script></TD>
							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtXchRate" TABINDEX="-1">
<%	End If %>								
							<TR>
								<TD CLASS="TD5" NOWRAP>선급금액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle2_txtPrpaymAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>								
								<TD CLASS="TD5" NOWRAP>선급금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle3_txtPrpaymLocAmt.js'></script></TD>
							</TR>
							<TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtPrpaymLocAmt" TABINDEX="-1">
<%	End If %>							
								<TD CLASS="TD5" NOWRAP>반제금액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle4_txtClsAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>								
								<TD CLASS="TD5" NOWRAP>반제금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle5_txtClsLocAmt.js'></script></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtClsLocAmt" TABINDEX="-1">
<%	End If %>							
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>청산금액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle6_txtSttlAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>								
								<TD CLASS="TD5" NOWRAP>청산금액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle7_txtSttlLocAmt.js'></script></TD>
							</TR>
							<TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtSttlLocAmt" TABINDEX="-1">
<%	End If %>								
								<TD CLASS="TD5" NOWRAP>잔액</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle8_txtBalAmt.js'></script></TD>
<%	If gIsShowLocal <> "N" Then	%>								
								<TD CLASS="TD5" NOWRAP>잔액(자국)</TD>
	                            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDoubleSingle9_txtBalLocAmt.js'></script></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtBalLocAmt" TABINDEX="-1">
<%	End If %>							
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>전표일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f6105ma1_fpDateTime1_txtGlDt.js'></script></TD>
							
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrpaymDesc" SIZE=50 MAXLENGTH=100 STYLE="TEXT_ALIGN:Left" tag="2X" ALT="비고"></TD>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCommand" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

