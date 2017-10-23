
<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1 %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Account Receivable
'*  3. Program ID           : a3112ma.asp
'*  4. Program Name         : 기초채권등록 
'*  5. Program Desc         :
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2003/01/07
'*  9. Modifier (First)     : You So Eun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
 -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'            1. 선 언 부 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc 선언   
' 기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit                 '☜: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
' .Constant는 반드시 대문자 표기.
' .변수 표준에 따름. prefix로 g를 사용함.
' .Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

'@PGM_ID
Const BIZ_PGM_QRY_ID = "a3112mb1.asp"							'☆: Head Query 비지니스 로직 ASP명 
Const BIZ_PGM_SAVE_ID = "a3112mb2.asp"							'☆: Save 비지니스 로직 ASP명 
Const BIZ_PGM_DEL_ID = "a3112mb3.asp"


Const TAB1 = 1													'☜: Tab의 위치 
Const TAB2 = 2

Dim  IsOpenPop													'Popup
Dim	 lgFormLoad
Dim	 lgQueryOk													' Queryok여부 (loc_amt =0 check)
Dim  lgstartfnc

Dim dtToday
dtToday = "<%=GetSvrDate%>"

'======================================================================================================
'            2. Function부 
'
' 내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
' 공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'               2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================

'======================================================================================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE						'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False								'Indicates that no value changed
	lgstartfnc=False
	lgFormLoad=True    
    lgQueryOk= False    
	lgstartfnc=False
	lgFormLoad=True
	lgQueryOk= False

End Sub
'======================================================================================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtArDt.text  =  UniConvDateAToB(dtToday, parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtDueDt.text  =  UniConvDateAToB(dtToday, parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtGlDt.text =  UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
'	frm1.cboArType.value = "NT" 
	frm1.txtDocCur.value = parent.gCurrency
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	frm1.txtDeptCd.value = parent.gDepart
	
	If gIsShowLocal <> "N" Then
		frm1.txtXchRate.text = "1"
	Else
		frm1.txtXchRate.value = "1"
	End if

	lgBlnFlgChgValue = False								'Indicates that no value changed 
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>    
End Sub


'======================================================================================================
'   Function Name : OpenPopUpgl()
'   Function Desc : 
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
 
	arrParam(0) = Trim(frm1.txtGlNo.value)											'회계전표번호 
	arrParam(1) = ""																'Reference번호 

	IsOpenPop = True
	  
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
End Function
 '------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtArDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
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
				.txtDeptCD.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtArDt.text = arrRet(3)
				call txtDeptCd_OnBlur()  
				.txtDeptCd.focus
        End Select
	End With
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
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :매출 S: 매입 T: 전체 
	Select Case iWhere
		Case 3
			arrParam(5) = "SOL"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 9
			arrParam(5) = "INV"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
		Case 4
			arrParam(5) = "PAYER"									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	End Select
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then     
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName	

	If IsOpenPop = True Then Exit Function 
 
	Select Case iWhere
		Case 0
			arrParam(0) = "채권정보팝업"										' 팝업 명칭 
			arrParam(1) = "A_OPEN_AR A,B_ACCT_DEPT B,B_BIZ_PARTNER C"				' TABLE 명칭 
			arrParam(2) = ""														' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = "A.DEPT_CD = B.DEPT_Cd AND A.DEAL_BP_CD = C.BP_CD AND A.AR_TYPE = " & FilterVar("NR", "''", "S") & "  "         ' Where Condition
			arrParam(5) = "채권번호"   
 
			arrField(0) = "A.Ar_NO"													' Field명(0)
			arrField(1) = "CONVERT(VARCHAR(40),A.Ar_DT)"							' Field명(1)
			arrField(2) = "B.DEPT_NM"												' Field명(2)
			arrField(3) = "A.DOC_CUR"												' Field명(3) 
			arrField(4) = "C.BP_FULL_NM"											' Field명(4) 
			arrField(5) = "CONVERT(VARCHAR(15),A.Ar_AMT)"							' Field명(5)
			arrField(6) = "CONVERT(VARCHAR(15),A.VAT_AMT)"							' Field명(6)
			 
			arrHeader(0) = "채권번호"											' Header명(0)
			arrHeader(1) = "채권일"												' Header명(1)
			arrHeader(2) = "부서명"												' Header명(2)
			arrHeader(3) = "거래통화"											' Header명(3)
			arrHeader(4) = "거래처명"											' Header명(4)
			arrHeader(5) = "채권금액"											' Header명(5)
			arrHeader(6) = "부가세금액"											' Header명(6)
		Case 1
			arrParam(0) = "계정코드팝업"										' 팝업 명칭 
			arrParam(1) = "A_Acct A, A_ACCT_GP B, A_JNL_ACCT_ASSN C"				' TABLE 명칭 
			arrParam(2) = Trim(strCode)												' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A.Acct_cd=C.Acct_CD " & _ 
			    "and C.trans_type = " & FilterVar("ar005", "''", "S") & "  and C.jnl_cd = " & FilterVar("AR", "''", "S") & " "					' Where Condition
			arrParam(5) = "계정코드"											' 조건필드의 라벨 명칭 

			arrField(0) = "A.Acct_CD"												' Field명(0)
			arrField(1) = "A.Acct_NM"												' Field명(1)
			arrField(2) = "B.GP_CD"													' Field명(2)
			arrField(3) = "B.GP_NM"													' Field명(3)
		 
			arrHeader(0) = "계정코드"											' Header명(0)
			arrHeader(1) = "계정코드명"											' Header명(1)
			arrHeader(2) = "그룹코드"											' Header명(2)
			arrHeader(3) = "그룹명"												' Header명(3)
		Case 3
			arrParam(0) = "주문처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "주문처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "주문처"							' Header명(0)
			arrHeader(1) = "주문처명"						' Header명(1)
		Case 4
			If UCase(frm1.txtPayBpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
			arrParam(0) = "수금처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "수금처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    		arrHeader(0) = "수금처"							' Header명(0)
			arrHeader(1) = "수금처명"						' Header명(1)		Case 5       
			
		Case 5
			arrParam(0) = "사업장팝업"											' 팝업 명칭 
			arrParam(1) = "B_Biz_AREA"												' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = ""														' Where Condition
			arrParam(5) = "사업장"		
 
			arrField(0) = "Biz_AREA_CD"												' Field명(0)
			arrField(1) = "Biz_AREA_NM"												' Field명(1)    
			 
			arrHeader(0) = "사업장"												' Header명(0)
			arrHeader(1) = "사업장명"											' Header명(1)
		Case 8
			arrParam(0) = "거래통화팝업"										' 팝업 명칭 
			arrParam(1) = "b_currency"												' TABLE 명칭 
			arrParam(2) = strCode													' Code Condition
			arrParam(3) = ""														' Name Cindition
			arrParam(4) = ""														' Where Condition
			arrParam(5) = "거래통화"    
 
			arrField(0) = "CURRENCY"												' Field명(0)
			arrField(1) = "CURRENCY_DESC"											' Field명(1)
			 
			arrHeader(0) = "거래통화"											' Header명(0)
			arrHeader(1) = "거래통화명"											' Header명(1)    
		Case 9
			arrParam(0) = "세금계산서발행처팝업"						' 팝업 명칭 
			arrParam(1) = "b_biz_partner"						' TABLE 명칭 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "세금계산서발행처"			
	
			arrField(0) = "BP_CD"								' Field명(0)
			arrField(1) = "BP_NM"								' Field명(1)
    
    
			arrHeader(0) = "세금계산서발행처"							' Header명(0)
			arrHeader(1) = "세금계산서발행처명"						' Header명(1)
		


		Case 10
			If  UCase(frm1.txtPayMethCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrHeader(0) = "결제방법"											' Header명(0)
			arrHeader(1) = "결제방법명"											' Header명(1)
			arrHeader(2) = "Reference"
			 
			arrField(0) = "B_Minor.MINOR_CD"										' Field명(0)
			arrField(1) = "B_Minor.MINOR_NM"										' Field명(1)
			arrField(2) = "b_configuration.REFERENCE"
			 
			arrParam(0) = "결제방법"											' 팝업 명칭 
			arrParam(1) = "B_Minor,b_configuration"									' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPayMethCd.Value)								' Code Condition
		 
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd= B_Minor.Major_Cd"  
			arrParam(5) = "결제방법"											' TextBox 명칭 
		Case 11
			if Trim(frm1.txtPayMethCd.Value) = "" then
				Call DisplayMsgBox("205152","X" , "결제방법","X")
				Exit Function
			End if

			If UCase(frm1.txtPayTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
 
			arrParam(0) = "입금유형"											' 팝업 명칭 
			arrParam(1) = "B_MINOR,B_CONFIGURATION," _
				& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD = " & FilterVar("B9004", "''", "S") & " "_
				& "And MINOR_CD= " & FilterVar(frm1.txtPayMethCd.value, "''", "S") & " And SEQ_NO>=2)C" ' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPayTypeCd.value)								' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  " _
			      & "AND B_CONFIGURATION.REFERENCE IN(" & FilterVar("RP", "''", "S") & " ," & FilterVar("R", "''", "S") & " )"					' Where Condition
			   
			arrParam(5) = "입금유형"											' TextBox 명칭 
	 
			arrField(0) = "B_MINOR.MINOR_CD"										' Field명(0)
			arrField(1) = "B_MINOR.MINOR_NM"										' Field명(1)
			  
			arrHeader(0) = "입금유형"											' Header명(0)
			arrHeader(1) = "입금유형명"											' Header명(1)  
	End Select    
 
	IsOpenPop = True
	 
	If iwhere = 0 Then  
		iCalledAspName = AskPRAspName("a3112ra1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3112ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
	   
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
	      "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")   
	End If
	 
	IsOpenPop = False
 
	If arrRet(0) = "" Then     
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtArNo.focus
			Case 1 
				.txtAcctCd.focus
			Case 3
				.txtDealBpCd.focus
			Case 4
				.txtPayBpCd.focus
			Case 5   
				.txtReportBizCd.focus
			Case 8
				.txtDocCur.focus
			Case 9
				.txtReportBpCd.focus
			Case 10
				.txtPayMethCd.focus
			Case 11 
			    .txtPayTypeCd.focus
		End Select    
	End With
 
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtArNo.value = arrRet(0)
				.txtArNo.focus
			Case 1 
				.txtAcctCd.value = arrRet(0)
				.txtAcctNm.value = arrRet(1)
				.txtAcctCd.focus
			Case 3
				.txtDealBpCd.value = arrRet(0)
				.txtDealBpNm.value = arrRet(1)
				Call txtDealBpCd_onChange()
				.txtDealBpCd.focus
			Case 4
				.txtPayBpCd.value = arrRet(0)
				.txtPayBpNm.value = arrRet(1)
				.txtPayBpCd.focus
			Case 5   
				.txtReportBizCd.value = arrRet(0)
				.txtReportBizNm.value = arrRet(1)
				.txtReportBizCd.focus
			Case 8
				.txtDocCur.value = arrRet(0)
				Call txtDocCur_OnChange()
				.txtDocCur.focus
			Case 9
			    .txtReportBpCd.value = arrRet(0)
				.txtReportBpNm.value = arrRet(1)
				.txtReportBpCd.focus
			Case 10
				.txtPayMethCd.Value = arrRet(0)
				.txtPayMethNm.Value = arrRet(1)
				.txtPayMethCd.focus
			Case 11 
				.txtPayTypeCd.value = arrRet(0)
			    .txtPayTypeNm.value = arrRet(1)               
			    .txtPayTypeCd.focus
		End Select    
	End With
 
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If 
End Function

'======================================================================================================
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================

'======================================================================================================
'            3. Event부 
' 기능: Event 함수에 관한 처리 
' 설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'======================================================================================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()															'Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
										parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
                         
    Call ggoOper.LockField(Document, "N")											'Lock  Suitable  Field    
    Call InitVariables()															'Initializes local global variables    

    Call SetToolbar("1110100000001111")												'버튼 툴바 제어 
	Call SetDefaultVal()
 
	frm1.txtArNo.focus
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'=======================================================================================================
'   Event Name : txtArDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtArDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtArDt.Action = 7   
        Call  txtArDt_OnBlur()    
        Call SetFocusToDocument("M")
		Frm1.txtArDt.Focus
		
    End If
End Sub
'==========================================================================================
'   Event Name : txtArDt_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtArDt_OnBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
   If lgstartfnc = False Then
		If lgFormLoad = True Then
			lgBlnFlgChgValue = True
			With frm1
	
				If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtArDt.Text <> "") Then
					'----------------------------------------------------------------------------------------
						strSelect	=			 " Distinct org_change_id "    		
						strFrom		=			 " b_acct_dept(NOLOCK) "		
						strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
						strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
						strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
						strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtArDt.Text, gDateFormat,""), "''", "S") & "))"			
	
					IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 			
					If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
						.txtDeptCd.value = ""
						.txtDeptNm.value = ""
						.hOrgChangeId.value = ""
						.txtDeptCd.focus
					End if

				End If
			End With
		'----------------------------------------------------------------------------------------
		End If
	End IF
  
	If lgQueryOk <> true then
		frm1.txtNetLocAmt.text = "0"
	End if
End Sub


'=======================================================================================================
'   Event Name : txtDealBpCd_onChange()
'   Event Desc :  
'=======================================================================================================
Sub  txtDealBpCd_onChange()

    lgBlnFlgChgValue = True
	If lgIntFlgMode <> parent.OPMD_UMODE Then 		
		frm1.txtPayBpCd.value = frm1.txtDealBpCd.value
		frm1.txtPayBpNm.value = frm1.txtDealBpNm.value
		frm1.txtReportBpCd.value = frm1.txtDealBpCd.value
		frm1.txtReportBpNm.value = frm1.txtDealBpNm.value
	End if

End Sub
'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_OnBlur()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtArDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtArDt.Text, gDateFormat,""), "''", "S") & "))"			
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
	End if	
		'----------------------------------------------------------------------------------------
End Sub

'=======================================================================================================
'   Event Name : txtArDt_Change()
'   Event Desc : 
'=======================================================================================================
Sub  txtArDt_Change() 
    lgBlnFlgChgValue = True

    If lgQueryOk <> True Then
		frm1.txtNetLocAmt.text = "0"
	End if
End Sub
'=======================================================================================================
'   Event Name : txTblDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txTblDt_DblClick(Button)
    If Button = 1 Then
        frm1.txTblDt.Action = 7        
    	Call SetFocusToDocument("M")
		Frm1.txTblDt.Focus
		
    End If
End Sub

'=======================================================================================================
'   Event Name : txTblDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txTblDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7     
        Call SetFocusToDocument("M")
		Frm1.txtDueDt.Focus           
    End If
End Sub
'=======================================================================================================
'   Event Name : txtGlDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtGlDt_DblClick(Button)
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
Sub  txtGlDt_Change()
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
'   Event Name : txtDueDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtDueDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtInvDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtInvDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInvDt.Action = 7  
        Call SetFocusToDocument("M")
		Frm1.txtInvDt.Focus                                
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInvDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtInvDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtCashAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtCashAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtCashLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtCashLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPrRcptAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtPrRcptAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPrRcptLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtPrRcptLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtNetAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtNetAmt_Change()
	lgBlnFlgChgValue = True

	If lgQueryOk <> True Then
		frm1.txtNetLocAmt.text = "0"
	End If	
End Sub

'=======================================================================================================
'   Event Name : txtNetLocAmt_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub  txtNetLocAmt_Change()
	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtPayDur_Change()
'   Event Desc : Single의 숫자필드가 바뀌었는지 check한다.
'=======================================================================================================
Sub txtPayDur_Change()
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'송장번호 입력시 송장일자 입력필수 
'======================================================================================================
Sub txtInvNo_OnBlur()
	If Trim(frm1.txtInvNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtInvDt, "N") 'N:Required, Q:Protected, D:Default
	End If
End Sub

'======================================================================================================
'선하증권번호 입력시 선하증권일자 입력필수 
'======================================================================================================
Sub txtBlNo_OnBlur()
	If Trim(frm1.txtBlNo.value) = "" Then
		Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
	Else
		Call ggoOper.SetReqAttr(frm1.txtBlDt, "N") 'N:Required, Q:Protected, D:Default
	End If
End Sub

Sub txTGlDt_Change()
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'            4. Common Function부 
' 기능: Common Function
' 설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================

'======================================================================================================
'            5. Interface부 
' 기능: Interface
' 설명: 각각의 Toolbar에 대한 처리를 행한다. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    lgstartfnc = True
    
    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then										'This function check indispensable field
       Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then  
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")     
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables()													'Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------                  
    Call DbQuery()															'☜: Query db data    
    FncQuery = True 
    lgstartfnc = False	    
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
     
    FncNew = False  
    lgstartfnc = True                                                         
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")												'Clear Condition Field
    Call ggoOper.LockField(Document, "N")												'Lock  Suitable  Field    
    Call InitVariables()																'Initializes local global variables
    call SetDefaultVal()
    
    frm1.txtArNo.Value = ""
    frm1.txtArNo.focus
    
    Call txtDocCur_OnChange()
    
    lgBlnFlgChgValue = False    

    FncNew = True 
    lgFormLoad = True							' tempgldt read
    lgstartfnc = False                                                         
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then											'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")						'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete               '☜: Delete db data
    
    FncDelete = True                                                        
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
 
    FncSave = False                                                         
    
    Err.Clear                                                               
    

    If lgBlnFlgChgValue = False Then										'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")				'⊙: Display Message(There is no changed data.)
		Exit Function
    End If
    
    If Len(frm1.txtNetAmt.Text) = 0 then
		Call DisplayMsgBox("970021","X",frm1.txtNetAmt.alt,"X")  
		Exit Function
    ElseIf UNICDbl(frm1.txtNetAmt.Text) = 0 then
		Call DisplayMsgBox("141704","X",frm1.txtNetAmt.alt,"X")  
		Exit Function
    End if
    
    If Not chkField(Document, "2") Then										'⊙: Check required field(Single area)
		Exit Function
    End If
    '================================================================================================
    '일자관계 체크 : LC발행일(txtLcDt)<=송장일(txtInvDt)<=선하증권일(txtBlDt)<=채권/채무일(txtArDt)
    '================================================================================================
    If frm1.txtBlDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtBlDt.Text,frm1.txtArDt.Text,frm1.txtBlDt.Alt,frm1.txtArDt.Alt, _
		                      "970025",frm1.txtBlDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			frm1.txtBlDt.focus
			Exit Function
		End If
    End If
    
    If frm1.txtInvDt.Text <> "" Then
		If frm1.txtBlDt.Text = "" Then
			If CompareDateByFormat(frm1.txtInvDt.Text,frm1.txtArDt.Text,frm1.txtInvDt.Alt,frm1.txtArDt.Alt, _
			                     "970025",frm1.txtInvDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   frm1.txtInvDt.focus
			   Exit Function
			End If
		Else
			If CompareDateByFormat(frm1.txtInvDt.Text,frm1.txtBlDt.Text,frm1.txtInvDt.Alt,frm1.txtBlDt.Alt, _
			                    "970025",frm1.txtInvDt.UserDefinedFormat,parent.gComDateType, true) = False Then
			   frm1.txtInvDt.focus
			   Exit Function
			End If
		End If
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																'☜: Save db data
    
    FncSave = True                                                       
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
 
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
    
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow() 

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    
End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next    
	Call parent.FncPrint()                                           
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function  FncPrev() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function  FncNext() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                          
	    		
	Set gActiveElement = document.activeElement    

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call FncExport(parent.C_SINGLEMULTI)
	    		
	Set gActiveElement = document.activeElement    

End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()

End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
 
	FncExit = False

	If lgBlnFlgChgValue = True Then														'⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")					'데이타가 변경되었습니다. 종료 하시겠습니까?
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
'========================================================================================================
Function  DbDelete() 
    DbDelete = False              
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtArNo=" & Trim(frm1.txtArNo.value)    '☜: 삭제 조건 데이타 
    
	Call RunMyBizASP(MyBizASP, strVal)          '☜: 비지니스 ASP 를 가동 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()																'삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "2")											'Clear Condition Field
    Call ggoOper.LockField(Document, "N")											'Lock  Suitable  Field    
    Call InitVariables()															'Initializes local global variables
    Call SetDefaultVal()
    
    frm1.txtArNo.Value = ""
    frm1.txtArNo.focus
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'☜: 
			strVal = strVal & "&txtArNo=" & Trim(.htxtArNo.value)					'조회 조건 데이타 
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'☜: 
			strVal = strVal & "&txtArNo=" & Trim(.txtArNo.value)					'조회 조건 데이타 
		End If
    End With

	Call RunMyBizASP(MyBizASP, strVal)												'☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk()
	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------  
		Call ggoOper.LockField(Document, "Q")										'This function lock the suitable field        
		Call SetToolbar("1111100000001111") 
		call InitVariables()
		
		lgQueryOk= True
				
		lgIntFlgMode = parent.OPMD_UMODE											'Indicates that current mode is Update mode
	 
		Call txtDocCur_OnChange()        
		Call txtDeptCd_OnBlur()
		If Trim(frm1.txtInvNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtInvDt, "N")								'N:Required, Q:Protected, D:Default
		End If
		If Trim(frm1.txtBlNo.value) = "" Then
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtBlDt, "N")								'N:Required, Q:Protected, D:Default
		End If
	 
		lgBlnFlgChgValue = False
		lgQueryOk= False
	End With 
End Function


'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
    Err.Clear 

	frm1.txtFlgMode.value = lgIntFlgMode         
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal ArNo)														'☆: 저장 성공후 실행 로직 
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtArNo.value = ArNo
	End If   
 
	Call ggoOper.ClearField(Document, "2")											'Clear Contents  Field
	Call InitVariables()															'Initializes local global variables
	Call DBquery()     
End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then                     
		Call CurFormatNumericOCX()
	End If    

	If lgQueryOk <> True Then
		If UCase(Trim(frm1.txtDocCur.Value)) <> UCase(Trim(parent.gCurrency)) Then
			frm1.txtXchRate.Text = "0" 
		Else			
			frm1.txtXchRate.Text = "1" 		
		End If			
		frm1.txtNetLocAmt.text = "0"
	End If   	
End Sub

'===================================== CurFormatNumericOCX()  =======================================
' Name : CurFormatNumericOCX()
' Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' 외상매출액 
		ggoOper.FormatFieldByObjectOfCur .txtNetAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 채권잔액 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' 환율 
		ggoOper.FormatFieldByObjectOfCur .txtXchRate, .txtDocCur.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!--'======================================================================================================
'            6. Tag부 
' 기능: Tag부분 설정 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
	<TABLE <%=LR_SPACE_TYPE_00%>>
		<TR>
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
									<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>기초채권등록</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
								</TR>
							</TABLE>
						</TD>
						<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
						<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>채권번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtArNo" ALT="채권번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:CALL OpenPopUp(frm1.txtArNo.Value, 0)"></TD>        
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>        
				<TR>
					<TD WIDTH="100%">     
				
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDealBpCd" ALT="주문처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value,3)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT NAME="txtDealBpNm" ALT="주문처" SIZE="20" tag = "24" ></TD>
									<TD CLASS=TD5 NOWRAP>송장번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvNo" ALT="송장번호" MAXLENGTH="50" SIZE=20 STYLE="TEXT-ALIGN:  left" tag="2XXXXU"></TD>       </TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>수금처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayBpCd" ALT="수금처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value,4)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT  NAME="txtpaybpnm"  ALT="수금처" SIZE="20" tag = "24" ></TD>
									<TD CLASS=TD5 NOWRAP>송장일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/a3112ma1_OBJECT3_txtInvDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>세금계산서발행처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReportBpCd" ALT="세금계산서발행처" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="21NXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenBp(frm1.txtReportBpCd.Value,9)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT  NAME="txtReportbpnm"  ALT="세금계산서발행처" SIZE="20" tag = "24" ></TD>        
									<TD CLASS=TD5 NOWRAP>선하증권번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txTblNo" ALT="선하증권번호" MAXLENGTH="35" SIZE=20 tag="2XXXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU"><IMG align=top name=btnCalType onclick="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
									<INPUT NAME="txtDeptNm" ALT="부서" SIZE="20" tag ="24" ></TD>
									<TD CLASS=TD5 NOWRAP>선하증권일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/a3112ma1_OBJECT4_txTblDt.js'></script></TD>               
								</TR>
								<TR><TD CLASS=TD5 NOWRAP>계정코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="22NXXU" ><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtAcctCd.value,1)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
										<INPUT NAME="txtAcctnm" ALT="계정코드명" MAXLENGTH="20"  tag  ="24"></TD>          
									<TD CLASS="TD5" nowrap>결제기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP>
													<script language =javascript src='./js/a3112ma1_fpDoubleSingle5_txtPayDur.js'></script>
												</TD>
												<TD NOWRAP>
													&nbsp;일
												</TD>
											</TR>
										</Table>
									</TD>
								</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>채권일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_OBJECT1_txtArDt.js'></script></TD>        
								<TD CLASS="TD5" nowrap>결제방법</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayMethCd" ALT="결제방법" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtPayMethCd.value, 10)">
									<INPUT TYPE=TEXT NAME="txtPayMethNm" ALT="결제방법" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>

					       </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>채권만기일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_OBJECT2_txtDueDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>입금유형</TD>
								<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtPayTypeCd" ALT="입금유형" style="HEIGHT: 20px; WIDTH: 80px" MAXLENGTH=5 tag="21NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" style="HEIGHT: 21px; WIDTH: 16px" NAME="btnPayMethod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtPayTypeCd.value, 11)">
									<INPUT TYPE=TEXT NAME="txtPayTypeNm" ALT="입금유형" style="HEIGHT: 20px; WIDTH: 150px" tag="24X" ></TD>
					       </TR>
					       <TR>
<% If gIsShowLocal <> "N" Then %>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: Left" tag ="22XXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtDocCur.Value,8)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON">
								&nbsp;&nbsp;환율<script language =javascript src='./js/a3112ma1_OBJECT5_txtXchRate.js'></script></TD>                
<% ELSE %>
									<INPUT TYPE=HIDDEN NAME="txtDocCur"   TABINDEX="-1">
									<INPUT TYPE=HIDDEN NAME="txtXchRate"  TABINDEX="-1">									
<% End If %>         
						        <TD CLASS=TD5 NOWRAP>대금결제참조</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPaymTerms" ALT="대금결제참조" MAXLENGTH="120" SIZE=30 STYLE="TEXT-ALIGN: left" tag ="21"></TD>        
							</TR>               
							<TR>
								<TD CLASS=TD5 NOWRAP>전표일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a3112ma1_OBJECT1_txtGlDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlNo" ALT="전표번호" SIZE="19" MAXLENGTH="18" STYLE="TEXT-ALIGN: Left" tag="24XXXU" ></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>외상매출액</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I456811636_txtNetAmt.js'></script></TD>
<% If gIsShowLocal <> "N" Then %>         
							    <TD CLASS=TD5 NOWRAP>외상매출액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I471998035_txtNetLocAmt.js'></script></TD>
							</TR>
							<TR>
<% ELSE %>
								<INPUT TYPE=HIDDEN NAME="txtNetLocAmt"   TABINDEX="-1">
<% End If %>       
								<TD CLASS=TD5 NOWRAP>채권잔액</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I623698359_txtBalAmt.js'></script></TD>
<% If gIsShowLocal <> "N" Then %>        
								<TD CLASS=TD5 NOWRAP>채권잔액(자국통화)</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a3112ma1_I870568823_txtBalLocAmt.js'></script></TD>
<% ELSE %>
										<INPUT TYPE=HIDDEN NAME="txtBalLocAmt"   TABINDEX="-1">
<% End If %>       
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="60" tag="2X" ></TD>
							    <TD CLASS=TD5 NOWRAP>프로젝트</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="프로젝트" MAXLENGTH=25 SIZE=25 tag="2X"></TD>
							</TR>
						</TABLE>
					</DIV>
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
		<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TabIndex="-1"></IFRAME>
	</TD>
 </TR>
</TABLE>
	<INPUT TYPE=hidden NAME="txtMode" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="htxtArNo" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="hAcctCd" tag="24" TabIndex="-1">
	<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>

