
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4205ra1
'*  4. Program Name         : 차입금번호팝업 
'*  5. Program Desc         : Popup of Loan No.
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.02.19
'*  8. Modified date(Last)  : 2001.11.10
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Oh, Soo Min
'* 11. Comment              :
'======================================================================================================= -->
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
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit                              '☜: indicates that All variables must be declared in advance

'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                             '☜: Popup status                           

Dim lgMark

Dim IsOpenPop                                                  '☜: 마크                                  
Dim CPGM_ID
'---------------  coding part(실행로직,Start)-----------------------------------------------------------
'   Call GetAdoFiledInf("F4205RA1","S","A")                        '☆: spread sheet 필드정보 query   -----
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no                                                               
                                                                
'--------------- 개발자 coding part(실행로직,End)-------------------------------------------------------------


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f4205rb1.asp"

Const C_MaxKey = 4

Dim arrReturn
Dim arrParent
Dim arrParam					

' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인 

	 '------ Set Parameters from Parent ASP ------ 
	arrParent		= window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam		= arrParent(1)
	
	If Trim("<%=Request("PGM")%>") = "F4235MA1"  Then
		top.document.title = "만기연장번호팝업"
	ElseIf Trim("<%=Request("PGM")%>") = "F4205MA1"  Then
		top.document.title = "거래처차입번호팝업"
	ElseIf Trim("<%=Request("PGM")%>") = "F4206MA1"  Then
		top.document.title = "거래처기초차입번호팝업"
	Else
		top.document.title = "차입금번호팝업"
	End If

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
    Redim arrReturn(0)
    
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo		= ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = PopupParent.OPMD_CMODE
    
	Self.Returnvalue = arrReturn


	' 권한관리 추가 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If
	
End Sub
'==========================================  2.1 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitComboBox()
		
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))    

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
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtLoanFromDt.Text = frDt
	frm1.txtLoanToDt.Text = toDt   
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
	
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE", "RA") %>	
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","RA") %>
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
Function OKClick()
	Dim ii 
	
	If frm1.vspdData.ActiveRow > 0 Then
		Redim arrReturn(C_MaxKey)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		For ii = 0 To C_MaxKey - 1
			frm1.vspdData.Col  = GetKeyPos("A",ii + 1)
			arrReturn(ii) = frm1.vspdData.Text
		Next
	End If

	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	frm1.vspddata.OperationMode = 3 
    Call SetZAdoSpreadSheet("F4205RA1","S","A","V20030407",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
	    .vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    .vspdData.ReDraw = True
    End With
End Sub

'**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** 

 '-----------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------- 

'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenSortPopup()

Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

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

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
'   ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)
	Call InitVariables														'⊙: Initializes local global variables
	Call InitComboBox()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'	Call ElementVisible(frm1.txtDummy, 0)	'InVisible
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

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************

'==========================================================================================
'   Event Name : DblClick
'   Event Desc :
'==========================================================================================
Sub txtLoanFromDt_DblClick(Button)
	if Button = 1 then
		frm1.fpLoanFrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpLoanFrDt.Focus
	End if
End Sub

Sub txtLoanToDt_DblClick(Button)
	if Button = 1 then
		frm1.fpLoanToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpLoanToDt.Focus
	End if
End Sub

Sub txtDueFromDt_DblClick(Button)
	if Button = 1 then
		frm1.fpDuefrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpDuefrDt.Focus
	End if
End Sub

Sub txtDueToDt_DblClick(Button)
	if Button = 1 then
		frm1.fpDueToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpDueToDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================

Sub txtLoanFromDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanToDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtLoanToDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanFromDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtDueFromDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanFromDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtDueToDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanFromDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
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

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

Sub vspdData_DblClick(ByVal Col, ByVal Row)
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
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function




Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
			arrParam(0) = frm1.txtDocCur.Alt								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		    arrField(0) = "CURRENCY"										' Field명(0)
		    arrField(1) = "CURRENCY_DESC"									' Field명(1)
'   
		    arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드명"									' Header명(1)
			
		Case 3
			arrParam(0) = "차입금번호팝업"
			arrParam(1) = "f_ln_info A"
			arrParam(2) = strCode
			arrParam(3) = ""
'			arrParam(4) = "A.CONF_FG IN ('C','E')"
			If Trim("<%=Request("PGM")%>") = "F4205MA1"  Then
				arrParam(4) = "A.LOAN_BASIC_FG = " & FilterVar("LN", "''", "S") & " "				
			ElseIf Trim("<%=Request("PGM")%>") = "F4206MA1"  Then
				arrParam(4) = "A.LOAN_BASIC_FG = " & FilterVar("LT", "''", "S") & " "				
			ElseIf Trim("<%=Request("PGM")%>") = "F4235MA1"  Then
				arrParam(4) = "A.LOAN_BASIC_FG = " & FilterVar("LR", "''", "S") & " "				
			Else
			End If
			arrParam(4) = arrParam(4) & "AND A.LOAN_PLC_TYPE = " & FilterVar("BP", "''", "S") & " "

			' 권한관리 추가 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			End If

			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			End If

			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			End If


			arrParam(5) = frm1.txtLoanNo.Alt
	
			arrField(0) = "A.Loan_NO"
			arrField(1) = "A.Loan_NM"
					    
			arrHeader(0) = frm1.txtLoanNo.Alt
			arrHeader(1) = "차입명"
		Case 5		'차입거래처 
			arrParam(0) = "거래처팝업"
			arrParam(1) = "B_BIZ_PARTNER A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = frm1.txtBpCd.Alt
	
			arrField(0) = "A.BP_CD"
			arrField(1) = "A.BP_NM"
			    
			arrHeader(0) = frm1.txtBpCd.Alt
			arrHeader(1) = frm1.txtBpNm.Alt

		Case 6		'차입용도 
			arrParam(0) = "차입용도팝업"
			arrParam(1) = "B_MINOR A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1000", "''", "S") & " "
			arrParam(5) = frm1.txtLoanType.Alt
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtLoanType.Alt
			arrHeader(1) = frm1.txtLoanTypeNm.Alt
		
		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanFromDt.focus
		Exit Function
	End If

	Select Case iWhere
		Case 0	'거래통화 
			frm1.txtDocCur.value = arrRet(0)
			frm1.txtDocCur.focus
		Case 3	'차입금번호 
			frm1.txtLoanNo.value = arrRet(0)	
			frm1.txtLoanNo.focus
		Case 5	'차입은행 
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNm.value = arrRet(1)
			frm1.txtBpCd.focus
		Case 6	'차입용도 
			frm1.txtLoanType.value = arrRet(0)
			frm1.txtLoanTypeNm.value = arrRet(1)
			frm1.txtLoanType.focus
	End Select
End Function

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
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
   
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
	
	If frm1.txtLoanFromDt.Text <> "" And frm1.txtLoanToDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtLoanFromDt.Text, frm1.txtLoanToDt.Text, frm1.txtLoanFromDt.Alt, frm1.txtLoanToDt.Alt, _
							"970025", frm1.txtLoanFromDt.UserDefinedFormat, popupparent.gComDateType, true) = False Then
				frm1.txtLoanFromDt.focus											'⊙: GL Date Compare Common Function
				Exit Function
		End if
	End If

	If frm1.txtDueFromDt.Text <> "" And frm1.txtDueToDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtDueFromDt.Text, frm1.txtDueToDt.Text, frm1.txtDueFromDt.Alt, frm1.txtDueToDt.Alt, _
					"970025", frm1.txtDueFromDt.UserDefinedFormat, popupparent.gComDateType, true) = False Then
			frm1.txtDueFromDt.focus											'⊙: GL Date Compare Common Function
			Exit Function
		End if
	End If
	
		
    '-----------------------
    'Query function call area
    '-----------------------
	
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call PopupParent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call PopupParent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call PopupParent.FncFind(PopupParent.C_MULTI , False)                                     '☜:화면 유형, Tab 유무 
End Function


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
    
    Err.Clear           
'    frm1.vspdData.MaxRows = 0                                                   '☜: Protect system from crashing                                                    '☜: Protect system from crashing
	Call LayerShowHide(1)       
	     
    With frm1
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
'	If lgIntFlgMode = Parent.OPMD_UMODE Then
'		strVal = BIZ_PGM_ID & "?txtLoanFromDt=" & Trim(.hLoanFromDt.value)
'		strVal = strVal & "&txtLoanToDt=" & Trim(.hLoanToDt.value)
'		strVal = strVal & "&txtDueFromDt=" & Trim(.hDueFromDt.value)
'		strVal = strVal & "&txtDueToDt=" & Trim(.hDueToDt.value)
'		strVal = strVal & "&txtLoanType=" & Trim(.hLoanType.value)
''		strVal = strVal & "&txtLoanType_Alt=" & Trim(.txtLoanType.Alt)
'		strVal = strVal & "&txtBpCd=" & Trim(.hBankLoanCd.value)
'		strVal = strVal & "&txtBpCd_Alt=" & Trim(.txtBpCd.Alt)
'	Else 
		strVal = BIZ_PGM_ID & "?txtLoanFromDt=" & Trim(.txtLoanFromDt.Text)
		strVal = strVal & "&txtLoanToDt=" & Trim(.txtLoanToDt.Text) 
		strVal = strVal & "&txtDocCur=" & Trim(.txtDocCur.Value)
		strVal = strVal & "&txtDueFromDt=" & Trim(.txtDueFromDt.Text)
		strVal = strVal & "&txtDueToDt=" & Trim(.txtDueToDt.Text)
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
		strVal = strVal & "&txtBpCd_Alt=" & Trim(.txtBpCd.Alt)		
		strVal = strVal & "&cboLoanFg=" & Trim(.cboLoanFg.value)		
		strVal = strVal & "&txtLoanType=" & Trim(.txtLoanType.value)
		strVal = strVal & "&txtLoanType_Alt=" & Trim(.txtLoanType.Alt)			
		strVal = strVal & "&txtLoanNo=" & Trim(.txtLoanNo.Value)
'	End If
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
		strVal = strVal & "&txtPgmId=" & Trim("<%=Request("PGM")%>")			 '☜: F4101MA1 (reference 조건 추가)
        strVal = strVal & "&lgPageNo="       & lgPageNo    
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))       

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
    lgBlnFlgChgValue = False                                                 'Indicates that no value changed
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
'	lgIntFlgMode = Parent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtLoanFromDt.focus
	End If

End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################



'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
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
						<TD CLASS=TD5 NOWRAP>차입일자</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanFrDt name=txtLoanFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="시작차입일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanToDt name=txtLoanToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료차입일자"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>거래통화</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" SIZE = "10" MAXLENGTH="3"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 0)"></TD>						
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>상환만기일자</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueFrDt name=txtDueFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="시작만기일자"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueToDt name=txtDueToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="종료만기일자"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>차입거래처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBpCd" SIZE="10" MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="차입거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 5)">
									                       <INPUT TYPE=TEXT NAME="txtBpNm" ALT="차입거래처명" SIZE=20 tag="24X"></TD>																		
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>장단기구분</TD>
						<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="장단기구분" STYLE="WIDTH: 135px" tag="11X"><OPTION VALUE=""></OPTION></SELECT></TD>
						<TD CLASS=TD5 NOWRAP>차입용도</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoanType" ALT="차입용도" SIZE="10" MAXLENGTH="2"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.value, 6)">&nbsp;<INPUT NAME="txtLoanTypeNm" ALT="차입용도명" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>						
					</TR>
					<TR>
<% IF Request("PGM") = "F4235MA1" THEN %>
						<TD CLASS="TD5" NOWRAP>만기연장번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="만기연장번호" SIZE="20" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanNo.value,3)"></TD>
<% ELSE%>
						<TD CLASS="TD5" NOWRAP>차입금번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="차입금번호" SIZE="20" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanNo.value,3)"></TD>
<% END IF %>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"><PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ONCLICK="FncQuery()"></IMG>
					&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)" ONCLICK="OkClick()"></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hBankLoanCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanType" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDueFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDueToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPgmId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

