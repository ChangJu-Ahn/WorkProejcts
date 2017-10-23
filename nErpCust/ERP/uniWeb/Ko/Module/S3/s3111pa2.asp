<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1 %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3111PA2
'*  4. Program Name         : 수주관리번호 팝업(proforma invoice용)
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : son bum yeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>수주번호</TITLE>
<!--
'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************
-->
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->
<%'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************%>
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim lgMark                                                  <%'☜: 마크                                  %>
Dim IscookieSplit 

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3111pb2.asp"
Const C_SHEETMAXROWS    = 25                                   '☆: Spread sheet에서 보여지는 row
Const C_SHEETMAXROWS_D  = 30                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value
Const gstPaytermsMajor = "B9004"
 
                                            '☆: Jump시 Cookie로 보낼 Grid value
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

<% '#########################################################################################################
'												2. Function부 
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### %>

<% '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= %>
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

<% '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'                 lgSort...로 시작하는 변수 영역에 sort대상 목록을 저장 
'                 IsPopUpR 변수영역에 sort 정보의 기본이 되는 값 저장 
'========================================================================================================= %>
Sub SetDefaultVal()
<%'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------%>
	frm1.txtSOFrDt.text = StartDate
	frm1.txtSOToDt.text = EndDate
<%'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------%>

End Sub

<%'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================%>
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
		'------ Developer Coding part (End )   -------------------------------------------------------------- 

End Sub

<%'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================%>
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S3112pa1","S","A","V20021106", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock 
     
End Sub


<%'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================%>
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    .vspdData.OperationMode = 5
    End With
End Sub
<% '**********************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'************************************************************************************** %>
<%
'++++++++++++++++++++++++++++++++++++++++++++  OpenBizPartner()  ++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenBizPartner()																				+
'+	Description : Business Partner PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenBizPartner()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)
			
		If lgIsOpenPop = True Then Exit Function
		
		lgIsOpenPop = True
			
		arrParam(0) = "주문처"							
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtBpCd.value)				
		arrParam(3) = ""									
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
		arrParam(5) = "주문처"							
		
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
		
		arrHeader(0) = "주문처"							
		arrHeader(1) = "주문처명"						
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
		lgIsOpenPop = False
		
		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetBizPartner(arrRet)
		End If
	End Function

<%
'++++++++++++++++++++++++++++++++++++++++++++++  OpenMinorCd()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenMinorCd()																				+
'+	Description : Minor Code PopUp Window Call															+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If lgIsOpenPop = True Then Exit Function

		lgIsOpenPop = True

		arrParam(0) = strPopPos								
		arrParam(1) = "B_Minor"								
		arrParam(2) = Trim(strMinorCD)						
		arrParam(3) = ""						            
		arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		
		arrParam(5) = strPopPos								

		arrField(0) = "Minor_CD"							
		arrField(1) = "Minor_NM"							

		arrHeader(0) = strPopPos							
		arrHeader(1) = strPopPos & "명"					

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetMinorCd(strMajorCd, arrRet)
		End If
	End Function

<%
'++++++++++++++++++++++++++++++++++++++++++++  OpenSalesGroup()  +++++++++++++++++++++++++++++++++=++++++
'+	Name : OpenSalesGroup()																				+
'+	Description : Sales Order Type PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function OpenSalesGroup()
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If lgIsOpenPop = True Then Exit Function

		lgIsOpenPop = True

		arrParam(0) = "영업그룹"								
		arrParam(1) = "B_SALES_GRP"									
		arrParam(2) = Trim(frm1.txtSalesGroup.value)						
		arrParam(3) = ""											
		arrParam(4) = ""											
		arrParam(5) = "영업그룹"								

		arrField(0) = "SALES_GRP"									
		arrField(1) = "SALES_GRP_NM"										

		arrHeader(0) = "영업그룹"								
		arrHeader(1) = "영업그룹명"								

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		lgIsOpenPop = False

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetSalesGroup(arrRet)
		End If
	End Function
<%
'+++++++++++++++++++++++++++++++++++++++++++++++  OpenSOType()  +++++++++++++++++++++++++++++++++++++++++
'+	Name : OpenSOType()																					+
'+	Description : Sales Order Type PopUp Window Call													+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
Function OpenSOType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수주형태"					
	arrParam(1) = "S_SO_TYPE_CONFIG"				
	arrParam(2) = Trim(frm1.txtSo_Type.value)		
	arrParam(3) = ""								
	arrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "				
	arrParam(5) = "수주형태"					
		
    arrField(0) = "SO_TYPE"							
    arrField(1) = "SO_TYPE_NM"						
	    
    arrHeader(0) = "수주형태"					
    arrHeader(1) = "수주형태명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSOType(arrRet)
	End If	
End Function



<%
'=======================================  2.4.2 POP-UP Return값 설정 함수  ==============================
'=	Name : Set???()																						=
'=	Description : Reference 및 POP-UP의 Return값을 받는 부분											=
'========================================================================================================
%>

<%
'+++++++++++++++++++++++++++++++++++++++++++  SetBizPartner()  ++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetBizPartner()																				+
'+	Description : Set Return array from Business Partner PopUp Window									+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetBizPartner(arrRet)
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
	End Function

<%
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetMinorCd()																					+
'+	Description : Set Return array from Minor Code PopUp Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetMinorCd(strMajorCd, arrRet)
		frm1.txtPay_terms.value = arrRet(0)
		frm1.txtPay_terms_nm.value = arrRet(1)
	End Function
<%
'+++++++++++++++++++++++++++++++++++++++++++++  SetMinorCd()  +++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetMinorCd()																					+
'+	Description : Set Return array from Minor Code PopUp Window											+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetSOType(arrRet)
		frm1.txtSo_Type.value = arrRet(0)
		frm1.txtSo_TypeNm.value = arrRet(1)
	End Function
<%
'++++++++++++++++++++++++++++++++++++++++++++++  SetSOType()  +++++++++++++++++++++++++++++++++++++++++++
'+	Name : SetSalesGroup()																				+
'+	Description : Set Return array from Sales Order Type PopUp Window									+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
	Function SetSalesGroup(arrRet)
		frm1.txtSalesGroup.Value = arrRet(0)
		frm1.txtSalesGroupNm.Value = arrRet(1)
	End Function	
<% '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ %>

<% '++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<% '#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################%>
<% '******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* %>
<% '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= %>
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
    Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub
<%
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
%>
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

<% '**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* %>

<% '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* %>
<%
'=====================================  3.2.2 btnApplicant_OnClick()  ===================================
'========================================================================================================
%>
	Sub btnBpCdOnClick()
		Call OpenBizPartner()
	End Sub
<%
'======================================  3.2.4 btnSOType_OnClick()  =====================================
'========================================================================================================
%>
	Sub btnSalesGroupOnClick()
		Call OpenSalesGroup()
	End Sub
<%
'======================================  3.2.2 btnPayTerms_OnClick()  ===================================
'=	Event Name : btnPayTerms_OnClick																	=
'=	Event Desc :																						=
'========================================================================================================
%>
	Sub btnSoTypeOnClick()
		Call OpenSOType()
	End Sub
<%
'======================================  3.2.2 btnPayTerms_OnClick()  ===================================
'=	Event Name : btnPayTerms_OnClick																	=
'=	Event Desc :																						=
'========================================================================================================
%>
	Sub btnPayTermsOnClick()
		Call OpenMinorCd(frm1.txtPay_terms.value, frm1.txtPay_terms_nm.value, "결제방법", gstPaytermsMajor)
	End Sub

<%
'======================================  3.2.2 vspdData_KeyPress()  =====================================
'=	Event Name : vspdData_KeyPress																		=
'=	Event Desc :																						=
'========================================================================================================
%>
    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

<%'==================================== 3.2.23 txtSOFrDt_DblClick()  =====================================
'   Event Name : txtSOFrDt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================%>
	Sub txtSOFrDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtSOFrDt.Action = 7 
	    End If
	End Sub
<%'==================================== 3.2.23 txtSOFrDt_DblClick()  =====================================
'   Event Name : txtSOFrDt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================%>
	Sub txtSOToDt_DblClick(Button)
	    If Button = 1 Then
	        frm1.txtSOToDt.Action = 7 
	    End If
	End Sub

<%'==================================== 3.2.23 txtDt_KeyPress()  ========================================
'   Event Name : txtDt_KeyPress
'   Event Desc : keyboard Operation
'=======================================================================================================%>
	Sub txtSOFrDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub

	Sub txtSOToDt_Keypress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 27 Then
			Call CancelClick()
		Elseif KeyAscii = 13 Then
			Call FncQuery()
		End if
	End Sub
<%
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
%>
	Function vspdData_DblClick(ByVal Col, ByVal Row)

        If Row = 0 Or  frm1.vspdData.MaxRows = 0 Then 
             Exit Function
        End If	
	
		If frm1.vspdData.MaxRows > 0 Then
			If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End Function
	
<%
'========================================  3.3.2 vspdData_LeaveCell()  ==================================
'=	Event Name : vspdData_LeaveCell																		=
'=	Event Desc :																						=
'========================================================================================================
%>

	Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
		With frm1.vspdData
			If Row >= NewRow Then
				Exit Sub
			End If

			If NewRow = .MaxRows Then
				If lgStrPrevKey <> "" Then							<% '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
					DbQuery
				End If
			End If
		End With
	End Sub
	
<%
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
%>
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    <% '----------  Coding part  -------------------------------------------------------------%>   
    If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크'
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			DbQuery
		End If
   End if
    
End Sub
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
	Function OKClick()
		
		dim arrReturn
		If frm1.vspdData.ActiveRow > 0 Then				
		
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = 1
			arrReturn = frm1.vspdData.Text

			Self.Returnvalue = arrReturn
		End If

		Self.Close()
	End Function
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
	Function CancelClick()
		Self.Close()
	End Function


<% '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### %>


<% '#########################################################################################################
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
'######################################################################################################### %>
<% '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* %>
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 커야 할때 **
	If ValidDateCheck(frm1.txtSOFrDt, frm1.txtSOToDt) = False Then Exit Function

	If frm1.rdoComfirmFlg1.checked = True Then
		frm1.txtRadio.value = "A"
	ElseIf frm1.rdoComfirmFlg2.checked = True Then
		frm1.txtRadio.value = "Y"
	ElseIf frm1.rdoComfirmFlg3.checked = True Then
		frm1.txtRadio.value = "N"
	End If			   	

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

<%
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
%>
Function FncPrint() 
    Call parent.FncPrint()
End Function

<%
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
%>
Function FncExcel() 
	Call parent.FncExport(C_MULTI)
End Function

<%
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
%>
Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

<%
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
%>
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

<% '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* %>
<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
		strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.value)	<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtSalesGroup.value)
		strVal = strVal & "&txtSo_Type=" & Trim(frm1.txtSo_Type.value)
		strVal = strVal & "&txtPay_terms=" & Trim(frm1.txtPay_terms.value)
		strVal = strVal & "&txtRadio=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&txtSOFrDt=" & Trim(frm1.txtSOFrDt.text)
		strVal = strVal & "&txtSoToDt=" & Trim(frm1.txtSoToDt.text)
		
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '☜: 한번에 가져올수 있는 데이타 건수 
		strVal = strVal & "&lgSelectListDT=" & lgSelectListDT

        strVal = strVal & "&lgTailList="     & MakeSql()
		strVal = strVal & "&lgSelectList="   & EnCoding(lgSelectList)
       
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
    
    DbQuery = True


End Function

<%
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
%>
Function DbQueryOk()														'☆: 조회 성공후 실행로직 

    '-----------------------
    'Reset variables area
    '-----------------------
'    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

End Function

<%
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
%>
<%
'========================================================================================
' Function Name : MakeSql()
' Function Desc : Order by 절과 group by 절을 만든다.
'========================================================================================
%>
Function MakeSql()
    Dim iStr,jStr
    Dim ii,jj
    Dim iFirst
    
    iFirst = "N"
    iStr   = ""  
    jStr   = ""      

    Redim  lgMark(0) 
    Redim  lgMark(UBound(lgFieldNM)) 
    lgMark(0) = ""
    
    For ii = 0 to C_MaxSelList - 1
        If lgPopUpR(ii,0) <> "" Then
           If lgTypeCD(0) = "G" Then
              For jj = 0 To UBound(lgFieldNM) - 1                                            <%'Sort 대상리스트   저장 %>
                  If lgMark(jj) <> "X" Then
                     If lgPopUpR(ii,0) = lgFieldCD(jj) Then
                        If iFirst = "Y" Then
                           iStr = iStr & " , "
                           jStr = jStr & " , " 
                        End If   
                        If CInt(Trim(lgNextSeq(jj))) >= 1 And CInt(Trim(lgNextSeq(jj))) <= UBound(lgFieldNM) Then
                           iStr = iStr & " " & lgPopUpR(ii,0) & " " & lgPopUpR(ii,1) & " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
                           jStr = jStr & " " & lgPopUpR(ii,0) & " " &  " , " & lgFieldCD(CInt(lgNextSeq(jj)) - 1)
                           lgMark(CInt(lgNextSeq(jj)) - 1) = "X"
                        Else
                          iStr = iStr & " " & lgPopUpR(ii,0) & " " & lgPopUpR(ii,1)
                          jStr = jStr & " " & lgPopUpR(ii,0) 
                        End If
                        iFirst = "Y"
                        lgMark(jj) = "X"
                     End If
                     
                  End If
              Next
           Else
              If iFirst = "Y" Then
                 iStr = iStr & " , "
                 jStr = jStr & " , " 
              End If   
              iStr = iStr & " " & lgPopUpR(ii,0) & " " & "DESC"
              iFirst = "Y"
           End If
              
        End If
    Next     
    
    If lgTypeCD(0) = "G" Then
       MakeSql =  "Group By " & jStr  & " Order By " & iStr 
    Else
       MakeSql = "Order By" & iStr
    End If   


End Function
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<% '#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################### %>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>주문처</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnBpCdOnClick()">&nbsp;
							<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnSalesGroupOnClick()">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
						</TD>
					</TR>
					<TR>	
						<TD CLASS=TD5 NOWRAP>수주형태</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtSo_Type" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoType" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnSoTypeOnClick()">&nbsp;
							<INPUT NAME="txtSo_TypeNm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24">
						</TD>
						<TD CLASS=TD5 NOWRAP>수주일</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s3111pa2_fpDateTime2_txtSOFrDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/s3111pa2_fpDateTime2_txtSoToDt.js'></script>
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>결제방법</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtPay_terms" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK="Vbscript:btnPayTermsOnClick()">&nbsp;
							<INPUT NAME="txtPay_terms_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="24">
						</TD>
						<TD CLASS=TD5 NOWRAP>확정여부</TD> 
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoComfirmFlg" TAG="11" VALUE="A" CHECKED ID="rdoComfirmFlg1"><LABEL FOR="rdoComfirmFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoComfirmFlg" TAG="11" VALUE="Y" ID="rdoComfirmFlg2"><LABEL FOR="rdoComfirmFlg2">확정</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoComfirmFlg" TAG="11" VALUE="N" ID="rdoComfirmFlg3"><LABEL FOR="rdoComfirmFlg3">미확정</LABEL>			
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
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s3111pa2_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
