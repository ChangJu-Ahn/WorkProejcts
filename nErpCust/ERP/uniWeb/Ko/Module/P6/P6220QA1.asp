<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : p6
'*  2. Function Name        : 금형수리내역조회(HB)
'*  3. Program ID           : p6220QA1
'*  4. Program Name         : 금형수리내역조회(HB)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/05/03
'*  8. Modified date(Last)  : 2005/07/20
'*  9. Modifier (First)     : Yoo Myung Sik
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

<%'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************%>
Const BIZ_PGM_ID = "P6220QB1.asp"												'☆: 비지니스 로직 ASP명 

<%'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================%>

Dim C_FAC_CAST_CD			'= 1
Dim C_CAST_NM				'= 2
Dim C_WORK_DT				'= 3
Dim C_MINOR_NM			'= 4
Dim C_INSP_TEXT			'= 5
Dim C_BP_NM				'= 6
Dim C_NAME				'= 7
Dim C_BIGO				'= 8

Const C_SHEETMAXROWS = 30

<% '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= %>
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size를 조사할 변수 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey
<% '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= %>
<% '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>
Dim IsOpenPop 
Dim lsDnNo 
Dim iDBSYSDate
Dim EndDate, StartDate,ACT_ROW,selChk,EndDate_,StartDate_

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = DateAdd("d", -7, EndDate)


<% '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### %>
<% '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= %>
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

<% '******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'********************************************************************************************************* %>
<% '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= %>
Sub SetDefaultVal()

	frm1.txtReqdlvyFromDt.text = StartDate
	frm1.txtReqdlvyToDt.text = Enddate
	Call BtnDisabled(1)
	selChk=false

End Sub

<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029(gCurrency, "I", "*") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    
		
	ggoSpread.Source = frm1.vspdData
		
	ggoSpread.Spreadinit	"V20021108",, parent.gAllowDragDropSpread    
		
	Call AppendNumberPlace("6", "5", "0")
		
	With frm1.vspdData
			
		.ReDraw = False
				  
		.MaxCols = C_BIGO + 1
		.MaxRows = 0
				
				
		Call ggoSpread.ClearSpreadData()	
				
		Call GetSpreadColumnPos("A")
	
		.ReDraw = false

	    ggoSpread.Source = frm1.vspdData			 

		ggoSpread.SSSetEdit		C_FAC_CAST_CD			, "금형코드"		, 10	
		ggoSpread.SSSetEdit		C_CAST_NM				, "금형명"			, 20
		ggoSpread.SSSetDate		C_WORK_DT				, "작업일자"		, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit		C_MINOR_NM				, "수리부위"		, 10
		ggoSpread.SSSetEdit		C_INSP_TEXT				, "점검내역"		, 40 
		ggoSpread.SSSetEdit		C_BP_NM					, "거래처"			, 15
		ggoSpread.SSSetEdit		C_NAME					, "작업자"			, 10
		ggoSpread.SSSetEdit		C_BIGO					, "비고"			, 10
		
				
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		

		ggoSpread.SpreadLockWithOddEvenRowColor()

		
		.ReDraw = True
    
    End With
    
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  	
	
	C_FAC_CAST_CD			= 1
	C_CAST_NM				= 2
	C_WORK_DT				= 3
	C_MINOR_NM				= 4
	C_INSP_TEXT				= 5
	C_BP_NM					= 6
	C_NAME					= 7
	C_BIGO					= 8

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
   
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)			
		
			C_FAC_CAST_CD			= iCurColumnPos(1)
			C_CAST_NM				= iCurColumnPos(2)
			C_WORK_DT				= iCurColumnPos(3)
			C_MINOR_NM				= iCurColumnPos(4)
			C_INSP_TEXT				= iCurColumnPos(5)			
			C_BP_NM					= iCurColumnPos(6)
			C_NAME					= iCurColumnPos(7)		
			C_BIGO					= iCurColumnPos(8)
		
    End Select    
End Sub

<%
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
%>
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

		ggoSpread.spreadlock		C_FAC_CAST_CD				, -1		
		ggoSpread.spreadlock		C_CAST_NM					, -1
		ggoSpread.spreadlock		C_WORK_DT					, -1
		ggoSpread.spreadlock		C_MINOR_NM					, -1
		ggoSpread.spreadlock		C_INSP_TEXT					, -1
		ggoSpread.spreadlock		C_BP_NM						, -1
		ggoSpread.spreadlock		C_NAME						, -1
		ggoSpread.spreadlock		C_BIGO						, -1
	
    .vspdData.ReDraw = True

    End With

End Sub

<% '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* %>

<% '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= %>
<% '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++%>

Function OpenRequried(ByVal iRequried)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried

	Case 1												
	
		arrParam(0) = "금형코드조회"					<%' 팝업 명칭 %>
		arrParam(1) ="Y_CAST"	<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtCastCd.value)		<%' Code Condition%>
		'arrParam(3) = Trim(frm1.txtDn_TypeNm.value)		<%' Name Cindition%>
		arrParam(4) = " " 
		arrParam(5) = "금형코드"			  	   <%' TextBox 명칭 %>

		arrField(0) = "CAST_CD"							<%' Field명(0)%>
		arrField(1) = "CAST_NM"							<%' Field명(1)%>

		arrHeader(0) = "금형코드"					<%' Header명(0)%>
		arrHeader(1) = "금형명칭"					<%' Header명(1)%>

			 
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRequried(arrRet,iRequried)
	End If	
	
End Function

<% '------------------------------------------  SetRequried()  --------------------------------------------------
'	Name : SetRequried()
'	Description : 거래처 Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetRequried(Byval arrRet,ByVal iRequried)

	Select Case iRequried
	Case 1
		
		frm1.txtCastCd.value = Trim(arrRet(0))
		frm1.txtCastNM.value = Trim(arrRet(1))	
			
	End Select
	
	lgBlnFlgChgValue=true
	

End Function


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

	Err.Clear

    Call LoadInfTB19029	

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec) 'condition
    
    Call ggoOper.LockField(Document, "N")      
    
    
	'----------  Coding part  -------------------------------------------------------------

	Call InitSpreadSheet

	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 

    Call InitVariables                                                      '⊙: Initializes local global variables

	Call SetDefaultVal
	
	frm1.txtCastCd.focus
	
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
'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
%>


<%
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
     '----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 	End If
End Sub

<%
'==========================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_MouseDown(Button , Shift , x , y)


    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

<%
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
%>
Sub vspdData_Change(ByVal Col , ByVal Row )


End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
    Call ggoSpread.ReOrderingSpreadData
    
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
    If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			DbQuery
		End If
    End if
    
End Sub

<%
'==========================================================================================
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() 시 Calendar Popup
'==========================================================================================
%>


<%
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 조회조건부의 OCX_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
%>

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
	Call BtnDisabled(1)
	Dim IntRetCD
	
	selChk=false

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 				
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
 										'⊙: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then
        Call RestoreToolBar()
        Exit Function
    End If      																'☜: Query db data
    
    FncQuery = True																'⊙: Processing is OK

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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
%>
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    
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
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
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

    DbQuery = False
    
    Err.Clear																	 '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
    
    
    With frm1

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001		
	
		strVal = strVal & "&txtCastCd=" & Trim(.txtCastCd.value)
		strVal = strVal & "&txtReqdlvyFromDt=" & Trim(.txtReqdlvyFromDt.text)
		strVal = strVal & "&txtReqdlvyToDt=" & Trim(.txtReqdlvyToDt.text)

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

Sub txtReqdlvyToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqdlvyToDt.Action = 7
	End If
End Sub

Sub txtReqdlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call fncquery()
End Sub

Sub txtReqdlvyFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqdlvyFromDt.Action = 7
	End If
End Sub

Sub txtReqdlvyFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call fncquery()
End Sub


Function DbQueryOk()														'☆: 조회 성공후 실행로직 

	lgIntFlgMode = parent.OPMD_UMODE	
	lgBlnFlgChgValue = False
    '-----------------------
    'Reset variables area
    '-----------------------
	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 
    Call ggoOper.LockField(Document, "2")									'⊙: This function lock the suitable field
End Function


'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	
    Dim strEbrFile
    Dim objName
    
	Dim var1

	dim strUrl
	dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)
	
	If frm1.vspdData.ActiveRow < 1 Then
		msgbox "먼저 출력할 금형코드를 클릭하십시요"
		exit function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_FAC_CAST_CD
		var1 = Trim(.Text)
	End With
	
	strUrl = "cast_cd|" & var1
	
	strEbrFile = "P6220OA1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
'----------------------------------------------------------------
' Print 함수에서 추가되는 부분 
'----------------------------------------------------------------
	call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()                                                    '☜: Protect system from crashing
    
    Dim strEbrFile
    Dim objName
    
	Dim var1

	
	dim strUrl
	dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)

	If frm1.vspdData.ActiveRow < 1 Then
		msgbox "먼저 출력할 금형코드를 클릭하십시요"
		exit function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_FAC_CAST_CD
		var1 = Trim(.Text)
	End With
	
	strUrl = "cast_cd|" & var1 
	
	ObjName = AskEBDocumentName("P6220OA1","ebr")

	call FncEBRPreview(objName, strUrl)

	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

<%
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
%>

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

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>					
					<TD CLASS="CLSMTABP" colspan=2>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>금형수리내역조회</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH= align=right></TD><TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>작업일자</TD>
									<TD CLASS=TD6 NOWRAP><TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/p6220qa1_fpDateTime1_txtReqdlvyFromDt.js'></script>
											&nbsp;~&nbsp;
											<script language =javascript src='./js/p6220qa1_fpDateTime1_txtReqdlvyToDt.js'></script>
											</TD>
										</TR>
													</TABLE></TD>
									<TD CLASS=TD5 NOWRAP>금형코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCastCd" ALT="금형코드" TYPE="Text" MAXLENGTH="13" SIZE=10 tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtCastNM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% >
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/p6220qa1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA Class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPostFlag" tag="14">

<INPUT TYPE=HIDDEN NAME="txtHDn_Type" tag="24">
<INPUT TYPE=HIDDEN NAME="hcastcd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSo_no" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHShip_to_party" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHReqGiDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHReqGiDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTrans_meth" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPostGiFlag" tag="24">


</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>