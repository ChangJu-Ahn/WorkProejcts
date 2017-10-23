
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5461ra1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2003.06.19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : ahn, do hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "a5461rb1.asp"												'☆: 비지니스 로직 ASP명 
Const C_GL_NO = 9
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 5                                    '☆: key count of SpreadSheet

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim lgIsOpenPop
Dim lgMaxFieldCount
Dim lgCookValue
Dim lgSaveRow
'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------

Dim arrReturn
Dim arrParent
Dim arrParam
Dim txtPayNo
Dim txtBankPayCd
Dim txtFrDt
Dim txtToDt

	 '------ Set Parameters from Parent ASP ------ 
	arrParent = window.DialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)
	
	top.document.title = "부가세참조팝업"
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    Redim arrReturn(0,1)

'    lgStrPrevKey		= ""
    lgPageNo			= ""
    lgIntFlgMode		= popupparent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue	= False                    'Indicates that no value changed
	lgSortKey			= 1
	lgSaveRow			= 0
	Self.Returnvalue	= arrReturn

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 서버로 부터 필드 정의 정보를 가져옴 
'                 lgSort...로 시작하는 변수 영역에 sort대상 목록을 저장 
'                 IsPopUpR 변수영역에 sort 정보의 기본이 되는 값 저장 
'========================================================================================================= 
Sub SetDefaultVal()

'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	Dim arrUser
	arrUser = split(arrParam(0), popupparent.gcolsep)
	
	frm1.txtFrDt.Text		= Trim(arrUser(0))
	frm1.txtToDt.Text		= Trim(arrUser(1))
	frm1.txtVatIoFg.value	= Trim(arrUser(2))
	frm1.txtVatIoNm.value	= Trim(arrUser(3))
	frm1.txtVatTypeCd.value	= Trim(arrUser(4))
	frm1.txtVatTypeNm.value	= Trim(arrUser(5))
	frm1.txtGlInputCd.value	= Trim(arrUser(6))
	frm1.txtGlInputNm.value	= Trim(arrUser(7))
	frm1.hIssuedDt.value	= Trim(arrUser(8))
	frm1.hBpCd.value		= Trim(arrUser(9))
	frm1.hBpNm.value		= Trim(arrUser(10))
	frm1.hBizAreaCd.value	= Trim(arrUser(11))
	frm1.hBizAreaNm.value	= Trim(arrUser(12))

'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
	
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "*", "NOCOOKIE", "RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","RA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("A5461RA101","S","A","V20030707",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetZAdoSpreadSheet("A5461RA102","S","A","V20030707",PopupParent.C_SORT_DBAGENT,frm1.vspdData1, C_MaxKey, "X","X")
    Call SetSpreadLock() 
    
End Sub

'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
    ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True

	ggoSpread.Source = frm1.vspdData1
    .vspdData1.ReDraw = False
    ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData1.ReDraw = True

    End With
End Sub

'==================================++++==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow)
End Sub

'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		frm1.txtFrDt.focus
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
	frm1.txtFrDt.focus

End Function

'===========================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'===========================================================================
Function OpenSortPopup()

Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		frm1.txtFrDt.focus
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
   frm1.txtFrDt.focus
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    '---------Developer Coding part (Start)----------------------------------------------------------------
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,popupparent.gDateFormat,popupparent.gComNum1000,popupparent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field

    Call InitVariables																	'⊙: Initializes local global variables
    Call SetDefaultVal
    Call InitSpreadSheet()

	Call CurFormatNumericOCX()

    Call FncQuery()
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function



'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call popupparent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call popupparent.FncExport(popupparent.C_MULTI)

    FncExcel = True  
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call popupparent.FncFind(popupparent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	
		Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = popupparent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = popupparent.SS_SCROLLBAR_BOTH
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncExit = True
    
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal

	Err.Clear                                                                   '☜: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)
	
	With frm1
		strVal = BIZ_PGM_ID
    '---------Developer Coding part (Start)----------------------------------------------------------------

		strVal = strVal & "?txtFrDt="		& .txtFrDt.text
		strVal = strVal & "&txtToDt="		& .txtToDt.text
		strVal = strVal & "&txtVatIoFg="	& Trim(.txtVatIoFg.value)
		strVal = strVal & "&txtVatTypeCd="	& Trim(.txtVatTypeCd.value)
		strVal = strVal & "&txtGlInputCd="	& Trim(.txtGlInputCd.value)
		strVal = strVal & "&txtIssuedDt="	& Trim(.hIssuedDt.value)
		strVal = strVal & "&txtBpCd="		& Trim(.hBpCd.value)
		strVal = strVal & "&txtBizAreaCd="	& Trim(.hBizAreaCd.value)
	'---------Developer Coding part (End)----------------------------------------------------------------
		
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
	lgBlnFlgChgValue	=False
	lgIntFlgMode		= popupparent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
	lgSaveRow			= 1

	Call CurFormatNumericOCX()

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtFrDt.focus
	End If

'	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
'	Call SetToolbar("110000000001111")										'⊙: 버튼 툴바 제어 
End Function


'========================================================================================================
' Name : CurFormatNumericOCX
' Desc : 
'========================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtVatLocAmt1, popupparent.gCurrency, popupparent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, popupparent.gDateFormat, popupparent.gComNum1000, popupparent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtGlLocAmt1,  popupparent.gCurrency, popupparent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, popupparent.gDateFormat, popupparent.gComNum1000, popupparent.gComNumDec
	End With

End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  이 부분에서 컬럼 추가하고 데이타 전송이 일어나야 합니다.   							=
'========================================================================================================
Function OKClick()
		
	Dim ii ,jj ,kk
	if frm1.vspdData.SelModeSelCount > 0 Then 			
		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To frm1.vspdData.MaxRows - 1
			frm1.vspdData.Row = ii + 1			
			If frm1.vspdData.SelModeSelected Then
				For jj = 1 To C_MaxKey 
					frm1.vspdData.Col	 = GetKeyPos("A",jj )		
					arrReturn(kk,jj) = frm1.vspdData.Text
				Next			
				kk = kk + 1
			End If
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

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim ii
    gMouseClickStatus = "SPC"   
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

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
	
End Sub

Sub vspdData1_Click(ByVal Col, ByVal Row)
	Dim ii
    gMouseClickStatus = "SP1C"   
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	Call SetSpreadColumnValue("A", frm1.vspdData1, Col, Row)
	
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : 
'==========================================================================================
'Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	 	 
    	If lgPageNo <> "" Then		
    		dbquery						
'           Call DisableToolBar(popupparent.TBC_QUERY)
'           If DbQuery = False Then
'              Call RestoreToolBar()
'              Exit Sub
'			End if
    	End If
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

 '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 
'============================================================
'회계전표 팝업 
'============================================================
Function OpenPopupGL(byval VatGl)

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", POPUPparent.VB_INFORMATION, "a5120ra1", "X")
		lgIsOpenPop = False
		Exit Function
	End If	

	If lgIsOpenPop = True Then Exit Function
	
	If gMouseClickStatus = "SP1C" then
		With frm1.vspdData1
			If .ActiveRow > 0 Then
				.Row = .ActiveRow
				.Col = C_GL_NO
				arrParam(0) = Trim(.Text)	'회계전표번호 
				arrParam(1) = ""			'Reference번호 
			Else
				Call DisplayMsgBox("900025","X","X","X")
				Exit Function
			End If
		End With
	Else
		With frm1.vspdData
			If .ActiveRow > 0 Then
				.Row = .ActiveRow
				.Col = C_GL_NO
				arrParam(0) = Trim(.Text)	'회계전표번호 
				arrParam(1) = ""			'Reference번호 
			Else
				Call DisplayMsgBox("900025","X","X","X")
				Exit Function
			End If
		End With
	End IF	

	lgIsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.POPUPparent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
End Function




Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub



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
						<TD CLASS="TD5" NOWRAP>계산서발행일자</TD>
						<TD CLASS="TD6" NOWRAP>
							<script language =javascript src='./js/a5461ra1_fpFrDt_txtFrDt.js'></script>&nbsp;~&nbsp;
						    <script language =javascript src='./js/a5461ra1_fpToDt_txtToDt.js'></script> &nbsp; &nbsp;
						</TD>
						<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME=txtGlInputCd ALT="전표입력경로" SIZE=10 MAXLENGTH=20 tag="14NXXU">
							<INPUT TYPE=TEXT NAME=txtGlInputNm ALT="전표입력경로명" SIZE="18"  tag="14" >
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>계산서유형</TD>
						<TD CLASS=TD6 nowrap>
							<INPUT TYPE=TEXT NAME=txtVatTypeCd ALT="계산서유형" SIZE=10 MAXLENGTH=20 tag="14NXXU">
							<INPUT TYPE=TEXT NAME=txtVatTypeNm ALT="계산서유형명" SIZE="18" tag="14" >
						</TD>
						<TD CLASS=TD5 NOWRAP>매입매출구분</TD>
						<TD CLASS=TD6 nowrap>
							<INPUT TYPE=TEXT NAME=txtVatIoFg ALT="매입매출구분" SIZE=10 MAXLENGTH=20 tag="14NXXU">
							<INPUT TYPE=TEXT NAME=txtVatIoNm ALT="매입매출구분" SIZE=18  tag="14" >
						</TD>
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
				<TR>
					<TD CLASS="TD5" NOWRAP>부가세합계금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/a5461ra1_fpDoubleSingle2_txtVatLocAmt1.js'></script>
						&nbsp; &nbsp; &nbsp;<BUTTON NAME="btnGl" CLASS="CLSMBTN" ONCLICK="vbscript:OpenPopUpGl('VAT')">회계전표팝업</BUTTON></TD>
					<TD CLASS="TD5" NOWRAP>전표합계금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/a5461ra1_fpDoubleSingle2_txtGlLocAmt1.js'></script>
						&nbsp; &nbsp; &nbsp;<BUTTON NAME="btnGl" CLASS="CLSMBTN" ONCLICK="vbscript:OpenPopUpGl('GL')">회계전표팝업</BUTTON>
						</TD>
				</TR>
				<TR HEIGHT=100%>
					<TD HEIGHT="100%" WIDTH="50%" colspan=2>
						<script language =javascript src='./js/a5461ra1_OBJECT1_vspdData.js'></script>
					</TD>
					<TD HEIGHT="100%" WIDTH="50%" colspan=2>
						<script language =javascript src='./js/a5461ra1_OBJECT1_vspdData1.js'></script>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ONCLICK="FncQuery()"></IMG>&nbsp;
					&nbsp;
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
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hIssuedDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBpNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hBizAreaCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBizAreaNm" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
