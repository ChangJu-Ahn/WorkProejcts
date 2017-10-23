
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Operation Popup															*
'*  3. Program ID			: p4711pa1.asp																*
'*  4. Program Name			: 이력번호 Popup															*
'*  5. Program Desc			: 이력번호 Popup															*
'*  7. Modified date(First)	: 2001/12/14																*
'*  8. Modified date(Last)	: 2002/12/11																*
'*  9. Modifier (First)     : Park, Bum-Soo																*
'* 10. Modifier (Last)		: Ryu Sung Won																*
'* 11. Comment 				:
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin																			*
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--#####################################################################################################
'#						1. 선 언 부																		#
'#####################################################################################################-->
<!--********************************************  1.1 Inc 선언  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 공통 Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit
	
'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================
Const BIZ_PGM_QRY_ID = "p4711pb1.asp"					<% '☆: 비지니스 로직 ASP명 %>
		
Dim C_BatchRunNo
Dim C_ExecStartDt
Dim C_ProdtOrderNoFrom
Dim C_ProdtOrderNoTo
Dim C_ItemCdFrom
Dim C_ItemCdTo
Dim C_WcCdFrom
Dim C_WcCdTo
Dim C_ShiftCdFrom
Dim C_ShiftCdTo
Dim C_ReportDtFrom
Dim C_ReportDtTo
Dim C_Status
Dim C_Success_Cnt
Dim C_Error_Cnt
Dim C_InsrtUserId
	
'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim arrReturn
Dim IsOpenPop
Dim lgNextNo
Dim lgPrevNo
Dim lgPlantCD
Dim ArrParent

ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable값 정의  ============================
'========================================================================================================
'----------------  공통 Global 변수값 정의  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						2. Function 부																	#
'#																										#
'#	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기술					#
'#	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.							#
'#						 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)			#
'########################################################################################################
'*******************************************  2.1 변수 초기화 함수  *************************************
'*	기능: 변수초기화																					*
'*	Description : Global변수 처리, 변수초기화 등의 작업을 한다.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_BatchRunNo		= 1
	C_ExecStartDt		= 2
	C_ProdtOrderNoFrom	= 3
	C_ProdtOrderNoTo	= 4
	C_ItemCdFrom		= 5
	C_ItemCdTo			= 6
	C_WcCdFrom			= 7
	C_WcCdTo			= 8
	C_ShiftCdFrom		= 9
	C_ShiftCdTo			= 10
	C_ReportDtFrom		= 11
	C_ReportDtTo		= 12
	C_Status			= 13
	C_Success_Cnt		= 14
	C_Error_Cnt			= 15
	C_InsrtUserId		= 16
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
	lgStrPrevKey = ""                           'initializes Previous Key		
	Self.Returnvalue = Array("")
End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE","PA") %>
	<% Call loadBNumericFormatA("Q", "P", "NOCOOKIE","PA") %>
End Sub
	
Function InitSetting()
	txtPlantCd.Value = ArrParent(1)
	txtPlantNm.Value = ArrParent(2)
	txtBatchRunNo.Value = ArrParent(3)
End Function
	
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
	
'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	
	vspdData.MaxCols = C_InsrtUserId + 1
	vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit		C_BatchRunNo,			"이력번호", 18
	ggoSpread.SSSetDate 	C_ExecStartDt,			"실행일", 12, 2, gDateFormat
	ggoSpread.SSSetEdit		C_ProdtOrderNoFrom,		"시작오더번호", 18
	ggoSpread.SSSetEdit		C_ProdtOrderNoTo,		"종료오더번호", 18
	ggoSpread.SSSetEdit		C_ItemCdFrom,			"시작품목", 18
	ggoSpread.SSSetEdit		C_ItemCdTo,				"종료품목", 18
	ggoSpread.SSSetEdit		C_WcCdFrom,				"시작작업장", 10
	ggoSpread.SSSetEdit		C_WcCdTo,				"종료작업장", 10
	ggoSpread.SSSetEdit		C_ShiftCdFrom,			"시작 Shift", 10
	ggoSpread.SSSetEdit		C_ShiftCdTo,			"종료 Shift", 10
	ggoSpread.SSSetDate 	C_ReportDtFrom,			"시작실적일", 12, 2, gDateFormat
	ggoSpread.SSSetDate 	C_ReportDtTo,			"종료실적일", 12, 2, gDateFormat
	ggoSpread.SSSetEdit		C_Success_Cnt,			"감안된실적수", 10
	ggoSpread.SSSetEdit		C_Error_Cnt,			"에러수", 10
	ggoSpread.SSSetEdit		C_InsrtUserId,			"실행자ID", 13
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	Call ggoSpread.SSSetColHidden(C_Status,C_Status, True)
    
    ggoSpread.SSSetSplit2(2)
	vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub
	

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

	ggoSpread.Source = vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BatchRunNo		= iCurColumnPos(1)
			C_ExecStartDt		= iCurColumnPos(2)
			C_ProdtOrderNoFrom	= iCurColumnPos(3)
			C_ProdtOrderNoTo	= iCurColumnPos(4)
			C_ItemCdFrom		= iCurColumnPos(5)
			C_ItemCdTo			= iCurColumnPos(6)
			C_WcCdFrom			= iCurColumnPos(7)
			C_WcCdTo			= iCurColumnPos(8)
			C_ShiftCdFrom		= iCurColumnPos(9)
			C_ShiftCdTo			= iCurColumnPos(10)
			C_ReportDtFrom		= iCurColumnPos(11)
			C_ReportDtTo		= iCurColumnPos(12)
			C_Status			= iCurColumnPos(13)
			C_Success_Cnt		= iCurColumnPos(14)
			C_Error_Cnt			= iCurColumnPos(15)
			C_InsrtUserId		= iCurColumnPos(16)
    End Select    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	If vspdData.MaxRows > 0 Then
		
		Dim intRowCnt
		Dim intColCnt
		Dim intSelCnt

		intSelCnt = 0
		Redim arrReturn(3)
		
		vspdData.Row = vspdData.ActiveRow

		If vspdData.SelModeSelected = True Then
			vspdData.Col = C_BatchRunNo
			arrReturn(0) = vspdData.Text
			vspdData.Col = C_Status
			arrReturn(1) = vspdData.Text
			vspdData.Col = C_Success_Cnt
			arrReturn(2) = vspdData.Text
			vspdData.Col = C_Error_Cnt
			arrReturn(3) = vspdData.Text
		End If

		Self.Returnvalue = arrReturn
		Self.Close()
	End If			
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=13 and vspdData.activeRow > 0 Then
 		Call OkClick()
			
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub	

'*******************************************  2.4 POP-UP 처리함수  **************************************
'*	기능: POP-UP																						*
'*	Description : POP-UP Call하는 함수 및 Return Value setting 처리										*
'********************************************************************************************************
'===========================================  2.4.1 POP-UP Open 함수()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개별 프로그램마다 필요한 개발자 정의 Procedure(Sub, Function, Validation & Calulation 관련 함수)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'########################################################################################################
'#						3. Event 부																		#
'#	기능: Event 함수에 관한 처리																		#
'#	설명: Window처리, Single처리, Grid처리 작업.														#
'#		  여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.								#
'#		  각 Object단위로 Grouping한다.																	#
'########################################################################################################
'********************************************  3.1 Window처리  ******************************************
'*	Window에 발생 하는 모든 Even 처리																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'⊙: Load table , B_numeric_format		
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call InitVariables
	Call InitSpreadSheet()
	Call InitSetting()
	Call FncQuery()
	vspdData.Row = 1
	vspdData.Col = 1
	Call SetFocusToDocument("M")
	vspdData.Focus
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
   FncQuery = False
	Call DbQuery()
   Fncquery = False
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

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = vspdData
    Call SetPopupMenuItemInf("0000111111")
    
    If vspdData.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################
'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	    
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkfield(Document, "1") Then									'⊙: This function check indispensable field
	   Exit Function
	End If
	    
    DbQuery = False                                                         <%'⊙: Processing is NG%>
	    
    vspdData.MaxRows = 0
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
    strVal = strVal & "&txtPlantCd=" & txtPlantCd.Value						<%'☆: 조회 조건 데이타 %>
    strVal = strVal & "&txtBatchRunNo=" & txtBatchRunNo.Value
	If rdoDeleteFlg1.checked = True Then
		strVal = strVal & "&txtrdoflag=" & "C"
	Else
		strVal = strVal & "&txtrdoflag=" & "R"
	End If

    Call RunMyBizASP(MyBizASP, strVal)					<%'☜: 비지니스 ASP 를 가동 %>
		
    DbQuery = True                                                          			<%'⊙: Processing is NG%>

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
		Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    lgIntFlgMode = PopupParent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode    
    
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>공장</TD>
						<TD CLASS=TD6 NOWRAP colspan=3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14XXXU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>이력번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBatchRunNo" SIZE=18 MAXLENGTH=18 tag="11XXXU"  ALT="이력번호"></TD>
						<TD CLASS=TD5 NOWRAP>취소구분</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoDeleteFlg" ID="rdoDeleteFlg1" CLASS="RADIO" tag="11" Value="Y"><LABEL FOR="rdoDeleteFlg1">예</LABEL>
						     				 <INPUT TYPE="RADIO" NAME="rdoDeleteFlg" ID="rdoDeleteFlg2" CLASS="RADIO" tag="11" Value="N" CHECKED><LABEL FOR="rdoDeleteFlg2">아니오</LABEL></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4711pa1_vspdData_vspdData.js'></script>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK = "FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
