<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Reference Popup For Production Order Detail List							*
'*  3. Program ID			: P4413RA1																			*
'*  4. Program Name			: 재작업오더이력															*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 2003/02/20																*
'*  8. Modified date(Last)	: 2003/02/20																*
'*  9. Modifier (First)     : Chen, Jae Hyun															*
'* 10. Modifier (Last)		: Chen, Jae Hyun																*
'* 11. Comment 				:
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin																			*
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--'####################################################################################################
'#						1. 선 언 부																		#
'#####################################################################################################-->
<!--'********************************************  1.1 Inc 선언  ****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ==================================
'=====================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script LANGUAGE="VBScript">

Option Explicit

'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

Const BIZ_PGM_ID = "p4413rb1.asp"							'☆: 비지니스 로직 ASP명 

Const C_SHEETMAXROWS = 30

Dim C_ReworkFlag		'= 1
Dim C_ProdtOrderNo		'= 2
Dim C_OprNo				'= 3
Dim C_ParentOrderNo		'= 4
Dim C_ParentOprNo		'= 5
Dim C_OrderStatus		'= 6
Dim C_ProdtOrderQty		'= 7
Dim C_ResultQty			'= 8
Dim C_ReworkQty			'= 9
Dim C_OrderUnit			'= 10
Dim C_PlanStartDt		'= 11
Dim C_PlanEndDt			'= 12
Dim C_DefectQty			'= 13
Dim C_InspDefectQty		'= 14
Dim C_TrackingNo		'= 15

'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->
Dim lgPlantCd
Dim lgItemCd
Dim lgProdtOrderNo
Dim lgOprNo

'*********************************************  1.3 변 수 선 언  ****************************************
'*	설명: Constant는 반드시 대문자 표기.																*
'********************************************************************************************************

Dim arrParent
Dim arrParam					
		
'------ Set Parameters from Parent ASP ------
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
lgPlantCd = arrParent(1)
lgItemCd = arrParent(2)
lgProdtOrderNo = arrParent(3)
lgOprNo = arrParent(4)
	
top.document.title = PopupParent.gActivePRAspName

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
	
	C_ReworkFlag		= 1
	C_ProdtOrderNo		= 2
	C_OprNo				= 3
	C_ParentOrderNo		= 4
	C_ParentOprNo		= 5
	C_OrderStatus		= 6
	C_ProdtOrderQty		= 7
	C_ResultQty			= 8
	C_ReworkQty			= 9
	C_OrderUnit			= 10
	C_PlanStartDt		= 11
	C_PlanEndDt			= 12
	C_DefectQty			= 13
	C_InspDefectQty		= 14
	C_TrackingNo		= 15

End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0
	lgStrPrevKey = ""
	Self.Returnvalue = Array("")
End Function

'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	txtPlantCd.value= lgPlantCd
	txtItemCd.value = lgItemCd
End Sub
	
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

    ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20030318",, PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	        
    vspdData.MaxCols = C_TrackingNo + 1
    vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit 	C_ReworkFlag,		"재작업", 8							' Rework Flag
	ggoSpread.SSSetEdit 	C_ProdtOrderNo,		"제조오더번호", 18						' Job Name
	ggoSpread.SSSetEdit 	C_OprNo,			"공정", 8							' Rework Flag
	ggoSpread.SSSetEdit 	C_ParentOrderNo,	"상위오더번호", 18						' Parent Order No
	ggoSpread.SSSetEdit 	C_ParentOprNo,		"상위공정", 8							' Parent Operation No
	ggoSpread.SSSetEdit 	C_OrderStatus,		"지시상태", 8					' Order Status
	ggoSpread.SSSetFloat	C_ProdtOrderQty,	"오더수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ResultQty,		"실적수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_ReworkQty,		"재작업수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit 	C_OrderUnit,		"오더단위", 8					' 
	ggoSpread.SSSetDate 	C_PlanStartDt,		"착수예정일", 11, 2, PopupParent.gDateFormat	' Planned Start Date
	ggoSpread.SSSetDate 	C_PlanEndDt,		"완료예정일", 11, 2, PopupParent.gDateFormat	' Planned Completion Date
	ggoSpread.SSSetFloat	C_DefectQty,		"실적불량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_InspDefectQty,	"품질불량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit 	C_TrackingNo,		"Tracking No.", 25						' Tracking No
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols, True)
	
	ggoSpread.SSSetSplit2(3)
	vspdData.ReDraw = True
	
	Call SetSpreadLock()
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
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
			
			C_ReworkFlag		= iCurColumnPos(1)
			C_ProdtOrderNo		= iCurColumnPos(2)
			C_OprNo				= iCurColumnPos(3)
			C_ParentOrderNo		= iCurColumnPos(4)
			C_ParentOprNo		= iCurColumnPos(5)
			C_OrderStatus		= iCurColumnPos(6)
			C_ProdtOrderQty		= iCurColumnPos(7)
			C_ResultQty			= iCurColumnPos(8)
			C_ReworkQty			= iCurColumnPos(9)
			C_OrderUnit			= iCurColumnPos(10)
			C_PlanStartDt		= iCurColumnPos(11)
			C_PlanEndDt			= iCurColumnPos(12)
			C_DefectQty			= iCurColumnPos(13)
			C_InspDefectQty		= iCurColumnPos(14)
			C_TrackingNo		= iCurColumnPos(15)

            
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

'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
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
	If keyAscii =27 Then
 		Call CancelClick()
	End If
End Sub	

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
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call SetDefaultVal
	Call InitVariables											'⊙: Initializes local global variables
	Call ggoOper.LockField(Document, "Q")						'⊙: This function lock the suitable field
	Call InitSpreadSheet()		
	If DbQuery = False Then
		Exit Sub
	End If
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
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
'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    '----------  Coding part  -------------------------------------------------------------
    'If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
	'	If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
	'		DbQuery
	'	End If
    'End if
End Sub

'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
    Err.Clear								'☜: Protect system from crashing
	    
    DbQuery = False							'⊙: Processing is NG
	    
    Call LayerShowHide(1)
	    
    Dim strVal
	
    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001		'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtPlantCd=" & lgPlantCd							'☆: 조회 조건 데이타 
    strVal = strVal & "&txtItemCd=" & lgItemCd								'☆: 조회 조건 데이타 
    strVal = strVal & "&txtProdtOrderNo=" & lgProdtOrderNo					'☆: 조회 조건 데이타 
    strVal = strVal & "&txtOprNo=" & lgOprNo								'☆: 조회 조건 데이타 
	
    Call RunMyBizASP(MyBizASP, strVal)								'☜: 비지니스 ASP 를 가동 

    DbQuery = True													'⊙: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(LngMaxRows)												'☆: 조회 성공후 실행로직 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
	    Call SetActiveCell(vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
	End If
	
    lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=50>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>공장</TD>
						<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14xxxU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="14xxxU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 tag="14"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>	
		<TD HEIGHT=110>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>
					<TR>
						<TD CLASS=TD5 NOWRAP>기존오더번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOriginalOrderNo" SIZE=18 MAXLENGTH=18 tag="24xxxU" ALT="기존오더번호"></TD>
						<TD CLASS=TD5 NOWRAP>지시상태</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderStatus" SIZE=10 tag="24xxxU" ALT="지시상태"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>오더수량</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4413ra1_I313834553_txtOrderQty.js'></script></TD>
						<TD CLASS=TD5 NOWRAP>오더단위</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderUnit" SIZE=5 MAXLENGTH=3 tag="24xxxU" ALT="오더단위"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>실적수량</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4413ra1_I294154381_txtProdQty.js'></script></TD>
						<TD CLASS=TD5 NOWRAP>불량수량</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4413ra1_I840054029_txtDefectQty.js'></script></TD>
					</TR>
					<TR>
						
						<TD CLASS=TD5 NOWRAP>품질불량</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4413ra1_I385527694_txtInspDefectQty.js'></script></TD>
						<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="24xxxU" ALT="Tracking No."></TD>					
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>착수예정일</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4413ra1_I694159668_txtPlanStratDt.js'></script></TD>
						<TD CLASS=TD5 NOWRAP>완료예정일</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/p4413ra1_I992619955_txtPlanEndDt.js'></script></TD>
					</TR>
				</TABLE>
			</FIELDSET>		
		</TD>	
	</TR>
	<TR>
		<TD HEIGHT=100%>
			<script language =javascript src='./js/p4413ra1_vspdData_vspdData.js'></script>
		</TD>
	</TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
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
