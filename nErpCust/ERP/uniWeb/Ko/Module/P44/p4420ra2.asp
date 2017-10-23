
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: 																			*
'*  3. Program ID			: p4420ra2.asp																*
'*  4. Program Name			: 제품생산월보(작업장별상세)												*
'*  5. Program Desc			: Reference Popup															*
'*  6. Comproxy List        : ADO : 189611sac															*
'*  7. Modified date(First)	: 2000/12/19																*
'*  8. Modified date(Last)	: 2002/12/12																			*
'*  9. Modifier (First)		: Park, Bumsoo																		*
'* 10. Modifier (Last)		: Ryu Sung Won																			*
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
Const BIZ_PGM_ID = "p4420rb2.asp"					<% '☆: 비지니스 로직 ASP명 %>
		
'Const C_SHEETMAXROWS = 30

Dim C_WcCd
Dim C_WcNm
Dim C_ProdQty
Dim C_GoodQty
Dim C_BadQty
Dim C_RcptQty
Dim C_Unit
'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgPlantCd
Dim lgWcCd
Dim lgShiftCd
Dim lgItemAcct
Dim lgItemCd
Dim lgItemNm
Dim lgWorkDt

'*********************************************  1.3 변 수 선 언  ****************************************
'*	설명: Constant는 반드시 대문자 표기.																*
'********************************************************************************************************
Dim arrParent
Dim arrParam					
		
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)	
lgPlantCd	= arrParent(1)				
lgWorkDt	= arrParent(2)		
lgItemCd	= arrParent(3)
lgItemNm	= arrParent(4)		
lgItemAcct	= arrParent(5)
lgWcCd		= arrparent(6)	
lgShiftCd	= arrParent(7)		

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
	C_WcCd		= 1
	C_WcNm		= 2
	C_ProdQty	= 3
	C_GoodQty	= 4
	C_BadQty	= 5
	C_RcptQty	= 6
	C_Unit		= 7
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
	
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	'lblTitle.innerHTML = "제품생산월보(작업장별상세)"
		
	Dim lWorkDt
	lWorkDt = UniConvDateAToB(lgWorkDt, PopupParent.gServerDateFormat, PopupParent.gDateFormatYYYYMM)
		
	frm1.txtItemCd.value  = lgItemCd
	frm1.txtItemNm.value  = lgItemNm
	frm1.txtWorkDt.text	  = lWorkDt
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
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

    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

	frm1.vspdData.ReDraw = False
	        
    frm1.vspdData.MaxCols = C_Unit + 1
    frm1.vspdData.MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit 	C_WcCd,			 "작업장", 10
    ggoSpread.SSSetEdit 	C_WcNm,			 "작업장명"	,15
    ggoSpread.SSSetFloat	C_ProdQty,       "총생산량"	,15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    ggoSpread.SSSetFloat	C_GoodQty,		 "양품수량"	,15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    ggoSpread.SSSetFloat	C_BadQty,		 "불량수량"	,15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    ggoSpread.SSSetFloat	C_RcptQty,		 "입고수량"	,15,PopupParent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
    ggoSpread.SSSetEdit 	C_Unit,			 "단위",8
    
    Call ggoSpread.SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)
    
    ggoSpread.SSSetSplit2(1)
	frm1.vspdData.ReDraw = True
	Call SetSpreadLock()
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
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
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_WcCd		= iCurColumnPos(1)
			C_WcNm		= iCurColumnPos(2)
			C_ProdQty	= iCurColumnPos(3)
			C_GoodQty	= iCurColumnPos(4)
			C_BadQty	= iCurColumnPos(5)
			C_RcptQty	= iCurColumnPos(6)
			C_Unit		= iCurColumnPos(7)
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
    frm1.vspdData.Redraw = False
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

Sub vspdData_KeyPress(keyAscii)
	If keyAscii=27 Then
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
	
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtWorkDt, gDateFormat, "2")
	Call InitVariables
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
		
	Call SetDefaultVal()
	Call InitSpreadSheet()
	If DbQuery = False Then	
		Exit Sub
	End If
	Call SetFocusToDocument("P")
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("0000111111")
    
    If frm1.vspdData.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
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
	If Row = 0 Then Exit Function

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call OKClick
		End If
	End If
End Function

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
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

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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
Function FncQuery()
   On Error Resume Next
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
    Dim iColumnLimit2

    If gMouseClickStatus = "SPCRP" Then
       iColumnLimit = 7

       ACol = frm1.vspddata.ActiveCol
       ARow = frm1.vspddata.ActiveRow

       If ACol > iColumnLimit Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
          Exit Function  
       End If

       frm1.vspddata.ScrollBars = PopupParent.SS_SCROLLBAR_NONE

       ggoSpread.Source = frm1.vspddata
    
       ggoSpread.SSSetSplit(ACol)    

       frm1.vspddata.Col = ACol
       frm1.vspddata.Row = ARow
    
       frm1.vspddata.Action = 0    
    
       frm1.vspddata.ScrollBars = PopupParent.SS_SCROLLBAR_BOTH
    End If   
End Function

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
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
	Dim strVal
        
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
    
	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001	
   	strVal = strVal & "&txtPlantCd=" & lgPlantCd				'☆: 조회 조건 데이타 
	strVal = strVal & "&txtItemCd=" & lgItemCd				'☆: 조회 조건 데이타		
	strVal = strVal & "&txtWorkDt=" & lgWorkDt
	strVal = strVal & "&cboItemAcct=" & lgItemAcct
	strVal = strVal & "&txtWcCd=" & lgWcCd
	strVal = strVal & "&txtShliftCd=" & lgShiftCd
	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
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
    If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
	    Call SetActiveCell(frm1.vspdData,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
	End If
    
    lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELL SPACING=0>
	<TR>
		<TD HEIGHT=5 WIDTH=100%></TD>
	</TR>	
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>작업장별상세내역</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>		
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
 					<TR>
 						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="14" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 tag="14" ALT="품목명">&nbsp</TD>
						<TD CLASS=TD5 NOWRAP>작업년월</TD> 
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/p4420ra2_I329253876_txtWorkDt.js'></script>
						</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE WIDTH="100%" HEIGHT="100%">
				<TR HEIGHT="100%">
					<TD WIDTH="100%">
						<script language =javascript src='./js/p4420ra2_fpSpread1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
