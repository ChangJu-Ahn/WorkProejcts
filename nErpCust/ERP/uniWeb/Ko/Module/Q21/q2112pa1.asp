<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2112PA1
'*  4. Program Name         : 
'*  5. Program Desc         : 품목별 검사기준 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "q2112pb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim C_InspItemCd
Dim C_InspItemNm
Dim C_InspOrder 
Dim C_InspMthdCd
Dim C_InspMthdNm
Dim C_InspUnitIndctnCd
Dim C_InspUnitIndctnNm
Dim C_InspSeries
Dim C_SampleQty 
Dim C_AccptncNumber
Dim C_RejtnNumber
Dim C_AccptncCoefficient
Dim C_MaxDefectRatio 
Dim C_InspSpec 
Dim C_LSL 
Dim C_USL
Dim C_MsmtEqpmtCd
Dim C_MsmtEqpmtNm
Dim C_MsmtUnit

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim ArrParent
Dim arrParam					 '--- First Parameter Group 
ReDim arrParam(1)
Dim arrReturn				 '--- Return Parameter Group 

Dim IsOpenPop          

 '------ Set Parameters from Parent ASP ------
ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam(0) = ArrParent(1)
arrParam(1) = ArrParent(2)

top.document.title = PopupParent.gActivePRAspName
'top.document.title = "검사기준 팝업"
 '--------------------------------------------- 

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
	lgSortKey    = 1                            '⊙: initializes sort direction
End Function

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
    C_InspItemCd		= 1
    C_InspItemNm		= 2
    C_InspOrder			= 3
    C_InspMthdCd		= 4
    C_InspMthdNm		= 5
    C_InspUnitIndctnCd	= 6
    C_InspUnitIndctnNm	= 7
    C_InspSeries		= 8
    C_SampleQty			= 9
    C_AccptncNumber		= 10
    C_RejtnNumber		= 11
    C_AccptncCoefficient = 12
    C_MaxDefectRatio	= 13
    C_InspSpec			= 14
    C_LSL				= 15
    C_USL				= 16
    C_MsmtEqpmtCd		= 17
    C_MsmtEqpmtNm		= 18
    C_MsmtUnit			= 19
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	txtInspReqNo.Value = arrParam(0)
	txtLotSize.Value = arrParam(1)
	
	Self.Returnvalue = Array("")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021216",,PopupParent.gAllowDragDropSpread    

	vspdData.ReDraw = False
	vspdData.MaxCols = C_MsmtUnit + 1
	vspdData.MaxRows = 0
	
	Call AppendNumberPlace("6", "3","0")
	Call AppendNumberPlace("7", "15","4")

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit C_InspItemCd, "검사항목코드", 15, 0, -1, 40
	ggoSpread.SSSetEdit C_InspItemNm, "검사항목명", 15, 0, -1, 40
	ggoSpread.SSSetFloat C_InspOrder, "검사순서", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec, , , "P"
	ggoSpread.SSSetEdit C_InspMthdCd, "검사방식코드", 15, 0, -1, 40
	ggoSpread.SSSetEdit C_InspMthdNm, "검사방식명", 15, 0, -1, 40
	ggoSpread.SSSetEdit C_InspUnitIndctnCd, "검사단위 품질표시코드", 10, 0, -1, 1
   	ggoSpread.SSSetEdit C_InspUnitIndctnNm, "검사단위 품질표시", 20, 0, -1, 40
	ggoSpread.SSSetFloat C_InspSeries, "차수", 7, "6", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	ggoSpread.SSSetFloat C_SampleQty, "시료수", 10, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
   	ggoSpread.SSSetFloat C_AccptncNumber, "합격판정개수", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
   	ggoSpread.SSSetFloat C_RejtnNumber, "불합격판정개수", 16, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
   	ggoSpread.SSSetFloat C_AccptncCoefficient, "합격판정계수", 15, "7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
   	ggoSpread.SSSetFloat C_MaxDefectRatio, "최대허용불량률", 16, "7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	ggoSpread.SSSetEdit C_InspSpec, "검사규격", 11, 2, -1, 40
	ggoSpread.SSSetFloat C_LSL, "하한규격", 15, "7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
   	ggoSpread.SSSetFloat C_USL, "상한규격", 16, "7", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
   	ggoSpread.SSSetEdit C_MsmtEqpmtCd, "측정기기코드", 15, 2, -1, 40
	ggoSpread.SSSetEdit C_MsmtEqpmtNm, "측정기기", 11, 2, -1, 40
	ggoSpread.SSSetEdit C_MsmtUnit, "측정단위", 11, 2, -1, 40

	Call ggoSpread.SSSetColHidden(C_InspUnitIndctnCd, C_InspUnitIndctnCd, True)
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	vspdData.ReDraw = True
	
	Call SetSpreadLock()
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method lock spreadsheet
'========================================================================================================
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

			C_InspItemCd			= iCurColumnPos(1)
			C_InspItemNm			= iCurColumnPos(2)
			C_InspOrder				= iCurColumnPos(3)
			C_InspMthdCd			= iCurColumnPos(4)
			C_InspMthdNm			= iCurColumnPos(5)
			C_InspUnitIndctnCd		= iCurColumnPos(6)
			C_InspUnitIndctnNm		= iCurColumnPos(7)
			C_InspSeries			= iCurColumnPos(8)
			C_SampleQty				= iCurColumnPos(9)
			C_AccptncNumber			= iCurColumnPos(10)
			C_RejtnNumber			= iCurColumnPos(11)
			C_AccptncCoefficient	= iCurColumnPos(12)
			C_MaxDefectRatio		= iCurColumnPos(13)
			C_InspSpec				= iCurColumnPos(14)
			C_LSL					= iCurColumnPos(15)
			C_USL					= iCurColumnPos(16)
			C_MsmtEqpmtCd			= iCurColumnPos(17)
			C_MsmtEqpmtNm			= iCurColumnPos(18)
			C_MsmtUnit				= iCurColumnPos(19)
			
    End Select    

End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt, iCurColumnPos
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 2)
	
		ggoSpread.Source = vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		vspdData.Row = vspdData.ActiveRow 
				
		For intColCnt = 0 To vspdData.MaxCols - 2
			vspddata.Col = iCurColumnPos(CInt(intColCnt + 1))
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call InitVariables
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call fncQuery()
End Sub
'==========================================  3.2.1 fncQuery=======================================
'========================================================================================================
Function FncQuery()
	vspdData.MaxRows = 0
	lgQueryFlag = "1"
	lgStrPrevKey1 = ""
	lgStrPrevKey2 = ""
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	If DbQuery = false then
		Exit Function
	End If
	fncQuery = True

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

	Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              ' 타이틀 cell을 dblclick했거나....
	   Exit Function
	End If
	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
	End If
End Function

Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" AND lgStrPrevKey2 <> ""  Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DBQuery = False Then
				'Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If   
	
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
	vspdData.Redraw = True
End Sub

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	
	Dim strVal
	Dim txtMaxRows
	
	'Show Processing Bar
    Call LayerShowHide(1)  

	DbQuery = False 
	txtMaxRows = vspdData.MaxRows
	
	strVal = BIZ_PGM_ID & "?txtInspReqNo=" & Trim(txtInspReqNo.Value)
	strVal = strVal & "&txtLotSize=" & Trim(txtLotSize.Value)
	strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
	strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtMaxRows=" & txtMaxRows		
	
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
	DbQuery = True 
	
End Function

Function DbQueryOk()								'☆: 조회 성공후 실행로직 

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR HEIGHT=*>
		<TD  WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사의뢰번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20 MAXLENGTH=18 tag="14" ALT="검사의뢰번호"></TD>
									<TD CLASS="TD5" NOWRAP>로트크기</TD>            
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotSize" SIZE=15 MAXLENGTH=15 ALT="LOT SIZE" tag="14" STYLE="Text-Align: Right"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD>
									<script language =javascript src='./js/q2112pa1_I816435408_vspdData.js'></script>
								</TD>
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
	<TR HEIGHT=20>
      		<TD WIDTH=100%>
      			<TABLE <%=LR_SPACE_TYPE_30%>>
        				<TR>        				
        					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/query_d.gif" Style="CURSOR: hand" ALT="Search" NAME="search" OnClick="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
