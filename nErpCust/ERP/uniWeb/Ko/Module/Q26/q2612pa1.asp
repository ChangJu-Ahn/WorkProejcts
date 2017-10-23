<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2612PA1
'*  4. Program Name         : 
'*  5. Program Desc         : 
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

Const BIZ_PGM_ID = "q2612pb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_MgmtNo		
Dim C_ItemCd		
Dim C_ItemNm		
Dim C_WcCd		
Dim C_WcNm		
Dim C_FrameDt
Dim C_PlanDt	
Dim C_CounterPlanFlag
Dim C_OccurDtFr		
Dim C_OccurDtTo		
Dim C_Framer		

Dim lgQueryFlag				<% '--- 1:New Query 0:Continuous Query %>

Dim lgInspClassCd
Dim lgPlantCd
Dim lgItemCd
Dim lgMgmtNo
Dim lgWcCd
Dim lgFrameDt1
Dim lgFrameDt2
Dim lgPlanDt1
Dim lgPlanDt2
Dim lgCounterPlanFlag

Dim hPlantCd
Dim hItemCd
Dim hMgmtNo
Dim hInspClassCd
Dim hWcCd
Dim hFrameDt1
Dim hFrameDt2
Dim hPlanDt1
Dim hPlanDt2

Dim ArrParent
Dim arrParam					<% '--- First Parameter Group %>
ReDim arrParam(0)
Dim arrReturn				<% '--- Return Parameter Group %>

Dim IsOpenPop          

ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam(0) = ArrParent(1)

top.document.title = PopupParent.gActivePRAspName
'top.document.title = "이상대책보고 현황 팝업"

Function InitVariables()
	lgSortKey    = 1                            '⊙: initializes sort direction
End Function

Sub initSpreadPosVariables()  
    C_MgmtNo			= 1
    C_ItemCd			= 2
    C_ItemNm			= 3
    C_WcCd				= 4
    C_WcNm				= 5
    C_FrameDt			= 6
    C_PlanDt			= 7
    C_CounterPlanFlag	= 8
    C_OccurDtFr			= 9
    C_OccurDtTo			= 10
    C_Framer			= 11
    
End Sub

Sub SetDefaultVal()
	txtMgmtNo.Value 		= arrParam(0)
	Self.Returnvalue = Array("")
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q","NOCOOKIE","PA") %>
End Sub

Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0001", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(cboInspClassCd , lgF0, lgF1, Chr(11))
End Sub

Sub InitSpreadSheet()
	Call initSpreadPosVariables()    

	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021216",,PopupParent.gAllowDragDropSpread    

	vspdData.ReDraw = False

	vspdData.MaxCols = C_Framer + 1
	vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit C_MgmtNo,"관리번호",20
	ggoSpread.SSSetEdit C_ItemCd,"품목코드",15
	ggoSpread.SSSetEdit C_ItemNm,"품목명",20
	ggoSpread.SSSetEdit C_WcCd,"작업장코드",14
	ggoSpread.SSSetEdit C_WcNm,"작업장명",20
	ggoSpread.SSSetEdit C_FrameDt,"작성일",10, 2
	ggoSpread.SSSetEdit C_PlanDt,"대책일",10, 2
	ggoSpread.SSSetEdit C_CounterPlanFlag,"대책여부", 11
	ggoSpread.SSSetEdit C_OccurDtFr,"발생기간(F)",15, 2
	ggoSpread.SSSetEdit C_OccurDtTo,"발생기간(T)",15, 2
	ggoSpread.SSSetEdit C_Framer,"작성자",10
	
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	vspdData.ReDraw = True
	
	Call SetSpreadLock()
End Sub

Sub SetSpreadLock()
	ggoSpread.Source = vspdData
			
	vspdData.ReDraw = False
	ggoSpread.SpreadLock -1, -1
	vspdData.ReDraw = True
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_MgmtNo			= iCurColumnPos(1)
			C_ItemCd			= iCurColumnPos(2)
			C_ItemNm			= iCurColumnPos(3)
			C_WcCd				= iCurColumnPos(4)
			C_WcNm				= iCurColumnPos(5)
			C_FrameDt			= iCurColumnPos(6)
			C_PlanDt			= iCurColumnPos(7)
			C_CounterPlanFlag	= iCurColumnPos(8)
			C_OccurDtFr			= iCurColumnPos(9)
			C_OccurDtTo			= iCurColumnPos(10)
			C_Framer			= iCurColumnPos(11)
			
    End Select    

End Sub

Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"											<%' 팝업 명칭 %>
	arrParam(1) = "B_PLANT"											<%' TABLE 명칭 %>
	arrParam(2) = Trim(txtPlantCd.Value)   									<%' Code Condition%>
	arrParam(3) = ""      											<%' Name Condition%>
	arrParam(4) = ""												<%' Where Condition%>
	arrParam(5) = "공장"											<%' TextBox 명칭 %>
		
   	arrField(0) = "PLANT_CD"	   										<%' Field명(0)%>
    	arrField(1) = "PLANT_NM"	   										<%' Field명(1)%>
    
    	arrHeader(0) = "공장코드"										<%' Header명(0)%>
    	arrHeader(1) = "공장명"											<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtPlantCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtPlantCd.Value    = arrRet(0)
		txtPlantNm.Value    = arrRet(1)
		txtPlantCd.Focus
	End If	

	Set gActiveElement = document.activeElement
	OpenPlant = true	
End Function

Function OpenItem()
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD
	
	'공장코드가 있는 지 체크 
	If Trim(txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705", "X", "X", "X")		<%'공장정보가 필요합니다 %>
		Exit Function	
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam1 = Trim(txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(txtItemNm.Value)	' Item Name
	arrParam5 = Trim(cboInspClassCd.Value)

	
	arrField(0) = 1 '"ITEM_CD"					<%' Field명(0)%>
    arrField(1) = 2 '"ITEM_NM"					<%' Field명(1)%>
    arrField(2) = 9 '"SPECIFICATION"				<%' Field명(1)%>
    arrField(3) = 6 '"BASIC_UNIT"					<%' Field명(1)%>
		
  	iCalledAspName = AskPRAspName("q1211pa2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtItemCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtItemCd.Value    = arrRet(0)		
		txtItemNm.Value    = arrRet(1)		
		txtItemCd.Focus
	End If	

	Set gActiveElement = document.activeElement	
	OpenItem = true
End Function

Function OpenWc()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'공장코드가 있는 지 체크 
	If Trim(txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705", "X", "X", "X")		<%'공장정보가 필요합니다 %>
		Exit Function	
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "작업장 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "P_WORK_CENTER"					<%' TABLE 명칭 %>
	arrParam(2) = Trim(txtWcCd.Value)					<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = "PLANT_CD = " & FilterVar(txtPlantCd.value, "''", "S") & "" 	<%' Where Condition%>
	arrParam(5) = "작업장"						<%' 조건필드의 라벨 명칭 %>	
	
    	arrField(0) = "Wc_CD"								<%' Field명(0)%>
    	arrField(1) = "Wc_NM"								<%' Field명(1)%>
    
    	arrHeader(0) = "작업장코드"					<%' Header명(0)%>
    	arrHeader(1) = "작업장명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	txtWcCd.Focus	
	If arrRet(0) = "" Then
		Exit Function
	Else
		txtWcCd.Value = arrRet(0)
		txtWcNm.Value = arrRet(1)
		txtWcCd.Focus
	End If	

	Set gActiveElement = document.activeElement	
	OpenWc = true	
End Function

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

Function CancelClick()
	Self.Close()
End Function

Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	Call InitVariables
	Call InitComboBox
	Call SetDefaultVal()
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

Sub txtFrameDt1_DblClick(Button)
    If Button = 1 Then
        txtFrameDt1.Action = 7
    End If
End Sub

Sub txtFrameDt2_DblClick(Button)
    If Button = 1 Then
        txtFrameDt2.Action = 7
    End If
End Sub

Sub txtPlanDt1_DblClick(Button)
    If Button = 1 Then
        txtPlanDt1.Action = 7
    End If
End Sub

Sub txtPlanDt2_DblClick(Button)
    If Button = 1 Then
        txtPlanDt2.Action = 7
    End If
End Sub

Function FncQuery()
	FncQuery = False
    	vspdData.MaxRows = 0

	lgQueryFlag = "1"
	
	lgPlantCd 	= Trim(txtPlantCd.Value)
	lgItemCd 		= Trim(txtItemCd.Value)
	lgInspClassCd	= Trim(cboInspClassCd.Value)
	lgMgmtNo 	= Trim(txtMgmtNo.Value)
	lgWcCd		= Trim(txtWcCd.Value)
	lgFrameDt1	= Trim(txtFrameDt1.Text)
	lgFrameDt2	= Trim(txtFrameDt2.Text)
	lgPlanDt1	= Trim(txtPlanDt1.Text)
	lgPlanDt2	= Trim(txtPlanDt2.Text)
	
	lgStrPrevKey = ""
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	if DbQuery = false then
		Exit Function
	End if
	
	fncQuery = True
End Function

Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

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
		If lgStrPrevKey <> "" Then
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

Function txtFrameDt1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

Function txtFrameDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

Function txtPlanDt1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

Function txtPlanDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	vspdData.Redraw = True
End Sub

Function DbQuery()
	
	Dim strVal
	Dim txtMaxRows
	
	DbQuery = False 
	
	If ValidDateCheck(txtFrameDt1, txtFrameDt2) = False Then
		Exit Function
	End If
	
	If ValidDateCheck(txtPlanDt1, txtPlanDt2) = False Then
		Exit Function
	End If
	
	'Show Processing Bar
    	Call LayerShowHide(1)  

	txtMaxRows = vspdData.MaxRows
	
	if lgStrPrevKey <> "" Then
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & hPlantCd
		strVal = strVal & "&txtItemCd=" & hItemCd
		strVal = strVal & "&txtInspClassCd=" & hInspClassCd
		strVal = strVal & "&txtMgmtNo=" & hMgmtNo
		strVal = strVal & "&txtWcCd=" & hWcCd
		strVal = strVal & "&txtFrameDt1=" & hFrameDt1
		strVal = strVal & "&txtFrameDt2=" & hFrameDt2
		strVal = strVal & "&txtPlanDt1=" & hPlanDt1
		strVal = strVal & "&txtPlanDt2=" & hPlanDt2
		strVal = strVal & "&txtCounterPlanFlag=" & lgCounterPlanFlag
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & txtMaxRows		
		
	else
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & lgPlantCd
		strVal = strVal & "&txtItemCd=" & lgItemCd
		strVal = strVal & "&txtInspClassCd=" & lgInspClassCd
		strVal = strVal & "&txtMgmtNo=" & lgMgmtNo
		strVal = strVal & "&txtWcCd=" & lgWcCd
		strVal = strVal & "&txtFrameDt1=" & lgFrameDt1
		strVal = strVal & "&txtFrameDt2=" & lgFrameDt2
		strVal = strVal & "&txtPlanDt1=" & lgPlanDt1
		strVal = strVal & "&txtPlanDt2=" & lgPlanDt2
		strVal = strVal & "&txtCounterPlanFlag=" & lgCounterPlanFlag
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & txtMaxRows	
		
	end if                                                        '⊙: Processing is NG
	
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="11XXXU" ></TD>								
									<TD CLASS="TD5" NOWRAP>검사분류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" TAG="11XXXU"><OPTION VALUE="" selected></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>관리번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMgmtNo" SIZE=20 MAXLENGTH=18 tag="11XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="품목" tag="11XXXU"><IMG align=top height=20 name=btnItemCd onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=20 ALT="작업장" tag="11XXXU"><IMG align=top height=20 name=btnWcCd onclick=vbscript:OpenWc() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
										<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>작성일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q2612pa1_fpDateTime1_txtFrameDt1.js'></script>&nbsp;~&nbsp; 
										<script language =javascript src='./js/q2612pa1_fpDateTime2_txtFrameDt2.js'></script>										
									</TD>			
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>대책일</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q2612pa1_fpDateTime3_txtPlanDt1.js'></script>&nbsp;~&nbsp; 
										<script language =javascript src='./js/q2612pa1_fpDateTime4_txtPlanDt2.js'></script>																				
									</TD>
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
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD>
									<script language =javascript src='./js/q2612pa1_I960403996_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
