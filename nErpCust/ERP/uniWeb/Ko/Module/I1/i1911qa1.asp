<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 전표차이발생수불조회 
'*  3. Program ID           : i1911qa1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/20
'*  8. Modified date(Last)  : 2003/05/20
'*  9. Modifier (First)     : Ahn Jung Je
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit	

Const BIZ_PGM_QRY_ID = "i1911qb1.asp"	

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
DIm lgStrPrevKey1
DIm SetComboList

Dim C_DocumentDt
Dim C_ItemDocumentNo
Dim C_TrnsType
Dim C_MoveType
Dim C_MoveTypeNm
Dim C_Amount
Dim C_subcntrct_mfg_cost_amt
Dim C_sales_amt
Dim C_Sum_Amt
Dim C_item_loc_amt
Dim C_Difference
Dim C_journal_post
Dim C_batch_no
Dim C_gl_no
Dim C_pos_dt
Dim C_bizarea

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE 
	lgStrPrevKey  = ""               
	lgStrPrevKey1 = ""               
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim StartDate
	
	StartDate = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtTrnsFrDt.Text = UNIDateAdd("m", -1, StartDate, Parent.gDateFormat)
	frm1.txtTrnsToDt.Text = StartDate

	frm1.txtTrnsFrDt.focus 
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "QA") %>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD =" & FilterVar("I0002", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTrnsType,lgF0,lgF1,Chr(11))
	SetComboList = lgF0 & Chr(12) & lgF1
	
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030520", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_bizarea + 1
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetDate  C_DocumentDt,		"수불일자",     16, 2,Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_ItemDocumentNo,	"수불번호",     15
		ggoSpread.SSSetEdit  C_TrnsType,		"수불구분",     10
		ggoSpread.SSSetEdit  C_MoveType,		"수불유형",     10
		ggoSpread.SSSetEdit  C_MoveTypeNm,		"수불유형명",   15
		ggoSpread.SSSetFloat C_Amount,			"수불금액",		15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_subcntrct_mfg_cost_amt,	"외주가공비",	15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_sales_amt,		"매출가계정",	15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_Sum_Amt,			"합계금액",		15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_item_loc_amt,	"전표금액",		15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_Difference,		"차이금액",		15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit  C_journal_post,    "전표유무", 10, 2
		ggoSpread.SSSetEdit  C_batch_no,        "배치번호",		15
		ggoSpread.SSSetEdit  C_gl_no,           "전표번호",		15
		ggoSpread.SSSetDate  C_pos_dt,			"수불일자",     16, 2,Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_bizarea,         "사업장",		15

 		Call ggoSpread.SSSetColHidden(C_pos_dt, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.ReDraw = true
		
	    ggoSpread.SSSetSplit2(2)  
    End With
    
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_DocumentDt			 = 1
	C_ItemDocumentNo		 = 2
	C_TrnsType			 	 = 3
	C_MoveType				 = 4
	C_MoveTypeNm			 = 5
	C_Amount				 = 6
	C_subcntrct_mfg_cost_amt = 7								
	C_sales_amt				 = 8
	C_Sum_Amt				 = 9
	C_item_loc_amt			 = 10
	C_Difference			 = 11
	C_journal_post			 = 12
	C_batch_no				 = 13
	C_gl_no					 = 14
	C_pos_dt				 = 15
	C_bizarea				 = 16
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_DocumentDt			 = iCurColumnPos(1)
		C_ItemDocumentNo		 = iCurColumnPos(2)
		C_TrnsType			 	 = iCurColumnPos(3)
		C_MoveType				 = iCurColumnPos(4)
		C_MoveTypeNm			 = iCurColumnPos(5)
		C_Amount				 = iCurColumnPos(6)
		C_subcntrct_mfg_cost_amt = iCurColumnPos(7)					
		C_sales_amt				 = iCurColumnPos(8)
		C_Sum_Amt				 = iCurColumnPos(9)
		C_item_loc_amt			 = iCurColumnPos(10)
		C_Difference			 = iCurColumnPos(11)
		C_journal_post			 = iCurColumnPos(12)
		C_batch_no				 = iCurColumnPos(13)
		C_gl_no					 = iCurColumnPos(14)
		C_pos_dt				 = iCurColumnPos(15)
		C_bizarea				 = iCurColumnPos(16)
 	End Select
End Sub

'------------------------------------------  OpenMoveDtlRef()  -------------------------------------------------
Function OpenMoveDtlRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1
	Dim Param2
	Dim Param3
	Dim Param4
	Dim Param5
	Dim Param6
	Dim Param7
	If IsOpenPop = True Then Exit Function

	ggoSpread.Source = frm1.vspdData    

	With frm1.vspdData	    
		If .MaxRows = 0 Then
		    Call DisplayMsgBox("169804","X", "X", "X")
			Exit Function
		else
		   Call .GetText(C_bizarea, .ActiveRow, Param1)
		   Call .GetText(C_ItemDocumentNo, .ActiveRow, Param3)
		   Call .GetText(C_DocumentDt,	.ActiveRow, Param4)
		   Call .GetText(C_pos_dt,		.ActiveRow, Param5)
		   Call .GetText(C_TrnsType,	.ActiveRow, Param6)
		   Call .GetText(C_MoveTypeNm,	.ActiveRow, Param7)
		End If	
    End With
    	
    if Param3 = "" then
       Call DisplayMsgBox("169804","X", "X", "X") 
    	Exit Function
    End If

	If CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ", " BIZ_AREA_CD = " & FilterVar(Param1, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
		lgF0 = Split(lgF0, Chr(11))
		Param2 = lgF0(0)
	End If
		
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I1711RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1711RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6,Param7), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.focus
		Exit Function
	End If	
	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
    Call ggoOper.LockField(Document, "N") 
    
    Call InitVariables                    
    Call SetdefaultVal
    Call InitComboBox
    Call InitSpreadSheet                   
	Call SetToolbar("11000000000011")      
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_DblClick(Button)
'=======================================================================================================
Sub txtTrnsFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTrnsFrDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtTrnsFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_KeyPress()
'=======================================================================================================
Sub txtTrnsFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsToDt_DblClick(Button)
'=======================================================================================================
Sub txtTrnsToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtTrnsToDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtTrnsToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsToDt_KeyPress()
'=======================================================================================================
Sub txtTrnsToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub

	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If (lgStrPrevKey <> "" and lgStrPrevKey1 <> "") Then					
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End If
	End if  
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111") 
 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 		
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

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim TempInspDt 
    FncQuery = False  
    
    Err.Clear         

	If Not chkField(Document, "1") Then	Exit Function
	
	If GetSetupMod(Parent.gSetupMod, "a") <> "Y" then
       Call DisplayMsgBox("169934","X", "X", "X")
       Exit Function
	End if
	
	If ValidDateCheck(frm1.txtTrnsFrDt, frm1.txtTrnsToDt) = False Then 
   		frm1.txtTrnsFrDt.focus 
		Set gActiveElement = document.activeElement
		Exit Function
	End If
    
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 

    Call InitVariables

    If DbQuery = False Then Exit Function

    FncQuery = True	
    
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()	
    FncExit = True
End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub  FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub 
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
End Sub 

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
	Dim strVal
    
	On error resume next
    Err.Clear        
    
    Call LayerShowHide(1)

    DbQuery = False
    With frm1    
		strVal = BIZ_PGM_QRY_ID &	"?txtMode="			& Parent.UID_M0001			& _	
									"&txtTrnsFrDt="     & Trim(.txtTrnsFrDt.Text)	& _
									"&txtTrnsToDt="     & Trim(.txtTrnsToDt.Text)	& _
									"&cboTrnsType="     & Trim(.cboTrnsType.value)	& _
									"&lgStrPrevKey="    & Trim(lgStrPrevKey)		& _
									"&lgStrPrevKey1="   & Trim(lgStrPrevKey1)		& _
									"&SetComboList="    & SetComboList				& _
									"&txtMaxRows="      & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)
    End With
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
    Call SetToolbar("11000000000111")
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>전표차이발생수불조회</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
		    		<TD WIDTH=* align=right><A href="vbscript:OpenMoveDtlRef()">수불상세정보</A> </TD>					
					<TD WIDTH=10></TD>
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
								<TD CLASS="TD5" NOWRAP>수불기간</TD>
								<TD CLASS="TD6" NOWRAP>
								    <script language =javascript src='./js/i1911qa1_fpDateTime1_txtTrnsFrDt.js'></script>
							        &nbsp;~&nbsp;
							        <script language =javascript src='./js/i1911qa1_fpDateTime2_txtTrnsToDt.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>수불구분</TD>
								<TD CLASS="TD6" NOWRAP>
								<SELECT Name="cboTrnsType" ALT="수불구분" STYLE="WIDTH: 100px" tag="11"><OPTION Value=""></OPTION></SELECT>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=*>
					<TD WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
								<script language =javascript src='./js/i1911qa1_OBJECT1_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

