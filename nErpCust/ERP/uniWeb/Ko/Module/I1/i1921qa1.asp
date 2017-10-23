<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name            : Inventory
'*  2. Function Name          : 
'*  3. Program ID             : i1921qa1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 월별전표생성수불조회 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/05/23
'*  8. Modified date(Last)    : 2003/05/23
'*  9. Modifier (First)       : Ahn Jung Je
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=  "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit

Const BIZ_PGM_QRY1_ID	= "i1921qb1.asp"
Const BIZ_PGM_QRY2_ID	= "i1921qb2.asp"

Dim C_ItemAcct
Dim C_StartInv
Dim C_MRAmt
Dim C_PRAmt
Dim C_ORAmt
Dim C_STRAmt
Dim C_PIAmt
Dim C_DIAmt
Dim C_OIAmt
Dim C_STIAmt
Dim C_CloseInv
Dim C_ItemAcctCD

Dim C_TrnsType
Dim C_MoveType
Dim C_MoveTypeNm
Dim C_RAmount
Dim C_IAmount
Dim C_CostDevy
Dim C_SubMFG

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop 

Dim lgOldRow
Dim strYear,strMonth,strDay

'==========================================  2.1.1 InitVariables()  =====================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE
	lgIntGrpCount     = 0	
	lgBlnFlgChgValue  = False
	lgLngCurRows      = 0
    lgOldRow		  = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	frm1.txtDocumentDt.Text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat) 
 	Call ggoOper.FormatDate(frm1.txtDocumentDt, parent.gDateFormat, 2)
	frm1.txtDocumentDt.focus
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "QA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	
	If pvSpdNo = "" Or pvSpdNo = "A" Then
	
		Call InitSpreadPosVariables(pvSpdNo)

		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20030522", , Parent.gAllowDragDropSpread
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1

			.ReDraw = false
				 
			.MaxCols = C_ItemAcctCD + 1					
			.MaxRows = 0
				
			Call GetSpreadColumnPos("A")
				
			ggoSpread.SSSetEdit     C_ItemAcct, "품목계정", 10,,,,2
			ggoSpread.SSSetFloat    C_StartInv, "기초재고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_MRAmt,    "생산입고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_PRAmt,    "구매입고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_ORAmt,    "기타입고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_STRAmt,   "재고이동(입고)", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_PIAmt,    "생산출고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_DIAmt,    "판매출고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_OIAmt,    "기타출고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_STIAmt,   "재고이동(출고)", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_CloseInv, "기말재고", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetEdit     C_ItemAcctCD, "품목계정", 10,,,,2

 			Call ggoSpread.SSSetColHidden(C_ItemAcctCD, .MaxCols, True)
			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetSplit2(2)
				
			.ReDraw = true
		End With
	End If
		
	If pvSpdNo = "" Or pvSpdNo = "B" Then

		Call InitSpreadPosVariables(pvSpdNo)

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20030522", , Parent.gAllowDragDropSpread
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2

			.ReDraw = false

			.MaxCols = C_SubMFG +1										
			.MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit     C_TrnsType,     "수불구분",   15,,,,2
			ggoSpread.SSSetEdit		C_MoveType,     "수불유형",	  10
			ggoSpread.SSSetEdit		C_MoveTypeNm,   "수불유형명", 20
			ggoSpread.SSSetFloat    C_RAmount,		"입고금액", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_IAmount,		"출고금액", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_CostDevy,		"부대비", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetFloat    C_SubMFG,		"외주가공비", 15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000, parent.gComNumDec

 			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

			ggoSpread.SpreadLockWithOddEvenRowColor()
			ggoSpread.SSSetSplit2(2)
				
			.ReDraw = true
		End With
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "" Or pvSpdNo = "A" Then
	' Grid 1(vspdData1) - Operation 
		C_ItemAcct	= 1
		C_StartInv  = 2
		C_MRAmt     = 3
		C_PRAmt     = 4
		C_ORAmt     = 5
		C_STRAmt    = 6
		C_PIAmt     = 7
		C_DIAmt     = 8
		C_OIAmt     = 9
		C_STIAmt    = 10
		C_CloseInv  = 11
		C_ItemAcctCD = 12
	End If
	If pvSpdNo = "" Or pvSpdNo = "B"  Then	
		' Grid 2(vspdData2) - Operation 
		C_TrnsType   = 1
		C_MoveType	 = 2
		C_MoveTypeNm = 3
		C_RAmount	 = 4
		C_IAmount	 = 5
		C_CostDevy	 = 6
		C_SubMFG	 = 7
	End If
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData1 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		' Grid 1(vspdData1) - Operation 
		C_ItemAcct	= iCurColumnPos(1)
		C_StartInv  = iCurColumnPos(2)
		C_MRAmt     = iCurColumnPos(3)
		C_PRAmt     = iCurColumnPos(4)
		C_ORAmt     = iCurColumnPos(5)
		C_STRAmt    = iCurColumnPos(6)
		C_PIAmt     = iCurColumnPos(7)
		C_DIAmt     = iCurColumnPos(8)
		C_OIAmt     = iCurColumnPos(9)
		C_STIAmt    = iCurColumnPos(10)
		C_CloseInv  = iCurColumnPos(11)
		C_ItemAcctCD= iCurColumnPos(12)
	
	Case "B"
 		ggoSpread.Source = frm1.vspdData2 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		' Grid 2(vspdData2) - Operation 
		C_TrnsType   = iCurColumnPos(1)
		C_MoveType	 = iCurColumnPos(2)
		C_MoveTypeNm = iCurColumnPos(3)
		C_RAmount	 = iCurColumnPos(4)
		C_IAmount	 = iCurColumnPos(5)
		C_CostDevy	 = iCurColumnPos(6)
		C_SubMFG	 = iCurColumnPos(7)
 	End Select
End Sub

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet("")    
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("11000000000011")
End Sub


'=======================================================================================================
'   Event Name : txtDocumentDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtTrnsFrDt_KeyPress()
'=======================================================================================================
Sub txtDocumentDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111") 

 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
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

Sub vspdData2_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111") 

 	gMouseClickStatus = "SP2C"   
 	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2
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
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
  	If frm1.vspdData1.MaxRows = 0 Then Exit Sub
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
  	If frm1.vspdData2.MaxRows = 0 Then Exit Sub
End Sub
 
'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
End Sub  

'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
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
    
    Select case gActiveSpdSheet.id
    case "vaSpread1"
		Call InitSpreadSheet("A")	
	case "vaSpread2"
		Call InitSpreadSheet("B")
	End Select
    
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'==========================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, Byval NewCol, Byval NewRow, Byval Cancel)
	
	If NewRow <= 0 Or Row = NewRow Then Exit Sub

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
	If DbDtlQuery(NewRow) = False Then Exit Sub

End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery

	FncQuery = False
	
	on Error resume next
	Err.Clear
	
	Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then	Exit Function	
    
    If GetSetupMod(Parent.gSetupMod, "a") <> "Y" then
       Call DisplayMsgBox("169934","X", "X", "X")
       Exit Function
	End if
    
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

	Call ExtractDateFrom(frm1.txtDocumentDt.Text,frm1.txtDocumentDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	
	If DbQuery = False Then	Exit Function
	
	FncQuery = False
	
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
    Call parent.FncExport(parent.C_MULTI)			
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , True) 
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
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	
    Err.Clear	
	    
    DbQuery = False								
	    
    Call LayerShowHide(1)
	    
    Dim strVal
	
 	strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001		
    strVal = strVal & "&strYear="			 & strYear
    strVal = strVal & "&strMonth="			 & strMonth

    Call RunMyBizASP(MyBizASP, strVal)		

    DbQuery = True                          

End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()		
	
    '-----------------------
    'Reset variables area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	lgOldRow = 1
    Call SetToolbar("11000000000111")
	
	Call DbDtlQuery(1)

End Function

'========================================================================================
' Function Name : DbDtlQuery
'========================================================================================
Function DbDtlQuery(Byval Row) 

	Dim strVal			'
   
	DbDtlQuery = False
			
	frm1.vspdData1.Row = Row
	frm1.vspdData1.Col = C_ItemAcctCD

	Call LayerShowHide(1)

	strVal = BIZ_PGM_QRY2_ID & "?txtMode="   & parent.UID_M0001
	strVal = strVal & "&txtItemAcct="         & Trim(frm1.vspdData1.Text)
    strVal = strVal & "&strYear="			 & strYear	
    strVal = strVal & "&strMonth="			 & strMonth	

	Call RunMyBizASP(MyBizASP, strVal)					

	DbDtlQuery = True

End Function

Function DbDtlQueryOk()									
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%> >
		</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE  <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월별전표생성수불조회</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
     			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> >
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
					<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%> >
						<TR>
							<TD CLASS="TD5" NOWRAP>수불년월</TD>
							<TD CLASS="TD6">
							<script language =javascript src='./js/i1921qa1_OBJECT1_txtDocumentDt.js'></script>
							</TD>
							<TD CLASS="TD656">
							</TD>
						</TR>
					</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="35%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/i1921qa1_vaSpread1_vspdData1.js'></script>	
								</TD>
							</TR>
							<TR HEIGHT="65%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/i1921qa1_vaSpread2_vspdData2.js'></script>
								</TD>	
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> >
		</TD>	
	</TR>
	<TR HEIGHT=20 >
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%> >
				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1">
		</IFRAME>
		</TD>	
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>	
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>	
	
	
	
	
	
