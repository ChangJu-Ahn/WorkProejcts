<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory
'*  2. Function Name          : 
'*  3. Program ID             : i1504qa1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 재고이동현황 조회(품목간)
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/07/01
'*  8. Modified date(Last)    : 2003/07/01
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
<%
EndDate   = GetSvrDate
%>

Const BIZ_PGM_QRY_ID	= "i1504qb1.asp"

Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_Unit
Dim C_MoveType
Dim C_DocumentDt
Dim C_Qty
Dim C_Price
Dim C_Amount
Dim C_SlCd
Dim C_TrnsItemCd
Dim C_TrnsItemNm
Dim C_TrnsSlCd
Dim C_ItemDocumentNo
Dim C_SeqNo

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop 
Dim lgOldRow


'==========================================  2.1.1 InitVariables()  =====================================
Sub InitVariables()
	lgIntFlgMode		= parent.OPMD_CMODE
	lgIntGrpCount		= 0	
	lgBlnFlgChgValue	= False
	lgLngCurRows		= 0
    lgOldRow			= 0
    lgStrPrevKeyIndex	= ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	Dim StartDate

	StartDate = UNIDateAdd("m", -1,"<%=EndDate%>", parent.gServerDateFormat)
	
	frm1.txtMovFrDt.Text = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtMovToDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
 	Call ggoOper.FormatDate(frm1.txtMovFrDt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtMovToDt, parent.gDateFormat, 1)
	
	 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
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
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030701", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
			
		.MaxCols = C_SeqNo+1					
		.MaxRows = 0
				
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "3", "0")
				
		ggoSpread.SSSetEdit     C_ItemCd,			"품목",			18
		ggoSpread.SSSetEdit		C_ItemNm,			"품목명",		20
		ggoSpread.SSSetEdit		C_Spec,				"규격",			20
		ggoSpread.SSSetEdit		C_TrackingNo,		"Tracking No.",	20
		ggoSpread.SSSetEdit 	C_LotNo,			"Lot No.",		12
		ggoSpread.SSSetFloat	C_LotSubNo,			"Lot순번",		10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit		C_Unit,				"단위",			10
		ggoSpread.SSSetEdit		C_MoveType,			"수불유형",		20
		ggoSpread.SSSetDate		C_DocumentDt,		"이동일자",		12, 2,Parent.gDateFormat
		ggoSpread.SSSetFloat    C_Qty,				"이동수량",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
		ggoSpread.SSSetFloat    C_Price,			"이동단가",		15, Parent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat    C_Amount,			"이동금액",		15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetEdit     C_SlCd,				"출고창고",		10
		ggoSpread.SSSetEdit     C_TrnsItemCd,		"이동품목",		10
		ggoSpread.SSSetEdit		C_TrnsItemNm,		"이동품목명",	20
		ggoSpread.SSSetEdit     C_TrnsSlCd,			"입고창고",		10
		ggoSpread.SSSetEdit     C_ItemDocumentNo,	"수불번호",		18
		ggoSpread.SSSetEdit     C_SeqNo,			"수불순번",		10
		 		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetSplit2(2)
		.ReDraw = true
	End With
		
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
		C_ItemCd			= 1
		C_ItemNm			= 2
		C_Spec				= 3
		C_TrackingNo		= 4
		C_LotNo				= 5
		C_LotSubNo			= 6
		C_Unit				= 7
		C_MoveType			= 8
		C_DocumentDt		= 9
		C_Qty				= 10
		C_Price				= 11
		C_Amount			= 12
		C_SlCd				= 13
		C_TrnsItemCd		= 14
		C_TrnsItemNm		= 15
		C_TrnsSlCd			= 16
		C_ItemDocumentNo	= 17
		C_SeqNo				= 18
End Sub

 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 		
		C_ItemCd			= iCurColumnPos(1)
		C_ItemNm			= iCurColumnPos(2)
		C_Spec				= iCurColumnPos(3)
		C_TrackingNo		= iCurColumnPos(4)
		C_LotNo				= iCurColumnPos(5)
		C_LotSubNo			= iCurColumnPos(6)
		C_Unit				= iCurColumnPos(7)
		C_MoveType			= iCurColumnPos(8)
		C_DocumentDt		= iCurColumnPos(9)
		C_Qty				= iCurColumnPos(10)
		C_Price				= iCurColumnPos(11)
		C_Amount			= iCurColumnPos(12)
		C_SlCd				= iCurColumnPos(13)
		C_TrnsItemCd		= iCurColumnPos(14)
		C_TrnsItemNm		= iCurColumnPos(15)
		C_TrnsSlCd			= iCurColumnPos(16)
		C_ItemDocumentNo	= iCurColumnPos(17)
		C_SeqNo				= iCurColumnPos(18)
	
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업" 
	arrParam(1) = "B_PLANT"    
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "공장"   
	 
	arrField(0) = "PLANT_CD" 
	arrField(1) = "PLANT_NM" 
	 
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)  
		frm1.txtPlantNm.Value    = arrRet(1)  
		frm1.txtPlantCd.focus
	End If  
End Function


 '------------------------------------------  OpenMovType()  -------------------------------------------------
Function OpenMovType()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If IsOpenPop = True Or UCase(frm1.txtMovType.ClassName)="PROTECTED" Then Exit Function 

	IsOpenPop = True

	arrParam(0) = "수불유형 팝업"     
	arrParam(1) = "B_MINOR A,I_MOVETYPE_CONFIGURATION B"      
	arrParam(2) = Trim(frm1.txtMovType.Value)
	arrParam(3) = ""
	arrParam(4) = "A.minor_cd = B.mov_type and A.major_cd = " & FilterVar("I0001", "''", "S") & " and B.trns_type = " & FilterVar("ST", "''", "S") & " and B.gui_control_flag3 = " & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "수불유형"
	 
	arrField(0) = "MINOR_CD" 
	arrField(1) = "MINOR_NM" 
	 
	arrHeader(0) = "수불유형"  
	arrHeader(1) = "수불유형명"  
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtMovType.focus
		Exit Function
	Else
		frm1.txtMovType.Value   = arrRet(0)
		frm1.txtMovTypeNm.Value = arrRet(1)
		frm1.txtMovType.focus
	End If
	Set gActiveElement = document.activeElement 
End Function


'------------------------------------------  OpenItemCd()  --------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Cd
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X") 
		frm1.txtPlantCd.focus
		Exit Function
	End If


	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목 팝업"
	arrParam(1) = "B_ITEM_BY_PLANT P,B_ITEM I"	
	arrParam(2) = Trim(frm1.txtItemCd.Value)	
	arrParam(3) = ""							
	arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "품목"			
	
	arrField(0) = "I.ITEM_CD"
	arrField(1) = "I.ITEM_NM"
	
	arrHeader(0) = "품목코드"
	arrHeader(1) = "품목명"	
	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value	=arrRet(0)
		frm1.txtItemNm.Value	=arrRet(1)
		frm1.txtItemCd.focus
	End If	
	
End Function


'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec) 
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("11000000000011")
End Sub


'=======================================================================================================
'   Event Name : txtMovFrDt_DblClick(Button)
'=======================================================================================================
Sub txtMovFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMovFrDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtMovFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtMovFrDt_KeyPress()
'=======================================================================================================
Sub txtMovFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtMovToDt_DblClick(Button)
'=======================================================================================================
Sub txtMovToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtMovToDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtMovToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtMovToDt_KeyPress()
'=======================================================================================================
Sub txtMovToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_LeaveCell
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then Exit Sub
		
		If NewRow = .MaxRows Then
			If lgStrPrevKeyIndex <> "" Then
				If DbQuery = False Then	Exit Sub
			End if
		End if
	
	End With
End Sub 


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub

		if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKeyIndex <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
				If DbQuery = False Then
					Call RestoreToolBar()
					Exit Sub
				End if
		End If
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
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row <= 0 Then Exit Sub
  	If frm1.vspdData.MaxRows = 0 Then Exit Sub
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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                    
    Err.Clear
    
    If Not chkField(Document, "1") Then Exit Function                                                           

	If Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("189220","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	Else
		If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		   
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.value = ""
			frm1.txtPlantCd.focus
			Exit function
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtPlantNm.value = lgF0(0)
	End If
	
	frm1.txtMovTypeNm.value = ""  
	If frm1.txtMovType.value <> "" Then
		If  CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND MINOR_CD = " & FilterVar(frm1.txtMovType.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
		 
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtMovTypeNm.value = lgF0(0)
		End If  
	End If
	

	frm1.txtItemNm.value = ""
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtItemNm.value = lgF0(0)
		End if
	End If
 
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
             
    Call InitVariables             
    '-----------------------
    'Query function call area
    '-----------------------
    Call SetToolbar("11000000000111")         
	If DbQuery = False Then Exit Function
    Set gActiveElement = document.activeElement   
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
	
    Dim strVal
    
    
	On error resume next
    Err.Clear        
    Call LayerShowHide(1)
    
    DbQuery = False
    With frm1    
		strVal = BIZ_PGM_QRY_ID &	"?txtMode="				& Parent.UID_M0001				& _	
									"&txtPlantCd="			& Trim(.txtPlantCd.value)		& _
									"&txtMovFrDt="			& Trim(frm1.txtMovFrDt.Text)	& _
									"&txtMovToDt="			& Trim(frm1.txtMovToDt.Text)	& _
									"&txtMovType="			& Trim(.txtMovType.value)		& _
									"&txtItemCd="			& Trim(.txtItemCd.value)		& _
									"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex				& _
									"&txtMaxRows="			& .vspdData.MaxRows		
		Call RunMyBizASP(MyBizASP, strVal)
    End With
    DbQuery = True
                             

End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()		
	frm1.vspdData.focus 
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고이동현황조회(품목간)</font></TD>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="14">	
								</TD>
								<TD CLASS="TD5" NOWRAP>이동일자</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/i1504qa1_OBJECT2_txtMovFrDt.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/i1504qa1_OBJECT3_txtMovToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCd" SIZE="18" MAXLENGTH="18" ALT="품목" tag="11XXXU"><IMG align=top height=20 name="btnItemNm" onclick="vbscript:OpenItemCd()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 MAXLENGTH=40 tag="14">
								</TD>
								<TD CLASS="TD5" NOWRAP>수불유형</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT Name="txtMovType" SIZE="5" MAXLENGTH="3"  ALT="수불유형" tag="11XXXU"><IMG align=top height=20 name=btnMovType onclick="vbscript:OpenMovType()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtMovTypeNm" size="30" tag="14">
								</TD>
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
								<script language =javascript src='./js/i1504qa1_OBJECT1_vspdData.js'></script>
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
	
	
	
	
	
