<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name          : Basis Architect															*
'*  2. Function Name        : Physical Inventory header Popup																*
'*  3. Program ID           : i2111pa1																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 실사번호팝업																	*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2003/06/02																*
'*  9. Modifier (First)     : Kim Nam Hoon																*
'* 10. Modifier (Last)      : Lee Seung Wook																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :																			*
'*                            2000/02/29 : Coding Start													*
'********************************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_ID = "i2111pb1.asp"							

Dim C_PhyInvNo
Dim C_DocumentDt
Dim C_SLCd
Dim C_SLNm
Dim C_PosIndctr
Dim C_PlantCd
Dim C_PlantNm

Dim arrParam
Dim arrReturn			

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgQueryFlag				
Dim lgDocumentNo
Dim lgDocSts
Dim lgSLCd
Dim lgFromDt
Dim lgToDt
Dim hlgPlantCd				
Dim hlgDocumentNo			
Dim hlgDocSts
Dim hlgSLCd
Dim hlgFromDt
Dim hlgToDt

Dim lsOpenPop

arrParam = window.dialogArguments
Set PopupParent = arrParam(0)

top.document.title = PopupParent.gActivePRAspName

Dim EndDate
Dim StartDate

   EndDate   = "<%=GetSvrDate%>"                                                 
   StartDate = UNIDateAdd("M", -1, EndDate, PopupParent.gServerDateFormat)       

'=======================================================================================================
' Function Name : LoadInfTB19029
'=======================================================================================================
Function LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	txtPhyInvNo.value = arrParam(1)
	lgDocSts          = arrParam(2)
	hlgPlantCd        = arrParam(3)
	txtSLCd.value     = arrParam(4)
	Self.Returnvalue  = Array("")
	txtToDt.Text      = UniConvDateAToB(EndDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat) 
	txtFromDt.Text    = UniConvDateAToB(StartDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat)
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()

 	Call InitSpreadPosVariables()
 		
	With vspdData
		
	ggoSpread.Source = vspdData
	
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread
	
	.ReDraw = False
    
	 .MaxCols = C_PlantNm +1
	 .MaxRows = 0

	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit C_PhyInvNo, 	"실사번호",		15
	ggoSpread.SSSetEdit C_DocumentDt,   "실사일자",		10,2	
	ggoSpread.SSSetEdit C_SLCd, 		"창고",			8,2
	ggoSpread.SSSetEdit C_SLNm, 		"창고명",		30
	ggoSpread.SSSetEdit C_PosIndctr, 	"Posting 구분", 12,2,,1
	ggoSpread.SSSetEdit C_PlantCd, 		"공장",			8,2
	ggoSpread.SSSetEdit C_PlantNm, 		"공장명",		30
	
	If lgDocSts = "PD" Then
  		Call ggoSpread.SSSetColHidden(C_PlantCd, C_PlantNm, True)	
	Else
 		Call ggoSpread.SSSetColHidden(C_PosIndctr, C_PlantNm, True)
 	End If
 	
 	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)	
	
	ggoSpread.SpreadLock -1, -1

	.ReDraw = True
   
   end With
	   
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_PhyInvNo		= 1
	C_DocumentDt	= 2
	C_SLCd			= 3
	C_SLNm			= 4
	C_PosIndctr		= 5
	C_PlantCd		= 6
	C_PlantNm		= 7
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_PhyInvNo		= iCurColumnPos(1)
		C_DocumentDt	= iCurColumnPos(2)	
		C_SLCd			= iCurColumnPos(3)
		C_SLNm			= iCurColumnPos(4)
		C_PosIndctr		= iCurColumnPos(5)
		C_PlantCd		= iCurColumnPos(6)
		C_PlantNm		= iCurColumnPos(7)

 	End Select
End Sub

'------------------------------------------  OpenSL()  -------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If hlgPlantCd = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		txtSLCd.focus  
		Exit Function
	End if

	If lsOpenPop = True Then Exit Function

	lsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(txtSLCd.Value)
	arrParam(3) = ""	
	arrParam(4) = "PLANT_CD = " & FilterVar(hlgPlantCd, "''", "S")		
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lsOpenPop = False
	
	If arrRet(0) = "" Then
		txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
	
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
Function SetSL(byRef arrRet)
	txtSLCd.Value    = arrRet(0)		
	txtSLNm.Value    = arrRet(1)
	txtSLCd.focus		
End Function

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt
	
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
	
		vspdData.Row = vspdData.ActiveRow
				
		vspdData.Col = C_PhyInvNo
		arrReturn(0) = vspdData.Text
		vspdData.Col = C_DocumentDt
		arrReturn(1) = vspdData.Text
		vspdData.Col = C_SLCd
		arrReturn(2) = vspdData.Text
		vspdData.Col = C_SLNm
		arrReturn(3) = vspdData.Text
		vspdData.Col = C_PosIndctr
		arrReturn(4) = vspdData.Text
		vspdData.Col = C_PlantCd
		arrReturn(5) = vspdData.Text
		vspdData.Col = C_PlantNm
		arrReturn(6) = vspdData.Text
				
		Self.Returnvalue = arrReturn
	End If
	
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")                             
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)                         	
	Call SetDefaultVal()
	Call InitSpreadSheet()
	
	Call FncQuery()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        txtFromDt.Action = 7
        Call SetFocusToDocument("M")        
        txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_Keypress(KeyAscii)
'=======================================================================================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_Change()
'=======================================================================================================
Sub txtFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        txtToDt.Action = 7
        Call SetFocusToDocument("M")        
        txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_Keypress(KeyAscii)
'=======================================================================================================
Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=======================================================================================================
'   Event Name : txtToDt_Change()
'=======================================================================================================
Sub txtToDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       
 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = vspdData
    
 	If vspdData.MaxRows = 0 Then Exit Sub

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

'========================================================================================
' Function Name : vspdData_KeyPress
'========================================================================================
Function vspdData_KeyPress(KeyAscii)
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================================================================
' Function Name : vspdData_TopLeftChange
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If OldLeft <> NewLeft Then Exit Sub
	
	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) And lgStrPrevKey <> "" Then
		DbQuery
	End if
End Sub
 
'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then Exit Sub
  	If vspdData.MaxRows = 0 Then Exit Sub
 	
	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
		End If
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
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
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
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
Function FncQuery()

    If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function

	ggoSpread.Source = vspdData
    ggoSpread.ClearSpreadData

	lgDocumentNo = Trim(txtPhyInvNo.Value)	
	lgSLCd = Trim(txtSLCd.Value)
	lgFromDt = txtFromDt.Text
	lgToDt = txtToDt.Text
	lgStrPrevKey = ""
	
	Call DbQuery()
End Function


'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	Dim strVal
	Dim txtMaxRows
    Call LayerShowHide(1)  
	
	DbQuery = False                                                        
	txtMaxRows = vspdData.MaxRows       
	
   if lgStrPrevKey <> "" Then
       
	strVal = BIZ_PGM_ID &	"?txtPhyInvNo="		& hlgDocumentNo	& _ 
							"&txtSLCd="         & hlgSLCd		& _	
							"&txtPlantCd="      & hlgPlantCd	& _
							"&txtFromDt="       & hlgFromDt		& _
							"&txtToDt="         & hlgToDt		& _	
							"&lgDocSts="        & hlgDocSts		& _
							"&lgStrPrevKey="    & lgStrPrevKey	& _
							"&txtMaxRows="      & txtMaxRows
   else 
   	strVal = BIZ_PGM_ID &	"?txtPhyInvNo="		& lgDocumentNo	& _
							"&txtSLCd="         & lgSLCd		& _
							"&txtPlantCd="      & hlgPlantCd	& _
							"&txtFromDt="       & lgFromDt		& _
							"&txtToDt="         & lgToDt		& _	
							"&lgDocSts="        & lgDocSts		& _
							"&lgStrPrevKey="    & lgStrPrevKey	& _
							"&txtMaxRows="      & txtMaxRows
   end if	

	Call RunMyBizASP(MyBizASP, strVal)			
	
	DbQuery = True                                                       
End Function

'********************************************  5.1 DbQueryOk()  *******************************************
Function DbQueryOk()							
  vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
	<TABLE <%=LR_SPACE_TYPE_00%>>
		<TR HEIGHT=*>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
										<TD CLASS="TD5">실사번호</TD>
										<TD CLASS="TD6"><INPUT TYPE="Text" Name="txtPhyInvNo" SIZE=18 MAXLENGTH=16  tag="11XXXU" ALT="실사번호" ></TD>
										<TD CLASS="TD5"></TD>
										<TD CLASS="TD6"></TD>
									</TR>
									<TR>
										<TD CLASS="TD5">실사일자</TD>
										<TD CLASS="TD6">
										<script language =javascript src='./js/i2111pa1_I931037026_txtFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/i2111pa1_I359441494_txtToDt.js'></script>
										</TD>                
										<TD CLASS="TD5">창고</TD>
										<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="11XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 tag="14"></TD>
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
					</TR>
					<TR>
						<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD WIDTH=100% HEIGHT=100%>
									<script language =javascript src='./js/i2111pa1_I293509095_vspdData.js'></script>
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
									<TD WIDTH=70% NOWRAP>		<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
									<TD WIDTH=30% ALIGN=RIGHT>	<IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
																<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
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
				<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
			</DIV>
		</BODY>
	</HTML>

