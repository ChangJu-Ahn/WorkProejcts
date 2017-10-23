<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s1912ra1.asp	
'*  4. Program Name         : 가용재고현황 
'*  5. Program Desc         : 가용재고현황 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit		

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

<!-- #Include file="../../inc/lgvariables.inc" --> 
Const BIZ_PGM_QRY_ID = "s1912rb1.asp"		

Dim C_Date				
Dim C_AvalReceiptQty	
Dim C_PlanReceiptQty	
Dim C_PlanGIQty		
Dim C_StockQty		

Const gstrPayTermsMajor = "B9004"
Const gstrIncoTermsMajor = "B9006"
Dim strReturn					
Dim arrParam	
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim gblnWinEvent			

'================================================================================================================
Function InitVariables()
	lgIntFlgMode = PopupParent.OPMD_CMODE					
	lgIntGrpCount = 0									
	lgStrPrevKey1 = ""									
	lgStrPrevKey2 = ""									
	'------ Coding part ------ 
	gblnWinEvent = False
		
	Redim strReturn(0)
	strReturn(0) = ""
	Self.Returnvalue = strReturn
End Function
	
'================================================================================================================
Sub initSpreadPosVariables()
	C_Date				= 1
	C_AvalReceiptQty	= 2
	C_PlanReceiptQty	= 3
	C_PlanGIQty			= 4
	C_StockQty			= 5
End Sub

'================================================================================================================
Sub SetDefaultVal()
	arrParam = arrParent(1)
	txtItem.Value = arrParam(0)
	txtItemNm.value = arrParam(1)
	txtPlant.value = arrParam(2)
	txtPlantNm.value = arrParam(3)
End Sub

'================================================================================================================
Private Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'================================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()		
	ggoSpread.Source = vspdData
		
	ggoSpread.Spreadinit "V20021214",,PopupParent.gAllowDragDropSpread
	vspdData.ReDraw = False
	vspdData.MaxCols = C_StockQty + 1
	vspdData.MaxRows = 0
		
	Call GetSpreadColumnPos("A")	 
		
	ggoSpread.SSSetDate		C_Date, "일자", 12, 2, PopupParent.gDateFormat
    ggoSpread.SSSetFloat	C_StockQty,"보유재고량" ,19,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetFloat	C_PlanReceiptQty,"입고예정량" ,19,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetFloat	C_PlanGIQty,"출고예정량" ,19,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetFloat	C_AvalReceiptQty,"가용재고량" , 19,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		
	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
	SetSpreadLock "", 0, -1, ""
	vspdData.ReDraw = True
	
End Sub

'================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Date				= iCurColumnPos(1)   
			C_AvalReceiptQty	= iCurColumnPos(2)   
			C_PlanReceiptQty	= iCurColumnPos(3)   
			C_PlanGIQty			= iCurColumnPos(4)
			C_StockQty			= iCurColumnPos(5)   	
    End Select    
End Sub


'================================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
	ggoSpread.Source = vspdData
			
	vspdData.ReDraw = False
			
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.ReDraw = True

End Sub

'================================================================================================================
Function CancelClick()
	Self.Close()
End Function


'================================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call SetDefaultVal
	Call InitVariables
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'================================================================================================================	
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000011111")
    gMouseClickStatus = "SPC"    
    Set gActiveSpdSheet = vspdData
    
    If vspdData.MaxRows = 0 Then             'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
       ggoSpread.Source = vspdData
       If lgSortKey = 1 Then
		  ggoSpread.SSSort Col				'Sort in Ascending
		  lgSortkey = 2
       Else
		  ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
		  lgSortkey = 1
	   End If
       Exit Sub
    End If	
End Sub

'================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	Call SetQuerySpreadColor(1)    

End Sub

'================================================================================================================
Function vspdData_KeyPress(KeyAscii)
   On Error Resume Next
   If KeyAscii = 27 Then
	  Call CancelClick()
   End If
End Function

'================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
		If CheckRunningBizProcess = True Then Exit Sub	
			
    	If lgStrPrevKey1 <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub

'================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub


'================================================================================================================
Function Document_onKeyDown()
	If window.event.keycode = 27 Then Self.Close
End Function
	
'================================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False									

	Err.Clear											

	Call ggoOper.ClearField(Document, "2")				
	Call InitVariables									
	Call DbQuery()										

	FncQuery = True										
End Function

'================================================================================================================
Function DbQuery()
	Err.Clear			

	DbQuery = False		

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001	
		strVal = strVal & "&txtItem=" & Trim(txtHItem.value)			
		strVal = strVal & "&txtPlant=" & Trim(txtHPlant.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001	
		strVal = strVal & "&txtItem=" & Trim(txtItem.value)				
		strVal = strVal & "&txtPlant=" & Trim(txtPlant.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	End If
				
	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True													
End Function
	
'================================================================================================================
	Function DbQueryOk()						
		lgIntFlgMode = PopupParent.OPMD_UMODE	
		vspdData.Focus
	End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>품목</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtItem" SIZE=18 MAXLENGTH=18 TAG="14XXXU" ALT="품목"></TD>
							<TD CLASS=TD5 NOWRAP>품목명</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtITemNm" SIZE=25 MAXLENGTH=50 TAG="14" ALT="품목명"></TD>
						</TR>	
						<TR>	
							<TD CLASS=TD5>공장</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="14XXXU" ALT="공장">&nbsp;
								<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="14">
							</TD>
							<TD CLASS=TD5>재고단위</TD>
							<TD CLASS=TD6>
								<INPUT TYPE=TEXT NAME="txtStockType" SIZE=10 MAXLENGTH=4 TAG="14XXXU" ALT="재고단위"></TD>
						</TR>
					</TABLE>
				</FIELDSET>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s1912ra1_vaSpread_vspdData.js'></script>
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
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHItem" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" Tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant" TAG="24">
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>