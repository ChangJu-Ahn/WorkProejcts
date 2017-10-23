<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : 																			*
'*  4. Program Name         : 																			*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2004/08/11																*
'*  8. Modified date(Last)  : 2004/08/11																*
'*  9. Modifier (First)     :            																*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

'================================================================================================================================
Const BIZ_PGM_ID 		= "U1111PB1.asp"

'================================================================================================================================
Dim C_PoNo
Dim C_PoSeqNo
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_POQty
Dim C_Unit
Dim C_PODlvyDt
Dim C_PlantCd
Dim C_PlantNm

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================

'================================================================================================================================
Dim IsOpenPop
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim EndDate, StartDate
'================================================================================================================================    
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
arrParam= arrParent(1)

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'================================================================================================================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
	lgKeyStream		 = ""
	
	frm1.vspdData.MaxRows = 0	
	
	IsOpenPop = False
	ReDim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================
Sub SetDefaultVal()
	
	Dim iCodeArr
		
	Err.Clear
	
	With frm1
		.txtFrPoDt.text = StartDate
		.txtToPoDt.text = EndDate
	
		.txtBpCd.value 	= arrParam(1)
		.txtBpNm.value 	= arrParam(2)
		
		.hdnPlantCd.value = arrParam(0)
	End With
	
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021224", ,PopupParent.gAllowDragDropSpread
		frm1.vspdData.OperationMode = 3
		
		.ReDraw = false
					
		.MaxCols = C_PlantNm + 1    
		.MaxRows = 0    
			
		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit 		C_PoNo,			"발주번호", 15
		ggoSpread.SSSetFloat 		C_PoSeqNo,		"순번",6,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_ItemCd,		"품목",15
		ggoSpread.SSSetEdit 		C_ItemNm,		"품목명",20
		ggoSpread.SSSetEdit 		C_Spec,			"규격",20
		ggoSpread.SSSetFloat 		C_POQty,		"발주수량",15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 		C_Unit,			"단위", 6
		ggoSpread.SSSetDate 		C_PODlvyDt,		"납기일"	,10, 2, PopupParent.gDateFormat		 
		ggoSpread.SSSetEdit 		C_PlantCd,		"공장", 6
		ggoSpread.SSSetEdit 		C_PlantNm,		"공장명", 10
			
		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
		ggoSpread.SSSetSplit2(2)
						
		Call SetSpreadLock()
						
		.ReDraw = true    
    
	End With
	   
End Sub
'================================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()
	C_PoNo			= 1
	C_PoSeqNo		= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_POQty			= 6
	C_Unit			= 7
	C_PODlvyDt		= 8
	C_PlantCd		= 9
	C_PlantNm		= 10
End Sub
'================================================================================================================================
Sub GetSpreadColumnPos()
      
    Dim iCurColumnPos
    
 	ggoSpread.Source = frm1.vspdData
		
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_PoNo			= iCurColumnPos(1)
	C_PoSeqNo		= iCurColumnPos(2)
	C_ItemCd		= iCurColumnPos(3)
	C_ItemNm		= iCurColumnPos(4)
	C_Spec			= iCurColumnPos(5)
	C_POQty			= iCurColumnPos(6)
	C_Unit			= iCurColumnPos(7)
	C_PODlvyDt		= iCurColumnPos(8)
	C_PlantCd		= iCurColumnPos(9)
	C_PlantNm		= iCurColumnPos(10)
End Sub    
'================================================================================================================================
Function OKClick()
	Dim intRowCnt
	Dim intColCnt
	Dim intSelCnt

	If frm1.vspdData.MaxRows > 0 Then
		intSelCnt = 0
		Redim arrReturn(0)
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow

		If frm1.vspdData.SelModeSelected = True Then
			frm1.vspdData.Col = C_PoNo
			arrReturn(0) = frm1.vspdData.Text
		End If

		Self.Returnvalue = arrReturn
	End If		
	Self.Close()
End Function	
'================================================================================================================================
Function CancelClick()
	Redim arrReturn(1)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'================================================================================================================================
Function OpenSortPopup()

	On Error Resume Next
	
End Function

'================================================================================================================================
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"										' 팝업 명칭 
	arrParam(1) = "B_Biz_Partner"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBpCd.Value)						' Code Condition
	arrParam(3) = ""
	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"	' Where Condition
	arrParam(5) = "공급처"										' TextBox 명칭 
	
    arrField(0) = "BP_CD"										' Field명(0)
    arrField(1) = "BP_NM"										' Field명(1)
    
    arrHeader(0) = "공급처"										' Header명(0)
    arrHeader(1) = "공급처명"									' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

'================================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029															'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	                                           
	Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field 
	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
		
	Call InitSpreadSheet()
		
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
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
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Sub
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================================================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrPoDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToPoDt.Focus
	End if
End Sub
'================================================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrPoDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToPoDt.text,PopupParent.gDateFormat,"")) And trim(.txtFrPoDt.text) <> "" And trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			.txtToPoDt.Focus()
			Exit Function
		End if   
	End with
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
        
	Call InitVariables
	
	If CheckRunningBizProcess = True Then Exit Function
    If DbQuery = False Then Exit Function

    FncQuery = True									
        
End Function
'================================================================================================================================
Function DbQuery()
	
	Dim strVal
	
	Err.Clear															'☜: Protect system from crashing

	DbQuery = False														'⊙: Processing is NG

    If LayerShowHide(1) = False Then Exit Function
    
    Call MakeKeyStream()
    
	strVal = BIZ_PGM_ID & "?txtMode="	& PopupParent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgPageNo
	
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

	DbQuery = True														'⊙: Processing is NG
End Function
'================================================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtBpCd.focus
	End If

End Function
'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()

	With frm1
		lgKeyStream = lgKeyStream & UCase(Trim(.hdnPlantCd.value))  & PopupParent.gColSep
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			lgKeyStream = lgKeyStream & Trim(.hdnFrPoDt.value)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & Trim(.hdnToPoDt.value)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.hdnBpCd.value))  & PopupParent.gColSep
		Else
			lgKeyStream = lgKeyStream & Trim(.txtFrPoDt.Text)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & Trim(.txtToPoDt.Text)  & PopupParent.gColSep
			lgKeyStream = lgKeyStream & UCase(Trim(.txtBpCd.value))  & PopupParent.gColSep

			.hdnFrPoDt.value		= .txtFrPoDt.Text
			.hdnToPoDt.value		= .txtToPoDt.Text
		End If
		
	End With
			 
End Sub    

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>발주일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/u1111pa1_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/u1111pa1_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
			 			<TD CLASS=TD5 NOWRAP>공급처</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="공급처명"></TD>
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
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/u1111pa1_vspdData_vspdData.js'></script>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnFrPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>