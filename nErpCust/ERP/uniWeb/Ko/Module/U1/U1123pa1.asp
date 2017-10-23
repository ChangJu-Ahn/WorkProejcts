<%@ LANGUAGE="VBSCRIPT" %>
<!--
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 구매입출고 관리 
'*  3. Program ID           : u1123pb1 
'*  4. Program Name         : Receipt No Popup ASP
'*  5. Program Desc         : 구매반품의 입출고번호 팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2003/05/28
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'*							  ADO Conv. 	
'**************************************************************************************
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
Option Explicit                                                             '☜: indicates that All variables must be declared in advance
'================================================================================================================================
Const BIZ_PGM_ID 		= "u1123pb1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 8                                           '☆: key count of SpreadSheet
'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================
Dim gblnWinEvent
Dim arrReturn
Dim arrParent
Dim arrParam
Dim chkFlg
Dim EndDate, StartDate

'================================================================================================================================
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)		
'================================================================================================================================
Function InitVariables()
	Redim arrReturn(0) 
	
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    
    lgSortKey        = 1   
    lgIntGrpCount	 = 0										<%'⊙: Initializes Group View Size%>
	gblnWinEvent	 = False
    
    Self.Returnvalue = arrReturn    
End Function
'================================================================================================================================
Sub SetDefaultVal()
	Err.Clear
	
	frm1.txtFrRcptDt.Text = StartDate
	frm1.txtToRcptDt.Text = EndDate
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("M4141PA1","S","A","V20030529",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 3 
End Sub
'================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End IF
End Sub
'================================================================================================================================
Function OKClick()
	If frm1.vspdData.ActiveRow > 0 Then	
		Redim arrReturn(0)
		arrReturn(0) = GetSpreadText(frm1.vspdData,GetKeyPos("A",1),frm1.vspdData.ActiveRow,"X","X")
	End If

	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'================================================================================================================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'================================================================================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
'================================================================================================================================
Function OpenConSItemDC(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	Select Case iWhere
		Case 0	'입출고유형 
		
			arrParam(1) = "M_Mvmt_type"
			arrParam(2) = Trim(frm1.txtMoveType.Value)
			arrParam(4) = "RET_FLG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLG=" & FilterVar("Y", "''", "S") & "  "
			If Trim(frm1.hdnRcptFlg.value) <> "" then 
				arrParam(4) = arrParam(4) & " And RCPT_FLG= " & FilterVar(frm1.hdnRcptFlg.value, "''", "S") & " "
			end if
			arrParam(5) = "입출고유형"			
	
			arrField(0) = "IO_Type_Cd"	
			arrField(1) = "IO_Type_NM"	
    
			arrHeader(0) = "입출고유형"		
			arrHeader(1) = "입출고유형명"
    
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		Case 1	'공급처 
			
			arrParam(1) = "B_BIZ_PARTNER"							' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtSupplierCd.Value)			' Code Condition
			arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & " "' Where Condition
			arrParam(5) = "공급처"								' TextBox 명칭 
	
			arrField(0) = "BP_Cd"									' Field명(0)
			arrField(1) = "BP_NM"									' Field명(1)
    
			arrHeader(0) = "공급처"								' Header명(0)
			arrHeader(1) = "공급처명"							' Header명(1)
    
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
		Case 2
	
			arrParam(1) = "B_Pur_Grp"				
			arrParam(2) = Trim(frm1.txtGroupCd.Value)
			arrParam(4) = ""			
			arrParam(5) = "구매그룹"			
			
		    arrField(0) = "PUR_GRP"	
		    arrField(1) = "PUR_GRP_NM"	
		    
		    arrHeader(0) = "구매그룹"		
		    arrHeader(1) = "구매그룹명"		
		    
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
	
	arrParam(0) = arrParam(5)												' 팝업 명칭	

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0:	frm1.txtMoveType.focus	
			Case 1: frm1.txtSupplierCd.focus	
			Case 2: frm1.txtGroupCd.focus	
		End Select
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function
'================================================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtMoveType.Value		= arrRet(0)		
				.txtMoveTypeNm.Value	= arrRet(1)
				.txtMoveType.focus	
			Case 1
				.txtSupplierCd.Value	= arrRet(0)		
				.txtSupplierNm.Value	= arrRet(1)	
				.txtSupplierCd.focus	
			Case 2
				.txtGroupCd.Value		= arrRet(0)		
				.txtGroupNm.Value		= arrRet(1)	
				.txtGroupCd.focus	
		End Select
	End With
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery()
End Sub
'================================================================================================================================
Sub txtToRcptDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToRcptDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtToRcptDt.focus
	End if
End Sub
'================================================================================================================================
Sub txtFrRcptDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrRcptDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtFrRcptDt.focus
	End if
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
    	
	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Sub	
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
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
Sub txtFrRcptDt_Keypress(KeyAscii)
	On Error Resume Next
	
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToRcptDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Function FncQuery() 
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrRcptDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToRcptDt.text,PopupParent.gDateFormat,"")) and Trim(.txtFrRcptDt.text)<>"" and Trim(.txtToRcptDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","입고일", "X")			
			.txtToRcptDt.Focus
			Exit Function
		End if   
    End with
    
	Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
        
End Function
'================================================================================================================================
Function DbQuery() 
	Dim strVal

	Err.Clear                                                               '☜: Protect system from crashing
    DbQuery = False
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	strVal = ""
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		    strVal = strVal & "&txtMvmtType=" & .hdnMvmtType.value
		    strVal = strVal & "&txtSupplier=" & .hdnSupplier.value
			strVal = strVal & "&txtFrRcptDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToRcptDt=" & .hdnToDt.value
		    strVal = strVal & "&txtGroup=" & .hdnGroup.value
		    strVal = strVal & "&txtRcptFlg=" & Trim(.hdnRcptFlg.value)
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		    strVal = strVal & "&txtMvmtType=" & Trim(.txtMoveType.value)
		    strVal = strVal & "&txtSupplier=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtFrRcptDt=" & Trim(.txtFrRcptDt.text)
			strVal = strVal & "&txtToRcptDt=" & Trim(.txtToRcptDt.text)
		    strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.Value)
		    strVal = strVal & "&txtRcptFlg=" & Trim(.hdnRcptFlg.value)	    
		end if 
		
			strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       
        Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True
    
End Function
'================================================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>
	
    lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtMoveType.focus
	End If
	
End Function

'================================================================================================================================
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
						<TD CLASS="TD5" NOWRAP>입출고유형</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입출고유형" NAME="txtMoveType" SIZE=10 MAXLENGTH=5 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 0">
											   <INPUT TYPE=TEXT Alt="입출고유형" NAME="txtMoveTypeNm" SIZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>입출고일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr NOWRAP>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입출고일 NAME="txtFrRcptDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=입출고일 NAME="txtToRcptDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="공급처" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 1">
					   			 	     	   <INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>구매그룹</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 2">
										 	   <INPUT TYPE=TEXT AlT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% ID=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnRcptFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
