<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        : Inspection Result Registraton No Popup									*										*
'*  3. Program ID           : m4131pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 																			*
'*  8. Modified date(Last)  : 2001/05/09																*
'*  9. Modifier (First)     : EverForever																			*
'* 10. Modifier (Last)      : EverForever																	*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'* 13. History              :																			*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit				
	
'================================================================================================================================
Const BIZ_PGM_ID 	= "m4131pb1_KO441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey      = 8                                           '☆: key count of SpreadSheet
'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================

Dim gblnWinEvent
Dim arrReturn					
Dim arrParent
Dim arrParam
Dim IsOpenPop
'================================================================================================================================ 
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

Dim EndDate, StartDate
EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'================================================================================================================================ 
Function InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                            'Indicates that current mode is Create mode
                
    gblnWinEvent = False
    Redim arrReturn(0)        
    Self.Returnvalue = arrReturn     
End Function
'================================================================================================================================ 
Sub SetDefaultVal()
	Err.Clear
	
	frm1.txtFrRegDt.Text = StartDate
	frm1.txtToRegDt.Text = EndDate
	frm1.txtRcptNo.value 	 = arrParam(0)
	frm1.txtGroupCd.value 	 = arrParam(1)

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
		frm1.txtGroupCd.Tag = left(frm1.txtGroupCd.Tag,1) & "4" & mid(frm1.txtGroupCd.Tag,3,len(frm1.txtGroupCd.Tag))
        frm1.txtGroupCd.value = lgPGCd
	End If

End Sub
'================================================================================================================================ 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Sub
'================================================================================================================================ 
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("M4131PA1","S","A","V2003528",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
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
Function openRcptNo()
	
	Dim strRet
	Dim arrParam(3)
	Dim iCalledAspName
	Dim IntRetCD
		
	If IsOpenPop = True Or UCase(frm1.txtRcptNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
	
	IsOpenPop = True		
 	
	arrParam(0) = ""'Trim(frm1.hdnSupplierCd.Value)
	arrParam(1) = ""'Trim(frm1.hdnGroupCd.Value)
	arrParam(2) = ""'Trim(frm1.hdnMvmtType.Value)		
	arrParam(3) = "I"
		
	iCalledAspName = AskPRAspName("M4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False
	
	If strRet(0) = "" Then
		frm1.txtRcptNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtRcptNo.value = strRet(0)
		frm1.txtRcptNo.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'================================================================================================================================ 
Function OpenGmNo()
	Dim arrRet
	Dim arrParam(0)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtGRNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True		
 			
	iCalledAspName = AskPRAspName("M4161PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M4161PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGRNo.focus	
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtGRNo.Value	= arrRet(0)	
		frm1.txtGRNo.focus	
		Set gActiveElement = document.activeElement	
	End If	

End Function
'================================================================================================================================ 
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	arrParam(4) = ""			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)	
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement		
	End If	

End Function 
'================================================================================================================================ 
Sub Form_Load()
    
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    Call InitVariables											  '⊙: Initializes local global variables
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery()
	
End Sub
'================================================================================================================================ 
Sub txtToRegDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToRegDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtToRegDt.focus
	End if
End Sub
'================================================================================================================================ 
Sub txtFrRegDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrRegDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtFrRegDt.focus
	End if
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
Sub txtFrRegDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================ 
Sub txtToRegDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
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
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrRegDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToRegDt.text,PopupParent.gDateFormat,"")) and Trim(.txtFrRegDt.text)<>"" and Trim(.txtToRegDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","등록일", "X")			
			.txtToRegDt.Focus
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

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
			strVal = strVal & "&txtRcptNo="			& .hdnRcptNo.value
			strVal = strVal & "&txtGRNo="			& .hdnGRNo.value	    	    
			strVal = strVal & "&txtFrRegDt="		& .hdnFrDt.value
			strVal = strVal & "&txtToRegDt="		& .hdnToDt.value
			strVal = strVal & "&txtGroup="			& .hdnGroup.value
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtRcptNo="			& Trim(.txtRcptNo.value)
			strVal = strVal & "&txtGRNo="			& Trim(.txtGRNo.value)	    	    
			strVal = strVal & "&txtFrRegDt="		& Trim(.txtFrRegDt.text)
			strVal = strVal & "&txtToRegDt="		& Trim(.txtToRegDt.text)
			strVal = strVal & "&txtGroup="			& Trim(.txtGroupCd.Value)
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function
'================================================================================================================================ 
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtRcptNo.focus
	End If

End Function
'================================================================================================================================ 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
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
						<TD CLASS="TD5" NOWRAP>입고번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입고번호" NAME="txtRcptNo" SIZE=32 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:openRcptNo()">
						<TD CLASS="TD5" NOWRAP>재고처리번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="재고처리번호" NAME="txtGRNo" SIZE=32 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGmNo()">
					</TR>
					<TR>										
						<TD CLASS="TD5" NOWRAP>등록일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr NOWRAP>
									<td NOWRAP>
										<script language =javascript src='./js/m4131pa1_fpDateTime1_txtFrRegDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m4131pa1_fpDateTime1_txtToRegDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>구매그룹</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
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
						<script language =javascript src='./js/m4131pa1_vaSpread_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnRcptNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGRNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
