<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1 %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3111ra4.asp																*
'*  4. Program Name         : 발주참조(매입세금계산서)					   					    	*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 000/05/10																*
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!--<TITLE>발주참조</TITLE> -->
<TITLE></TITLE>
<!--
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "m3111rb4_KO441.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 16                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                    
Dim lgCookValue 
Dim IsOpenPop  
Dim gblnWinEvent

Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParam	

Dim lgStrPrevKey2

Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
						
	frm1.vspdData.MaxRows = 0	
			
	gblnWinEvent = False
	ReDim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
	
	
'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	frm1.txtFrPoDt.Text = StartDate
	frm1.txtToPoDt.Text = EndDate

End Sub
	

'==========================================  LoadInfTB19029()  ==========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA")%>
End Sub
'==========================================  InitSpreadSheet()  ==========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3111RA4","S","A","V20030328",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
    Call SetSpreadLock 
	frm1.vspdData.OperationMode = 3
End Sub
'==========================================  SetSpreadLock()  ==========================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	
	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim arrReturn(frm1.vspdData.MaxCols - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To frm1.vspdData.MaxCols - 1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
				
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'------------------------------------------  OpenPoType()  -------------------------------------------------
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주형태"						<%' 팝업 명칭 %>
	arrParam(1) = "M_CONFIG_PROCESS"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPotypeCd.Value)	<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)	<%' Name Cindition%>
	arrParam(4) = "import_flg=" & FilterVar("N", "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "발주형태"							<%' TextBox 명칭 %>
	
    arrField(0) = "PO_TYPE_CD"					<%' Field명(0)%>
    arrField(1) = "PO_TYPE_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "발주형태"						<%' Header명(0)%>
    arrHeader(1) = "발주형태명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPotypeCd.focus
		Exit Function
	Else
		frm1.txtPotypeCd.Value = arrRet(0)
		frm1.txtPotypeNm.Value = arrRet(1)
		frm1.txtPotypeCd.focus
	End If	
	Set gActiveElement = document.activeElement
End Function
'------------------------------------------  OpenSupplier()  --------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"						<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	<%' Name Cindition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
	arrParam(5) = "공급처"							<%' TextBox 명칭 %>
	
    arrField(0) = "BP_Cd"					<%' Field명(0)%>
    arrField(1) = "BP_NM"					<%' Field명(1)%>
    
    arrHeader(0) = "공급처"						<%' Header명(0)%>
    arrHeader(1) = "공급처명"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenGroup()  --------------------------------------------
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주담당"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "발주담당"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "발주담당"		
    arrHeader(1) = "발주담당명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)		
		frm1.txtGroupCd.focus
	End If	
	Set gActiveElement = document.activeElement
End Function 

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call InitVariables							
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call MM_preloadimages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'------------------------------------------  OpenSortPopup()  --------------------------------------------
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

'=========================================  OCX_EVENT  ===================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToPoDt.focus
	End if
End Sub

Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrPoDt.focus
	End if
END SUB
'=========================================  3.3.1 vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
		Exit Function
	End If
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================  3.3.2 vspdData_KeyPress()  ===================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If CheckRunningBizProcess = True Then
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

'==========================================================================================
'   Event Name : DateOcx_Keypress
'==========================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub


Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=====================================  FncQuery()  ========================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        

	If ValidDateCheck(frm1.txtFrPoDt, frm1.txtToPoDt) = False Then Exit Function

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData						
	Call InitVariables												

    If DbQuery = False Then Exit Function

    Call DbQuery									

    FncQuery = True									
    Set gActiveElement = document.activeElement    
End Function

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>
		

	Dim strVal
		
	if LayerShowHide(1) = false then
		exit function
	end if

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPotype=" & Trim(frm1.hdnPoType.value)
		strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplier.value)
		strVal = strVal & "&txtFrPoDt=" & Trim(frm1.hdnFrDt.value)
		strVal = strVal & "&txtToPoDt=" & Trim(frm1.hdnToDt.value)
		strVal = strVal & "&txtGroup=" & Trim(frm1.hdnGroup.value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPotype=" & Trim(frm1.txtPoTypeCd.value)
		strVal = strVal & "&txtSupplier=" & Trim(frm1.txtSupplierCd.value)
		strVal = strVal & "&txtFrPoDt=" & Trim(frm1.txtFrPoDt.Text)
		strVal = strVal & "&txtToPoDt=" & Trim(frm1.txtToPoDt.Text)
		strVal = strVal & "&txtGroup=" & Trim(frm1.txtGroupCd.value)
	End if 
			
    strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
    strVal =	 strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
    strVal =	 strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal =	 strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

	Call RunMyBizASP(MyBizASP, strVal)								<%'☜: 비지니스 ASP 를 가동 %>
		
	DbQuery = True														<%'⊙: Processing is NG%>

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPotypeCd.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->	
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
						<TD CLASS="TD5" NOWRAP>발주형태</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="발주형태" NAME="txtPotypeCd" MAXLENGTH=5 SIZE=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPotype()">
											   <INPUT TYPE=TEXT AlT="발주형태" NAME="txtPotypeNm" SIZE=20 tag="14X" ></TD>
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="공급처" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
					   			 	     	   <INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>발주일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3111ra4_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3111ra4_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>발주담당</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="발주담당" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
										 	   <INPUT TYPE=TEXT AlT="발주담당" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
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
						<script language =javascript src='./js/m3111ra4_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" Tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">


</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
