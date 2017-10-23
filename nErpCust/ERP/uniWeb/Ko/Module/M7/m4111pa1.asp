<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4111PA1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 검사결과등록-결과등록번호-입고번호팝업									*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/05/28																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : KIM JIN HA																*			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'******************************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit                              '☜: indicates that All variables must be declared in advance
'================================================================================================================================
Const BIZ_PGM_QRY_ID = "m4111pb1.asp"                               '☆: Biz Logic ASP Name
Const C_MaxKey          = 8                                           '☆: key count of SpreadSheet
Const C_LC_NO			= 1
'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam

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
	Dim arrParent
		
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                           'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
		
	arrParent = window.dialogArguments
	gblnWinEvent = False
        
    arrReturn = ""
    Redim arrReturn(0)  
    Self.Returnvalue = arrReturn     
End Function
'================================================================================================================================
Sub SetDefaultVal()
	Err.Clear
	
	frm1.txtFrRcptDt.text = StartDate
	frm1.txtToRcptDt.text = EndDate
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M4111PA1","S","A","V20030328",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
    Redim arrReturn(0)
    
    If frm1.vspdData.ActiveRow < 1 Then 
		Self.Close()
		Exit Function
	End if
	
    With frm1
		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Col = GetKeyPos("A",C_LC_NO)
		arrReturn(0) = .vspdData.Text
	End with
		
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
	
	'If lgIsOpenPop = True Then Exit Function
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
		Case 1
			arrParam(0) = "공급처"							<%' 팝업 명칭 %>
			arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
			arrParam(2) = Trim(frm1.txtSupplierCd.Value)	    <%' Code Condition%>
			'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	    <%' Name Cindition%>
			arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & " "							<%' Where Condition%>
			arrParam(5) = "공급처"								<%' TextBox 명칭 %>
				
			arrField(0) = "BP_CD"									<%' Field명(0)%>
			arrField(1) = "BP_NM"									<%' Field명(1)%>
				    
			arrHeader(0) = "공급처"								<%' Header명(0)%>
			arrHeader(1) = "공급처명"							<%' Header명(1)%>
		Case 2					
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
		case 3
			arrParam(0) = "입고형태"	
			'arrParam(1) = "M_Mvmt_type"
			arrParam(1) = "( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("N", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and ((b.RCPT_FLG=" & FilterVar("Y", "''", "S") & "  AND b.RET_FLG=" & FilterVar("N", "''", "S") & " ) or (b.RET_FLG=" & FilterVar("N", "''", "S") & "  And b.SUBCONTRA_FLG=" & FilterVar("N", "''", "S") & " )) ) c"				
			arrParam(2) = Trim(frm1.txtMvmtType.Value)
			arrParam(3) = Trim(frm1.txtMvmtTypeNm.Value)
			
			'arrParam(4) = "((RCPT_FLG='Y' AND RET_FLG='N') or (RET_FLG='N' And SUBCONTRA_FLG='N')) AND USAGE_FLG='Y' "
			arrParam(5) = "입고형태"			
					
			arrField(0) = "IO_Type_Cd"	
			arrField(1) = "IO_Type_NM"	
				    
			arrHeader(0) = "입고형태"		
			arrHeader(1) = "입고형태명"
	End Select

	arrParam(0) = arrParam(5)								<%' 팝업 명칭 %>

	Select Case iWhere
	Case 1,2,3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

    arrParam(0) = arrParam(5)	
    gblnWinEvent = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtSupplierCd.focus			  
			Case 2
			    frm1.txtGroupCd.focus			
			case 3
			    frm1.txtMvmtType.focus	
		End Select	
		Set gActiveElement = document.activeElement			
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
		
End Function
'================================================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtSupplierCd.Value = arrRet(0)
				.txtSupplierNm.Value = arrRet(1)
				.txtSupplierCd.focus			  
			Case 2
			    .txtGroupCd.Value = arrRet(0)
			    .txtGroupNm.Value = arrRet(1)
			    .txtGroupCd.focus			
			case 3
			    .txtMvmtType.Value = arrRet(0) 
			    .txtMvmtTypeNm.Value = arrRet(1)
			    .txtMvmtType.focus	
		End Select	
		
		Set gActiveElement = document.activeElement			  
	End With
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
    call FncQuery() 
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
Sub txtFrRcptDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrRcptDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtFrRcptDt.focus
	End If
End Sub
'================================================================================================================================
Sub txtToRcptDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToRcptDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtToRcptDt.focus
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
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
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
	
	Err.Clear														'☜: Protect system from crashing
	
	DbQuery = False													'⊙: Processing is NG
	If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		   
		    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001		    
		    strVal = strVal & "&txtMvmtType=" & Trim(frm1.hdnMvmtType.value)
		    strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplier.value)
			strVal = strVal & "&txtFrRcptDt=" & Trim(frm1.hdnFrRcptDt.value)
			strVal = strVal & "&txtToRcptDt=" & Trim(frm1.hdnToRcptDt.value)
		    strVal = strVal & "&txtGroup=" & Trim(frm1.hdnGroup.value)
		    strVal = strVal & "&txtInspFlag=" & Trim(frm1.hdnInspFlag.value)		
		else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
		    strVal = strVal & "&txtMvmtType=" & Trim(frm1.txtMvmtType.value)
		    strVal = strVal & "&txtSupplier=" & Trim(frm1.txtSupplierCd.value)
			strVal = strVal & "&txtFrRcptDt=" & Trim(frm1.txtFrRcptDt.text)
			strVal = strVal & "&txtToRcptDt=" & Trim(frm1.txtToRcptDt.text)
		    strVal = strVal & "&txtGroup=" & Trim(frm1.txtGroupCd.Value)
		    strVal = strVal & "&txtInspFlag=" & frm1.hdnInspFlag.value	
		End if
		        
			strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

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
		frm1.vspdData.focus
	End If

End Function
'================================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
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
						<TD CLASS="TD5" NOWRAP>입고형태</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 3">
											   <INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>입고일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr NOWRAP>
									<td NOWRAP>
										<script language =javascript src='./js/m4111pa1_fpDateTime1_txtFrRcptDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m4111pa1_fpDateTime1_txtToRcptDt.js'></script>
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
						<script language =javascript src='./js/m4111pa1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnInspFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrRcptDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToRcptDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
