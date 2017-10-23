 <%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 세금계산서출력 
'*  3. Program ID           : S5111oa5
'*  4. Program Name         : 세금계산서출력 
'*  5. Program Desc         : 세금계산서 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/07/18
'*  8. Modified date(Last)  : 2000/07/18
'*  9. Modifier (First)     : 손범열 
'* 10. Modifier (Last)      : Son bum yeol
'* 11. Comment              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Const C_PopSoldToParty = 1

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Dim gblnWinEvent

'=========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.TAX_BILL_NO.focus 
	frm1.cboApType.value = "양식출력"
	frm1.txtIssuedFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtIssuedToDt.Text = EndDate
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "OA") %>	
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

Function OpenTaxbillNoPop()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
		
	iCalledAspName = AskPRAspName("s5311pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5311pa1", "x")
		gblnWinEvent = False
		exit Function
	end if
	
	strRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	gblnWinEvent = False

	frm1.TAX_BILL_NO.focus
			
	If strRet <> "" Then Call SetTaxbillNoPop(strRet)
End Function

'=========================================
Function SetTaxbillNoPop(arrRet)
	frm1.TAX_BILL_NO.Value = arrRet
		
	Call TAX_BILL_NO_OnChange()
End Function

'=========================================
Sub InitComboBox()
	With frm1
	     .cboApType.value = "양식출력"
		Call SetCombo(.cboApType, "양식출력","양식출력")
		Call SetCombo(.cboApType, "무양식출력","무양식출력")                              
    End With
End Sub

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables
    Call GetValue_ko441()														'⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
    Call InitComboBox
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Sub btnTaxbillNo_OnClick()
	Call OpenTaxbillNoPop()
End Sub

'=========================================
Sub txtIssuedFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssuedFromDt.Focus
    End If
End Sub

'=========================================
Sub txtIssuedToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssuedToDt.Focus
    End If
End Sub

' 세금계산서관리번호가 변경되는 경우, 과세여부를 참조한다.
'========================================
Sub TAX_BILL_NO_OnChange()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrTemp

	Err.Clear
	 
	If Trim(frm1.TAX_BILL_NO.value) = "" Then
		frm1.rdoTaxFlagY.checked = True
		frm1.rdoTaxFlagN.checked = False
		ggoOper.SetReqAttr frm1.rdoTaxFlagN, "N"
		ggoOper.SetReqAttr frm1.rdoTaxFlagY, "N"		
	Else
		iStrSelectList = " REFERENCE "
		iStrFromList = " S_TAX_BILL_HDR LEFT OUTER JOIN (SELECT MINOR_CD, REFERENCE FROM B_CONFIGURATION WHERE MAJOR_CD = " & FilterVar("B9001", "''", "S") & " AND SEQ_NO = 9) TMP ON S_TAX_BILL_HDR.VAT_TYPE = TMP.MINOR_CD "
		iStrWhereList = "TAX_BILL_NO =  " & FilterVar(frm1.TAX_BILL_NO.value , "''", "S") & ""
		    
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrTemp = Split(iStrRs, Chr(11))
			If iArrTemp(1) = "Y" Then	'--- 면세세금계산서 
				frm1.rdoTaxFlagN.checked = True
				frm1.rdoTaxFlagY.checked = False
			Else '--- 과세세금계산서 
				frm1.rdoTaxFlagY.checked = True
				frm1.rdoTaxFlagN.checked = False	
			End IF
			ggoOper.SetReqAttr frm1.rdoTaxFlagN, "Q"
			ggoOper.SetReqAttr frm1.rdoTaxFlagY, "Q"
		End If
	End if
End Sub

'========================================
Function btnPreview_OnClick()
	Call BtnPrint("N")
End Function

'========================================
Function btnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'========================================
Function BtnPrint(ByVal pvStrPrint) 
	Dim iStrUrl
	Dim vargBizArea,vargPlant,vargSalesGrp,vargSalesOrg

    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtIssuedFromDt, frm1.txtIssuedToDt) = False Then Exit Function

	If UCase(frm1.TAX_BILL_NO.value) = "" Then
		iStrUrl = "TAX_BILL_NO|%"
	Else
		iStrUrl = "TAX_BILL_NO|" & Replace(Trim(frm1.TAX_BILL_NO.value), "'" ,  "''")
	End If

	If lgBACd <> "" Then
		vargBizArea = " AND S_TAX_BILL_HDR.BIZ_AREA_CD =  " & FilterVar(Trim(UCase(lgBACd)), "" ,  "S")
	Else
		vargBizArea = ""
	End If
	'If lgPLCd <> "" Then
	'	vargPlant = " AND S_DN_DTL.PLANT_CD = " & FilterVar(Trim(UCase(lgPLCd)), "" ,  "S")
	'Else
	'	vargPlant = ""
	'End If
	If lgSGCd <> "" Then
		vargSalesGrp = " AND S_TAX_BILL_HDR.SALES_GRP = " & FilterVar(Trim(UCase(lgSGCd)), "" ,  "S")
	Else
		vargSalesGrp = ""
	End If
	If lgSOCd <> "" Then
		vargSalesOrg = " AND S_TAX_BILL_HDR.SALES_ORG = " & FilterVar(Trim(UCase(lgSOCd)), "" ,  "S")
	Else
		vargSalesOrg = ""
	End If


	iStrUrl = iStrUrl & "|ISSUED_FROM_DT|" & UniConvDateToYYYYMMDD(frm1.txtIssuedFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	iStrUrl = iStrUrl & "|ISSUED_TO_DT|" & UniConvDateToYYYYMMDD(frm1.txtIssuedToDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	istrUrl = istrUrl & "|gBizArea|" & vargBizArea 
	'istrUrl = istrUrl & "|gPlant|" & vargPlant 
	istrUrl = istrUrl & "|gSalesGrp|" & vargSalesGrp
	istrUrl = istrUrl & "|gSalesOrg|" & vargSalesOrg 

	If frm1.rdoTaxFlagY.checked = true then	'--- 과세세금계산서 
		If frm1.cboApType.Value = "양식출력" then
			ObjName = AskEBDocumentName("s5111oa5_KO441","ebr")
		Else
			ObjName = AskEBDocumentName("s5111oa6_KO441","ebr")
		End If
	Else
		If frm1.cboApType.Value = "양식출력" then
			ObjName = AskEBDocumentName("s5111oa8_KO441","ebr")
		Else
			ObjName = AskEBDocumentName("s5111oa9_KO441","ebr")
		End If	
	End if
		
	If pvStrPrint = "N" Then
		' 미리보기 
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		' 출력 
		Call FncEBRprint(EBAction, ObjName, iStrUrl)
	End If
End Function

'========================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================
Function FncExit()
	FncExit = True
End Function

'========================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>세금계산서관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="TAX_BILL_NO" ALT="세금계산서관리번호" SIZE=20 MAXLENGTH="18" TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxbillNo" align=top TYPE="BUTTON" ><div style="display:none"><input type="text" name="none"></div></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>세금계산서양식선택</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboApType" ALT="세금계산서양식선택" STYLE="Width: 150px;" tag="22"></SELECT></TD>                                                    
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>발행일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
										<TD>
											<script language =javascript src='./js/s5111oa5_fpDateTime1_txtIssuedFromDt.js'></script>
										</TD>
										<TD>
												&nbsp;~&nbsp;
										</TD>
										<TD>
											<script language =javascript src='./js/s5111oa5_fpDateTime2_txtIssuedToDt.js'></script>
										</TD>
										</TABLE>				
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>세금계산서종류</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoTaxFlag" id="rdoTaxFlagY" value="Y" tag = "21" checked>
										<label for="rdoPostFlagY">과세세금계산서</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoTaxFlag" id="rdoTaxFlagN" value="N" tag = "21">
										<label for="rdoPostFlagN">면세세금계산서</label></TD>									
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD>    
						    <BUTTON NAME="BtnPreview" CLASS="CLSSBTN" Flag=1>미리보기</BUTTON>&nbsp;
						    <BUTTON NAME="BtnPrint" CLASS="CLSSBTN" Flag=1>인쇄</BUTTON>
						</TD>

					</TR>
								
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX ="-1">
    <input type="hidden" name="dbname" TABINDEX ="-1">
    <input type="hidden" name="filename" TABINDEX ="-1">
    <input type="hidden" name="condvar" TABINDEX ="-1">
	<input type="hidden" name="date" TABINDEX ="-1">
</FORM>
</BODY>
</HTML>
