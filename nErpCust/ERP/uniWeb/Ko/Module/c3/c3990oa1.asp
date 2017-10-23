
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : Costing
'*  2. Function Name        : 
'*  3. Program ID           : c3990oa1.asp
'*  4. Program Name         : 원가추이 
'*  5. Program Desc         : Template
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2004/05/
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho Ig sung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    
End Sub

Sub SetDefaultVal()


	Dim strFrDate, strToDate
	strFrDate	= Parent.gFiscStart		' 당기시작일 
	strToDate	= "<%=GetSvrDate%>"		' 현재일 

	frm1.txtFrYyyymm.text	= UniConvDateAToB(strFrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtToYyyymm.text	= UniConvDateAToB(strToDate,Parent.gServerDateFormat,Parent.gDateFormat)

	Call ggoOper.FormatDate(frm1.txtFrYyyymm, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToYyyymm, Parent.gDateFormat, 2)
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

' 모듈 코드로 입력 예:원가 "C"
<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "PA") %>
End Sub

Sub InitComboBox()

	'품목계정 Combo
	Call CommonQueryRs(" a.MINOR_CD,a.MINOR_NM "," B_MINOR a,b_item_acct_inf b ","a.MAJOR_CD = " & parent.FilterVar("P1001","''","S") & " and a.minor_cd = b.item_acct and b.item_acct_group <> " & parent.FilterVar("6MRO","''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboItemAcct ,lgF0  ,lgF1  ,Chr(11))

End Sub


Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "공장팝업"
			arrParam(1) = "B_PLANT"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장"			
	
			arrField(0) = "PLANT_CD"
			arrField(1) = "PLANT_NM"
			 
			arrHeader(0) = "공장코드"
			arrHeader(1) = "공장명"
		Case 1,2
			arrParam(0) = "품목팝업"
			arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "a.ITEM_CD = b.item_cd"

			IF frm1.txtPlantCd.value <> "" Then 
				arrParam(4) = arrParam(4) & " and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
			END IF

			IF frm1.cboItemAcct.value  <> "" Then 
				arrParam(4) = arrParam(4) & " and b.item_acct = " & FilterVar(frm1.cboItemAcct.value, "''", "S")
			END IF

			arrParam(5) = "품목"			
	
			arrField(0) = "a.ITEM_CD"
			arrField(1) = "a.ITEM_NM"
			arrField(2) = "a.SPEC"
			 
			arrHeader(0) = "품목코드"
			arrHeader(1) = "품목명"
			arrHeader(1) = "SPEC"
		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				frm1.txtPlantCd.focus
			Case 1
				frm1.txtFrItemCd.focus
			Case 2
				frm1.txtToItemCd.focus
			Case Else
		End Select
		Exit Function

	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

Function SetPopup(Byval arrRet, Byval iWhere)
	
	With frm1
	
		Select Case iWhere
			Case 0
				.txtPlantCd.focus
    			.txtPlantCd.value = arrRet(0)
    			.txtPlantNm.value = arrRet(1)
			Case 1
    			.txtFrItemCd.focus
    			.txtFrItemCd.value = arrRet(0)
    			.txtFrItemNm.value = arrRet(1)
			Case 2
    			.txtToItemCd.focus
    			.txtToItemCd.value = arrRet(0)
    			.txtToItemNm.value = arrRet(1)
			Case Else
				Exit Function
		End Select
	End With

End Function

Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitVariables
    Call SetDefaultVal

	Call InitComboBox
    Call SetToolbar("10000000000011")
    
    frm1.txtFrYyyymm.focus 
    Set gActiveElement = document.activeElement	
    
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

' 날짜에서 엔터키 입력시 미리보기 실행 
Sub txtFrYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrYyyymm.focus
	End If
End Sub

Sub txtToYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtToYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToYyyymm.focus
	End If
End Sub

Sub txtFrYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Sub txtToYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Function FncQuery()
	' 엔터키 입력시 미리보기 실행 
    FncBtnPreview()
End Function

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strFrYyyymm, strToYyyymm, strItemAcct, strPlantCd, strFrItemCd, strToItemCd
	Dim	strFrYear, strFrMonth, strFrDay, strToYear, strToMonth, strToDay

	SetPrintCond = False

	if frm1.PrintOpt1.checked = True then
		StrEbrFile = "c3990oa1a"
	elseif frm1.PrintOpt2.checked = True then
		StrEbrFile = "c3990oa1b"
	elseif frm1.PrintOpt3.checked = True then
		StrEbrFile = "c3990oa1c"
	elseif frm1.PrintOpt4.checked = True then
		StrEbrFile = "c3990oa1d"
	else
		Exit Function
	end if

	' 시작월 종료월 비교 
    If CompareDateByFormat(frm1.txtFrYyyymm.Text,frm1.txtToYyyymm.Text,frm1.txtFrYyyymm.Alt,frm1.txtToYyyymm.Alt, _
	 "970024", frm1.txtFrYyyymm.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtToYyyymm.Focus
		Exit Function
	End If

	' 시작품목 종료품목 비교 
	If Trim(UCase(frm1.txtFrItemCd.value)) > Trim(UCase(frm1.txtToItemCd.value)) Then
		Call DisplayMsgBox("970024","X", frm1.txtFrItemCd.Alt , frm1.txtToItemCd.Alt)
		frm1.txtToItemCd.Focus
		Exit Function
	End If

	Call ExtractDateFrom(frm1.txtFrYyyymm.Text,frm1.txtFrYyyymm.UserDefinedFormat,Parent.gComDateType,strFrYear,strFrMonth,strFrDay)
	Call ExtractDateFrom(frm1.txtToYyyymm.Text,frm1.txtToYyyymm.UserDefinedFormat,Parent.gComDateType,strToYear,strToMonth,strToDay)

	strFrYyyymm		= strFrYear & strFrMonth
	strToYyyymm		= strToYear & strToMonth

	strItemAcct		= Trim(UCase(frm1.cboItemAcct.value))
	strPlantCd		= Trim(UCase(frm1.txtPlantCd.value))
	strFrItemCd		= Trim(UCase(frm1.txtFrItemCd.value))
	strToItemCd		= Trim(UCase(frm1.txtToItemCd.value))
	
	if strItemAcct = "" then
		strItemAcct = "%"
	End if	

	if strPlantCd = "" then
		strPlantCd = "%"
		frm1.txtPlantNm.value = ""
	End if	

	if strFrItemCd = "" then
		strFrItemCd = ""
		frm1.txtFrItemNm.value = ""
	End if	

	if strToItemCd = "" then
		strToItemCd = "ZZZZZZZZZZZZZZZZZZ"
		frm1.txtToItemNm.value = ""
	End if	

	strUrl	= strUrl & "fr_yyyymm|"		& strFrYyyymm
	strUrl	= strUrl & "|to_yyyymm|"	& strToYyyymm
	strUrl	= strUrl & "|item_acct|"	& strItemAcct
	strUrl	= strUrl & "|plant_cd|"		& strPlantCd
	strUrl	= strUrl & "|fr_item_cd|"	& strFrItemCd
	strUrl	= strUrl & "|to_item_cd|"	& strToItemCd

	SetPrintCond = True
End Function

Function FncBtnPrint() 

    Dim StrEbrFile, strUrl
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPrint(EBAction,ObjName,strUrl)	

End Function

Function FncBtnPreview() 
    
    Dim StrEbrFile, strUrl

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPreview(ObjName , strUrl)
	
End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

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
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원가추이</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">총평균단가</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">입고단가</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt3" VALUE="N" tag="25"><LABEL FOR="PrintOpt3">작업단계별단가</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt4" VALUE="N" tag="25"><LABEL FOR="PrintOpt4">원가요소별단가</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/c3990oa1_fpDateTime1_txtFrYyyymm.js'></script>&nbsp;~&nbsp;
														<script language =javascript src='./js/c3990oa1_fpDateTime2_txtToYyyymm.js'></script></TD>
							</TR>

							<TR>
								<TD CLASS="TD5" NOWRAP>품목계정</TD>
								<TD CLASS="TD6" NOWRAP><SELECT ID="cboItemAcct" NAME="cboItemAcct" ALT="품목계정" STYLE="WIDTH: 120px" tag="11X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtPlantCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtFrItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="시작품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtFrItemCd.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtFrItemNm" SIZE=30 tag="14"></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>&nbsp;~&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtToItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="종료품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtToItemCd.value, 2">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=30 tag="14"></TD>
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
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
<!--						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>	-->
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="strUrl" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

