
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 품목별배부내역 출력 
'*  3. Program ID           : c3606oa1
'*  4. Program Name         : 품목별배부내역 출력 
'*  5. Program Desc         : 품목별배부내역 출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2004/05/07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Tae Soo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================  -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_QRY_ID = "c3606ob1.asp"

Dim lgBlnFlgChgValue
Dim lgIntFlgMode
Dim lgIntGrpCount

Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
         
End Sub

Sub SetDefaultVal()
	Dim StartDate

	StartDate	= "<%=GetSvrDate%>"
'	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)

	frm1.txtFromYyyymm.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtFromYYYYMM, Parent.gDateFormat, 2)
	
	
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "PA") %>
End Sub

Function OpenPopup(Byval param, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1
			arrParam(0) = "공장 팝업"
			arrParam(1) = "B_PLANT"
			arrParam(2) = Trim(frm1.txtPlantCd.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장"
	
			arrField(0) = "plant_cd"
			arrField(1) = "plant_nm"
    
			arrHeader(0) = "공장"
			arrHeader(1) = "공장명"	
		Case 2
			arrParam(0) = "품목계정 팝업"
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"
			arrParam(2) = Trim(frm1.txtitemAcct.value)
			arrParam(3) = ""
			arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group in ('1FINAL','2SEMI') "
			arrParam(5) = "품목계정"
	
			arrField(0) = "minor_CD"
			arrField(1) = "minor_NM"
    
			arrHeader(0) = "품목계정"
			arrHeader(1) = "품목계정명"
		Case 3
			arrParam(0) = "코스트센터 팝업"
			arrParam(1) = "B_COST_CENTER"
			arrParam(2) = Trim(frm1.txtCostCd1.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "코스트센터"
	
			arrField(0) = "COST_CD"
			arrField(1) = "COST_NM"
    
			arrHeader(0) = "코스트센터코드"
			arrHeader(1) = "코스트센터명"
		Case 4
			arrParam(0) = "코스트센터 팝업"
			arrParam(1) = "B_COST_CENTER"
			arrParam(2) = Trim(frm1.txtCostCd2.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "코스트센터"
	
			arrField(0) = "COST_CD"
			arrField(1) = "COST_NM"
    
			arrHeader(0) = "코스트센터코드"
			arrHeader(1) = "코스트센터명"
		Case 5
			arrParam(0) = "코스트센터 팝업"
			arrParam(1) = "B_COST_CENTER"
			arrParam(2) = Trim(frm1.txtCostCd3.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "코스트센터"
	
			arrField(0) = "COST_CD"
			arrField(1) = "COST_NM"
    
			arrHeader(0) = "코스트센터코드"
			arrHeader(1) = "코스트센터명"	
		Case 6
			arrParam(0) = "계정 팝업"
			arrParam(1) = "A_ACCT"
			arrParam(2) = Trim(frm1.txtAcctCd.Value)
			arrParam(3) = ""
			arrParam(4) = "temp_fg_3 LIKE " & FilterVar("m%", "''", "S") & " "
			arrParam(5) = "코스트센터"
	
			arrField(0) = "ACCT_CD"
			arrField(1) = "ACCT_NM"
    
			arrHeader(0) = "계정"
			arrHeader(1) = "계정명"	
		Case 7
			arrParam(0) = "품목 팝업"
			arrParam(1) = "B_ITEM"
			arrParam(2) = Trim(frm1.txtItemCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "품목"
	
			arrField(0) = "ITEM_CD"
			arrField(1) = "ITEM_NM"
    
			arrHeader(0) = "품목"
			arrHeader(1) = "품목명"									
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	  If iWhere = 1 Then
		frm1.txtPlantCd.focus	
	  ElseIf iWhere = 2 Then
		frm1.txtitemAcct.focus
	  ElseIf iWhere = 3 Then
	    frm1.txtCostCd1.focus
	  ElseIf iWhere = 4 Then
	    frm1.txtCostCd2.focus
	  ElseIf iWhere = 5 Then
	    frm1.txtCostCd3.focus	    
	  ElseIf iWhere = 6 Then
	    frm1.txtAcctCd.focus	    
	  ElseIf iWhere = 7 Then
	    frm1.txtItemCd.focus	    

      End If	    	
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If
		
End Function

Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
	Select Case iWhere
		Case 1
			frm1.txtPlantCd.focus
			frm1.txtPlantCd.value	= arrRet(0)		
			frm1.txtPlantNm.value	= arrRet(1)		
		Case 2
			frm1.txtItemAcct.focus
			frm1.txtItemAcct.value	= arrRet(0)		
			frm1.txtItemAcctNm.value= arrRet(1)		
		Case 3
			frm1.txtCostCd1.focus
			frm1.txtCostCd1.Value	= arrRet(0)		
			frm1.txtCostNm1.Value	= arrRet(1)		
		Case 4
			frm1.txtCostCd2.focus
			frm1.txtCostCd2.Value	= arrRet(0)		
			frm1.txtCostNm2.Value	= arrRet(1)		
		Case 5
			frm1.txtCostCd3.focus
			frm1.txtCostCd3.Value	= arrRet(0)		
			frm1.txtCostNm3.Value	= arrRet(1)		
		Case 6
			frm1.txtAcctCd.focus
			frm1.txtAcctCd.Value	= arrRet(0)		
			frm1.txtAcctNm.Value	= arrRet(1)		
		Case 7
			frm1.txtItemCd.focus
			frm1.txtItemCd.Value	= arrRet(0)		
			frm1.txtItemNm.Value	= arrRet(1)														
		Case Else
			Exit Function
	End select	

End Function

Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitVariables                                        
    Call SetDefaultVal
    Call SetToolbar("10000000000011")
    
    frm1.txtFromYyyymm.focus 
    Set gActiveElement = document.activeElement	
    
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

Sub txtFromYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromYyyymm.focus
	End If
End Sub


Sub txtFromYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub


Function FncQuery() 
    FncBtnPreview()
End Function

Sub SetPrintCond(StrEbrFile, strFromyyyymm )
	Dim	strYear, strMonth, strDay

	if frm1.PrintOpt1.Checked = True then
		StrEbrFile = "c3606oa1"
	elseif frm1.PrintOpt2.Checked = True then
		StrEbrFile = "c3606oa2"
	elseif frm1.PrintOpt3.Checked = True then
		StrEbrFile = "c3606oa3"
	elseif frm1.PrintOpt4.Checked = True then
		StrEbrFile = "c3606oa4"
	end if

	Call ExtractDateFrom(frm1.txtFromYyyyMm.Text,frm1.txtFromYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	strFromyyyymm	= strYear & strMonth

End Sub

Function FncBtnPrint() 
 
    Dim StrEbrFile
	Dim condvar
	dim stryyyymm, strworkstep, stritemacct

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call BtnDisabled(1)
	

	Call BatchExe()

    
End Function




Function PrintExecOk()

	dim strFromyyyymm
    Dim StrEbrFile
	Dim condvar


	Call SetPrintCond(StrEbrFile, strFromyyyymm)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & strFromyyyymm
	condvar	= condvar & "|spid|" & Trim(frm1.txtSpId.value)
	
	call FncEBRprint(EBAction, ObjName, strUrl)
	
	Call BtnDisabled(0)
End Function




Function FncBtnPreview() 
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call BtnDisabled(1)
	
	Call BatchExe()
	

	
End Function

'========================================================================================
' Function Name : PrevExecOk()
' Function Desc : BOM Temp 테이블에 데이터 생성이 성공하면 EasyBase를 Open한다.
'========================================================================================
Function PrevExecOk()

	dim strFromyyyymm
    Dim StrEbrFile
	Dim condvar


	Call SetPrintCond(StrEbrFile, strFromyyyymm)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & strFromyyyymm
	condvar	= condvar & "|spid|" & Trim(frm1.txtSpId.value)
	
	call FncEBRPrevIew(ObjName , condvar)	
	
	
	Call BtnDisabled(0)

End Function


Function BatchExe()
    Dim strVal
    Dim strYear,strMonth,strDay
    Dim strFromyyyymm
    Dim strPlant,stritemacct,strCostCd,strFromCostCd,strDstbCostCd,strAcctCd,strItemCd
    
	Call ExtractDateFrom(frm1.txtFromYyyyMm.Text,frm1.txtFromYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	strFromyyyymm	= strYear & strMonth


	strplant	= Trim(UCase(frm1.txtPlantCd.value))
	stritemacct	= Trim(UCase(frm1.txtItemAcct.value))
	strCostCd= Trim(UCase(frm1.txtCostCd1.value))
	strFromCostCd= Trim(UCase(frm1.txtCostCd2.value))
	strDstbCostCd= Trim(UCase(frm1.txtCostCd3.value))
	strAcctCd= Trim(UCase(frm1.txtAcctCd.value))
	strItemCd= Trim(UCase(frm1.txtItemCd.value))

	
	
	
	if strPlant = "" then
		strPlant = "%"
		frm1.txtPlantNm.value	= ""
	End if	
	
	if stritemacct = "" then
		stritemacct = "%"
		frm1.txtItemAcctNm.value	= ""
	End if	

	if strCostCd = "" then
		strCostCd = "%"
		frm1.txtCostNm1.value = ""
	End if	

	if strFromCostCd = "" then
		strFromCostCd = "%"
		frm1.txtCostNm2.value = ""
	End if	

	if strDstbCostCd = "" then
		strDstbCostCd = "%"
		frm1.txtCostNm3.value = ""
	End if	

	if strAcctCd = "" then
		strAcctCd = "%"
		frm1.txtAcctNm.value = ""
	End if	

	if strItemCd = "" then
		strItemCd = "%"
		frm1.txtItemNm.value = ""
	End if	



    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001

    strVal = strVal & "&txtFromYYYYMM=" & strFromyyyymm
	strVal = strVal & "&txtPlantCd=" & strPlant
    strVal = strVal & "&txtItemAcctCd=" & stritemacct
    strVal = strVal & "&txtCostCd1=" & strCostCd
    strVal = strVal & "&txtCostCd2=" & strFromCostCd
    strVal = strVal & "&txtCostCd3=" & strDstbCostCd        
    strVal = strVal & "&txtAcctCd=" & strAcctCd        
    strVal = strVal & "&txtItemCd=" & strItemCd   
    	
    Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

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
<TABLE CLASS="BatchTB3" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별배부내역출력</font></td>
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
			<TABLE <%=LR_SPACE_TYPE_60%>>
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">품목기준 집계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">품목기준 상세</LABEL></SPAN>
								</TD>
								
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt3" VALUE="N" tag="25"><LABEL FOR="PrintOpt3">발생코스트센터기준 집계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt4" VALUE="N" tag="25"><LABEL FOR="PrintOpt4">발생코스트센터기준 상세</LABEL></SPAN>
								</TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/c3606oa1_fpDateTime1_txtFromYyyymm.js'></script>								
								</TD>								
							</TR>
							<TR>	
								<TD CLASS="TD5">공장</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtPlantCd.value,1)">
									 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">품목계정</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtitemAcct" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtitemAcct.value,2)">
									 <INPUT TYPE=TEXT ID="txtItemAcctNm" NAME="txtItemAcctNm" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">제조코스트센터</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="제조코스트센터"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCostCd1.value,3)">
									 <INPUT TYPE=TEXT ID="txtCostNm1" NAME="txtCostNm1" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">발생코스트센터</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostCd2" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="발생코스트센터"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCostCd2.value,4)">
									 <INPUT TYPE=TEXT ID="txtCostNm2" NAME="txtCostNm2" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">배부코스트센터</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostCd3" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="배부코스트센터"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd3" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCostCd3.value,5)">
									 <INPUT TYPE=TEXT ID="txtCostNm3" NAME="txtCostNm3" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">계정코드</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="계정코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCd.value,6)">
									 <INPUT TYPE=TEXT ID="txtAcctNm" NAME="txtAcctNm" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">품목코드</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtItemCd.value,7)">
									 <INPUT TYPE=TEXT ID="txtItemNm" NAME="txtItemNm" SIZE=30 tag="14X">
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
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
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
	<INPUT TYPE="HIDDEN" NAME="txtSpId" tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

