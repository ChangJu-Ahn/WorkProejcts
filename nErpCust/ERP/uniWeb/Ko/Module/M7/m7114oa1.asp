<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : Offer Sheet
'*  5. Program Desc         : Offer Sheet
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/06/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit	

<!-- #Include file="../../inc/lgvariables.inc" -->
       
Dim lblnWinEvent
Dim IsOpenPop          
Dim lblnFlag

Dim StartDate, EndDate
StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat) 
EndDate   = UniConvDateAToB("<%=GetSvrDate%>"  ,Parent.gServerDateFormat,Parent.gDateFormat)
'================================================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                 
    lgBlnFlgChgValue = False                  
    lgIntGrpCount = 0                         

End Sub
'================================================================================================================================
Sub SetDefaultVal()

	frm1.txtFrPoDt.Text	= StartDate
    frm1.txtToPoDt.Text	= EndDate
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
	Set gActiveElement = document.activeElement
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "OA") %>
End Sub
'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"			
	arrParam(1) = "B_Biz_Partner"		
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""
	arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "		
	arrParam(5) = "공급처"		
	
	arrField(0) = "BP_Cd"			
	arrField(1) = "BP_NM"			

	arrHeader(0) = "공급처"		
	arrHeader(1) = "공급처명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'================================================================================================================================
Function OpenPoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주형태"					
	arrParam(1) = "M_CONFIG_PROCESS"			
	arrParam(2) = Trim(frm1.txtPoType.Value)	
	arrParam(3) = ""
	arrParam(4) = ""							
	arrParam(5) = "발주형태"					
	
    arrField(0) = "PO_TYPE_CD"						
    arrField(1) = "PO_TYPE_NM"						
        
    arrHeader(0) = "발주형태"					
    arrHeader(1) = "발주형태명"					
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoType.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoType.Value = arrRet(0)
		frm1.txtPoTypeNm.Value = arrRet(1)
		frm1.txtPoType.focus
		Set gActiveElement = document.activeElement
	End If	
End Function
'================================================================================================================================
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
		Set gActiveElement = document.activeElement
	End If	

End Function 
'================================================================================================================================
Sub Form_Load()
  
    Call LoadInfTB19029             
    Call ggoOper.LockField(Document, "N")  
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                     
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")		
    frm1.txtPlantCd.focus 
    Set gActiveElement = document.activeElement
End Sub
'================================================================================================================================
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtFrPoDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtToPoDt.Focus
	End if
End Sub
'================================================================================================================================
 Function ChkKeyField()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	strWhere = " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
	
	Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","공장","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	
	frm1.txtPlantNm.value = strDataNm(0)
End Function
'================================================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)   
    Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
	Dim var1,var2,var3,var4, var5, var6
    
    On Error Resume Next 	
    Err.Clear
    
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
    End if
    
	With frm1
	     If CompareDateByFormat(.txtFrPoDt.text,.txtToPoDt.text,.txtFrPoDt.Alt,.txtToPoDt.Alt, _
                   "970025",.txtFrPoDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrPoDt.text) <> "" And Trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","발주일","X")			
			Exit Function
		End if   
	End with
	
	lngPos = 0
	
	var1 = UCase(frm1.txtPlantCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrPoDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToPoDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	
	If UCase(frm1.txtSupplierCd.value) = "" Then
		var4 = "%"
	Else
		var4 = UCase(frm1.txtSupplierCd.value)
	End If
	
	If UCase(frm1.txtPoType.value) = "" Then
		var5 = "%"
	Else
		var5 = UCase(frm1.txtPoType.value)
	End If
	
	If UCase(frm1.txtPurGrpCd.value) = "" Then
		var6 = "%"
	Else
		var6 = UCase(frm1.txtPurGrpCd.value)
	End If
		     		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
	strUrl = strUrl & "plant|" & var1
	strUrl = strUrl & "|fr_dt|" & var2
	strUrl = strUrl & "|to_dt|" & var3
	strUrl = strUrl & "|bp_cd|" & var4
	strUrl = strUrl & "|Po_type|" & var5
	strUrl = strUrl & "|Pur_grp|" & var6
	
	If frm1.rdoSelFlg(0).checked = True then								'전체 
		ObjName = AskEBDocumentName("m7114oa1","ebr")
	Else																'미입고 
		objName = AskEBDocumentName("m7114oa2","ebr")
	End If
	
	Call FncEBRprint(EBAction, ObjName, strUrl)
	
	Call BtnDisabled(0)	
	
	Set gActiveElement = document.activeElement	
End Function
'================================================================================================================================
Function BtnPreview() 
	Dim var1,var2,var3,var4,var5,var6
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	
	On Error Resume Next                       
    Err.Clear
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If
	
	IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement	
		Exit Function
    End if
    
	with frm1
	     If CompareDateByFormat(.txtFrPoDt.text,.txtToPoDt.text,.txtFrPoDt.Alt,.txtToPoDt.Alt, _
                   "970025",.txtFrPoDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrPoDt.text) <> "" And Trim(.txtToPoDt.text) <> "" then
			Call DisplayMsgBox("17a003","X","발주일","X")			
			Exit Function
		End if   
	End with

	
	var1 = UCase(frm1.txtPlantCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrPoDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToPoDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	
	If UCase(frm1.txtSupplierCd.value) = "" Then
		var4 = "%"
	Else
		var4 = UCase(frm1.txtSupplierCd.value)
	End If
	
	If UCase(frm1.txtPoType.value) = "" Then
		var5 = "%"
	Else
		var5 = UCase(frm1.txtPoType.value)
	End If
	
	If UCase(frm1.txtPurGrpCd.value) = "" Then
		var6 = "%"
	Else
		var6 = UCase(frm1.txtPurGrpCd.value)
	End If
	
	strUrl = strUrl & "plant|" & var1
	strUrl = strUrl & "|fr_dt|" & var2
	strUrl = strUrl & "|to_dt|" & var3
	strUrl = strUrl & "|bp_cd|" & var4
	strUrl = strUrl & "|po_type|" & var5
	strUrl = strUrl & "|Pur_grp|" & var6
	
	If frm1.rdoSelFlg(0).checked = True then								'전체 
		ObjName = AskEBDocumentName("m7114oa1","ebr")
	Else																	'미입고 
		objName = AskEBDocumentName("m7114oa2","ebr")
	End If
	
	Call FncEBRPreview(ObjName, strUrl)
	
	Call BtnDisabled(0)	
	
	Set gActiveElement = document.activeElement		
End Function
'================================================================================================================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement	
End Function
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>								
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주대비입고현황출력</font></td>
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
					<TD HEIGHT=20 WIDTH=100%>						
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 ALT="공장" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발주일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m7114oa1_fpDateTime1_txtFrPoDt.js'></script> ~
													   <script language =javascript src='./js/m7114oa1_fpDateTime2_txtToPoDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 ALT="공급처" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
													   <INPUT TYPE=TEXT NAME="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="공급처" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발주형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주형태"  NAME="txtPoType" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoType()">
													   <INPUT TYPE=TEXT NAME="txtPoTypeNm" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>					   
							</TR>					   														   
							<TR>
								<TD CLASS="TD5" NOWRAP>선택</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="전체" NAME="rdoSelFlg" ID="rdoSelFlg0" CLASS="RADIO" value = "A" tag="11" checked ><label for="rdoSelFlg0">&nbsp;전체&nbsp;&nbsp;</label>
													   <INPUT TYPE=radio AlT="미입고" NAME="rdoSelFlg" ID="rdoSelFlg1" CLASS="RADIO" value = "N" tag="11"><label for="rdoSelFlg1">&nbsp;미입고&nbsp;&nbsp;</label></TD>
							</TR>								
						</TABLE>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD Width = 10>&nbsp</TD>
					<TD Valign=top>				
					    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;		    
					    <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>		
					</TD>
					<TD Width = 10>&nbsp</TD>
				</TR>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m3112mb1.asp" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
