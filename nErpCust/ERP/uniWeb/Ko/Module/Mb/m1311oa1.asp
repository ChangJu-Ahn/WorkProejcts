<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m1311oa1
'*  4. Program Name         : 외주PL출력 
'*  5. Program Desc         : 외주PL출력(외주처별, 품목별)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/06/05
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

'================================================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0        
         
End Sub
'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %>
End Sub
'================================================================================================================================
Function OpenBpCd1()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "외주처"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd1.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "외주처"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "외주처"						
    arrHeader(1) = "외주처명"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBpCd1.Value = arrRet(0)
		frm1.txtBpNm1.Value = arrRet(1)
		frm1.txtBpCd1.focus
		Set gActiveElement = document.activeElement
	End If		
		
End Function
'================================================================================================================================
Function OpenBpCd2()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "외주처"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd2.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "외주처"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "외주처"						
    arrHeader(1) = "외주처명"					
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd2.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBpCd2.Value = arrRet(0)
		frm1.txtBpNm2.Value = arrRet(1)
		frm1.txtBpCd2.focus
		Set gActiveElement = document.activeElement
	End If		
		
End Function
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
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
	frm1.txtItemCd1.value=""
	frm1.txtItemNm1.value=""
	frm1.txtItemCd2.value=""
	frm1.txtItemNm2.value=""
	
End Function
'================================================================================================================================
Function OpenItemCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd1.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	    
	iCalledAspName = AskPRAspName("B1B11PA3")					
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd1.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd1.Value = arrRet(0)
		frm1.txtItemNm1.Value = arrRet(1)
		frm1.txtItemCd1.focus
		Set gActiveElement = document.activeElement
	End If	

End Function
'================================================================================================================================
Function OpenItemCd2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName	

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd2.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec
	    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd2.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd2.Value = arrRet(0)
		frm1.txtItemNm2.Value = arrRet(1)
		frm1.txtItemCd2.focus
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
	Dim intCnt
	dim var1,var2,var3,var4,var5
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
	If UCase(frm1.txtBpCd1.value) = "" Then
		var1 = "%"
	Else
		var1= UCase(frm1.txtBpCd1.value)
	End If
	
	If UCase(frm1.txtBpCd2.value) = "" Then
		var2 = "ZZZZZZZZZZ"
	Else
		var2= UCase(frm1.txtBpCd2.value)
	End If	
	
	If UCase(frm1.txtItemCd1.value) = "" Then
		var3 = "%"
	Else
		var3= UCase(frm1.txtItemCd1.value)
	End If
	
	If UCase(frm1.txtItemCd2.value) = "" Then
		var4 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var4= UCase(frm1.txtItemCd2.value)
	End If
	
	If UCase(frm1.txtPlantCd.value) = "" Then
		var5 = "%"
	Else
		var5= UCase(frm1.txtPlantCd.value)
	End If
	
	strUrl = strUrl & "frsupplier|"		& var1
	strUrl = strUrl & "|tosupplier|"	& var2
	strUrl = strUrl & "|fritem|"		& var3
	strUrl = strUrl & "|toitem|"		& var4
	strUrl = strUrl & "|plant|"			& var5		

	if frm1.rdoflg1.checked = True then
		ObjName = AskEBDocumentName("m1311oa1","ebr")
		Call FncEBRprint(EBAction, ObjName, strUrl)
	else
		ObjName = AskEBDocumentName("m1311oa2","ebr")
		Call FncEBRprint(EBAction, ObjName, strUrl)
	End if
		
	Call BtnDisabled(0)	
		
End Function
'================================================================================================================================
sub btnPreview() 
	On Error Resume Next                       
    Err.Clear                                                       
    
    Dim strVal
    dim var1,var2,var3,var4,var5	
	dim strUrl
	dim arrParam, arrField, arrHeader

   
	If UCase(frm1.txtBpCd1.value) = "" Then
		var1 = "%"
	Else
		var1= UCase(frm1.txtBpCd1.value)
	End If
	
	If UCase(frm1.txtBpCd2.value) = "" Then
		var2 = "ZZZZZZZZZZ"
	Else
		var2= UCase(frm1.txtBpCd2.value)
	End If	
	
	If UCase(frm1.txtItemCd1.value) = "" Then
		var3 = "%"
	Else
		var3= UCase(frm1.txtItemCd1.value)
	End If
	
	If UCase(frm1.txtItemCd2.value) = "" Then
		var4 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var4= UCase(frm1.txtItemCd2.value)
	End If
	
	If UCase(frm1.txtPlantCd.value) = "" Then
		var5 = "%"
	Else
		var5= UCase(frm1.txtPlantCd.value)
	End If
	
			
	strUrl = strUrl & "frsupplier|"		& var1
	strUrl = strUrl & "|tosupplier|"	& var2
	strUrl = strUrl & "|fritem|"		& var3
	strUrl = strUrl & "|toitem|"		& var4
	strUrl = strUrl & "|plant|"			& var5			


	if frm1.rdoflg1.checked = True then		
		ObjName = AskEBDocumentName("m1311oa1","ebr")
		Call FncEBRPreview(ObjName, strUrl)
	else		
		ObjName = AskEBDocumentName("m1311oa2","ebr")
		Call FncEBRPreview(ObjName, strUrl)
	End if
	
	Call BtnDisabled(0)	
		
End Sub
'================================================================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'================================================================================================================================
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>외주P/L</font></td>
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
								<TD CLASS="TD5" NOWRAP>발행유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="발행유형" NAME="rdoflg" id = "rdoflg1" Value="Y"  checked tag="12"><label for="rdoflg1">&nbsp;외주처별&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="발행유형" NAME="rdoflg" id = "rdoflg2" Value="N"  tag="12"><label for="rdoflg2">&nbsp;품목별&nbsp;</label>
								</TD>									
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>외주처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd1"   SIZE=10 MAXLENGTH=10 ALT="외주처" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd1()">
													   <INPUT TYPE=TEXT NAME="txtBpNm1" SIZE=20 MAXLENGTH=18 ALT="외주처" tag="14"> ~</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP> <INPUT TYPE=TEXT NAME="txtBpCd2"  SIZE=10 MAXLENGTH=10 ALT="외주처" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd2()">
													   <INPUT TYPE=TEXT NAME="txtBpNm2" SIZE=20 MAXLENGTH=18 ALT="외주처" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 ALT="공장" tag="14"></TD>
																						   
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=10 MAXLENGTH=18 ALT="품목" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd1()">
													   <INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=20 MAXLENGTH=18 ALT="품목" tag="14"> ~</TD>
							</TR> 
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>						   
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=10 MAXLENGTH=18 ALT="품목" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd2()">
													   <INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=20 MAXLENGTH=18 ALT="품목" tag="14"></TD>					   
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" tabindex = -1>
    <input type="hidden" name="dbname" tabindex = -1>
    <input type="hidden" name="filename" tabindex = -1>
    <input type="hidden" name="condvar" tabindex = -1>
	<input type="hidden" name="date" tabindex = -1>
</FORM>
</BODY>
</HTML>
