<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111oa1
'*  4. Program Name         : 공급처별발주현황 
'*  5. Program Desc         : 공급처별발주현황 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
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
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VbScript">
Option Explicit													

Dim StartDate
Dim EndDate
EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)  


Dim lgBlnFlgChgValue           
Dim lgIntFlgMode               
Dim lgIntGrpCount              

'==========================================  1.2.3 Global Variable값 정의  ===============================
       
Dim lblnWinEvent
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                  
    lgBlnFlgChgValue = False                   
    lgIntGrpCount = 0                          
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPoFrDt.Text	= StartDate 
	frm1.txtPoToDt.Text	= EndDate 

'	frm1.txtPlantCd.value=parent.gPlant
'	frm1.txtPlantNm.value=parent.gPlantNm
	
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd1, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd2, "Q") 
		frm1.txtPurGrpCd1.Tag = left(frm1.txtPurGrpCd1.Tag,1) & "4" & mid(frm1.txtPurGrpCd1.Tag,3,len(frm1.txtPurGrpCd1.Tag))
		frm1.txtPurGrpCd2.Tag = left(frm1.txtPurGrpCd2.Tag,1) & "4" & mid(frm1.txtPurGrpCd2.Tag,3,len(frm1.txtPurGrpCd2.Tag))
        frm1.txtPurGrpCd1.value = lgPGCd
        frm1.txtPurGrpCd2.value = lgPGCd
	End If
End Sub

'=====================================  LoadInfTB19029()  ====================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "OA") %>
End Sub

'------------------------------------------  OpenSupplier()  -------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"			
	arrParam(1) = "B_Biz_Partner"		
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
'	arrParam(3) = Trim(frm1.txtSupplierNm.Value)	
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "		
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
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
		lgBlnFlgChgValue = True
	End If	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
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
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)		
		frm1.txtPlantCd.focus
		lgBlnFlgChgValue = True
	End If	
End Function
Function OpenPurGrpCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd1.className = "protected" Then Exit Function
	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd1.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd1.focus
		Exit Function
	Else
		frm1.txtPurGrpCd1.Value = arrRet(0)
		frm1.txtPurGrpNm1.Value = arrRet(1)
		frm1.txtPurGrpCd1.focus
	End If	
End Function

Function OpenPurGrpCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd2.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd2.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd2.focus
		Exit Function
	Else
		frm1.txtPurGrpCd2.Value = arrRet(0)
		frm1.txtPurGrpNm2.Value = arrRet(1)
		frm1.txtPurGrpCd2.focus
	End If	
End Function  

'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPoToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoToDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	on error resume next
    Call LoadInfTB19029                   
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                    
	Call GetValue_ko441()
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")	

    'frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement
End Sub

'===================================  FncPrint()  ========================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'===================================  FncFind()  ========================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False) 
	Set gActiveElement = document.activeElement
End Function
Function FncQuery() 
End Function
'===================================  FncBtnPrint()  ========================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt

	dim var1,var2,var3,var4,var5
    Dim ObjName
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If

    IF ChkKeyField() = False Then 
		Exit Function
    End if

	On Error Resume Next                  
	
	lngPos = 0

	with frm1
		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			Exit Function
		End if   
	End with		
		
	var1 = FilterVar(Trim(UCase(frm1.txtSupplierCd.value)), "" ,  "SNM") 
	var2 = FilterVar(Trim(UCase(frm1.txtPurGrpCd1.value)), "" ,  "SNM") 
	var3 = FilterVar(Trim(UCase(frm1.txtPurGrpCd2.value)), "" ,  "SNM") 
	var4 = UniConvDateToYYYYMMDD(frm1.txtPoFrDt.Text, Parent.gDateFormat, Parent.gServerDateType)'uniCdate(frm1.txtFrDt.text)
	var5 = UniConvDateToYYYYMMDD(frm1.txtPoToDt.Text, Parent.gDateFormat, Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
	
	strUrl = strUrl & "bpcd|"		& var1 
	strUrl = strUrl & "|pur_grp1|"	& var2
	strUrl = strUrl & "|pur_grp2|"	& var3
	strUrl = strUrl & "|PoFrDt|"	& var4
	strUrl = strUrl & "|PoToDt|"	& var5	
	strUrl = strUrl & "|loc_cur|"	& parent.gCurrency

	ObjName = AskEBDocumentName("m3111oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
	
	Call BtnDisabled(0)
		
	Set gActiveElement = document.activeElement
End Function
'===================================  BtnPreview()  ========================================
Function BtnPreview() 
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

    IF ChkKeyField() = False Then 
		Exit Function
    End if

	dim var1,var2,var3,var4,var5
	
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim ObjName

	with frm1
		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			Exit Function
		End if   
	End with		

	var1 = FilterVar(Trim(UCase(frm1.txtSupplierCd.value)), "" ,  "SNM") 
	var2 = FilterVar(Trim(UCase(frm1.txtPurGrpCd1.value)), "" ,  "SNM") 
	var3 = FilterVar(Trim(UCase(frm1.txtPurGrpCd2.value)), "" ,  "SNM") 
	var4 = UniConvDateToYYYYMMDD(frm1.txtPoFrDt.Text, Parent.gDateFormat, Parent.gServerDateType)'uniCdate(frm1.txtFrDt.text)
	var5 = UniConvDateToYYYYMMDD(frm1.txtPoToDt.Text, Parent.gDateFormat, Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)

	strUrl = strUrl & "bpcd|"		& var1 
	strUrl = strUrl & "|pur_grp1|"	& var2
	strUrl = strUrl & "|pur_grp2|"	& var3
	strUrl = strUrl & "|PoFrDt|"	& var4
	strUrl = strUrl & "|PoToDt|"	& var5
	strUrl = strUrl & "|loc_cur|"	& parent.gCurrency

	ObjName = AskEBDocumentName("m3111oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)
	Call BtnDisabled(0)

	Set gActiveElement = document.activeElement
End Function
'===================================  FncExit()  ========================================
Function FncExit()
    FncExit = True
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	ChkKeyField = true
	
'	strWhere = " PLANT_CD = '" & FilterVar(frm1.txtPlantCd.value, "","SNM") & "' "
'	
'	Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'	
'	IF Len(lgF0) < 1 Then 
'		Call DisplayMsgBox("17a003","X","공장","X")
'		frm1.txtPlantCd.focus 
'		frm1.txtPlantNm.value = ""
'		ChkKeyField = False
'		Exit Function
'	End If
'	
'	strDataNm = split(lgF0,chr(11))
'	frm1.txtPlantNm.value = strDataNm(0)
	
	strWhere = " BP_CD =  " & FilterVar(frm1.txtSupplierCd.value, "''", "S") & "  "
	
	Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","공급처","X")
		frm1.txtSupplierCd.focus 
		frm1.txtSupplierNm.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	frm1.txtSupplierNm.value = strDataNm(0)
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공급처별발주현황</font></td>
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
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSupplierCd"  SIZE=10 MAXLENGTH=10 ALT="공급처" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
													   <INPUT TYPE=TEXT NAME="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="공급처" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrpCd1" SIZE=10 MAXLENGTH=10 ALT="구매그룹" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd1()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm1" SIZE=20 MAXLENGTH=18 ALT="구매그룹" tag="14">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrpCd2" SIZE=10 MAXLENGTH=10 ALT="구매그룹" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd2()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm2" SIZE=20 MAXLENGTH=18 ALT="구매그룹" tag="14"></TD>					   
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발주일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111oa1_fpDateTime2_txtPoFrDt.js'></script>&nbsp;~&nbsp;
														<script language =javascript src='./js/m3111oa1_fpDateTime2_txtPoToDt.js'></script></TD>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex="-1" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
