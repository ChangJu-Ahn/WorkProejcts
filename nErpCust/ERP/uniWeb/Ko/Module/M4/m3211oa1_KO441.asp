<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : Offer Sheet
'*  5. Program Desc         : Offer Sheet
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/20
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Lee Eun Hee
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
<!-- 
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)  


Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
      
Dim lblnWinEvent
Dim IsOpenPop          

<!-- '==========================================  2.1.1 InitVariables()  ======================================
-->
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0        
End Sub

<!-- '==========================================  2.2.1 SetDefaultVal()  ========================================
-->
Sub SetDefaultVal()
	frm1.txtFrDt.Text	=  StartDate
	frm1.txtToDt.Text	=  EndDate
End Sub

<!--'=======================================  LoadInfTB19029()   ====================================-->
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "OA") %>
End Sub

<!-- '------------------------------------------  OpenPayterm()  -------------------------------------------------!-->
Function OpenPayterm()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "결제방법"	
	arrParam(1) = "B_minor,b_configuration"				
	arrParam(2) = Trim(frm1.txtPaytermCd.Value)
	arrParam(3) = ""
	arrParam(4) = "b_minor.Major_Cd=" & FilterVar("B9004", "''", "S") & "" 
	arrParam(4) =	arrParam(4) & " and b_minor.minor_cd=b_configuration.minor_cd" 
	arrParam(4) =	arrParam(4) & " and (b_configuration.REFERENCE = " & FilterVar("L", "''", "S") & "  or b_configuration.REFERENCE = " & FilterVar("M", "''", "S") & " )"		
	arrParam(5) = "결제방법"			
	
    arrField(0) = "b_minor.minor_cd"	
    arrField(1) = "b_minor.minor_nm"	
    
    arrHeader(0) = "결제방법"		
    arrHeader(1) = "결제방법명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPaytermCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPaytermCd.Value= arrRet(0)		
		frm1.txtPaytermNm.value= arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtPaytermCd.focus
		Set gActiveElement = document.activeElement
	End If	
	
End Function

<!-- '==========================================  3.1.1 Form_Load()  ======================================-->
Sub Form_Load()
    
    Call LoadInfTB19029                                                     
    Call ggoOper.LockField(Document, "N")                                   
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)

    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables                                                      
    Call SetToolbar("1000000000001111")										
    
    frm1.txtPaytermCd.focus 
    Set gActiveElement = document.activeElement
    
End Sub
<!--
'=====================================  Form_QueryUnload()  ===========================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

<!--
'==========================================================================================
'   Event Name : txtFrDt  	 
'==========================================================================================
-->
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtFrDt.focus
	End if
End Sub
<!--
'==========================================================================================
'   Event Name : txtToDt  	 
'==========================================================================================
-->
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
   		frm1.txtToDt.focus
	End if
End Sub
<!-- '*******************************  5.1 Toolbar(Main)에서 호출되는 Function ************************ -->
Function fncQuery()
Set gActiveElement = document.activeElement
End Function
Function fncSave()
Set gActiveElement = document.activeElement
End Function
<!--
'=================================  FncPrint()   ==================================================
-->
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
<!--
'=================================  FncFind()   ==================================================
-->
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)   
    Set gActiveElement = document.activeElement
End Function
<!--
'=================================  FncBtnPrint()   ==================================================
-->
Function FncBtnPrint() 
	Dim StrUrl
	dim var1,var2,var3,var4
	Dim ObjName
	    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If

    IF ChkKeyField() = False Then 
		Exit Function
    End if
	
	with frm1
	     If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
		'if (UNICDate(.txtFrDt.text) > Parent.UNICDate(.txtToDt.text)) And trim(.txtFrDt.text) <> "" And trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","개설기간", "X")			
			Exit Function
		End if   
	End with
	
	On Error Resume Next                                                    
	
	var1 = Trim(frm1.txtPaytermCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType)'UNICDate(frm1.txtFrDt.text)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'UNICDate(frm1.txtToDt.text)

	strUrl = strUrl & "PAY_METHOD|" & var1 & "|AFromOpenDt|" & var2 & "|AToOpenDt|" & var3

    If lgBACd<>"" Then
        strUrl = strUrl & "|FR_BIZ_AREA|" & lgBACd 
        strUrl = strUrl & "|TO_BIZ_AREA|" & lgBACd 
    Else
        strUrl = strUrl & "|FR_BIZ_AREA|" & "" 
        strUrl = strUrl & "|TO_BIZ_AREA|" & "ZZZZZZZZZZ" 
    End If

    If lgPGCd<>"" Then
        strUrl = strUrl & "|FR_PUR_GRP|" & lgPGCd 
        strUrl = strUrl & "|TO_PUR_GRP|" & lgPGCd 
    Else
        strUrl = strUrl & "|FR_PUR_GRP|" & "" 
        strUrl = strUrl & "|TO_PUR_GRP|" & "ZZZZZZZZZZ" 
    End If

    If lgPOCd<>"" Then
        strUrl = strUrl & "|FR_PUR_ORG|" & lgPOCd 
        strUrl = strUrl & "|TO_PUR_ORG|" & lgPOCd 
    Else
        strUrl = strUrl & "|FR_PUR_ORG|" & "" 
        strUrl = strUrl & "|TO_PUR_ORG|" & "ZZZZZZZZZZ" 
    End If
	
	
'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------
	'call FncEBRprint(EBAction, "m3211oa1.ebr", strUrl)
	ObjName = AskEBDocumentName("m3211oa1_KO441","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------
	Set gActiveElement = document.activeElement
End Function
<!--
'=================================  BtnPreview()   ==================================================
-->
Function BtnPreview() 
'On Error Resume Next                                                    '☜: Protect system from crashing
    Dim ObjName
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

    IF ChkKeyField() = False Then 
		Exit Function
    End If
	
	With frm1
	     If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
		'if (UNICDate(.txtFrDt.text) > Parent.UNICDate(.txtToDt.text)) And trim(.txtFrDt.text) <> "" And trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","개설기간", "X")			
			Exit Function
		End If   
	End With

	Dim var1,var2,var3
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
		
	var1 = Trim(frm1.txtPaytermCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType)'UNICDate(frm1.txtFrDt.text)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'UNICDate(frm1.txtToDt.text)
	
	strUrl = strUrl & "PAY_METHOD|" & var1 & "|AFromOpenDt|" & var2 & "|AToOpenDt|" & var3 
	
    If lgBACd<>"" Then
        strUrl = strUrl & "|FR_BIZ_AREA|" & lgBACd 
        strUrl = strUrl & "|TO_BIZ_AREA|" & lgBACd 
    Else
        strUrl = strUrl & "|FR_BIZ_AREA|" & "" 
        strUrl = strUrl & "|TO_BIZ_AREA|" & "ZZZZZZZZZZ" 
    End If

    If lgPGCd<>"" Then
        strUrl = strUrl & "|FR_PUR_GRP|" & lgPGCd 
        strUrl = strUrl & "|TO_PUR_GRP|" & lgPGCd 
    Else
        strUrl = strUrl & "|FR_PUR_GRP|" & "" 
        strUrl = strUrl & "|TO_PUR_GRP|" & "ZZZZZZZZZZ" 
    End If

    If lgPOCd<>"" Then
        strUrl = strUrl & "|FR_PUR_ORG|" & lgPOCd 
        strUrl = strUrl & "|TO_PUR_ORG|" & lgPOCd 
    Else
        strUrl = strUrl & "|FR_PUR_ORG|" & "" 
        strUrl = strUrl & "|TO_PUR_ORG|" & "ZZZZZZZZZZ" 
    End If
	
'	call FncEBRPreview("m3211oa1.ebr", strUrl)	
	ObjName = AskEBDocumentName("m3211oa1_KO441","ebr")
	Call FncEBRPreview(ObjName, strUrl)
End Function
<!--
'=================================  FncExit()   ==================================================
-->
Function FncExit()	
    FncExit = True
    Set gActiveElement = document.activeElement
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
Function ChkKeyField()
	
   Dim strDataCd, strDataNm
   Dim strWhere 
   Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
   Err.Clear                                       
	
   ChkKeyField = true

   strWhere = " MAJOR_CD=" & FilterVar("B9004", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtPaytermCd.value, "''", "S") & "  "
	
   Call CommonQueryRs(" MINOR_NM "," B_MINOR ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
   IF Len(lgF0) < 1 Then 
   	Call DisplayMsgBox("17a003","X","결제방법","X")
   	frm1.txtPaytermCd.focus
   	frm1.txtPaytermNm.value = ""
   	ChkKeyField = False
   	Exit Function
   End If
	
   strDataNm = split(lgF0,chr(11))
	
   frm1.txtPaytermNm.value = strDataNm(0)
	
End Function


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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수입LC관리대장</font></td>
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
								<TD CLASS="TD5" NOWRAP>결제방법</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPaytermCd" SIZE=10 MAXLENGTH=2 ALT="결제방법" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayterm()">
													   <INPUT TYPE=TEXT NAME="txtPaytermNm" SIZE=20 MAXLENGTH=18 ALT="결제방법" tag="14X"></TD>					
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>개설기간</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td NOWRAP>
												<script language =javascript src='./js/m3211oa1_fpDateTime1_txtFrDt.js'></script>
											</td>
											<td NOWRAP>
												~
											</td>
											<td NOWRAP>
												<script language =javascript src='./js/m3211oa1_fpDateTime2_txtToDt.js'></script>
											</td>
										</tr>
									</table>
								</TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m3112mb1.asp" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
