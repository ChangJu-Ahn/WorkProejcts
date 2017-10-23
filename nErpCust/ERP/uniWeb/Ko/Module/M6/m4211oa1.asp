
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
'*  8. Modified date(Last)  : 2003/05/29
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
<!-- '******************************************  1.1 Inc 선언   **************************************** !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ===================================== !-->
<!--meta http-equiv="Content-type" content="text/html; charset=euc-kr"-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ==================================== !-->

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

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

<!-- '==========================================  2.1.1 InitVariables()  ==================================!-->
Sub InitVariables()
	
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
End Sub
<!-- '==========================================  2.2.1 SetDefaultVal()  ================================ !-->
Sub SetDefaultVal()	
	frm1.txtFrDt.Text	= StartDate
    frm1.txtToDt.Text	= EndDate
    frm1.txtSupplierCd.focus
    Set gActiveElement = document.activeElement
End Sub

<!--'====================================================================================== !-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "OA") %>
End Sub

<!-- '------------------------------------------  OpenSupplier()  --------------------------------------- !-->

Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래선"			
	arrParam(1) = "B_Biz_Partner"		
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	<%' Name Cindition%>
	
	arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "			
	arrParam(5) = "거래선"		
	
	arrField(0) = "BP_Cd"			
	arrField(1) = "BP_NM"			

	arrHeader(0) = "거래선"		
	arrHeader(1) = "거래선명"	
    
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
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement	
		lgBlnFlgChgValue = True
	End If	
End Function

<!-- '==========================================  3.1.1 Form_Load()  ==================================== !-->
Sub Form_Load()
   
    Call LoadInfTB19029             
    Call ggoOper.LockField(Document, "N")      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                         
    
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")
    
    frm1.txtSupplierCd.focus 
    Set gActiveElement = document.activeElement
End Sub
<!--
'========================================  Form_QueryUnload()  ========================================
-->
Sub Form_QueryUnload(Cancel, UnloadMode)
  
End Sub

<!--
'==========================================================================================
'   Event Name : txtFrDt  , txtToDt	 
'==========================================================================================
-->
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrDt.focus
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.focus
	End if
End Sub

<!-- '*******************************  5.1 Toolbar(Main)에서 호출되는 Function ***************************** !-->
Function fncQuery()
	Call BtnPreview()
End Function
Function fncSave()

End Function
<!--
'======================================  FncPrint()  =========================================
-->
Function FncPrint() 
	Call parent.FncPrint()
End Function
<!--
'======================================  FncFind()  =========================================
-->
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False) 
End Function
'======================================  ChkKeyField()  =========================================
Function ChkKeyField()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	strWhere = " BP_CD =  " & FilterVar(frm1.txtSupplierCd.value, "''", "S") & "  AND Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") "
	
	Call CommonQueryRs(" BP_NM "," B_Biz_Partner ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","거래선","X")
		frm1.txtSupplierNm.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	
	frm1.txtSupplierNm.value = strDataNm(0)
	
End Function

<!--'=====================  FncBtnPrint()  =============================================================!-->
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
	dim var1,var2,var3,var4
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If

	IF ChkKeyField() = False Then 
		frm1.txtSupplierCd.focus
		Exit Function
    End if
	
	With frm1
	     If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","통관기간","X")			
			Exit Function
		End if   
	End with
	
	On Error Resume Next                  
	
	lngPos = 0
	
	var1 = UCase(frm1.txtSupplierCd.value)
	var2 =  UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	var3 =  UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType) 'uniCdate(frm1.txtToDt.text)
            		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
	strUrl = strUrl & "bp_cd|" & var1 & "|fr_dt|" & var2 & "|to_dt|" & var3 
'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------	
	ObjName = AskEBDocumentName("m4211oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)		
		
End Function
<!--'=====================  BtnPreview()  =============================================================!-->
Function BtnPreview() 
	dim var1,var2,var3
	dim strUrl
	dim arrParam, arrField, arrHeader
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	IF ChkKeyField() = False Then 
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
    End if
    
    With frm1
         If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
		'if (UniCdate(.txtFrDt.text) > UniCdate(.txtToDt.text)) And trim(.txtFrDt.text) <> "" And trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003","X","통관기간","X")			
			Exit Function
		End if   
	End with

	var1 = UCase(frm1.txtSupplierCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtFrDt.text)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
			
	strUrl = strUrl & "bp_cd|" & var1 & "|fr_dt|" & var2 & "|to_dt|" & var3 
	
	'call FncEBRPreview("m4211oa1.ebr", strUrl)
	
	'2002/12/10
	'표준변경 
	'''''''''''''''''''''''''''''''''''''''''''''
	ObjName = AskEBDocumentName("m4211oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)
	
	Call BtnDisabled(0)	

		
End Function
<!--'=====================  FncExit()  =============================================================!-->
Function FncExit()
	    FncExit = True
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관관리대장</font></td>
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
								<TD CLASS="TD5" NOWRAP>거래선</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 ALT="거래선" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
													   <INPUT TYPE=TEXT NAME="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="거래선" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>통관기간</TD>
								<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m4211oa1_fpDateTime1_txtFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m4211oa1_fpDateTime2_txtToDt.js'></script>
												</td>
											<tr>
										</table>
							         </TD>
	                            </TR>
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
