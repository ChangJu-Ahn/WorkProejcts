<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 
'*  3. Program ID           : S3211OA1
'*  4. Program Name         : L/C 관리대장 출력 
'*  5. Program Desc         : L/C 관리대장 출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/07/16
'*  8. Modified date(Last)  : 2000/07/16
'*  9. Modifier (First)     : Cho Sung Hyun
'* 10. Modifier (Last)      : Son bum Yeol
'* 11. Comment              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

Const gstrPayTermsMajor = "B9004"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          
Dim gblnWinEvent

'===========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           'initializes Group View Size  
         
End Sub

'===========================================================================================================
Sub SetDefaultVal()
	frm1.txtLCFromDt.text = StartDate
	frm1.txtLCToDt.text = EndDate
	frm1.txtPayMeth.focus   
End Sub

'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=========================================================================================================== 
Function OpenConPop()
		
		Dim arrRet
		Dim arrParam(5), arrField(6), arrHeader(6)

		If gblnWinEvent = True Then Exit Function
		gblnWinEvent = True
		
		arrParam(0) = "수출L/C관리대장"								
		arrParam(1) = "b_minor,b_configuration"						
		arrParam(2) = Trim(frm1.txtPayMeth.Value)					
		arrParam(3) = ""											
		arrParam(4) = "b_minor.MINOR_CD = b_configuration.MINOR_CD AND b_minor.MAJOR_CD = " & FilterVar(gstrPayTermsMajor, "''", "S") & " AND b_configuration.REFERENCE = " & FilterVar("M", "''", "S") & " "
		arrParam(5) = "결제방법"													

		arrField(0) = "b_minor.Minor_CD"														
		arrField(1) = "b_minor.Minor_NM"														

		arrHeader(0) = "결제방법"													
		arrHeader(1) = "결제방법명"													
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		gblnWinEvent = False
		
		frm1.txtPayMeth.focus 

		If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetConPop(arrRet)
		End If
	End Function
 
'===========================================================================================================
Function SetConPop(Byval arrRet)
	With frm1	
		.txtPayMeth.Value	= arrRet(0)
		.txtPayMethNm.Value	= arrRet(1)
		.txtPayMeth.focus
	End With
End Function

'===========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
	Call InitVariables														
    <% '----------  Coding part  -------------------------------------------------------------%>
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										
End Sub


'===========================================================================================================
Sub txtLCFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLCFromDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtLCFromDt.Focus
    End If
End Sub

Sub txtLCToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLCToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtLCToDt.Focus
    End If
End Sub

'===========================================================================================================
Function FncPrint() 
   Call parent.FncPrint()
End Function

'===========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)   
End Function

'===========================================================================================================
Function FncQuery() 
    FncQuery = true
End Function

'===========================================================================================================
Function BtnPrint() 
	Dim strUrl     
	
	If ValidDateCheck(frm1.txtLCFromDt, frm1.txtLCToDt) = False Then Exit Function
    
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
	Dim var1, var2, var3
	
	If UCase(frm1.txtPayMeth.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtPayMeth.value)), "" ,  "SNM")
	End If

		var2 = UniConvDateToYYYYMMDD(frm1.txtLCFromDt.text,Parent.gDateFormat,Parent.gServerDateType)

		var3 = UniConvDateToYYYYMMDD(frm1.txtLCToDt.text,Parent.gDateFormat,Parent.gServerDateType)
	
	strUrl = strUrl & "PayMeth|" & var1 & "|LCFromDt|" & var2 & "|LCToDt|" & var3 

	OBjName = AskEBDocumentName("s3211oa1","ebr")    
	Call FncEBRprint(EBAction, OBjName, strUrl)
	
End Function


'===========================================================================================================
Function BtnPreview() 
    
	If ValidDateCheck(frm1.txtLCFromDt, frm1.txtLCToDt) = False Then Exit Function
    
    If Not chkField(Document, "1") Then					
       Exit Function
    End If

	Dim var1, var2, var3
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
		
	If UCase(frm1.txtPayMeth.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtPayMeth.value)), "" ,  "SNM")
	End If

		var2 = UniConvDateToYYYYMMDD(frm1.txtLCFromDt.text,Parent.gDateFormat,Parent.gServerDateType)

		var3 = UniConvDateToYYYYMMDD(frm1.txtLCToDt.text,Parent.gDateFormat,Parent.gServerDateType)
		
	
	strUrl = strUrl & "PayMeth|" & var1 & "|LCFromDt|" & var2 & "|LCToDt|" & var3 

	OBjName = AskEBDocumentName("s3211oa1","ebr")    
	Call FncEBRPreview(OBjName, strUrl)			
End Function

'===========================================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C 관리대장</font></td>
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
									<TD CLASS=TD5 NOWRAP>결제방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayMeth" ALT="결제방법" SIZE=10 MAXLENGTH=5 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPop" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPop">&nbsp;
									                     <INPUT TYPE=TEXT NAME="txtPayMethNm" SIZE=20 MAXLENGTH=10 TAG="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>개설일</TD>						
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s3211oa1_fpDateTime1_txtLCFromDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s3211oa1_fpDateTime2_txtLCToDt.js'></script>
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
						    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
						</TD>

					</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX="-1">
    <input type="hidden" name="dbname" TABINDEX="-1">
    <input type="hidden" name="filename" TABINDEX="-1">
    <input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>
