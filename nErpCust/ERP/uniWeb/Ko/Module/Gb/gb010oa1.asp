
<%@ LANGUAGE="VBSCRIPT" %>
<!--
**********************************************************************************************
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        : 
*  3. Program ID           : GB010MA1
*  4. Program Name         : 사업부별손익추이표 출력 
*  5. Program Desc		   : 사업부별손익추이표 출력 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/11/27
*  8. Modified date(Last)  : 2001/11/27
*  9. Modifier (First)     : Kwon Ki Soo
* 10. Modifier (Last)      : Kwon Ki Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim IsOpenPop
Dim lgOldRow

'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
End Sub

'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date
	
	frm1.txtYyyymm.focus()    
	frm1.txtYyyymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
	Call ggoOper.FormatDate(frm1.txtYyyymm, parent.gDateFormat, 3)
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("P", "G", "NOCOOKIE","PA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitVariables 

    Call SetDefaultVal
    Call SetToolbar("1000000000000111")
   
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================
Function FncQuery()
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncQuery = True                                                              '☜: Processing is OK

End Function

'========================================================================================
Function txtGrade_onKeyPress(Key)    
    
    frm1.action = "../../blank.htm"       
    
End Function
	
'=======================================================================================================
Function FncBtnPrint() 
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim var1,var2
    Dim strYear, strMonth, strDay
    	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field%>
       Exit Function
    End If
	
	StrEbrFile = "GB010oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" BIZ_UNIT_NM "," B_BIZ_UNIT "," BIZ_UNIT_CD =   " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    		
	End if	
	
    '--출력조건을 지정하는 부분 수정 
	
	condvar = "YYYY|" & var1
	condvar = condvar & "|BIZ_UNIT_CD|" & var2		
	   
	Call FncEBRPrint(EBAction,ObjName,condvar)
	
End Function

'========================================================================================
Function FncBtnPreview()
On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field%>
       Exit Function
    End If
	
	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	Dim var1, var2	
    Dim strYear, strMonth, strDay
        	
	StrEbrFile = "gb010oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
	var1 = Trim(frm1.txtYyyymm.year)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
	    Call CommonQueryRs(" BIZ_UNIT_NM "," B_BIZ_UNIT "," BIZ_UNIT_CD =   " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    	
	End if	
					
	strUrl = strUrl & parent.gEbEnginePath
	strUrl = strUrl & "ExecuteWinReport?"
	strUrl = strUrl & "uname=" & parent.gEbUserName & "&"
	strUrl = strUrl & "dbname=" & parent.gEbDbName & "&"
	strUrl = strUrl & "filename=" & parent.gEbPkgRptPath & "\" & parent.gLang & "\ebr\" & StrEbrFile
	
	condvar = "YYYY|" & var1
	condvar = condvar & "|BIZ_UNIT_CD|" & var2
	
	
	Call FncEBRPreview(ObjName,condvar)
	
End Function

'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================================
Function OpenPopUp()
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


			arrParam(0) = "사업부"					' 팝업 명칭 
			arrParam(1) = "B_BIZ_UNIT"						' TABLE 명칭 
			arrParam(2) = UCase(Trim(frm1.txtFr_dept_cd.Value))	' Code Condition
			arrParam(3) = ""							' Name Cindition
			'arrParam(4) = ""	
			arrParam(5) = "사업부"			
	
   			arrField(0) = "BIZ_UNIT_CD"	     				' Field명(0)
			arrField(1) = "BIZ_UNIT_NM"			    		' Field명(1)
		
			arrHeader(0) = "사업부"					' Header명(0)
			arrHeader(1) = "사업부명"				' Header명(1)
    
    
	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtFr_dept_cd.focus
		Exit Function
	Else
	   frm1.txtFr_dept_cd.focus	
	   Frm1.txtFr_dept_cd.value = arrRet(0)
	   frm1.txtFr_dept_nm.value = arrRet(1)
	End If	

End Function

'========================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사업부별 손익추이표</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
								<TD CLASS=TD5  NOWRAP>대상년월</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/gb010oa1_fpLoanDtFr_txtYyyymm.js'></script>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>사업부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="사업부코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp()">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="사업부코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU"></TD>
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()">미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()">인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=20><IFRAME type=hidden NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>
