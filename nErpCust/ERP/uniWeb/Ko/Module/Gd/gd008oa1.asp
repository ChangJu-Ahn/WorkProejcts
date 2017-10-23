
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
*  1. Module Name       : P&L Mgmt.
*  2. Function Name     : 
*  3. Program ID        : gc0070a1
*  4. Program Name      : 영업조직 손익추이표출력 
*  5. Program Desc      : 영업조직 손익추이표출력 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/03/08
*  8. Modified date(Last)  : 2002/03/08
*  9. Modifier (First)     : Jang Yoon Ki
* 10. Modifier (Last)      : Jang Yoon Ki
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>

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
	Dim StartDate
	StartDate = "<%=GetSvrDate%>"
	
    frm1.txtYyyymm.focus()    
	frm1.txtYyyymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtYyyymm, parent.gDateFormat, 3) 
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("P", "G", "NOCOOKIE", "PA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
 	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitVariables 
    
'    Call ggoOper.FormatDate(frm1.txtpay_yymm, parent.gDateFormat, 2)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    
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
    	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

	StrEbrFile = "gd008oa1"
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
			
'	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" SALES_ORG_NM "," B_SALES_ORG "," SALES_ORG =  " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    		
	End if	
	
	'--출력조건을 지정하는 부분 수정 - 끝 %>
	
	condvar = "YYYY|" & var1
	condvar = condvar & "|SALES_ORG|" & var2
	
	Call FncEBRPrint(EBAction,ObjName,condvar)	

End Function


'========================================================================================
Function FncBtnPreview()
'On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
	
	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
	Dim var1, var2	
    Dim strYear, strMonth, strDay
         	
	StrEbrFile = "gd008oa1"
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
		
'	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	
	var1 = Trim(frm1.txtYyyymm.Text)
	var2 = Trim(UCase(frm1.txtFr_dept_cd.value))	
	
    if var2 = "" then
		var2 = "%"
		frm1.txtFr_dept_nm.value = ""
	else
		Call CommonQueryRs(" SALES_ORG_NM "," B_SALES_ORG "," SALES_ORG =  " & FilterVar(frm1.txtFr_dept_cd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  frm1.txtFr_dept_nm.value = ""
		else   
		  frm1.txtFr_dept_nm.value = Trim(Replace(lgF0,Chr(11),""))
		end if    	    	
	End if	
					
	condvar = "YYYY|" & var1
	condvar = condvar & "|SALES_ORG|" & var2
	
	Call FncEBRPreview(ObjName,condvar)

End Function

'========================================================================================================
Function FncPrint()
	Call FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
Function FncFind() 
	Call FncFind(parent.C_SINGLE, False)
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


			arrParam(0) = "영업조직"					' 팝업 명칭 
			arrParam(1) = "B_SALES_ORG"						' TABLE 명칭 
			arrParam(2) = UCase(Trim(frm1.txtFr_dept_cd.Value))	' Code Condition
			arrParam(3) = ""							' Name Cindition
			'arrParam(4) = ""	
			arrParam(5) = "영업조직"			
	
   			arrField(0) = "SALES_ORG"	     				' Field명(0)
			arrField(1) = "SALES_ORG_NM"			    		' Field명(1)
		
			arrHeader(0) = "영업조직"					' Header명(0)
			arrHeader(1) = "영업조직명"				' Header명(1)
    
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>영업조직별 손익추이표</font></td>
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
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/gd008oa1_fpLoanDtFr_txtYyyymm.js'></script>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>영업조직</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="영업조직코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp()">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="영업조직코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU"></TD>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag=1>인쇄</BUTTON>

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

