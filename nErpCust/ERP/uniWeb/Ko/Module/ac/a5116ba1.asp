<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Account Management
'*  3. Program ID           : A5116BA1
'*  4. Program Name         : 결산마감 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->			<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================================================================================

Const BIZ_PGM_ID = "a5116bb1.asp"  

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim IsOpenPop          

'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A", "NOCOOKIE", "BA") %>
End Sub

'========================================================================================================= 
Function fnButtonExec()
    Dim strVal       
    Dim WorkDt
    Dim strYYYYMM1,strYear1,strMonth1,strDay1
    Dim strYYYYMM2,strYear2,strMonth2,strDay2
    Dim strTarget
    Dim intRetCD

  '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then                             '⊙: Check contents area
       Exit Function
    End If

    
    strYYYYMM1 = frm1.txtFromdt.value
	

  
	intRetCd = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")
	If intRetCd = VBNO Then
		Exit Function
	End IF
     
    Call LayerShowHide(1) 
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002
    
    if frm1.Rb_WK1.checked = true then
		strVal = strVal & "&txtRadio=" & "1"
    elseif frm1.Rb_WK2.checked = true then
		strVal = strVal & "&txtRadio=" & "2"
    elseif frm1.Rb_WK3.checked = true then
		strVal = strVal & "&txtRadio=" & "3"
	else
		strVal = strVal & "&txtRadio=" & "4"	
	end if        
	
    strVal = strVal & "&txtFromdt=" & strYYYYMM1
    strVal = strVal & "&txtTodt=" & strYYYYMM1    
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
Function fnButtonExecOk()
	Dim IntRetCD

	IntRetCD =DisplayMsgBox("990000","X","X","X")
    Call ggoOper.ClearField(Document, "2")
    Call LayerShowHide(0)
    Call SetLastCloseMnth
    frm1.Rb_WK1.checked = true
End Function

'========================================================================================================= 
Sub SetDefaultVal()
	Dim strSvrDate
	strSvrDate = "<%=GetSvrDate%>"
		
	Call SetLastCloseMnth

	lgBlnFlgChgValue = False
End Sub


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
   
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call SetDefaultVal
	frm1.txtFromdt.focus 
End Sub


Sub SetLastCloseMnth()
    Dim strSelect, strFrom, strWhere
    Dim IntRetCD
	Dim arrVal1
	Dim arrVal2

	strSelect	=			 " CONVERT(VARCHAR(6), DATEADD(m,-1, MAX(fisc_yr + fisc_mnth)+" & FilterVar("01", "''", "S") & " ), 112),CONVERT(VARCHAR(6), DATEADD(m,0, MAX(fisc_yr + fisc_mnth)+" & FilterVar("01", "''", "S") & " ), 112) "
	strFrom		=			 " a_gl_sum (NOLOCK) "
	strWhere	=			 " fisc_dt = " & FilterVar("00", "''", "S") & "  "

	Call CommonQueryRs( strSelect , strFrom ,  strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Replace(lgF0, Chr(11),"") = ""  Then
		Call CommonQueryRs( "convert(VARCHAR(6),min(gl_dt),112)" , "a_gl" ,  "(dr_amt <> 0 and cr_amt <> 0 and dr_loc_amt <> 0 and cr_loc_amt <> 0)", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		frm1.txtLastCloseMnth.value = Replace(lgF0, Chr(11),"")
		frm1.txtFromDt.value		=  Replace(lgF0, Chr(11),"")   
		Exit Sub
	Else 
		arrVal1 = Split(lgF0, Chr(11))
		arrVal2 = Split(lgF1, Chr(11))  
		frm1.txtLastCloseMnth.value = arrVal1(0)
		frm1.txtFromDt.value = arrVal2(0)
		
	End If
	
End Sub

Sub Rb_WK1_ONCLICK()
	
	Dim strSelect, strFrom, strWhere
	
	strSelect	=			 " CONVERT(VARCHAR(6), DATEADD(m,-1, MAX(fisc_yr + fisc_mnth)+" & FilterVar("01", "''", "S") & " ), 112),CONVERT(VARCHAR(6), DATEADD(m,0, MAX(fisc_yr + fisc_mnth)+" & FilterVar("01", "''", "S") & " ), 112) "
	strFrom		=			 " a_gl_sum (NOLOCK) "
	strWhere	=			 " fisc_dt = " & FilterVar("00", "''", "S") & "  "

	Call CommonQueryRs( strSelect , strFrom ,  strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	
	if Replace(lgF0, Chr(11),"") <> "" then
	   frm1.txtFromDt.value		=  left(replace(UniDateAdd("m", 0, frm1.txtLastCloseMnth.value + 01 ,parent.gServerDateFormat),"-",""),6)
	Else
		Call CommonQueryRs( "convert(VARCHAR(6),min(gl_dt),112)" , "a_gl" ,  "(dr_amt <> 0 and cr_amt <> 0 and dr_loc_amt <> 0 and cr_loc_amt <> 0)", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		frm1.txtLastCloseMnth.value = Replace(lgF0, Chr(11),"")
		frm1.txtFromDt.value		=  Replace(lgF0, Chr(11),"")   
	End if
		
End Sub

Sub Rb_WK2_ONCLICK()

	Dim strSelect, strFrom, strWhere
	
	strSelect	=			 " CONVERT(VARCHAR(6), DATEADD(m,-1, MAX(fisc_yr + fisc_mnth)+" & FilterVar("01", "''", "S") & " ), 112),CONVERT(VARCHAR(6), DATEADD(m,0, MAX(fisc_yr + fisc_mnth)+" & FilterVar("01", "''", "S") & " ), 112) "
	strFrom		=			 " a_gl_sum (NOLOCK) "
	strWhere	=			 " fisc_dt = " & FilterVar("00", "''", "S") & "  "

	Call CommonQueryRs( strSelect , strFrom ,  strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Replace(lgF0, Chr(11),"") <> "" then
	   frm1.txtFromDt.value		=  frm1.txtLastCloseMnth.value
	Else
	   Call DisplayMsgBox("121290","X","X","X")
	   frm1.Rb_WK1.checked = true
	End if	
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================
Function FncPrint() 
    On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_SINGLE)
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>결산마감및이월</font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업구분</TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked><LABEL FOR=Rb_WK1>결산마감</LABEL>&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2><LABEL FOR=Rb_WK2>마감취소</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>최종마감년월</TD>
								<TD CLASS="TD6"><INPUT NAME="txtLastCloseMnth" ALT="최종마감년월"   MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="24X" class=protected readonly=true tabindex="-1"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6"><INPUT NAME="txtFromDt" ALT="작업년월"   MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="24X" class=protected readonly=true tabindex="-1"></TD>
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
					<TD><BUTTON NAME="btn배치" CLASS="CLSMBTN" OnClick="VBScript:Call fnButtonExec()" Flag=1>실 행</BUTTON></TD>		        		
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

