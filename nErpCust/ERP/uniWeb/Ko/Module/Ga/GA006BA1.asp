<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3980ba1
'*  4. Program Name         : 매출/매출원가 집계 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/01/09
'*  8. Modified date(Last)  : 2001/03/5
'* 11. Comment              :
'=======================================================================================================  -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<Script Language="VBScript">

Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "GA006BB1.asp"

Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgLngCurRows = 0
    
End Sub

Sub SetDefaultVal()

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA")%>
End Sub

Function ExeReflect() 
	Dim strVal
	Dim strYyyymm
	Dim	strYear, strMonth, strDay
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If

	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
	ExeReflect = False
    
	On Error Resume Next

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    strYYYYMM = strYear & strMonth	

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtYyyymm=" & strYyyymm

	Call RunMyBizASP(MyBizASP, strVal)

	frm1.txtSpread.value =  strVal

	
	ExeReflect = True
    
End Function

Function ExeReflectOk()
Dim IntRetCD

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	'window.status = "작업 완료"

End Function

Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables
	Call SetDefaultVal
	Call SetToolbar("10000000000000")

	frm1.txtYyyymm.focus
	Set gActiveElement = document.activeElement			
    
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub

Function FncQuery()

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
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출/매출원가집계</font></td>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/ga006ba1_fpDateTime1_txtYyyymm.js'></script>
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
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD Width = 10> &nbsp </TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

