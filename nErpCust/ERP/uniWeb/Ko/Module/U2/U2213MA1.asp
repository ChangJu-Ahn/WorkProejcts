
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: SCM
'*  2. Function Name		: Firm Purchase Receipt Management
'*  3. Program ID			: Mscm1dma1.asp
'*  4. Program Name			: 반품등록 
'*  5. Program Desc			: 
'*  6. Comproxy List		: 
'*  7. Modified date(First)	: 2005/03/28
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: NHG
'* 10. Modifier (Last)		: 
'* 11. Comment	
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "U2213MA1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit									'☜: indicates that All variables must be declared in advance

Dim LocSvrDate
Dim lgStrColorFlag
Dim lgStrPrevKey1
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

LocSvrDate = "<%=GetSvrDate%>"

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	'****************************
	'List Minor code(Order Type)
	'****************************
	<%
	Dim iData
    iData = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P3211' ")
	Response.write "Call SetCombo3(frm1.cboOrderType, """ &  iData & """) " & vbCrLf
	%>
	frm1.cboOrderType.value = "" 

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     				'⊙: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                          			'⊙: Lock  Suitable  Field
	
    Call InitSpreadSheet()                                                    				'⊙: Setup the Spread sheet

	Call SetDefaultVal
	
    Call InitVariables																		'⊙: Initializes local global variables

    Call InitSpreadComboBox()

	Set gActiveElement = document.activeElement 
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>반품등록(납품처)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
			 						<TD CLASS=TD5 NOWRAP>업체</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="14xxxU" ALT="업체">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=10 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant(frm1.txtplantcd.value)">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>						 					 
									<TD CLASS=TD5 NOWRAP>반품일자</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/u2213ma1_OBJECT1_txtRetFrDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/u2213ma1_OBJECT2_txtRetToDt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>납품업체</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd2" SIZE=10 MAXLENGTH=10 tag="11xxxU" ALT="납품업체"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDlvyno" align=top TYPE="BUTTON"ONCLICK="vbscript:Openbpcd2(FRM1.txtBpCd2.VALUE)">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBpNm2" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>상태</TD>					 
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=radio Class="Radio" ALT="상태" NAME="rdoflg" id = "rdoflg3" Value="A" tag="11" checked><label for="rdoflg3">&nbsp;전체&nbsp;</label>
														 <INPUT TYPE=radio Class="Radio" ALT="상태" NAME="rdoflg" id = "rdoflg1" Value="P" tag="11"><label for="rdoflg1">&nbsp;계획&nbsp;</label>
														 <INPUT TYPE=radio Class="Radio" ALT="상태" NAME="rdoflg" id = "rdoflg2" Value="O" tag="11"><label for="rdoflg2">&nbsp;실행&nbsp;</label>
														 <INPUT TYPE=radio Class="Radio" ALT="상태" NAME="rdoflg" id = "rdoflg4" Value="E" tag="11"><label for="rdoflg2">&nbsp;확정&nbsp;</label></TD>
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT="100%">
								<TD WIDTH="100%" colspan = "4">
									<script language =javascript src='./js/u2213ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBPCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hDvFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hDvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hrdoAppflg" tag="24"><INPUT TYPE=HIDDEN NAME="hPoFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hPoToDt" tag="24">
<P ID="divTextArea"></P>
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname">
	<input type="hidden" name="dbname">
	<input type="hidden" name="filename">
	<input type="hidden" name="condvar">
	<input type="hidden" name="date">                 
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>