<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : SCM
'*  2. Function Name        : SCM Order Management
'*  3. Program ID           : Mscm1dqa1
'*  4. Program Name         : ������Ȳ��ȸ(��ǰó)
'*  5. Program Desc         : ����Ȯ���� SCM���� ��ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004/07/28
'*  8. Modified date(Last)  : 2004/08/09
'*  9. Modifier (First)     : nhg
'* 10. Modifier (Last)      : nhg
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 
'*                            
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="U2214QA1.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit				
'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================

Dim LocSvrDate
LocSvrDate = "<%=GetSvrDate%>"
						   
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'================================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029								
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")           
	Call InitSpreadSheet("*")
	Call SetDefaultVal	
    Call InitVariables
	
	If parent.gPlant <> "" and frm1.txtPlantCd.Value = "" Then
		frm1.txtPlantCd.Value = parent.gPlant
		frm1.txtPlantNm.Value = parent.gPlantNm
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement 
	End If
	
End Sub
'================================================================================================================================

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������Ȳ��ȸ(��ǰó))</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=500>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
			 						<TD CLASS=TD5 NOWRAP>��ü</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="14xxxU" ALT="��ü">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="��ü��"></TD>
									<TD CLASS=TD5 NOWRAP>��ǰó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="��ǰó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/u2214qa1_OBJECT3_txtDvFrDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/u2214qa1_OBJECT4_txtDvToDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg2" Value="N" checked tag="11"><label for="rdoAppflg2">&nbsp;����&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="��ǰ" NAME="rdoAppflg" id = "rdoAppflg3" Value="Y" tag="11"><label for="rdoAppflg3">&nbsp;��ǰ&nbsp;</label></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/u2214qa1_OBJECT1_txtPoFrDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/u2214qa1_OBJECT2_txtPoToDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰâ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="��ǰâ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSlInfo frm1.txtSlCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=28 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/u2214qa1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/u2214qa1_B_vspdData2.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBpCd" tag="24"><INPUT TYPE=HIDDEN NAME="hDvFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hDvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hrdoAppflg" tag="24"><INPUT TYPE=HIDDEN NAME="hPoFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hPoToDt" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>