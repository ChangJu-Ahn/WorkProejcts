<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		:
'*  3. Program ID			: p4912ma1
'*  4. Program Name			: �۾��Ϻ�����Ʈ 
'*  5. Program Desc			:
'*  6. Comproxy List		: +
'*  7. Modified date(First)	: 2005-01-17
'*  8. Modified date(Last) 	:
'*  9. Modifier (First) 	: Yoon, Jeong Woo
'* 10. Modifier (Last)		:
'* 11. Comment				:
'* 12. History              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>                 <!--�޴��� �Է��ϴ� ȭ�� �ڵ� khk-->
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ��� -->
<!--'==========================================  1.1.1 Style Sheet  ==========================================
'========================================================================================================= -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!--Link�� ���� css������ ��� khk?????-->

<!--'==========================================  1.1.2 ���� Include   ========================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4912ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit														'��: indicates that All variables must be declared in advance

Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
StartDate = UNIDateAdd("D",-10,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
EndDate = UNIDateAdd("D", 20,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'===========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

'    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

	Call SetToolbar("1100000000001111")										'��: ��ư ���� ���� 

    Call InitSpreadSheet                                                    '��: Setup the Spread sheet

	Call DefaultSumValue
	Call SetDefaultVal
    Call InitVariables														'��: Initializes local global variables

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Ucase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�۾��Ϻ�����Ʈ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�۾�����</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/p4912ma1_I306724661_txtFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4912ma1_I908196461_txtToDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProdOrderNo() "></TD>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="12xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConWC()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="14" ALT="�۾����"></TD>
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
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/p4912ma1_I257290583_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
							</TR>
							<TR>
							<TD WIDTH=100% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>�հ�</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD6Y6 NOWRAP>�ҷ���</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle1_txtDocAmt.js'></script></TD>
										<TD CLASS=TD6Y6 NOWRAP>���Լ�</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle2_txtDocAmt.js'></script></TD>
										<TD CLASS=TD6Y6 NOWRAP>�ϼ���</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle3_txtDocAmt.js'></script></TD>
										<TD CLASS=TD6Y6 NOWRAP>ǥ�ذ���</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle4_txtDocAmt.js'></script></TD>
									</TR>
									<TR>
										<TD CLASS=TD6Y6 NOWRAP>�۾�����</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle5_txtDocAmt.js'></script></TD>
										<TD CLASS=TD6Y6 NOWRAP>��������</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle6_txtDocAmt.js'></script></TD>
										<TD CLASS=TD6Y6 NOWRAP>���ǰ���</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle7_txtDocAmt.js'></script></TD>
										<TD CLASS=TD6Y6 NOWRAP>��Ÿ����</TD>
										<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4912ma1_fpDoubleSingle8_txtDocAmt.js'></script></TD>
									</TR>
								</TABLE>
							</FIELDSET>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hToDt" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>