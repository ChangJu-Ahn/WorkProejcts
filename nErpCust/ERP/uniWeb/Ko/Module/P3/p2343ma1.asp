<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2343ma1
'*  4. Program Name         : MRP����/������� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Hyun Jae 
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p2343ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->

'=========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=========================================================================================================
Sub LoadInfTB19029() 
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
End Sub 

'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 

Sub Form_Load()
	Call LoadInfTB19029()
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")

    Call SetToolbar("10000000000011")

    Call InitVariables    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		
		gLookUpEnable = True
		
		If DBQuery = False Then 
           Call RestoreToolBar()
           Exit Sub
        End If 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
    End If
    
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP����/�������</font></td>
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
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="23XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>MRP �����ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMRPHisNo" SIZE=18 MAXLENGTH=18 tag="24" ALT="MRP �����ȣ"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2343ma1_I450466549_txtFixExecFromDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Ȯ�������Ⱓ</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2343ma1_I830675631_txtFixExecToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���������Ⱓ</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2343ma1_I722488570_txtPlanExecToDt.js'></script>
								</TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������� ���ȿ���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg1 CHECKED><LABEL FOR=rdoAvailInvFlg1>������</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg2><LABEL FOR=rdoAvailInvFlg2>���Ⱦ���</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������� ���ȿ���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg1 CHECKED><LABEL FOR=rdoSafeInvFlg1>������</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg2><LABEL FOR=rdoSafeInvFlg2>���Ⱦ���</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2343ma1_I368479198_txtStartDt.js'></script>									
								</TD>		
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ȹ������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2343ma1_I709968984_txtOrderQty.js'></script>
								</TD>		
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Status</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStatus" SIZE=20 MAXLENGTH=10 tag="24" ALT="Status"></TD>		
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
					<TD><BUTTON NAME="btnApprove" CLASS="CLSMBTN" Flag=1 onclick="ApproveMRP()">MRP ����</BUTTON>&nbsp;<BUTTON NAME="btnCancel" CLASS="CLSMBTN" Flag=1 onclick="CancelMRP()">MRP �������</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>							
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
