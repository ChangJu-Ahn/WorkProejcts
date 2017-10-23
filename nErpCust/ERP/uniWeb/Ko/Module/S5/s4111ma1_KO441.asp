<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S4111MA1
'*  4. Program Name         : ���ϵ�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41111MaintDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2003/08/22
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="S4111ma1_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'=====================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=22>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="changeTabs(TAB1)">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ϵ��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="changeTabs(TAB2)">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ �� �������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenSORef">��������</A>&nbsp;<A href="vbscript:OpenDNReqRef">���Ͽ�û����</A></TD>
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
									<TD CLASS="TD5" NOWRAP>���Ϲ�ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConDnNo" ALT="���Ϲ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConDnNo()"></TD>
									<TD CLASS="TDT"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>���Ϲ�ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtDnNo" ALT="���Ϲ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlantCd" ALT="����" TYPE="Text" SIZE=10 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="�����"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtPlanned_gi_dt" CLASS=FPDTYYYYMMDD tag="22X1" ALT="�������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMovType" ALT="��������" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtMovTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="��ǰó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShipToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup C_PopShiptoParty">&nbsp;<INPUT NAME="txtShip_to_partyNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="��ǰó��"></TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_Type" ALT="��������" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtSo_TypeNm" ALT="�������¸�" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvMgr" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInvMgr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup C_PopInvMgr">&nbsp;<INPUT NAME="txtInvMgrNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="������ڸ�"></TD>
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" ALT="�����׷�" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtSales_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��۹��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_meth" ALT="��۹��" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransMeth" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup C_PopTransMeth">&nbsp;<INPUT NAME="txtTrans_meth_nm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="��۹����"></TD>
									<TD CLASS=TD5 NOWRAP>���������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtActGi_dt" ALT="���������" TYPE="Text" style="text-align=center" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_no" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkSoNo" tag="25" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkSoNo"> ���ֹ�ȣ����</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>����ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGoods_mv_no" ALT="����ȣ" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5 NOWRAP>������ǰ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtArriv_dt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="������ǰ��" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDlvy_dt" ALT="������" TYPE="Text" style="text-align=center" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ�ð�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtArriv_Tm" TYPE="Text" ALT="��ǰ�ð�" MAXLENGTH="10" SIZE=40 tag="21"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>							
								<TR>	
									<TD CLASS=TD5 NOWRAP>���</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="120" SIZE=91 ALT="���" tag="21"></TD>
								</TR>
	                            <% Call SubFillRemBodyTD5656(7) %>
							</TABLE>
						</DIV>

						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰó��������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSTP_Inf_No" ALT="��ǰó��������ȣ" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<BUTTON NAME = "btnShipToPlceRef" CLASS="CLSMBTN">��ǰó����������</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIP_cd" TYPE="Text" ALT="�����ȣ" MAXLENGTH="12" SIZE=20 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
									<TD CLASS=TD5 NOWRAP>�μ��ڸ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiver" TYPE="Text" ALT="�μ��ڸ�" MAXLENGTH="50" SIZE=35 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ�ּ�</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1_Dlv" TYPE="Text" ALT="��ǰ�ּ�" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2_Dlv" TYPE="Text" ALT="��ǰ�ּ�" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR3_Dlv" TYPE="Text" ALT="��ǰ�ּ�" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ���</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtShip_to_place"  ALT="��ǰ���" TYPE="Text" MAXLENGTH="30" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ȭ��ȣ1</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No1" TYPE="Text" ALT="��ȭ��ȣ1" MAXLENGTH="20" SIZE=35 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>��ȭ��ȣ2</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No2" TYPE="Text" ALT="��ȭ��ȣ2" MAXLENGTH="20" SIZE=35 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrnsp_Inf_No" ALT="���������ȣ" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<BUTTON NAME = "btnTrnsMethRef" CLASS="CLSMBTN">�����������</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5>���ȸ��</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtTransCo" SIZE=20 MAXLENGTH=50 TAG="21XXXX" ALT="���ȸ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransCo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTransCo()"></TD>
									<TD CLASS=TD5 NOWRAP>�ΰ��ڸ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSender" TYPE="Text" ALT="�ΰ��ڸ�" MAXLENGTH="50" SIZE=35 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5>������ȣ</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 TAG="21XXXX" ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVehicleNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenVehicleNo()"></TD>
									<TD CLASS=TD5 NOWRAP>�����ڸ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDriver" TYPE="Text" ALT="�����ڸ�" MAXLENGTH="50" SIZE=35 tag="21"></TD>
								</TR>
	                            <% Call SubFillRemBodyTD5656(6) %>
							</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck()">���ϳ������</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtChkSoNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtTempSoNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHCntryCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtlgBlnChgValue1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtlgBlnChgValue2" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHRefRoot" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
