<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111ma1
'*  4. Program Name         : ���ֵ�� 
'*  5. Program Desc         : ���ֵ�� 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="m3111ma1.vbs"></SCRIPT>

<SCRIPT  LANGUAGE="VBScript" >

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029
    Call AppendNumberRange("0","0","999")
    
    call initFormatField()        							
    Call SetDefaultVal
    Call InitVariables
    '----------  Coding part  -------------------------------------------------------------
    Call Changeflg
    Call CookiePage(0)
	Call changeTabs(TAB1)
    
	gIsTab     = "Y"
	gTabMaxCnt = 2
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����Ϲ�����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" onMouseOver="vbscript:SetClickflag" onMouseOut="vbscript:ResetClickflag" onFocus="vbscript:SetClickflag" onBlur="vbscript:ResetClickflag">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ֹ�������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT Class = required TYPE=TEXT NAME="txtPoNo" SIZE=32  MAXLENGTH=18 ALT="���ֹ�ȣ" tag="12NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS=TD6></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR height="*">
					<TD WIDTH=100% valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ֹ�ȣ" NAME="txtPoNo2"  MAXLENGTH=18 SIZE=34 tag="21XXXU" STYLE = "text-transform:uppercase"></TD>
									<TD CLASS="TD5" NOWRAP>Ȯ������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="����Ȯ��" NAME="rdoRelease" CLASS="RADIO" checked tag="24" ONCLICK="vbscript:SetChangeflg()"><label for="rdoRelease">&nbsp;��Ȯ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="����Ȯ��" NAME="rdoRelease" CLASS="RADIO" ONCLICK="vbscript:setChangeflg()" tag="24"><label for="rdoRelease">&nbsp;Ȯ��&nbsp;</label></TD>
								</TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="��������" NAME="txtPotypeCd"  MAXLENGTH=5 SIZE=10 tag="23NXXU" ONChange="vbscript:ChangePotype()" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPotype()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="��������" NAME="txtPotypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier()" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								<TR>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDateTime1_txtPoDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" MAXLENGTH=4 SIZE=10 tag="22NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
														   <INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���󳳱���</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDateTime2_txtDvryDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>								
								<TR>
									<TD CLASS="TD5" NOWRAP>ȭ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="ȭ��" NAME="txtCurr" MAXLENGTH=3 SIZE=10 tag="22NXXU" onChange="ChangeCurr()" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurr()">
														   <INPUT TYPE=HIDDEN AlT="ȭ��" NAME="txtCurrNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>ȯ��</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle2_txtXch.js'></script></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ּ��ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle1_txtPoAmt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>���ּ��ڱ��ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle1_txtPoLocAmt.js'></script></td>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>�����ѱݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle1_txtGrossPoAmt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>�������ڱ��ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle1_txtGrossPoLocAmt.js'></script></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>VAT</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVattype" ALT="VAT"  MAXLENGTH=5 SIZE=10 tag="21NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVat()">
														   <INPUT TYPE=TEXT AlT="VAT" NAME="txtVatTypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>VAT�ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle4_txtVatAmt.js'></script></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>VAT��</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle3_txtVatrt.js'></script>&nbsp;&nbsp;%</TD>
																<TD CLASS="TD5" nowrap>VAT���Ա���</TD>
								    <TD CLASS="TD6" nowrap>
								    <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT���Ա���" CLASS="RADIO" checked id="rdoVatFlg1" tag="21X"><label for="rdoVatFlg">���� </label>
									<INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT���Ա���" CLASS="RADIO" id="rdoVatFlg2"  tag="21X"><label for="rdoVatFlg">����&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="�������" NAME="txtPayTermCd"  MAXLENGTH=5 SIZE=10 tag="22NXXU" OnChange="VBScript:changePayterm()" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayMeth()">
														   <INPUT TYPE=TEXT AlT="�������" NAME="txtPayTermNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 >
														   <INPUT TYPE=HIDDEN AlT="�������" NAME="txtReference" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>�����Ⱓ</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP><script language =javascript src='./js/m3111ma1_fpDoubleSingle2_txtPayDur.js'></script>
												</TD>
												<TD NOWRAP>
													&nbsp;��
												</TD>
											</TR>
										</Table>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��������" NAME="txtPayTypeCd"  MAXLENGTH=5 SIZE=10 tag="21NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayType()">
														   <INPUT TYPE=TEXT AlT="��������" NAME="txtPayTypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>���ձ��ſ���</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="���ձ��ſ���" NAME="rdoMergPurFlg" CLASS="RADIO" tag="21" id="rdoMergPurFlg1" ONCLICK="vbscript:SetChangeflg()"><label for="rdoMergPurFlg1">&nbsp;YES&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="���ձ��ſ���" NAME="rdoMergPurFlg" CLASS="RADIO" checked id="rdoMergPurFlg2" ONCLICK="vbscript:setChangeflg()" tag="21"><label for="rdoMergPurFlg2">&nbsp;NO&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����ó�������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó�������" NAME="txtSuppSalePrsn" MAXLENGTH=50 SIZE=34 tag="21"></TD>
									<TD CLASS="TD5" NOWRAP>��޿���ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��޿���ó" NAME="txtTel" MAXLENGTH=30 SIZE=34 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">��ݰ�������</TD>
									<TD CLASS="TD6" colspan=3 width=100% NOWRAP><INPUT TYPE=TEXT AlT="��ݰ�������" Size="90" NAME="txtPayTermstxt" MAXLENGTH=120 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">���</TD>
									<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT TYPE=TEXT  NAME="txtRemark" ALT="���" tag = "21" SIZE=90 MAXLENGTH=120></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(4)%>
							</TABLE>
						</div>
						<!--�ι�° �� -->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>OFFER�ۼ���</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDateTime1_txtOffDt.js'></script></TD>
							        <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
							        <TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ȿ��</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma1_fpDateTime3_txtExpiryDt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>INVOICE NO.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInvNo" MAXLENGTH=50 SIZE=34 ALT="INVOICE NO." MAXLENGTH=20 tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT NAME="txtIncotermsCd" ALT ="��������"  MAXLENGTH=5 SIZE=10 tag="32NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9006')">
														   <INPUT TYPE=TEXT NAME="txtIncotermsNm" ALT ="��������" SIZE=20 tag="34X" CLASS = protected readonly = True TabIndex = -1  ></TD>
									<TD CLASS="TD5" NOWRAP>��۹��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT NAME="txtTransCd"  MAXLENGTH=5 SIZE=10 ALT ="��۹��" tag="32NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9009')">
														   <INPUT TYPE=TEXT NAME="txtTransNm" SIZE=20 ALT ="��۹��" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�۱�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBankCd"  MAXLENGTH=10 SIZE=10 ALT ="�۱�����" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBank()">
														   <INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 ALT ="�۱�����" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>�ε����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDvryPlce" MAXLENGTH=5 SIZE=10 ALT="�ε����" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9095')">
														   <INPUT TYPE=TEXT NAME="txtDvryPlceNm" SIZE=20 ALT ="�ε����" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT NAME="txtApplicantCd" MAXLENGTH=10 SIZE=10 ALT ="������" tag="32NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz('Appl')">
														   <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 ALT ="������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtManuCd" MAXLENGTH=10 SIZE=10 ALT ="������" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz('Manu')">
														   <INPUT TYPE=TEXT NAME="txtManuNm" SIZE=20 ALT ="������" tag="34X" CLASS = protected readonly = True TabIndex = -1  ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAgentCd"  MAXLENGTH=10 SIZE=10 ALT ="������" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz('Agent')">
														   <INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 ALT ="������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOrigin"  MAXLENGTH=5 SIZE=10 ALT="������" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9094')">
														   <INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 ALT ="������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPackingCd" MAXLENGTH=5 SIZE=10 ALT ="��������" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9007')">
														   <INPUT TYPE=TEXT NAME="txtPackingNm" SIZE=20 ALT ="��������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>�˻���</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspectCd" MAXLENGTH=5 SIZE=10 ALT ="�˻���" tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9008')">
														   <INPUT TYPE=TEXT NAME="txtInspectNm" SIZE=20 ALT ="�˻���" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDisCity" MAXLENGTH=5 ALT="��������" SIZE=10 tag="31NXXU"  STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9096')">
														   <INPUT TYPE=TEXT NAME="txtDisCityNm" SIZE=20 ALT ="��������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDisPort" MAXLENGTH=5 ALT="������" SIZE=10 tag="31XXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9092')">
														   <INPUT TYPE=TEXT NAME="txtDisPortNm" SIZE=20 ALT ="������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoadPort" MAXLENGTH=5 ALT="������" SIZE=10 tag="31NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9092-1')">
														   <INPUT TYPE=TEXT NAME="txtLoadPortNm" SIZE=20 ALT ="������" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtShipment" MAXLENGTH=70 ALT="��������" SIZE=34 tag="31"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(11)%>
							</TABLE>
						</DIV>
					</TD>	
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td align="Left"><a><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">Ȯ��</button></a>
									 <Div  STYLE="DISPLAY: none"><a><button name="btnSend" id="btnSend" class="clsmbtn" ONCLICK="Sending()">�ֹ����߼�</button></a></Div>
					</td>   
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">���ֳ������</a> | <a href="VBSCRIPT:CookiePage(2)">�����</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRelease" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCurr" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBLflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCCflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdvatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIssueType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMergPurFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaintNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnxchrateop" tag="2">
<INPUT TYPE=HIDDEN NAME="hdclsflg" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
