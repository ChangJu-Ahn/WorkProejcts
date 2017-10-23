<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� smj
'*  2. Function Name        : �������/��ǰ��� 
'*  3. Program ID           : S4311MA1
'*  4. Program Name         : �������/��ǰ��� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S31111MaintSoHdrSvr, S31119LookupSoHdrSvr
'*  7. Modified date(First) : 2002/03/22
'*  8. Modified date(Last)  : 2003/10/14
'*  9. Modifier (First)     : Sung MiJung
'* 10. Modifier (Last)      : Hwang seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/09 : ..........
'*                            -2000/05/09 : ǥ�ؼ����������� 
'*                            -2000/09/04 : 4Th Coding
'*                            -2001/12/18 : Date ǥ�� ���� 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="S4311ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                               

Dim iDBSYSDate
Dim EndDate, StartDate

Dim lblnWinEvent   '������ �߰� 
Dim interface_Account   '������ �߰� 


iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'=========================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
    <% Call LoadBNumericFormatA( "I", "*", "NOCOOKIE", "MA") %>
End Sub

</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" --> 

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ChangeTabs(TAB1)">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������/��ǰ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ChangeTabs(TAB2)">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���� �� ǰ������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>     
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ChangeTabs(TAB3)">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ �� �������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>     
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD6"><INPUT NAME="txtConDn_no" ALT="���Ϲ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnNo"></TD>
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
						<!-- ù��° �� ���� -->
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>���Ϲ�ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtDnNo" ALT="���Ϲ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDn_Type" ALT="��������" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="23XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDNType()">&nbsp;<INPUT NAME="txtDn_TypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSold_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="�ֹ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp 0" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtSold_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�Ǹ�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeal_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="�Ǹ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 3">&nbsp;<INPUT NAME="txtDeal_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR> 
									 <TD CLASS=TD5 NOWRAP>��ǰó</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="��ǰó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp 1" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtShip_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD> 
									 <TD CLASS=TD5 NOWRAP>�����׷�</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtSales_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>          
								</TR>        
								<TR>         
									 <TD CLASS=TD5 NOWRAP>���ݽŰ�����</LABEL></TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRequried 4">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									 <TD CLASS=TD5 NOWRAP>�������</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_terms" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 5">&nbsp;<INPUT NAME="txtPay_terms_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									 <TD CLASS="TD5" NOWRAP>������</TD>
									 <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="������"></OBJECT>');</SCRIPT></TD>
									 <TD CLASS="TD5" NOWRAP>�������</TD>
									 <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtPlannedGIDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="�������"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>��ݰ�������</TD>
									 <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txt_Payterms_txt" TYPE="Text" MAXLENGTH="120" SIZE=80 tag="21" ALT="��ݰ�������"></TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>���</TD>
									 <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="120" SIZE=80 tag="21" ALT="���"></TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>VAT����</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtVat_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="VAT����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 6">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									 <TD CLASS=TD5 NOWRAP>VAT��</TD>
									 <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtVat_rate" ALT = "VAT��" CLASS=FPDS140 tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL></TD>         
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
									 <TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoVat_Inc_flag" id="rdoVat_Inc_flag1" value="1" tag = "21" checked>
											<label for="rdoVat_Inc_flag1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoVat_Inc_flag" id="rdoVat_Inc_flag2" value="2" tag = "21">
											<label for="rdoVat_Inc_flag2">����</label></TD>
									 <TD CLASS=TD5 NOWRAP>�ݾ�</TD>
									 <TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtNet_amt" CLASS=FPDS140 tag="24X2Z" ALT = "�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;&nbsp;
												</TD>
												<TD>
													<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="ȭ��">
												</TD>
											</TR>
										</TABLE>
										</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT�������</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoVat_Calc_Type" id="rdoVat_Calc_Type1" value="1" tag = "21" checked>
										 <label for="rdoVat_Calc_Type1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoVat_Calc_Type" id="rdoVat_Calc_Type2" value="2" tag = "21">
										 <label for="rdoVat_Calc_Type2">����</label></TD>
									<TD CLASS=TD5 NOWRAP>VAT�ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtVat_amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
									</TD>  
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��۹��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_Meth" ALT="��۹��" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 4">&nbsp;<INPUT NAME="txtTrans_Meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�ѱݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
									 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtTot_amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
									</TD>  
						        </TR> 
   								<TR>
									<TD CLASS=TD5 NOWRAP>������ǰ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtArriv_dt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="������ǰ��" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>��ǰ�ð�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtArriv_Tm" TYPE="Text" ALT="��ǰ�ð�" MAXLENGTH="10" SIZE=36 tag="21"></TD>
								</TR>									
							        <%Call SubFillRemBodyTD5656(4)%>
							</TABLE>
						</DIV>
       
      
						<!-- �ι�° �� ���� -->
						<DIV ID="TabDiv"  STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="����" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS="TD5" NOWRAP>â��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtSlCd" ALT="â��" TYPE="Text" MAXLENGTH=7 SiZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenSl()">&nbsp;<INPUT NAME="txtSlNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR> 
								<TR>
									<TD CLASS=TD5 NOWRAP>�������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvMgr" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInvMgr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInvMgrPopUp">&nbsp;<INPUT NAME="txtInvMgrNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="������ڸ�"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ļ��۾�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">����ä��</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="21" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">���ݰ�꼭</LABEL>
									</TD>
									<TD CLASS=TD5 NOWRAP>�ѱݾ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtTotal_Amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;&nbsp;</TD>
								</TR>
							    <TR>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCol_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRequried 2">&nbsp;<INPUT NAME="txtCol_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>���ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
								        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtCol_amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE" ALT="���ݾ�"></OBJECT>');</SCRIPT>       
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���������</TD>
									<TD CLASS=TD6 NOWRAP>
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtGI_Dt" CLASS=FPDTYYYYMMDD tag="24X1" ALT="�����" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>����ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtGINo" ALT="����ȣ" TYPE="Text" MAXLENGTH=18 SiZE=22 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR> 
								<TR>
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>

							</TABLE>
						</DIV>
						
						<!-- ����° �� ���� -->
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
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiver" TYPE="Text" ALT="�μ��ڸ�" MAXLENGTH="50" SIZE=36 tag="21"></TD>
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
									 <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDlvyPlace" TYPE="Text" MAXLENGTH="30" SIZE=91 ALT="��ǰ���" tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ȭ��ȣ1</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No1" TYPE="Text" ALT="��ȭ��ȣ1" MAXLENGTH="20" SIZE=37 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>��ȭ��ȣ2</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No2" TYPE="Text" ALT="��ȭ��ȣ2" MAXLENGTH="20" SIZE=37 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSender" TYPE="Text" ALT="�ΰ��ڸ�" MAXLENGTH="50" SIZE=37 tag="21"></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5>������ȣ</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 TAG="21XXXX" ALT="������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVehicleNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenVehicleNo()"></TD>							
									<TD CLASS=TD5 NOWRAP>�����ڸ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDriver" TYPE="Text" ALT="�����ڸ�" MAXLENGTH="50" SIZE=37 tag="21"></TD>
								</TR>
								   <%Call SubFillRemBodyTD5656(6)%>
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
					<TD>
						<BUTTON NAME="btnPosting" CLASS="CLSMBTN">���</BUTTON>&nbsp;
						<BUTTON NAME="btnPostCancel" CLASS="CLSMBTN">������</BUTTON>&nbsp;
			         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">��ǥ��ȸ</BUTTON>
					</TD>     
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
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioDnParcel" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHDNNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="RdoConfirm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSO_TYPE" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdrStateFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtDtlStateFlg" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtArFlag" tag="24" TABINDEX="-1">	<!-- DB����, ���������� ��ϵ� ��� 'Y', �׷��� ���� ��� 'N' -->
<INPUT TYPE=HIDDEN NAME="txtVATFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRetItemFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRetBillFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtExportFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHTransit_LT" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHCntryCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtlgBlnChgValue1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtlgBlnChgValue2" tag="24" TABINDEX="-1">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm" TABINDEX="-1"></IFRAME>
</DIV> 
</BODY>
</HTML>
