<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name		: Production
'*  2. Function Name	: 
'*  3. Program ID		: p1401ma2.asp
'*  4. Program Name		: BOM ��ȸ 
'*  5. Program Desc		:
'*  6. Component List	: 
'*  7. Modified date(First)	: 2000/04/18
'*  8. Modified date(Last)	: 2002/11/19
'*  9. Modifier (First)		: Im Hyun Soo
'* 10. Modifier (Last)		: Hong Chang Ho
'* 11. Comment		:
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p1401ma2.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			'��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False				'��: Indicates that no value changed
    lgIntGrpCount = 0					'��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
	lgSelNode = ""
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "MA")%>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029													'��: Load table , B_numeric_format
	Call AppendNumberPlace("6", "5", "0")
	Call AppendNumberPlace("7", "2", "2")
	Call AppendNumberPlace("8", "11", "6")	
	
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")								'��: Lock  Suitable  Field
   
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11000000000011")
    Call SetDefaultVal
	Call InitVariables													'��: Initializes local global variables
	Call InitTreeImage	
		
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>BOM��ȸ</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtBaseDt CLASSID=<%=gCLSIDFPDT%> tag="12X1" ALT="������"></OBJECT>');</SCRIPT>
										</OBJECT>
									</TD>
									<TD CLASS=TD5 ROWSPAN=2 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoSrchType" ID="rdoSrchType1" CLASS="RADIO" tag="1X" Value="2" CHECKED><LABEL FOR="rdoSrchType1">������</LABEL>
														
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()" >&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>BOM Type</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBomNo" SIZE=5 MAXLENGTH=3 tag="12XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBomNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo"></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoSrchType" ID="rdoSrchType2" CLASS="RADIO" tag="1X" Value="4"><LABEL FOR="rdoSrchType2">������</LABEL></TD>
									
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
								<TD HEIGHT=* WIDTH=50%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=uniTree1 width=100% height=100% <%=UNI2KTV_IDVER%>> <PARAM NAME="ImageWidth" VALUE="16"> <PARAM NAME="ImageHeight" VALUE="16"> <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7"> <PARAM NAME="LabelEdit" VALUE="1"> </OBJECT>');</SCRIPT>
								</TD>
								<TD HEIGHT=* WIDTH=50% VAlign=Top>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18  tag="24" ALT="��ǰ��"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ���</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40  tag="24" ALT="��ǰ���"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=25 tag="24" ALT="ǰ�����"><INPUT TYPE=HIDDEN NAME="txtItemAcct" tag="24" ALT="ǰ�����"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ǰ��԰�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE=40 tag="24" ALT="ǰ��԰�"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemFromDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="������" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
																										&nbsp;~&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemToDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="������" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
												</TD>	
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>BOM Type / ����</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD><INPUT TYPE=TEXT NAME="txtBomNo1" SIZE=5 MAXLENGTH=3  tag="24" ALT="BOM Type"></TD>
															<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtBOMDesc" SIZE=30 MAXLENGTH=40  tag="24" ALT="BOM ����"></TD>
														</TR>
													</TABLE>
												</TD>									
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>������</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDrawNo" SIZE=40 tag=24 ALT="������"></TD>
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>���� / ����</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD><INPUT TYPE=TEXT NAME="txtLevel" SIZE=8  tag="24" ALT="����"></TD>
															<TD>&nbsp;
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtItemSeq CLASS=FPDS65 title=FPDOUBLESINGLE SIZE="15" MAXLENGTH="3" ALT="����" tag="24X6Z"> </OBJECT>');</SCRIPT>
															</TD>
														</TR>
													</TABLE>
												</TD>									
											</TR>

											<TR>
												<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtChildItemQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X8Z" ALT="��ǰ����ؼ�" MAXLENGTH="15" SIZE="15"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemUnit" SIZE=4 MAXLENGTH=3  tag="24" STYLE="Text-Transform: uppercase" ALT="��ǰ�����">
															</TD>
														</TR>
													</TABLE>
												</TD>														
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtPrntItemQty CLASS=FPDS140 title=FPDOUBLESINGLE SIZE=15 MAXLENGTH=15 ALT="��ǰ����ؼ�" tag="24X8Z"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>
																&nbsp;<INPUT TYPE=TEXT NAME="txtPrntItemUnit" align=top SIZE=4 MAXLENGTH=3  tag="24" STYLE="Text-Transform: uppercase" ALT="��ǰ�����">
															</TD>	
														</TR>
													</TABLE>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>����L/T</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSafetyLt CLASS=FPDS140 title=FPDOUBLESINGLE SIZE="15" MAXLENGTH="3" ALT="����L/T" tag="24X6Z"> </OBJECT>');</SCRIPT>
															</TD>
															<TD valign=bottom>
																&nbsp;��
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Loss��(%)</TD>
												<TD CLASS=TD6 NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtLossRate CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X7Z" ALT="Loss��" MAXLENGTH="15" SIZE="15"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�����󱸺�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoSupplyFlg" ID="rdoSupplyFlg1" CLASS="RADIO" tag="24X" Value="F" CHECKED><LABEL FOR="rdoSupplyFlg1">����</LABEL>
												     				 <INPUT TYPE="RADIO" NAME="rdoSupplyFlg" ID="rdoSupplyFlg2" CLASS="RADIO" tag="24X" Value="C"><LABEL FOR="rdoSupplyFlg2">����</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>���</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRemark" SIZE=40  tag="24" ALT="���"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="" name=txtValidFromDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24X1" ALT="��ȿ�Ⱓ" MAXLENGTH="10" SIZE="10"> </OBJECT>');</SCRIPT>
													&nbsp;~&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id="" name=txtValidToDt1 CLASS=FPDTYYYYMMDD title=FPDATETIME SIZE="10" MAXLENGTH="10" ALT="��ȿ�Ⱓ" tag="24X1"> </OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>���躯���ȣ</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo" SIZE=20 tag="24" ALT="���躯���ȣ"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>���躯�泻��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNDescription" SIZE=40 tag=24 ALT="���躯�泻��"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>���躯��ٰ�</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNReasonCd" SIZE=40 tag="24" ALT="���躯��ٰ�"></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txthBomNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHdnItemAcct" tag="14">
<INPUT TYPE=HIDDEN NAME="txtSrchType" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>