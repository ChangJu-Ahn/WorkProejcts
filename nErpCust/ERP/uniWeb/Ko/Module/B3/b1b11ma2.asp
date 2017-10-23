<%@ LANGUAGE="VBScript" %>												   
<!--'****************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : b1b11ma2.asp																*
'*  4. Program Name         : Item By Plant ��ȸ ASP													*
'*  5. Program Desc         : 																			*
'*  6. Component List       :
'*  7. Modified date(First) : 2000/12/14																*
'*  8. Modified date(Last)  : 2002/11/14																*
'*  9. Modifier (First)     : Jung Yu Kyung																*
'* 10. Modifier (Last)      : Hong Chang Ho																*
'* 11. Comment              :																			*
'********************************************************************************************************-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "b1b11ma2.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "MA")%>
End Sub

Sub InitComboBox()
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
    
End Sub

Sub Form_Load()
	Call LoadInfTB19029																
	
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)

	Call ggoOper.LockField(Document, "N")											
	
	Call InitSpreadSheet
	Call SetDefaultVal		
	Call InitComboBox
	Call InitVariables
	gSelframeFlg = TAB1
	Call changeTabs(TAB1)
	Call SetToolbar("11000000000011")		
	gTabMaxCnt = 3
    gIsTab = "Y"										
	
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

</Script>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
-->
<BODY SCROLL="NO" TABINDEX="-1">
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
						<TABLE CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���庰ǰ����ȸ</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU"  ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenConItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14" ALT="ǰ���"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" ALT="����" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>���ޱ���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProcType" ALT="���ޱ���" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="11X1"> </OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="11X1"> </OBJECT>');</SCRIPT>					
									</TD>
									<TD CLASS=TD5 NOWRAP>��ȿ����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailableItem" tag="1X" CHECKED ID="rdoAvailableItem1" VALUE="A"><LABEL FOR="rdoAvailableItem1">��ü</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailableItem" tag="1X" ID="rdoAvailableItem2" VALUE="Y"><LABEL FOR="rdoAvailableItem2">��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailableItem" tag="1X" ID="rdoAvailableItem3" VALUE="N"><LABEL FOR="rdoAvailableItem3">�ƴϿ�</LABEL></TD>
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
						<!--<TABLE WIDTH="100%" HEIGHT="100%">-->
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>								
								<!-- TreeView AREA -->
								<TD HEIGHT=100% WIDTH=40%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<!-- DATA AREA -->
								<TD WIDTH="60%" HEIGHT="100%">
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
										<TR HEIGHT=23>
											<TD WIDTH="100%">
												<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH="100%" border=0>
													<TR>
														<TD WIDTH=10>&nbsp;</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
																<TR>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ���Ϲ�����</font></td>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
															    </TR>
															</TABLE>
														</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
																<TR>
																	<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP��������</font></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
															    </TR>
															</TABLE>
														</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
																<TR>
																	<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���/ǰ������</font></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
															    </TR>
															</TABLE>
														</TD>
														<TD WIDTH=*>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD WIDTH="100%" CLASS="TB2">
												<!-- ù��° �� ���� -->
												<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR> <!-- Data Sheet -->
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>																				
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="24" ALT="ǰ��"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=50 tag="24" ALT="ǰ���"></TD>													
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAccount" ALT="ǰ�����" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																		<TD CLASS=TD5 NOWRAP>ǰ��԰�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=50 tag="24" ALT="ǰ��԰�"></TD>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���ش���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasicUnit" ALT="���ش���" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�����ǰ��Ŭ����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemClass" ALT="�����Ŭ����" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���ޱ���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProcType" ALT="���ޱ���" SIZE=20 tag="24"></SELECT></TD>
																	</TR>																				
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdEnv" ALT="��������" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>MPSǰ��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMPSItem" tag="24" ID="rdoMPSItem1" VALUE="Y"><LABEL FOR="rdoMPSItem1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMPSItem" tag="24" ID="rdoMPSItem2" VALUE="N"><LABEL FOR="rdoMPSItem2">�ƴϿ�</LABEL></TD>
																	</TR>												
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Tracking����</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="24" ID="rdoTrackingItem1" VALUE="Y"><LABEL FOR="rdoTrackingItem1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="24" ID="rdoTrackingItem2" VALUE="N"><LABEL FOR="rdoTrackingItem2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�ܰ�������</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoCollectFlg" tag="24" ID="rdoCollectFlg1" VALUE="Y"><LABEL FOR="rdoCollectFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoCollectFlg" tag="24" ID="rdoCollectFlg2" VALUE="N"><LABEL FOR="rdoCollectFlg2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�۾���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWorkCenter" SIZE=20 MAXLENGTH=7 tag="24" ALT="�۾���"></TD>													
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��ȿ����</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailable" tag="24" ID="rdoAvailable1" VALUE="Y"><LABEL FOR="rdoAvailable1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailable" tag="24" ID="rdoAvailable2" VALUE="N"><LABEL FOR="rdoAvailable2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ǥ��ST</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT  name=txtStdTime SIZE=20 tag="24" ALT="ǥ��ST" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ATP L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT  name=txtAtpLt SIZE=20 tag="24" ALT="ATP L/T" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidFromDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="������" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>																																			&nbsp;~&nbsp;
																			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidToDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="������" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
																		</TD>
																	</TR>																	
																</TABLE>										
															</TD>
														</TR>
													</TABLE>
												</DIV>
												<!-- �ι�° �� ���� -->
												<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no> 
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR> <!-- Data Sheet -->
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB3" CELLSPACING=0>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������������</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMRPFlg" tag="24" ID="rdoMRPFlg1" VALUE="Y"><LABEL FOR="rdoMRPFlg1">��</LABEL>
																			<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMRPFlg" tag="24" ID="rdoMRPFlg2" VALUE="N"><LABEL FOR="rdoMRPFlg2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>													
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderFrom" ALT="������������" SIZE=20 tag="24" ></TD>													
																		<TD CLASS=TD5 NOWRAP>Lot Sizing</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLotSizing" ALT="Lot Sizing" SIZE=20 tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�ҿ䷮�ø�����</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRoundFlg" tag="24" ID="rdoRoundFlg1" VALUE="Y"><LABEL FOR="rdoRoundFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRoundFlg" tag="24" ID="rdoRoundFlg2" VALUE="N"><LABEL FOR="rdoRoundFlg2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>�ø��Ⱓ</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtRoundPeriod SIZE=20 ALT="�ø��Ⱓ" tag="24" STYLE="TEXT-ALIGN: right"></TD>												
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�ִ��������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMaxOrderQty SIZE=20 ALT="�ִ��������" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>���� L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtOffsetLt SIZE=20 ALT="���� L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�ּҿ�������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMinOrderQty SIZE=20 ALT="�ּҿ�������" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>�ø���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtRoundQty SIZE=20 ALT="�ø���" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtFixOrderQty SIZE=20 ALT="������������" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>���Ҽ�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtLineNo SIZE=20 ALT="���μ�" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Į����Ÿ��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtCalType SIZE=5 ALT="Į����Ÿ��" tag="24"></TD>
																		<TD CLASS=TD5 NOWRAP>MRP �����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMRPMgr" ALT="MRP �����" SIZE=20 tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>��������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdMgr" ALT="��������" SIZE=20 tag="24"></TD>													
																		<TD CLASS=TD5 NOWRAP>���� L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtVarLT SIZE=20 ALT="���� L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>																					
																	<TR>																					
																		<TD CLASS=TD5 NOWRAP>Damper����</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDamperFlg" tag="24" ID="rdoDamperFlg1" VALUE="Y"><LABEL FOR="rdoDamperFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDamperFlg" tag="24" ID="rdoDamperFlg2" VALUE="N"><LABEL FOR="rdoDamperFlg2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>Damper �ּ���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtDamperMinQty SIZE=20 ALT="Damper �ּ���" tag="24" STYLE="TEXT-ALIGN: right"></TD>											
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMfgOrderUnit" SIZE=5  tag="24"  ALT="������������"></TD>
																		<TD CLASS=TD5 NOWRAP>���ſ�������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrderUnit" SIZE=5 MAXLENGTH=3 tag="24"  ALT="���ſ�������"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�������� L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMfgOrderLT SIZE=20 ALT="�������� L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>���ſ��� L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtPurOrderLT SIZE=20 ALT="�������� L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>												
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>����ǰ��ҷ���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMfgScrapRate SIZE=20 ALT="����ǰ��ҷ���" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>����ǰ��ҷ���</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtPurScrapRate ALT="���źҷ���" tag="24" size= 20 STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD5 NOWRAP>��������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=20 MAXLENGTH=18 tag="24" ALT="��������"></TD>
																	</TR>
																				
																</TABLE>								
															</TD>
														</TR>
													</TABLE>
												</DIV>
												<!-- ����° �� ���� -->
												<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR> <!-- Data Sheet -->
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB3" CELLSPACING=0>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�԰�â��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=20 MAXLENGTH=7 tag="24" ALT="�԰�â��"></TD>
																		<TD CLASS=TD5 NOWRAP>�����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueType" ALT="�����" STYLE="aling: right;" tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���â��</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueSLCd" SIZE=20 MAXLENGTH=7 tag="24" ALT="���â��"></TD>
																		<TD CLASS=TD5 NOWRAP>������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueUnit" SIZE=5 MAXLENGTH=3 tag="24"  ALT="��������"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Lot No.����</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotNoFlg" tag="24" ID="rdoLotNoFlg1" VALUE="Y"><LABEL FOR="rdoLotNoFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotNoFlg" tag="24" ID="rdoLotNoFlg2" VALUE="N"><LABEL FOR="rdoLotNoFlg2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>�������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtSFStockQty SIZE=20 ALT="�������" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtReorderPnt SIZE=20 ALT="������" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>�������üũ</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInvCheckFlg" tag="24" ID="rdoInvCheckFlg1" VALUE="Y"><LABEL FOR="rdoInvCheckFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInvCheckFlg" tag="24" ID="rdoInvCheckFlg2" VALUE="N"><LABEL FOR="rdoInvCheckFlg2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���԰���뿩��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOverRcptFlg" tag="24" ID="rdoOverRcptFlg1" VALUE="Y"><LABEL FOR="rdoOverRcptFlg1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOverRcptFlg" tag="24" ID="rdoOverRcptFlg2" VALUE="N"><LABEL FOR="rdoOverRcptFlg2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>���԰������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtOverRcptRate SIZE=20 ALT="���԰������" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���ǻ��ֱ�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtCycleCntPerd SIZE=20  tag="24" ALT="���ǻ��ֱ�" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>ǰ��ABC����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT  NAME="txtABCFlg" SIZE=5 ALT="ǰ��ABC����" tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�������</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInvMgr" SIZE=20 ALT="�������" tag="24"></TD>
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>���԰˻翩��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPurInspType" tag="24" ID="rdoPurInspType1" VALUE="Y"><LABEL FOR="rdoPurInspType1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPurInspType" tag="24" ID="rdoPurInspType2" VALUE="N"><LABEL FOR="rdoPurInspType2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>�����˻翩��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMfgInspType" tag="24" ID="rdoMfgInspType1" VALUE="Y"><LABEL FOR="rdoMfgInspType1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMfgInspType" tag="24" ID="rdoMfgInspType2" VALUE="N"><LABEL FOR="rdoMfgInspType2">�ƴϿ�</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�����˻翩��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFinalInspType" tag="24" ID="rdoFinalInspType1" VALUE="Y"><LABEL FOR="rdoFinalInspType1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFinalInspType" tag="24" ID="rdoFinalInspType2" VALUE="N"><LABEL FOR="rdoFinalInspType2">�ƴϿ�</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>���ϰ˻翩��</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueInspType" tag="24" ID="rdoIssueInspType1" VALUE="Y"><LABEL FOR="rdoIssueInspType1">��</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueInspType" tag="24" ID="rdoIssueInspType2" VALUE="N"><LABEL FOR="rdoIssueInspType2">�ƴϿ�</LABEL></TD>
																	</TR>
				     												<TR>
																		<TD CLASS=TD5 NOWRAP>�����˻� L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT SIZE=20 name=txtMfgInspLT tag="24" ALT="�����˻� L/T" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>���Ű˻� L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT SIZE=20 name=txtPurInspLT tag="24" ALT="���Ű˻� L/T" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>������ �˻�����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMPSMgr" ALT="MPS �����" SIZE=20 tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>
																		<TD CLASS=TD5 NOWRAP>���Ž� �˻�����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT SIZE=20  NAME="txtInspecMgr" ALT="�˻�����" tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�ܰ�����</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPrcCtrlInd" ALT="�ܰ�����" SIZE=15 tag="24"></TD>
																		<TD CLASS=TD5 NOWRAP>ǥ�شܰ�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtStdPrice SIZE=20 ALT="ǥ�شܰ�" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>�̵���մܰ�</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMoveAvgPrice SIZE=20 ALT="�̵���մܰ�" tag="24" STYLE="TEXT-ALIGN: right"></TD>											
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
												</DIV>
											</TD>
										</TR>
									</TABLE>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemAccunt" tag="24">
<INPUT TYPE=HIDDEN NAME="hProcType" tag="24"><INPUT TYPE=HIDDEN NAME="hAvailableItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
