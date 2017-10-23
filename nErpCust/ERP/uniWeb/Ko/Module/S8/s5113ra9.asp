<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S5113RA9
'*  4. Program Name         : B/L ������																*
'*  5. Program Desc         : ���� B/L��� ���� ASP														*
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/08/12																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn TaeHee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : ȭ�� design												*
'*							  2. 2000/04/18 : Coding Start												*
'*							  2. 2002/08/12 : Ado														*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>B/L ������</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					<% '��: indicates that All variables must be declared in advance %>
'========================================================================================================
Const BIZ_PGM_ID = "s5113rb9.asp"				<% '��: �����Ͻ� ���� ASP�� %>
'========================================================================================================
Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3
'========================================================================================================
DIm gSelframeFlg					<% '���� TAB�� ��ġ�� ��Ÿ���� Flag %>

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
'========================================================================================================
Function InitVariables()
	
	Dim arrParam

	arrParam = arrParent(1)

	frm1.txtBLNo.value = arrParam(0)
	frm1.txtBLDocNo.value = arrParam(1)
	
	Self.Returnvalue = ""
		
End Function
'********************************************************************************************************
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q","S","NOCOOKIE","PA") %>
End Sub	
'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
		
	Call changeTabs(TAB1)
		
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
		
	gSelframeFlg = TAB2
End Function
	
Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
		
	Call changeTabs(TAB3)
		
	gSelframeFlg = TAB3
End Function
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029																<% '��: Load table , B_numeric_format %>
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call ggoOper.LockField(Document, "N")											<% '��: Lock  Suitable  Field %>
	Call InitVariables
	Call changeTabs(TAB1)

	If Trim(frm1.txtBLNo.value) <> "" Then
		Call DbQuery()
	End If
	frm1.txtLocCurrency.value = PopupParent.gCurrency
	frm1.txtLocCurrency1.value = PopupParent.gCurrency
End Sub
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub	
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'B/L �ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'B/L �ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt1, .txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'B/L �ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtMoney, .txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub
'========================================================================================================
	Function DbQuery()
		Err.Clear															<%'��: Protect system from crashing%>

		DbQuery = False														<%'��: Processing is NG%>

		Dim strVal

		If LayerShowHide(1) = False Then
			Exit Function
		End If

		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)			<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&txtLocCurrency=" & PopupParent.gCurrency					<%'��: ��ȸ ���� ����Ÿ %>
		Call RunMyBizASP(MyBizASP, strVal)									<%'��: �����Ͻ� ASP �� ���� %>
	
		DbQuery = True														<%'��: Processing is NG%>
	End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE CLASS="BasicTB" CELLSPACING=0>
			<TR>
				<TD HEIGHT=5>&nbsp;<% ' ���� ���� %></TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE CLASS="BasicTB" CELLSPACING=0>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������Ÿ</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ä������</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD WIDTH=500>&nbsp;</TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<TABLE CLASS="BasicTB">
						<TR>
							<TD HEIGHT=5 WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS=TD5 NOWRAP>B/L ������ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=20 MAXLENGTH=18 TAG="14"></TD>
											<TD CLASS=TD5 NOWRAP>B/L��ȣ</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="14XXXU" ALT="B/L��ȣ"></TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=100% WIDTH=100%>
							<!-- ù��° �� ���� -->
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE CLASS="BasicTB">	
										<TR>
											<TD HEIGHT=2 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSONo" TYPE=TEXT SIZE=20 TAG="24XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime2_txtBLIssueDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>B/L�ݾ�</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtDocAmt.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��">
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>ȯ��</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtXchRate.js'></script></TD>
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>B/L�ڱ��ݾ�</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtLocAmt.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��">
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>��۹��</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="��۹��">&nbsp;<INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>��������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="��������">&nbsp;<INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="24"></TD></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePort" ALT="������" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>�����׷�</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="�����׷�">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>					
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime2_txtLoadingDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�������ҹ��</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="�������ҹ��">&nbsp;<INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>B/L�������</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtBLIssueCnt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>B/L�������</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtBLIssuePlce" ALT="B/L�������" TYPE=TEXT MAXLENGTH=30 SIZE=80 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>	
								<!-- �ι�° �� ���� -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=5 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>VESSEL��</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL��" TYPE=TEXT MAXLENGTH=34 SIZE=34 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>������ȣ</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVoyageNo" MAXLENGTH=20 SIZE=34 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>����ȸ��</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtForwarder" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtForwarderNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>���ڱ���</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVesselCntry" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtVesselCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�������</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtReceiptPlce" ALT="�������" TYPE=TEXT MAXLENGTH=35 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>�ε����</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtDeliveryPlce" ALT="�ε����" TYPE=TEXT MAXLENGTH=50 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>����������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="����������" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>����������</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime4_txtDischgeDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>ȯ������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTranshipCntry" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtTranshipCntryNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>ȯ����</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime5_txtTranshipDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>��������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="��������">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>�����尹��</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtTotPackingCnt.js'></script></TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>�����������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackingTxt" ALT="�����������" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>�����̳ʼ�</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtContainerCnt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>���߷�</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtGrossWeight.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>�߷�����</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWeightUnit" ALT="�߷�����" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�ѿ���</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtGrossVolumn.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>��������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVolumnUnit" ALT="��������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>����������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="����������" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�����������</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtFreightPlce" ALT="�����������" TYPE=TEXT MAXLENGTH=30 SIZE=80 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>
								<!-- ����° �� ���� -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=5 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>���ݽŰ�����</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizArea" ALT="���ݽŰ�����" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>����ä������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillType" ALT="����ä������" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>����ó</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayer" ALT="����ó" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtPayerNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>����ó</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBilltoParty" ALT="����ó" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBilltoPartyNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>Ȯ������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="24X" VALUE="Y" ID="rdoPostingflg1"><LABEL FOR="rdoPostingflg1">Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" VALUE="N" TAG="24X" CHECKED ID="rdoPostingflg2"><LABEL FOR="rdoPostingflg2">��Ȯ��</LABEL></TD>
														<TD CLASS=TD5 NOWRAP>���ݸ�����</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime8_txtPayDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>���ݿ����׷�</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtToSalesGroup" ALT="���ݿ����׷�" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtToSalesGroupNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>B/L�ݾ�</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtDocAmt1.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��">
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>B/L�ڱ��ݾ�</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtLocAmt1.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="ȭ��">
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>�Ѽ��ݾ�</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtMoney.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayType" ALT="�Ա�����" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>�������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="������">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" ALT="�����Ⱓ" STYLE="TEXT-ALIGN: right" TYPE=TEXT MAXLENGTH=3 SIZE=5 TAG="24X7">&nbsp;��</TD>
														<TD CLASS=TD5 NOWRAP>��������</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermstxt" ALT="��������" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>���</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark" ALT="���" TYPE=TEXT MAXLENGTH=35 SIZE=80 TAG="24"></TD>
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
			<TR>
				<TD HEIGHT=30>
					<TABLE CLASS="basicTB" CELLSPACING=0>
						<TR>
							<TD WIDTH=* ALIGN=RIGHT>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" 
							     onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHLCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtCCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHBLNo" TAG="24">
		<INPUT TYPE=HIDDEN NAME="txtRefFlg" TAG="24">
	</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
