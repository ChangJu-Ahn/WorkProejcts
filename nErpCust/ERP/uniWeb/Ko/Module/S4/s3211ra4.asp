<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ma4.asp																*
'*  4. Program Name         : L/C 상세정보(L/C등록에서)													*
'*  5. Program Desc         : L/C 상세정보(L/C등록에서)													*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2000/08/29																*
'*  9. Modifier (First)     : An ChangHwan 																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'*																										*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>L/C 상세정보</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID = "s3211rb4.asp"			

Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3
Const TAB4 = 4

Dim arrReturn					
Dim gSelframeFlg
Dim gblnWinEvent				
Dim arrParam
Dim arrParent
Dim PopupParent

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0					
	lgStrPrevKey = ""					
	gblnWinEvent = False
End Function

'========================================================================================================	
Sub SetDefaultVal()
	With frm1
		arrParam = arrParent(1)

		.txtLCNo.value = arrParam(0)
		.txtSONo.value = arrParam(1)

		gblnWinEvent = False
		Self.Returnvalue = ""
	End With
End Sub	
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "*", "NOCOOKIE", "RA") %>
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
	
Function ClickTab4()
	If gSelframeFlg = TAB4 Then Exit Function
		
	Call changeTabs(TAB4)
		
	gSelframeFlg = TAB4
End Function

'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
Sub Form_Load()
		
	Call LoadInfTB19029																
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											
	Call InitVariables
	Call SetDefaultVal
	Call DbQuery()
	Call changeTabs(TAB1)
	
End Sub
	
'========================================================================================================
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	If   LayerShowHide(1) = False Then
	    Exit Function 
	End If
		
	Call ggoOper.ClearField(Document, "2")								
	Call InitVariables													

	frm1.txtLocCurrency.value = PopupParent.gCurrency
		
	Dim strVal
	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)					
		
	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function
	
'========================================================================================================
Function DbQueryOk()													

	Call ggoOper.LockField(Document, "Q")								

	If gSelframeFlg <> TAB1 Then
		Call ClickTab1()
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE CLASS="BasicTB" CELLSPACING=0>
			<TR>
				<TD HEIGHT=5>&nbsp;</TD>
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C 금액정보</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>L/C 선적정보</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>	
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구비서류</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab4()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>은행및기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
					<TABLE CLASS="BasicTB">    
						<TR>
							<TD HEIGHT=5 WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD"> 
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS=TD5 NOWRAP>L/C관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCNo"  SIZE=20 MAXLENGTH=18 TAG="14XXXU" ALT="L/C관리번호"></TD>
											<TD CLASS=TD5 NOWRAP>수주번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSONo" SIZE=20  MAXLENGTH=18 TAG="14XXXU" ALT="수주번호"></TD>
										</TR>	
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=2 WIDTH=100%></TD>
						</TR>
					  	<TR>
							<TD WIDTH=100% HEIGHT=100%>
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">	
									<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
										<TR>	
											<TD CLASS=TD5 NOWRAP>L/C번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LC번호" TYPE=TEXT MAXLENGTH=35 SIZE=20 TAG="14XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>통지번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvNo" SIZE=20 MAXLENGTH=35 TAG="24XXXU" ALT="통지번호"></TD>
										</TR>		
										<TR>
											<TD CLASS=TD5 NOWRAP>L/C유형</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCType" SIZE=10 MAXLENGTH=4 STYLE="TEXT-ALIGN: left" TAG="14XXXU" ALT="L/C유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtLCTypeNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>통지일</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra4_fpDateTime1_txtAdvDt.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>통지은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdvBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="24XXXU" ALT="통지은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAdvbank" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtAdvBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>유효일</TD>
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra4_fpDateTime2_txtExpireDt.js'></script></TD>
										</TR>									
										<TR>
											<TD CLASS=TD5 NOWRAP>개설은행</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="24XXXU" ALT="개설은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>개설일</TD>						
											<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra4_fpDateTime3_txtOpenDt.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>개설금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=2 TAG="24XXXU" ALT="화폐">&nbsp;
														</TD>
														<TD>
															&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s3211ra4_fpDoubleSingle1_txtDocAmt.js'></script>
														</TD>
													</TR>
												</TABLE>
											</TD>	
											<TD CLASS=TD5 NOWRAP>개설자국금액</TD>
											<TD CLASS=TD6 NOWRAP>
												<TABLE CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=2 TAG="24XXXU">
														</TD>
														<TD>
															&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s3211ra4_fpDoubleSingle2_txtLocAmt.js'></script>
														</TD>
													</TR>
												</TABLE>
										</TR>	
										<TR>							
											<TD CLASS=TD5 NOWRAP>환율</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/s3211ra4_fpDoubleSingle3_txtXchRate.js'></script>
											</TD>
											<TD CLASS=TD5 NOWRAP>수입자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수입자">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>과부족허용율</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/s3211ra4_fpDoubleSingle4_txttolerance.js'></script>&nbsp;%
											</TD>
											<TD CLASS=TD5 NOWRAP>수출자</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
										</TR>	
										<TR>	
											<TD CLASS=TD5 NOWRAP>가격조건</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoTerms" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="가격조건">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtIncoTermsNm" SIZE=20 TAG="24"></TD>										
											<TD CLASS=TD5 NOWRAP>영업그룹</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="영업그룹">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>결제방법</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="결제방법">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
											<TD CLASS=TD5 NOWRAP>결제기간</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayDur" SIZE=5 STYLE="TEXT-ALIGN: right" MAXLENGTH=3 TAG="24" ALT="결제기간">&nbsp;DAYS</TD>
										</TR>								
									</TABLE>
								</DIV>
								
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>최종선적일자</TD>	
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra4_fpDateTime4_txtLatestShipDt.js'></script></TD>
														<TD CLASS=TD6 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운송방법</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="운송방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransport" align=top TYPE="BUTTON"ON>&nbsp;<INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>환적허용</TD>
														<TD CLASS=TD6 COLSPAN=3>
																<INPUT TYPE="RADIO" CLASS="RADIO" VALUE="Y" NAME="rdoTranshipment" TAG="24" CHECKED ID="rdoTranshipment1"><LABEL FOR="rdoTranshipment1">Y</LABEL>&nbsp;&nbsp;&nbsp;
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTranshipment" TAG="24" VALUE="N" ID="rdoTranshipment2"><LABEL FOR="rdoTranshipment2">N</LABEL>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>분할선적허용</TD>
														<TD CLASS=TD6 COLSPAN=3>
															<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="24" VALUE="Y" CHECKED ID="rdoPartailShip1"><LABEL FOR="rdoPartailShip1">Y</LABEL>&nbsp;&nbsp;&nbsp;
															<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="24" VALUE="N" ID="rdoPartailShip2"><LABEL FOR="rdoPartailShip2">N</LABEL>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>선적항</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>도착항</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtDischgePort" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
													</TR>				
													<TR>
														<TD CLASS=TD5 NOWRAP>인도장소</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDeliveryPlce" ALT="운송방법" TYPE=TEXT MAXLENGTH=30 SIZE=35 TAG="24"></TD>
													</TR>	
													<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3>&nbsp;</TD>
													</TR> 
													<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3>&nbsp;</TD>
													</TR> 
													<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3>&nbsp;</TD>
													</TR> 
													<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3>&nbsp;</TD>
													</TR> 							
												</TABLE>
											</TD>		
										</TR>
									</TABLE>
								</DIV>
					
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>서류제시기간</TD>
														<TD CLASS=TD6 NOWRAP>
															<script language =javascript src='./js/s3211ra4_fpDoubleSingle5_txtFileDt.js'></script>&nbsp;DAYS
														</TD>
														<TD CLASS=TD5 NOWRAP>서류제시기간 참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFileDtTxt" TYPE=TEXT MAXLENGTH=35 SIZE=25 TAG="24" ALT="서류제시기간 참조"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>상업송장</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvCnt" ALT="COMMERCIAL INVOICE" TYPE=TEXT MAXLENGTH=2 SIZE=5  STYLE="TEXT-ALIGN: right" TAG="24X7">&nbsp;부</TD>
														<TD CLASS=TD5 NOWRAP>포장명세서</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackList" ALT="PACKING LIST" TYPE=TEXT MAXLENGTH=2 SIZE=5  STYLE="TEXT-ALIGN: right" TAG="24X7">&nbsp;부</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP><INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="24" VALUE="Y" NAME="chkCertOriginFlg" ID="chkCertOriginFlg"></TD>
														<TD CLASS=TD6 NOWRAP><LABEL FOR="chkCertOriginFlg">원산지증명서</LABEL></TD>
														<TD CLASS=TD5 NOWRAP>B/L종류</TD>
														<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoBLAwFlg" TAG="24" VALUE="Y" CHECKED ID="rdoBLAwFlg1">
															<LABEL FOR="rdoBLAwFlg">BILL OF LADING</LABEL>
															<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoBLAwFlg" TAG="24" VALUE="N" ID="rdoBLAwFlg2">
															<LABEL FOR="rdoBLAwFlg">AIRWAY BILL</LABEL>
														</TD>
													</TR>	
													<TR>	
														<TD CLASS=TD5 NOWRAP>운임지불여부</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="운임지불여부">&nbsp;<INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
														
														<TD CLASS=TD5 NOWRAP>통지처</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNotifyParty" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="통지처">&nbsp;<INPUT TYPE=TEXT NAME="txtNotifyPartyNm" SIZE=20 TAG="24"></TD>	
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수탁자</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtConsignee"  SIZE=80 MAXLENGTH=80 TAG="24" ALT="수탁자"></TD>	
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>부보조건</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtInsurPolicy" ALT="보험부보조건" TYPE=TEXT MAXLENGTH=80 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>구비서류</TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc1"  TYPE=TEXT MAXLENGTH=80 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc2"  TYPE=TEXT MAXLENGTH=80 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc3"  TYPE=TEXT MAXLENGTH=80 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc4"  TYPE=TEXT MAXLENGTH=80 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDoc5"  TYPE=TEXT MAXLENGTH=80 SIZE=80 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>
						
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>지급은행</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayBank" TYPE=TEXT SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="지급은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayBank" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtPayBankNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>RENEGO은행</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRenegoBank" TYPE=TEXT SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="RENEGO은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRenegoBank" ALIGN=TOP TYPE="BUTTOMN">&nbsp;<INPUT TYPE=TEXT NAME="txtRenegoBankNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>확인은행</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConfirmBank" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="확인은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConfirmBank" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtConfirmBankNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>은행지시사항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBankTxt" SIZE=35 MAXLENGTH=70 TAG="24" ALT="은행지시사항"></TD>
													</TR>	
													<TR>
														<TD CLASS=TD5 NOWRAP>양도허용여부</TD>
														<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoTransfer" TAG="24" VALUE="Y" CHECKED ID="rdoTransfer1"><LABEL FOR="rdoTransfer">Y</LABEL>
															<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoTransfer" TAG="24" VALUE="N" ID="rdoTransfer2"><LABEL FOR="rdoTransfer">N</LABEL>
														</TD>
														<TD CLASS=TD5 NOWRAP>신용공여주체</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCreditCore" SIZE=10 MAXLENGTH=4 TAG="24" ALT="신용공여주체"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCreditCore" ALIGN=TOP TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtCreditCoreNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수수료 부담자</TD>
														<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoChargeCd" TAG="24" VALUE="Y" CHECKED ID="rdoChargeCd1"><LABEL FOR="rdoTransfer">Applicant</LABEL>
															<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoChargeCd" TAG="24" VALUE="N" ID="rdoChargeCd2"><LABEL FOR="rdoTransfer">Beneficiary</LABEL>
														</TD>	
														<TD CLASS=TD5 NOWRAP>수수료 참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChargeTxt" SIZE=35 MAXLENGTH=30 TAG="24" ALT="수수료 참조사항"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>대금결제 참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPaymentTxt" SIZE=35 MAXLENGTH=30 TAG="24" ALT="대금 결제참조"></TD>
														<TD CLASS=TD5 NOWRAP>선적기일 참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShipment" SIZE=35 MAXLENGTH=30 TAG="24" ALT="선적기일 참조사항"></TD>
													</TR>											
													<TR>
														<TD CLASS=TD5 NOWRAP>선통지참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPreAdvRef" ALT="선통지 참조사항" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>운송회사</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransportComp" ALT="운송회사" TYPE=TEXT MAXLENGTH=30 SIZE=35 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>원산지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="원산지" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON">&nbsp;<INPUT NAME="txtOriginNm" ALT="원산지명" TYPE=TEXT MAXLENGTH=30 SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>원산지국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="원산지국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>대행자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="대행자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>제조자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="제조자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>기타참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark" ALT="기타참조" TYPE=TEXT MAXLENGTH=70 SIZE=35 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>AMEND일</TD>
														<TD CLASS=TD6 NOWRAP>
															<script language =javascript src='./js/s3211ra4_fpDateTime5_txtAmendDt.js'></script>
														</TD>
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
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
			</TR>
		</TABLE>
	<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtHSoNo" TAG="24">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
		<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>           

