<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Sales & Distribution														*
'*  2. Function Name        :																			*
'*  3. Program ID           : S5113RA9
'*  4. Program Name         : B/L 상세정보																*
'*  5. Program Desc         : 수출 B/L등록 참조 ASP														*
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2002/08/12																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Ahn TaeHee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/18 : 화면 design												*
'*							  2. 2000/04/18 : Coding Start												*
'*							  2. 2002/08/12 : Ado														*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>B/L 상세정보</TITLE>
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

Option Explicit					<% '☜: indicates that All variables must be declared in advance %>
'========================================================================================================
Const BIZ_PGM_ID = "s5113rb9.asp"				<% '☆: 비지니스 로직 ASP명 %>
'========================================================================================================
Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3
'========================================================================================================
DIm gSelframeFlg					<% '현재 TAB의 위치를 나타내는 Flag %>

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
	Call LoadInfTB19029																<% '⊙: Load table , B_numeric_format %>
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call ggoOper.LockField(Document, "N")											<% '⊙: Lock  Suitable  Field %>
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
		'B/L 금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'B/L 금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt1, .txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'B/L 금액 
		ggoOper.FormatFieldByObjectOfCur .txtMoney, .txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub
'========================================================================================================
	Function DbQuery()
		Err.Clear															<%'☜: Protect system from crashing%>

		DbQuery = False														<%'⊙: Processing is NG%>

		Dim strVal

		If LayerShowHide(1) = False Then
			Exit Function
		End If

		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtBLNo=" & Trim(frm1.txtBLNo.value)			<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtLocCurrency=" & PopupParent.gCurrency					<%'☆: 조회 조건 데이타 %>
		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>
	
		DbQuery = True														<%'⊙: Processing is NG%>
	End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE CLASS="BasicTB" CELLSPACING=0>
			<TR>
				<TD HEIGHT=5>&nbsp;<% ' 상위 여백 %></TD>
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>선적정보</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>선적기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권정보</font></td>
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
											<TD CLASS=TD5 NOWRAP>B/L 관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBLNo" SIZE=20 MAXLENGTH=18 TAG="14"></TD>
											<TD CLASS=TD5 NOWRAP>B/L번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLDocNo" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="14XXXU" ALT="B/L번호"></TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=100% WIDTH=100%>
							<!-- 첫번째 탭 내용 -->
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE CLASS="BasicTB">	
										<TR>
											<TD HEIGHT=2 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>수주번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSONo" TYPE=TEXT SIZE=20 TAG="24XXXU"></TD>
														<TD CLASS=TD5 NOWRAP>L/C번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT SIZE=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>발행일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime2_txtBLIssueDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>B/L금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtDocAmt.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐">
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>환율</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtXchRate.js'></script></TD>
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>B/L자국금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtLocAmt.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐">
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>운송방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransport" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="운송방법">&nbsp;<INPUT TYPE=TEXT NAME="txtTransportNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>수입자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수입자">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>선적항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>가격조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="가격조건">&nbsp;<INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="24"></TD></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>도착항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePort" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>영업그룹</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="영업그룹">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>					
														<TD CLASS=TD5 NOWRAP>선적일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime2_txtLoadingDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>수출자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운임지불방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFreight" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="운임지불방법">&nbsp;<INPUT TYPE=TEXT NAME="txtFreightNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>B/L발행통수</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtBLIssueCnt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>B/L발행장소</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtBLIssuePlce" ALT="B/L발행장소" TYPE=TEXT MAXLENGTH=30 SIZE=80 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>	
								<!-- 두번째 탭 내용 -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=5 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>대행자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="대행자">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>제조자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="제조자">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>VESSEL명</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL명" TYPE=TEXT MAXLENGTH=34 SIZE=34 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>항차번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVoyageNo" MAXLENGTH=20 SIZE=34 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>선박회사</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtForwarder" SIZE=10 MAXLENGTH=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtForwarderNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>선박국적</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVesselCntry" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtVesselCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수취장소</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtReceiptPlce" ALT="수취장소" TYPE=TEXT MAXLENGTH=35 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>인도장소</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtDeliveryPlce" ALT="인도장소" TYPE=TEXT MAXLENGTH=50 SIZE=80 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>최종목적지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="최종목적지" TYPE=TEXT MAXLENGTH=30 SIZE=34 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>도착예정일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime4_txtDischgeDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>환적국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTranshipCntry" SIZE=10 MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtTranshipCntryNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>환적일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime5_txtTranshipDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>포장조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="포장형태">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>총포장갯수</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtTotPackingCnt.js'></script></TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>포장참고사항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPackingTxt" ALT="포장참고사항" TYPE=TEXT MAXLENGTH=50 SIZE=34 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>컨테이너수</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtContainerCnt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>총중량</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtGrossWeight.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>중량단위</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWeightUnit" ALT="중량단위" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>총용적</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtGrossVolumn.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>용적단위</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVolumnUnit" ALT="용적단위" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>원산지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="원산지" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>원산지국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="원산지국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운임지불장소</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtFreightPlce" ALT="운임지불장소" TYPE=TEXT MAXLENGTH=30 SIZE=80 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>
								<!-- 세번째 탭 내용 -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=5 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizArea" ALT="세금신고사업장" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>매출채권형태</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillType" ALT="매출채권형태" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBillTypeNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수금처</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayer" ALT="수금처" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtPayerNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>발행처</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBilltoParty" ALT="발행처" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBilltoPartyNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>확정여부</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" TAG="24X" VALUE="Y" ID="rdoPostingflg1"><LABEL FOR="rdoPostingflg1">확정</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostingflg" VALUE="N" TAG="24X" CHECKED ID="rdoPostingflg2"><LABEL FOR="rdoPostingflg2">미확정</LABEL></TD>
														<TD CLASS=TD5 NOWRAP>수금만기일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDateTime8_txtPayDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수금영업그룹</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtToSalesGroup" ALT="수금영업그룹" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtToSalesGroupNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>B/L금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtDocAmt1.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐">
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>B/L자국금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>	
																<TR>
																	<TD><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtLocAmt1.js'></script></TD>
																	<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency1" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐">
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>총수금액</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5113ra9_fpDoubleSingle3_txtMoney.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>입금유형</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayType" ALT="입금유형" TYPE=TEXT MAXLENGTH=4 SIZE=10 TAG="24">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>결제방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="결재방법">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>결제기간</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayDur" ALT="결제기간" STYLE="TEXT-ALIGN: right" TYPE=TEXT MAXLENGTH=3 SIZE=5 TAG="24X7">&nbsp;일</TD>
														<TD CLASS=TD5 NOWRAP>결제조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermstxt" ALT="결제조건" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>비고</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark" ALT="비고" TYPE=TEXT MAXLENGTH=35 SIZE=80 TAG="24"></TD>
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
