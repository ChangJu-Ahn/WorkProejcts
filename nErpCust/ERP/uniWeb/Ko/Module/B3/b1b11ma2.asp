<%@ LANGUAGE="VBScript" %>												   
<!--'****************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : b1b11ma2.asp																*
'*  4. Program Name         : Item By Plant 조회 ASP													*
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
'#						6. TAG 부																		#
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공장별품목조회</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbScript:OpenConItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14" ALT="품목명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목계정</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" ALT="계정" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>조달구분</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProcType" ALT="조달구분" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>종료일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="11X1"> </OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="11X1"> </OBJECT>');</SCRIPT>					
									</TD>
									<TD CLASS=TD5 NOWRAP>유효구분</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailableItem" tag="1X" CHECKED ID="rdoAvailableItem1" VALUE="A"><LABEL FOR="rdoAvailableItem1">전체</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailableItem" tag="1X" ID="rdoAvailableItem2" VALUE="Y"><LABEL FOR="rdoAvailableItem2">예</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailableItem" tag="1X" ID="rdoAvailableItem3" VALUE="N"><LABEL FOR="rdoAvailableItem3">아니오</LABEL></TD>
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
																	<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목일반정보</font></td>
																	<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
															    </TR>
															</TABLE>
														</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
																<TR>
																	<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP기준정보</font></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
															    </TR>
															</TABLE>
														</TD>
														<TD CLASS="CLSMTABP">
															<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
																<TR>
																	<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
																	<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고/품질정보</font></td>
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
												<!-- 첫번째 탭 내용 -->
												<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR> <!-- Data Sheet -->
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>																				
																	<TR>
																		<TD CLASS=TD5 NOWRAP>품목</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="24" ALT="품목"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>품목명</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=50 tag="24" ALT="품목명"></TD>													
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>품목계정</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAccount" ALT="품목계정" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																		<TD CLASS=TD5 NOWRAP>품목규격</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=50 tag="24" ALT="품목규격"></TD>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>기준단위</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasicUnit" ALT="기준단위" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>집계용품목클래스</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemClass" ALT="집계용클래스" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>조달구분</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProcType" ALT="조달구분" SIZE=20 tag="24"></SELECT></TD>
																	</TR>																				
																	<TR>
																		<TD CLASS=TD5 NOWRAP>생산전략</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdEnv" ALT="생산전략" SIZE=20 tag="24"></SELECT></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>MPS품목</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMPSItem" tag="24" ID="rdoMPSItem1" VALUE="Y"><LABEL FOR="rdoMPSItem1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMPSItem" tag="24" ID="rdoMPSItem2" VALUE="N"><LABEL FOR="rdoMPSItem2">아니오</LABEL></TD>
																	</TR>												
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Tracking여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="24" ID="rdoTrackingItem1" VALUE="Y"><LABEL FOR="rdoTrackingItem1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="24" ID="rdoTrackingItem2" VALUE="N"><LABEL FOR="rdoTrackingItem2">아니오</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>단공정여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoCollectFlg" tag="24" ID="rdoCollectFlg1" VALUE="Y"><LABEL FOR="rdoCollectFlg1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoCollectFlg" tag="24" ID="rdoCollectFlg2" VALUE="N"><LABEL FOR="rdoCollectFlg2">아니오</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>작업장</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWorkCenter" SIZE=20 MAXLENGTH=7 tag="24" ALT="작업장"></TD>													
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>유효구분</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailable" tag="24" ID="rdoAvailable1" VALUE="Y"><LABEL FOR="rdoAvailable1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAvailable" tag="24" ID="rdoAvailable2" VALUE="N"><LABEL FOR="rdoAvailable2">아니오</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>표준ST</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT  name=txtStdTime SIZE=20 tag="24" ALT="표준ST" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>ATP L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT  name=txtAtpLt SIZE=20 tag="24" ALT="ATP L/T" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>유효기간</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidFromDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="시작일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>																																			&nbsp;~&nbsp;
																			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidToDt CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="종료일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
																		</TD>
																	</TR>																	
																</TABLE>										
															</TD>
														</TR>
													</TABLE>
												</DIV>
												<!-- 두번째 탭 내용 -->
												<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no> 
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR> <!-- Data Sheet -->
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB3" CELLSPACING=0>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>오더생성여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																			<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMRPFlg" tag="24" ID="rdoMRPFlg1" VALUE="Y"><LABEL FOR="rdoMRPFlg1">예</LABEL>
																			<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMRPFlg" tag="24" ID="rdoMRPFlg2" VALUE="N"><LABEL FOR="rdoMRPFlg2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>													
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>오더생성구분</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderFrom" ALT="오더생성구분" SIZE=20 tag="24" ></TD>													
																		<TD CLASS=TD5 NOWRAP>Lot Sizing</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLotSizing" ALT="Lot Sizing" SIZE=20 tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>소요량올림구분</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRoundFlg" tag="24" ID="rdoRoundFlg1" VALUE="Y"><LABEL FOR="rdoRoundFlg1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRoundFlg" tag="24" ID="rdoRoundFlg2" VALUE="N"><LABEL FOR="rdoRoundFlg2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>올림기간</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtRoundPeriod SIZE=20 ALT="올림기간" tag="24" STYLE="TEXT-ALIGN: right"></TD>												
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>최대오더수량</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMaxOrderQty SIZE=20 ALT="최대오더수량" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>분할 L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtOffsetLt SIZE=20 ALT="분할 L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>최소오더수량</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMinOrderQty SIZE=20 ALT="최소오더수량" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>올림수</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtRoundQty SIZE=20 ALT="올림수" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>고정오더수량</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtFixOrderQty SIZE=20 ALT="고정오더수량" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>분할수</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtLineNo SIZE=20 ALT="라인수" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>칼렌다타입</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtCalType SIZE=5 ALT="칼렌다타입" tag="24"></TD>
																		<TD CLASS=TD5 NOWRAP>MRP 담당자</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMRPMgr" ALT="MRP 담당자" SIZE=20 tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>생산담당자</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdMgr" ALT="생산담당자" SIZE=20 tag="24"></TD>													
																		<TD CLASS=TD5 NOWRAP>가변 L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtVarLT SIZE=20 ALT="가변 L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>																					
																	<TR>																					
																		<TD CLASS=TD5 NOWRAP>Damper여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDamperFlg" tag="24" ID="rdoDamperFlg1" VALUE="Y"><LABEL FOR="rdoDamperFlg1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDamperFlg" tag="24" ID="rdoDamperFlg2" VALUE="N"><LABEL FOR="rdoDamperFlg2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>Damper 최소율</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtDamperMinQty SIZE=20 ALT="Damper 최소율" tag="24" STYLE="TEXT-ALIGN: right"></TD>											
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>제조오더단위</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMfgOrderUnit" SIZE=5  tag="24"  ALT="제조오더단위"></TD>
																		<TD CLASS=TD5 NOWRAP>구매오더단위</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrderUnit" SIZE=5 MAXLENGTH=3 tag="24"  ALT="구매오더단위"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>제조오더 L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMfgOrderLT SIZE=20 ALT="제조오더 L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>구매오더 L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtPurOrderLT SIZE=20 ALT="제조오더 L/T" tag="24" STYLE="TEXT-ALIGN: right"></TD>												
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>제조품목불량율</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMfgScrapRate SIZE=20 ALT="제조품목불량율" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>구매품목불량율</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtPurScrapRate ALT="구매불량율" tag="24" size= 20 STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD5 NOWRAP>구매조직</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=20 MAXLENGTH=18 tag="24" ALT="구매조직"></TD>
																	</TR>
																				
																</TABLE>								
															</TD>
														</TR>
													</TABLE>
												</DIV>
												<!-- 세번째 탭 내용 -->
												<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
													<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
														<TR> <!-- Data Sheet -->
															<TD WIDTH=100% HEIGHT=* valign=top>
																<TABLE CLASS="TB3" CELLSPACING=0>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>입고창고</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=20 MAXLENGTH=7 tag="24" ALT="입고창고"></TD>
																		<TD CLASS=TD5 NOWRAP>출고방법</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueType" ALT="출고방법" STYLE="aling: right;" tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>출고창고</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueSLCd" SIZE=20 MAXLENGTH=7 tag="24" ALT="출고창고"></TD>
																		<TD CLASS=TD5 NOWRAP>출고단위</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueUnit" SIZE=5 MAXLENGTH=3 tag="24"  ALT="오더단위"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>Lot No.관리</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotNoFlg" tag="24" ID="rdoLotNoFlg1" VALUE="Y"><LABEL FOR="rdoLotNoFlg1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotNoFlg" tag="24" ID="rdoLotNoFlg2" VALUE="N"><LABEL FOR="rdoLotNoFlg2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>안전재고량</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtSFStockQty SIZE=20 ALT="안전재고량" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>발주점</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtReorderPnt SIZE=20 ALT="발주점" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>가용재고체크</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInvCheckFlg" tag="24" ID="rdoInvCheckFlg1" VALUE="Y"><LABEL FOR="rdoInvCheckFlg1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInvCheckFlg" tag="24" ID="rdoInvCheckFlg2" VALUE="N"><LABEL FOR="rdoInvCheckFlg2">아니오</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>과입고허용여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOverRcptFlg" tag="24" ID="rdoOverRcptFlg1" VALUE="Y"><LABEL FOR="rdoOverRcptFlg1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoOverRcptFlg" tag="24" ID="rdoOverRcptFlg2" VALUE="N"><LABEL FOR="rdoOverRcptFlg2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>과입고허용율</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtOverRcptRate SIZE=20 ALT="과입고허용율" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>재고실사주기</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtCycleCntPerd SIZE=20  tag="24" ALT="재고실사주기" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>품목ABC구분</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT  NAME="txtABCFlg" SIZE=5 ALT="품목ABC구분" tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>재고담당자</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInvMgr" SIZE=20 ALT="재고담당자" tag="24"></TD>
																		<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
																		<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>수입검사여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPurInspType" tag="24" ID="rdoPurInspType1" VALUE="Y"><LABEL FOR="rdoPurInspType1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPurInspType" tag="24" ID="rdoPurInspType2" VALUE="N"><LABEL FOR="rdoPurInspType2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>공정검사여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMfgInspType" tag="24" ID="rdoMfgInspType1" VALUE="Y"><LABEL FOR="rdoMfgInspType1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMfgInspType" tag="24" ID="rdoMfgInspType2" VALUE="N"><LABEL FOR="rdoMfgInspType2">아니오</LABEL></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>최종검사여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFinalInspType" tag="24" ID="rdoFinalInspType1" VALUE="Y"><LABEL FOR="rdoFinalInspType1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFinalInspType" tag="24" ID="rdoFinalInspType2" VALUE="N"><LABEL FOR="rdoFinalInspType2">아니오</LABEL></TD>
																		<TD CLASS=TD5 NOWRAP>출하검사여부</TD>
																		<TD CLASS=TD6 NOWRAP>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueInspType" tag="24" ID="rdoIssueInspType1" VALUE="Y"><LABEL FOR="rdoIssueInspType1">예</LABEL>
																					<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueInspType" tag="24" ID="rdoIssueInspType2" VALUE="N"><LABEL FOR="rdoIssueInspType2">아니오</LABEL></TD>
																	</TR>
				     												<TR>
																		<TD CLASS=TD5 NOWRAP>제조검사 L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT SIZE=20 name=txtMfgInspLT tag="24" ALT="제조검사 L/T" STYLE="TEXT-ALIGN: right"></TD>
																		<TD CLASS=TD5 NOWRAP>구매검사 L/T</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT SIZE=20 name=txtPurInspLT tag="24" ALT="구매검사 L/T" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>제조시 검사담당자</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMPSMgr" ALT="MPS 담당자" SIZE=20 tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>
																		<TD CLASS=TD5 NOWRAP>구매시 검사담당자</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT SIZE=20  NAME="txtInspecMgr" ALT="검사담당자" tag="24"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>단가구분</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPrcCtrlInd" ALT="단가구분" SIZE=15 tag="24"></TD>
																		<TD CLASS=TD5 NOWRAP>표준단가</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtStdPrice SIZE=20 ALT="표준단가" tag="24" STYLE="TEXT-ALIGN: right"></TD>
																	</TR>
																	<TR>
																		<TD CLASS=TD5 NOWRAP>이동평균단가</TD>
																		<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT name=txtMoveAvgPrice SIZE=20 ALT="이동평균단가" tag="24" STYLE="TEXT-ALIGN: right"></TD>											
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
