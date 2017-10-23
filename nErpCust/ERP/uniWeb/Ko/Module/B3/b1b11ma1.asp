
<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : b1b11ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Item By Plant 등록 ASP													*
'*  6. Component List       : 																			*
'*  7. Modified date(First) : 2000/04/06																*
'*  8. Modified date(Last)  : 2004/03/19																*
'*  9. Modifier (First)     : Kim GyoungDon																*
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "b1b11ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim BaseDate, StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

Sub Form_Load()
	Call LoadInfTB19029																 '⊙: Load table , B_numeric_format 
	
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "4", "0")
	
    Call FormatDATEField(frm1.txtValidFromDt)
    Call FormatDATEField(frm1.txtValidToDt)
    
    Call FormatDoubleSingleField(frm1.txtCycleCntPerd)
    Call FormatDoubleSingleField(frm1.txtStdPrice)
    Call FormatDoubleSingleField(frm1.txtPrevStdPrice)
    Call FormatDoubleSingleField(frm1.txtMoveAvgPrice)
    Call FormatDoubleSingleField(frm1.txtReorderPoint)
    Call FormatDoubleSingleField(frm1.txtRoundPeriod)
    Call FormatDoubleSingleField(frm1.txtMfgOrderLT)
    Call FormatDoubleSingleField(frm1.txtPurOrderLT)
    
	Call LockObjectField(frm1.txtValidFromDt,"R")
	Call LockObjectField(frm1.txtValidToDt,"R")
	Call LockObjectField(frm1.txtStdPrice,"R")
	Call LockObjectField(frm1.txtCycleCntPerd,"R")
	Call LockObjectField(frm1.txtMfgOrderLT,"R")
	Call LockObjectField(frm1.txtPrevStdPrice,"P")
	Call LockObjectField(frm1.txtMoveAvgPrice,"P")
	Call LockObjectField(frm1.txtReorderPoint,"P")
	Call LockObjectField(frm1.txtRoundPeriod,"P")
	
	Call SetToolbar("11101000000011")												 '⊙: 버튼 툴바 제어 
	Call InitComboBox
	Call SetDefaultVal
	Call InitVariables
	 'Plant Code, Plant Name Setting 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
	Call SetCookieVal
	Call ggoOper.SetReqAttr(frm1.txtReorderPoint, "Q")		'회색 
End Sub

Sub InitComboBox()
	   'Call SetCombo(frm1.cboABCFlg,"A","A")
	   'Call SetCombo(frm1.cboABCFlg,"B","B")
	   'Call SetCombo(frm1.cboABCFlg,"C","C")
	'ABC FLAG SEARCH B_MINOR 2005-03-18 LSW
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("I1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboABCFlg, lgF0, lgF0, Chr(11))
	   
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1018", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboPrcCtrlInd, lgF0, lgF1, Chr(11))
	    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))
			
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboProcType, lgF0, lgF1, Chr(11))
	    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1008", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboMatType, lgF0, lgF1, Chr(11))
			
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1004", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboProdEnv, lgF0, lgF1, Chr(11))
	    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboOrderFrom, lgF0, lgF1, Chr(11))
			
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1005", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboLotSizing, lgF0, lgF1, Chr(11))
	    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1016", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboIssueType, lgF0, lgF1, Chr(11))

End Sub

</SCRIPT>
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공장별품목정보등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=88%>
		<TD WIDTH=100% CLASS="Tab11">
			<!-- 첫번째 탭 내용 -->
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" CLASS=required STYLE="text-transform:uppercase" SIZE=10 MAXLENGTH=4 tag="12XXXU"  ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtPlantNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" CLASS=required STYLE="text-transform:uppercase" SIZE=20 MAXLENGTH=18 tag="12XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConItemCd()">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtItemNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=25 tag="14" ALT="품목명"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR> <!-- Data Sheet -->
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=50% valign=top>
									<FIELDSET>			
										<LEGEND>기본정보</LEGEND>
											<TABLE CLASS="TB2" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>품목</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" CLASS=required STYLE="text-transform:uppercase" SIZE=25 MAXLENGTH=18 tag="23XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConItemCd1()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()"></TD>													
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>품목명</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" CLASS=protected READONLY=true TABINDEX="-1" SIZE=40 MAXLENGTH=40 tag="24" ALT="품목명"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>품목계정</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" CLASS=required ALT="계정" STYLE="Width: 168px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>조달구분</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProcType" CLASS=required ALT="조달구분" STYLE="Width: 168px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>자재Type</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboMatType" CLASS=required ALT="자재Type" STYLE="Width: 145px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>Tracking여부</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="21" ID="rdoTrackingItem1" VALUE="Y"><LABEL FOR="rdoTrackingItem1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoTrackingItem" tag="21" CHECKED ID="rdoTrackingItem2" VALUE="N"><LABEL FOR="rdoTrackingItem2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>유효기간</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/b1b11ma1_I228718996_txtValidFromDt.js'></script>	&nbsp;~&nbsp;
														<script language =javascript src='./js/b1b11ma1_I957478350_txtValidToDt.js'></script>					
													</TD>
												</TR>																								
											</TABLE>		
									</FIELDSET>
									<FIELDSET>	
										<LEGEND>재고정보</LEGEND>
											<TABLE CLASS="TB2" CELLSPACING=0>	
												<TR>
													<TD CLASS=TD5 NOWRAP>입고창고</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" CLASS=required STYLE="text-transform:uppercase" SIZE=15 MAXLENGTH=7 tag="22XXXU" ALT="입고창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>출고방법</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboIssueType" CLASS=required ALT="출고방법" STYLE="Width: 133px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>출고창고</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueSLCd" CLASS=required STYLE="text-transform:uppercase" SIZE=15 MAXLENGTH=7 tag="22XXXU" ALT="출고창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIssueSLCd()"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>출고단위</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIssueUnit" CLASS=required STYLE="text-transform:uppercase" SIZE=5 MAXLENGTH=3 tag="22XXXU"  ALT="출고단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrderUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIssueUnit()"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>Lot No.관리</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotNoFlg" tag="21" ID="rdoLotNoFlg1" VALUE="Y"><LABEL FOR="rdoLotNoFlg1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoLotNoFlg" tag="21" CHECKED ID="rdoLotNoFlg2" VALUE="N"><LABEL FOR="rdoLotNoFlg2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>재고실사주기</TD>
													<TD CLASS=TD6 NOWRAP>
														<TABLE CELLPADDING=0 CELLSPACING=0>
															<TR>
																<TD>
																	<script language =javascript src='./js/b1b11ma1_I494900238_txtCycleCntPerd.js'></script>
																</TD>
																<TD valign=bottom>&nbsp;일
																</TD>
															</TR>
														</TABLE>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>품목ABC구분</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboABCFlg" CLASS=required ALT="품목ABC구분" STYLE="Width: 98px;" tag="22"><OPTION Value=""></OPTION></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>단가구분</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboPrcCtrlInd" CLASS=required ALT="단가구분" STYLE="Width: 145px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>표준단가</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/b1b11ma1_I480284720_txtStdPrice.js'></script>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>전월표준단가</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/b1b11ma1_I624930834_txtPrevStdPrice.js'></script>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>이동평균단가</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/b1b11ma1_I507343810_txtMoveAvgPrice.js'></script>
													</TD>
												</TR>
											</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET>	
										<LEGEND>계획정보</LEGEND>
											<TABLE CLASS="TB2" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>생산전략</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdEnv" CLASS=required ALT="생산전략" STYLE="Width: 140px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>MPS품목</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMPSItem" tag="21" CHECKED ID="rdoMPSItem1" VALUE="Y"><LABEL FOR="rdoMPSItem1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMPSItem" tag="21" ID="rdoMPSItem2" VALUE="N"><LABEL FOR="rdoMPSItem2">아니오</LABEL></TD>
												</TR>	
												<TR>
													<TD CLASS=TD5 NOWRAP>오더생성여부</TD>
													<TD CLASS=TD6 NOWRAP>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMRPFlg" tag="21" CHECKED ID="rdoMRPFlg1" VALUE="Y"><LABEL FOR="rdoMRPFlg1">예</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMRPFlg" tag="21" ID="rdoMRPFlg2" VALUE="N"><LABEL FOR="rdoMRPFlg2">아니오</LABEL></TD>
												</TR>												
												<TR>
													<TD CLASS=TD5 NOWRAP>오더생성구분</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboOrderFrom" CLASS=protected READONLY=true TABINDEX="-1" ALT="오더생성구분" STYLE="Width: 140px;" tag="24"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>발주점</TD>
													<TD CLASS=TD6 NOWRAP>
														<script language =javascript src='./js/b1b11ma1_I754800927_txtReorderPoint.js'></script>												
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>Lot Sizing</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="cboLotSizing" CLASS=required ALT="Lot Sizing" STYLE="Width: 168px;" tag="22"></SELECT></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>올림기간</TD>
													<TD CLASS=TD6 NOWRAP>
														<TABLE CELLSPACING=0 CELLPADDING=3>													
															<TR>
																<TD>																
																	<script language =javascript src='./js/b1b11ma1_I497091442_txtRoundPeriod.js'></script>
																</TD>
																<TD valign=bottom>&nbsp;일
																</TD>
															</TR>
														</TABLE>
													</TD>
												</TR>
											</TABLE>
									</FIELDSET>
									<FIELDSET>
										<LEGEND>오더정보</LEGEND>	
											<TABLE CLASS="TB2" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>단공정여부</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoCollectFlg" tag="21" ID="rdoCollectFlg1" VALUE="Y"><LABEL FOR="rdoCollectFlg1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoCollectFlg" tag="21" CHECKED ID="rdoCollectFlg2" VALUE="N"><LABEL FOR="rdoCollectFlg2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>작업장</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWorkCenter" CLASS=protected  READONLY=true TABINDEX="-1" SIZE=15 MAXLENGTH=7 tag="24XXXU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWorkCenter" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenWorkCenter()">&nbsp;<INPUT TYPE=HIDDEN NAME="txtWcNm" SIZE=40 tag="24" ALT="주작업장명"></TD>													
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>제조오더단위</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMfgOrderUnit" CLASS=required STYLE="text-transform:uppercase" SIZE=5 MAXLENGTH=3 tag="22XXXU"  ALT="제조오더단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMfgOrderUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMfgUnit()"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>제조오더 L/T</TD>
													<TD CLASS=TD6 NOWRAP>
														<TABLE CELLSPACING=0 CELLPADDING=1>													
															<TR>
																<TD>
																	<script language =javascript src='./js/b1b11ma1_I209984302_txtMfgOrderLT.js'></script>
																</TD>
																<TD valign=bottom>&nbsp;일
																</TD>
															</TR>
														</TABLE>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>구매오더단위</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrderUnit"  STYLE="text-transform:uppercase" SIZE=5 MAXLENGTH=3 tag="21XXXU"  ALT="구매오더단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrderUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurUnit()"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>구매오더 L/T</TD>
													<TD CLASS=TD6 NOWRAP>
														<TABLE CELLSPACING=0 CELLPADDING=0>													
															<TR>
																<TD>
																	<script language =javascript src='./js/b1b11ma1_I370762509_txtPurOrderLT.js'></script>
																</TD>
																<TD valign=bottom>&nbsp;일
																</TD>
															</TR>															
														</TABLE>
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>구매조직</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg"  STYLE="text-transform:uppercase" SIZE=15 MAXLENGTH=4 tag="21XXXU" ALT="구매조직"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurOrg()">&nbsp;<INPUT TYPE=HIDDEN NAME="txtPurOrgNm" SIZE=30 tag="24"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>공정검사여부</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMfgInspType" tag="21" ID="rdoMfgInspType1" VALUE="Y"><LABEL FOR="rdoMfgInspType1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoMfgInspType" tag="21" CHECKED ID="rdoMfgInspType2" VALUE="N"><LABEL FOR="rdoMfgInspType2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>수입검사여부</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPurInspType" tag="21" ID="rdoPurInspType1" VALUE="Y"><LABEL FOR="rdoPurInspType1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPurInspType" tag="21" CHECKED ID="rdoPurInspType2" VALUE="N"><LABEL FOR="rdoPurInspType2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>최종검사여부</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFinalInspType" tag="21" ID="rdoFinalInspType1" VALUE="Y"><LABEL FOR="rdoFinalInspType1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoFinalInspType" tag="21" CHECKED ID="rdoFinalInspType2" VALUE="N"><LABEL FOR="rdoFinalInspType2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>출하검사여부</TD>
													<TD CLASS=TD6 NOWRAP>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueInspType" tag="21" ID="rdoIssueInspType1" VALUE="Y"><LABEL FOR="rdoIssueInspType1">예</LABEL>
																<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueInspType" tag="21" CHECKED ID="rdoIssueInspType2" VALUE="N"><LABEL FOR="rdoIssueInspType2">아니오</LABEL></TD>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:PlantItemDetail">공장별품목상세정보</A>&nbsp;|&nbsp;<A href="vbscript:AltItem">대체품 등록</A>&nbsp;|&nbsp;<A href="vbscript:LotControl">로트 관리</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPhantomFlg" TAG="24"><INPUT TYPE=HIDDEN NAME="txtBasicUnit" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
