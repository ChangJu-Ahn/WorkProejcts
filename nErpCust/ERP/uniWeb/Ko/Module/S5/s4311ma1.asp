<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 smj
'*  2. Function Name        : 예외출고/반품등록 
'*  3. Program ID           : S4311MA1
'*  4. Program Name         : 예외출고/반품등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S31111MaintSoHdrSvr, S31119LookupSoHdrSvr
'*  7. Modified date(First) : 2002/03/22
'*  8. Modified date(Last)  : 2003/10/14
'*  9. Modifier (First)     : Sung MiJung
'* 10. Modifier (Last)      : Hwang seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/09 : ..........
'*                            -2000/05/09 : 표준수정사항적용 
'*                            -2000/09/04 : 4Th Coding
'*                            -2001/12/18 : Date 표준 적용 
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

Dim lblnWinEvent   '박정순 추가 
Dim interface_Account   '박정순 추가 


iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>예외출고/반품</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ChangeTabs(TAB2)">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수금 및 품목정보</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>     
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ChangeTabs(TAB3)">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납품 및 운송정보</font></td>
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
									<TD CLASS="TD5" NOWRAP>출하번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConDn_no" ALT="출하번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnNo"></TD>
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
						<!-- 첫번째 탭 내용 -->
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>출하번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtDnNo" ALT="출하번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDn_Type" ALT="출하형태" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="23XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDNType()">&nbsp;<INPUT NAME="txtDn_TypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>주문처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSold_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp 0" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtSold_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>판매유형</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeal_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 3">&nbsp;<INPUT NAME="txtDeal_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR> 
									 <TD CLASS=TD5 NOWRAP>납품처</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="납품처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp 1" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtShip_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD> 
									 <TD CLASS=TD5 NOWRAP>영업그룹</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtSales_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>          
								</TR>        
								<TR>         
									 <TD CLASS=TD5 NOWRAP>세금신고사업장</LABEL></TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxBizAreaCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRequried 4">&nbsp;<INPUT NAME="txtTaxBizAreaNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									 <TD CLASS=TD5 NOWRAP>결제방법</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_terms" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 5">&nbsp;<INPUT NAME="txtPay_terms_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									 <TD CLASS="TD5" NOWRAP>납기일</TD>
									 <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtDlvyDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="납기일"></OBJECT>');</SCRIPT></TD>
									 <TD CLASS="TD5" NOWRAP>출고예정일</TD>
									 <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtPlannedGIDt" CLASS=FPDTYYYYMMDD tag="22X1" Title="FPDATETIME" ALT="출고예정일"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>대금결제참조</TD>
									 <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txt_Payterms_txt" TYPE="Text" MAXLENGTH="120" SIZE=80 tag="21" ALT="대금결제참조"></TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>비고</TD>
									 <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="120" SIZE=80 tag="21" ALT="비고"></TD>
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>VAT유형</TD>
									 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtVat_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="VAT유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 6">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									 <TD CLASS=TD5 NOWRAP>VAT율</TD>
									 <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtVat_rate" ALT = "VAT율" CLASS=FPDS140 tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL></TD>         
								</TR>
								<TR>
									 <TD CLASS=TD5 NOWRAP>VAT포함구분</TD>
									 <TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoVat_Inc_flag" id="rdoVat_Inc_flag1" value="1" tag = "21" checked>
											<label for="rdoVat_Inc_flag1">별도</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoVat_Inc_flag" id="rdoVat_Inc_flag2" value="2" tag = "21">
											<label for="rdoVat_Inc_flag2">포함</label></TD>
									 <TD CLASS=TD5 NOWRAP>금액</TD>
									 <TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtNet_amt" CLASS=FPDS140 tag="24X2Z" ALT = "금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;&nbsp;
												</TD>
												<TD>
													<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="화폐">
												</TD>
											</TR>
										</TABLE>
										</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT적용기준</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoVat_Calc_Type" id="rdoVat_Calc_Type1" value="1" tag = "21" checked>
										 <label for="rdoVat_Calc_Type1">개별</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoVat_Calc_Type" id="rdoVat_Calc_Type2" value="2" tag = "21">
										 <label for="rdoVat_Calc_Type2">통합</label></TD>
									<TD CLASS=TD5 NOWRAP>VAT금액</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtVat_amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
									</TD>  
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_Meth" ALT="운송방법" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 4">&nbsp;<INPUT NAME="txtTrans_Meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>총금액</TD>
									<TD CLASS=TD6 NOWRAP>
									 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtTot_amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>       
									</TD>  
						        </TR> 
   								<TR>
									<TD CLASS=TD5 NOWRAP>실제납품일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtArriv_dt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="실제납품일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>납품시간</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtArriv_Tm" TYPE="Text" ALT="납품시간" MAXLENGTH="10" SIZE=36 tag="21"></TD>
								</TR>									
							        <%Call SubFillRemBodyTD5656(4)%>
							</TABLE>
						</DIV>
       
      
						<!-- 두번째 탭 내용 -->
						<DIV ID="TabDiv"  STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlant" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenPlant()">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtSlCd" ALT="창고" TYPE="Text" MAXLENGTH=7 SiZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenSl()">&nbsp;<INPUT NAME="txtSlNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR> 
								<TR>
									<TD CLASS=TD5 NOWRAP>재고담당자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvMgr" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="재고담당자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInvMgr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInvMgrPopUp">&nbsp;<INPUT NAME="txtInvMgrNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="재고담당자명"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>후속작업여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">매출채권</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="21" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">세금계산서</LABEL>
									</TD>
									<TD CLASS=TD5 NOWRAP>총금액</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtTotal_Amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;&nbsp;</TD>
								</TR>
							    <TR>
									<TD CLASS=TD5 NOWRAP>수금유형</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCol_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" ALT="수금유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRequried 2">&nbsp;<INPUT NAME="txtCol_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>수금액</TD>
									<TD CLASS=TD6 NOWRAP>
								        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtCol_amt" CLASS=FPDS140 tag="24X2Z" Title="FPDOUBLESINGLE" ALT="수금액"></OBJECT>');</SCRIPT>       
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>실제출고일</TD>
									<TD CLASS=TD6 NOWRAP>
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtGI_Dt" CLASS=FPDTYYYYMMDD tag="24X1" ALT="출고일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>출고번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtGINo" ALT="출고번호" TYPE="Text" MAXLENGTH=18 SiZE=22 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR> 
								<TR>
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
										 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>

							</TABLE>
						</DIV>
						
						<!-- 세번째 탭 내용 -->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처상세정보번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSTP_Inf_No" ALT="납품처상세정보번호" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<BUTTON NAME = "btnShipToPlceRef" CLASS="CLSMBTN">납품처상세정보참조</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>우편번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIP_cd" TYPE="Text" ALT="우편번호" MAXLENGTH="12" SIZE=20 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
									<TD CLASS=TD5 NOWRAP>인수자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiver" TYPE="Text" ALT="인수자명" MAXLENGTH="50" SIZE=36 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품주소</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1_Dlv" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2_Dlv" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR3_Dlv" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR> 
									 <TD CLASS=TD5 NOWRAP>납품장소</TD>
									 <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtDlvyPlace" TYPE="Text" MAXLENGTH="30" SIZE=91 ALT="납품장소" tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전화번호1</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No1" TYPE="Text" ALT="전화번호1" MAXLENGTH="20" SIZE=37 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>전화번호2</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No2" TYPE="Text" ALT="전화번호2" MAXLENGTH="20" SIZE=37 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운송정보번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrnsp_Inf_No" ALT="운송정보번호" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<BUTTON NAME = "btnTrnsMethRef" CLASS="CLSMBTN">운송정보참조</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5>운송회사</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtTransCo" SIZE=20 MAXLENGTH=50 TAG="21XXXX" ALT="운송회사"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransCo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTransCo()"></TD>
									<TD CLASS=TD5 NOWRAP>인계자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSender" TYPE="Text" ALT="인계자명" MAXLENGTH="50" SIZE=37 tag="21"></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5>차량번호</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 TAG="21XXXX" ALT="차량번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVehicleNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenVehicleNo()"></TD>							
									<TD CLASS=TD5 NOWRAP>운전자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDriver" TYPE="Text" ALT="운전자명" MAXLENGTH="50" SIZE=37 tag="21"></TD>
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
						<BUTTON NAME="btnPosting" CLASS="CLSMBTN">출고</BUTTON>&nbsp;
						<BUTTON NAME="btnPostCancel" CLASS="CLSMBTN">출고취소</BUTTON>&nbsp;
			         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>
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
<INPUT TYPE=HIDDEN NAME="txtArFlag" tag="24" TABINDEX="-1">	<!-- DB상태, 수금유형이 등록된 경우 'Y', 그렇지 않은 경우 'N' -->
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
