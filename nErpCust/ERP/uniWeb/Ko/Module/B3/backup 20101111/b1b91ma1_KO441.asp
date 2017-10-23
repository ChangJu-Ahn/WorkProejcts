<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b01ma1.asp
'*  4. Program Name         : Entry Item
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/25
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : Kim Gyoung-Don
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="b1b01ma1_KO441.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

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
	
	Call LoadInfTB19029
	
    Call FormatDATEField(frm1.txtValidFromDt)
    Call FormatDATEField(frm1.txtValidToDt)
    
    Call FormatDoubleSingleField(frm1.txtWeight)
    Call FormatDoubleSingleField(frm1.txtGrossWeight)
    Call FormatDoubleSingleField(frm1.txtCBM)
    Call FormatDoubleSingleField(frm1.txtVatRate)
    
    Call LockObjectField(frm1.txtValidFromDt, "R")
    Call LockObjectField(frm1.txtValidToDt, "R")
    Call LockObjectField(frm1.txtVatRate, "P")
    Call LockHTMLField(frm1.rdoPhoto1,"P")
    Call LockHTMLField(frm1.rdoPhoto2,"P")


    '----------  Coding part  -------------------------------------------------------------
    Call SetCookieVal
    
    Call SetToolbar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal    
	Call InitVariables																'⊙: Initializes local global variables
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement
	
End Sub

Sub InitComboBox()
    On Error Resume Next
    Err.Clear
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemClass, lgF0, lgF1, Chr(11))
    
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목정보등록</font></td>
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
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 CLASS=required STYLE="text-transform:uppercase" tag="12XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=50 tag="14"></TD>
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
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=50%  valign=top>
									<FIELDSET>
										<LEGEND>일반정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" CLASS=required STYLE="text-transform:uppercase" SIZE=25 MAXLENGTH=18 tag="23XXXU" ALT="품목"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목명</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" CLASS=required SIZE=40 MAXLENGTH=40 tag="22" ALT="품목명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목정식명칭</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemDesc" SIZE=40 MAXLENGTH=60 tag="21" ALT="품목정식명칭"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>단위</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUnit" CLASS=required STYLE="text-transform:uppercase" SIZE=5 MAXLENGTH=3 tag="22XXXU" ALT="단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenUnit()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목계정</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" CLASS=required STYLE="text-transform:uppercase; Width: 168px;" ALT="품목계정" tag="22"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목그룹</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" STYLE="text-transform:uppercase" SIZE=20 MAXLENGTH=10 tag="21XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroup" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목그룹명</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=40 tag="24" ALT="품목그룹명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Phantom구분</TD>
												<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE="RADIO" NAME="rdoPhantomType" ID="rdoPhantomType1" Value="Y" CLASS="RADIO" tag="2X"><LABEL FOR="rdoPhantomType1">예</LABEL>
															<INPUT TYPE="RADIO" NAME="rdoPhantomType" ID="rdoPhantomType2" Value="N" CLASS="RADIO" tag="2X" CHECKED><LABEL FOR="rdoPhantomType2">아니오</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>통합구매구분</TD>
												<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg1" Value="Y" CLASS="RADIO" tag="2X"><LABEL FOR="rdoUnifyPurFlg1">예</LABEL>
															<INPUT TYPE="RADIO" NAME="rdoUnifyPurFlg" ID="rdoUnifyPurFlg2" Value="N" CLASS="RADIO" tag="2X" CHECKED><LABEL FOR="rdoUnifyPurFlg2">아니오</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>기준품목</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasisItemCd" STYLE="text-transform:uppercase" SIZE=25 MAXLENGTH=18 tag="21XXXU" ALT="기준품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBasisItemCd()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>기준품목명</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasisItemNm" CLASS=protected READONLY=true TABINDEX="-1" SIZE=40 tag="24" ALT="기준품목명"></TD>
											</TR>		
											<TR>
												<TD CLASS=TD5 NOWRAP>집계용품목클래스</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemClass" ALT="집계용품목클래스" STYLE="Width: 168px;" tag="21"><OPTION VALUE=""></OPTION></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효구분</TD>
												<TD CLASS=TD6 NOWRAP>
															<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg1" Value="Y" CLASS="RADIO" tag="2X" CHECKED><LABEL FOR="rdoValidFlg1">예</LABEL>
															<INPUT TYPE="RADIO" NAME="rdoValidFlg" ID="rdoValidFlg2" Value="N" CLASS="RADIO" tag="2X"><LABEL FOR="rdoValidFlg2">아니오</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/b1b01ma1_I714172488_txtValidFromDt.js'></script> &nbsp;~&nbsp;
													<script language =javascript src='./js/b1b01ma1_I235279250_txtValidToDt.js'></script>																
												</TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET>
										<LEGEND>품목규격정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목규격</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec" SIZE=40 MAXLENGTH=50 tag="21" ALT="품목규격"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Net중량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/b1b01ma1_I740263408_txtWeight.js'></script>&nbsp;
													<INPUT TYPE=TEXT NAME="txtWeightUnit" SIZE=5 MAXLENGTH=3 tag="21XXXU" ALT="중량단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenWeightUnit()">
													</OBJECT>												
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Gross중량</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/b1b01ma1_I164147603_txtGrossWeight.js'></script>&nbsp;
													<INPUT TYPE=TEXT NAME="txtGrossWeightUnit" SIZE=5 MAXLENGTH=3 tag="21XXXU" ALT="중량단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrossWeightUnit" align = top TYPE="BUTTON"ONCLICK="vbscript:OpenGrossWeightUnit()">
													
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>CBM(부피)</TD>
												<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b1b01ma1_I144735008_txtCBM.js'></script></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>MES 품목코드</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCBMInfo" CLASS=required STYLE="text-transform:uppercase" SIZE=25 MAXLENGTH=50 tag="23XXXU" style="background:#FFE5CB"  ALT="MES 품목코드"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>도면번호</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDrawNo" SIZE=25 MAXLENGTH=20 tag="21" ALT="도면번호"></TD>
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<LEGEND>기타</LEGEND>
											<TABLE CLASS="TB2" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>HS코드</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHSCd" STYLE="text-transform:uppercase" SIZE=20 MAXLENGTH=20 tag="21XXXU" ALT="HS코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnHsCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenHsCd()"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>HS단위</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtHSUnit" CLASS=protected READONLY=true TABINDEX="-1" SIZE=5 MAXLENGTH=3 tag="24"  ALT="HS단위"></TD>
												</TR>
												<TR>	
													<TD CLASS=TD5 NOWRAP>사진유무</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPhoto" ID="rdoPhoto1" Value="Y" CLASS="RADIO" tag="24X"><LABEL FOR="rdoPhoto1">예</LABEL>
												 						 <INPUT TYPE="RADIO" NAME="rdoPhoto" ID="rdoPhoto2" Value="N" CLASS="RADIO" tag="24X" CHECKED><LABEL FOR="rdoPhoto2">아니오</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>VAT유형</TD>
													<TD CLASS=TD6 NOWRAP>
														<INPUT NAME="txtVatType" STYLE="text-transform:uppercase" TYPE="Text"  MAXLENGTH="5" SIZE=10  ALT="VAT유형" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBillHdr">
														<INPUT NAME="txtVatTypeNm" CLASS=protected READONLY=true TABINDEX="-1" TYPE="Text" MAXLENGTH="25" SIZE=25 tag="24">
													</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>VAT율</TD>
													<TD CLASS=TD6 NOWRAP>
														<TABLE CELLSPACING=0 CELLPADDING=0>
															<TR>
																<TD>
																	<script language =javascript src='./js/b1b01ma1_I440002822_txtVatRate.js'></script>
																	&nbsp;<LABEL><b>%</b></LABEL>
																</TD>
																
															</TR>
														</TABLE>
													</TD>
												</TR>			
												<TR>
													<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
													<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												</TR>			
												<TR>
													<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
													<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
													<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpItemImage">품목사진등록</A>&nbsp;|&nbsp;<A href="vbscript:JumpItemByPlant">공장별 품목등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUsrId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtItemByPlantFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hCBMInfo" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
