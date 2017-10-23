<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: Production Order Management
'*  3. Program ID			: p4111ma1.asp
'*  4. Program Name			: Manage Production Order (Single)
'*  5. Program Desc			: Create, Update, Delete Production Order
'*  6. Comproxy List		: 
'*     Biz Aps  List		: +p4111mb1.asp		LookUp Production Order Header
'*							  +p4111mb2.asp		Manage Production Order
'*							  +p4111mb3.asp		LookUp Item By Plant
'*							  +p4111mb4.asp		Release Production Order
'*							  +p2350ma1.asp		Order Explosion		
'*  7. Modified date(First)	: 2000/03/29
'*  8. Modified date(Last)	: 2002/07/09
'*  9. Modifier (First)		: Kim, Gyoung-Don
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'					1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우 -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4111ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit									 '☜: indicates that All variables must be declared in advance 

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

'========================================= 2.1.2 LoadInfTB19029() ==================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================  2.1.3 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : ComboBox에 데이타 Setting
'====================================================================================================
Sub InitComboBox()

    Call SetCombo(frm1.cboReWork, "N", "작업")
    Call SetCombo(frm1.cboReWork, "Y", "재작업")		'⊙: InitCombo 에서 해야 되는데 임시로 넣은 것임 
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P3211", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtOrderType, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtStatus, lgF0, lgF1, Chr(11))
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1401", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtBOMNo, lgF0, lgF1, Chr(11))

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1015", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtProdMgr, lgF0, lgF1, Chr(11))

	frm1.txtOrderType.value = "" 
	frm1.txtStatus.value = ""

End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	
    Call LoadInfTB19029															'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("11101000000011")
    
    Call InitVariables		'⊙: Initializes local global variables
    Call InitComboBox
	Call SetDefaultVal
	
	If ReadCookie("txtPGMID") <> "" Then
		frm1.txtPlantCd.Value		= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value		= ReadCookie("txtPlantNm")
		frm1.txtProdOrderNo.value	= ReadCookie("txtProdOrderNo")
		WriteCookie "txtPlantCd", ""
		WriteCookie "txtPlantNm", ""
		WriteCookie "txtProdOrderNo", ""
		WriteCookie "txtPGMID", ""
		Call LookUpInvClsDt
	Else
		If ReadCookie("txtPlantCd") <> "" Then
			lgReworkMode = "Y"
			Call SetCookieVal
			If Trim(frm1.txtTrackingNo) = "*" Or Trim(frm1.txtTrackingNo) = "" Then
				Call ggoOper.SetReqAttr(frm1.txtTrackingNo, "Q")
			Else
				Call ggoOper.SetReqAttr(frm1.txtTrackingNo, "N")
			End If
			Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"D")
			Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"D")
			'Call txtPlantCd_OnChange
		Else		
			If parent.gPlant <> "" Then
				frm1.txtPlantCd.value = parent.gPlant
				frm1.txtPlantNm.value = parent.gPlantNm
				frm1.txtProdOrderNo.focus
				Set gActiveElement = document.activeElement
				Call LookUpInvClsDt
				Call txtPlantCd_OnChange
			Else
				frm1.txtPlantCd.focus 
				Set gActiveElement = document.activeElement
			End If
			Call ggoOper.SetReqAttr(frm1.txtTrackingNo, "Q")
			Call ggoOper.SetReqAttr(frm1.txtParentOrderNo,"Q")
			Call ggoOper.SetReqAttr(frm1.txtParentOprNo,"Q")	
		End If
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>제조오더관리(Single)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenStockRef()">재고현황</A> | <A href="vbscript:OpenPartRef()">부품내역</A> | <A href="vbscript:OpenOprRef()">공정내역</A></TD>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				    <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100% colspan=2>
									<FIELDSET valign=top>
										<LEGEND>일반정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo1" SIZE=18 MAXLENGTH=18 tag="21xxxU" ALT="제조오더번호"></TD>
												<TD CLASS=TD5 NOWRAP>지시상태</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="txtStatus" ALT="지시상태" STYLE="Width: 98px;" tag="24"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="23xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="24"></TD>
												<TD CLASS=TD5 NOWRAP>라우팅</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRouting" SIZE=12 MAXLENGTH=7 tag="23xxxU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCtr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutingNo()"></TD>												
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>규격</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpecification" SIZE=40 MAXLENGTH=50 tag="24" ALT="규격"></TD>
												<TD CLASS=TD5 NOWRAP>비고</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRemark" SIZE=30 MAXLENGTH=20 tag="21" ALT="비고"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>오더수량</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtOrderQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="22X3Z" ALT="오더수량" MAXLENGTH="15" SIZE="10" id=fpDoubleSingle2></OBJECT>');</SCRIPT>
												</TD>
												<TD CLASS=TD5 NOWRAP>작업지시 C/C</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=17 MAXLENGTH=10 tag="23XXXU" ALT="작업지시 C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCtr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCtr()">&nbsp;<INPUT NAME="txtCostNm" MAXLENGTH="20" SIZE=30 ALT ="코스트센타명" tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>오더단위</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUnit" SIZE=5 MAXLENGTH=3 tag="23xxxU" ALT="오더단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnit" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenUnit()"></TD>
												<TD CLASS=TD5 NOWRAP>BOM Type</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="txtBOMNo" ALT="BOM Type" STYLE="Width: 98px;" tag="24"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>기준수량</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBaseOrderQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X3" ALT="기준수량" MAXLENGTH="15" SIZE="10" id=fpDoubleSingle1></OBJECT>');</SCRIPT>
												</TD>
												<TD CLASS=TD5 NOWRAP>재작업</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboReWork" ALT="재작업" STYLE="Width: 98px;" tag="22"><OPTION VALUE=""></OPTION></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>기준단위</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBaseUnit" SIZE=5 MAXLENGTH=3 tag="24xxxU" ALT="재고단위"></TD>
												<TD CLASS=TD5 NOWRAP>상위오더번호</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtParentOrderNo" SIZE=18 MAXLENGTH=18 tag="22xxxU" ALT="상위오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnParentOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenParentOrderNo()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>창고</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=15 MAXLENGTH=7 tag="23xxxU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 tag="24" ALT="창고명"></TD>
												<TD CLASS=TD5 NOWRAP>상위공정</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtParentOprNo" SIZE=10 MAXLENGTH=3 tag="22xxxU" ALT="상위공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnParentOpr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenParentOprNo()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="22xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNoBtn" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
												<TD CLASS=TD5 NOWRAP>지시구분</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="txtOrderType" ALT="지시구분" STYLE="Width: 98px;" tag="24"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP></TD>
												<TD CLASS=TD6 NOWRAP></TD>
												<TD CLASS=TD5 NOWRAP>계획오더번호</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlanOrderNo" SIZE=18 MAXLENGTH=12 tag="24" ALT="계획오더번호"></TD>
											</TR>
										</TABLE>
									</FIELDSET>	
								</TD>		
							</TR>	
							<TR>
								<TD WIDTH=50% valign=top>
									<FIELDSET>
										<LEGEND>일정</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>착수예정일</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPlanStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="착수예정일"></OBJECT>');</SCRIPT>
													</TD>
												</TR>												
												<TR>
													<TD CLASS=TD5 NOWRAP>완료예정일</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPlanEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="완료예정일"></OBJECT>');</SCRIPT>
													</TD>
												</TR>
												<TR>	
													<TD CLASS=TD5 NOWRAP>착수계획일정</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPlannedStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="착수계획일정" ></OBJECT>');</SCRIPT>
												</TR>
												<TR>	
													<TD CLASS=TD5 NOWRAP>완료계획일정</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPlannedEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="완료계획일정" ></OBJECT>');</SCRIPT>
												</TR>
												<TR>	
													<TD CLASS=TD5 NOWRAP>작업지시일</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtReleaseDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="24" ALT="작업지시일"></OBJECT>');</SCRIPT>
												</TR>
											</TABLE>
									</FIELDSET>
								</TD>
								<TD WIDTH=50% valign=top>
									<FIELDSET>
										<LEGEND>품목생산 기준정보</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>제조 L/T</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdLT" SIZE=10 STYLE="TEXT-ALIGN: right" tag="24" ALT="제조 L/T"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>최대LOT수</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtMaxLotQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X3" ALT="최대LOT수" MAXLENGTH="15" SIZE="10" ></OBJECT>');</SCRIPT>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>최소LOT수</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtMinLotQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X3" ALT="최소LOT수" MAXLENGTH="15" SIZE="10" ></OBJECT>');</SCRIPT>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>올림수</TD>
													<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRoundingQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X3" ALT="올림수" MAXLENGTH="15" SIZE="10" ></OBJECT>');</SCRIPT>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>생산담당자</TD>
													<TD CLASS=TD6 NOWRAP><SELECT NAME="txtProdMgr" ALT="생산담당자" STYLE="Width: 98px;" tag="24"></SELECT></TD>
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
					<TD><BUTTON NAME="btnRelease" ONCLICK="vbscript:ReleaseOrder()" CLASS="CLSMBTN">제조오더확정</BUTTON></TD>
					<TD WIDTH=* Align=right><A href="vbscript:JumpOrderRun()">제조오더전개</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hOprCostFlag" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
