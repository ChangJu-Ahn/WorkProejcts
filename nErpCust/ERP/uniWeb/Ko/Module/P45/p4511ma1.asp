<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4511ma1
'*  4. Program Name			: 생산입고등록 
'*  5. Program Desc			:
'*  6. Comproxy List		: 
'*  7. Modified date(First)	: 2000/04/21
'*  8. Modified date(Last) 	: 2002/07/18
'*  9. Modifier (First) 	: Park, Bum Soo
'* 10. Modifier (Last)		: Kang Seong Moon
'* 11. Comment				:
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'      Park Kye Jin         : 착수계획일정/완료계획일정/실완료일 삭제(2003.04.07)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************** -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우 -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4511ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'☜: indicates that All variables must be declared in advance

Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)		'새로바뀐 날짜 형식											
StartDate = UNIDateAdd("D",-10,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 처음 날짜 
EndDate = UNIDateAdd("D", 20,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
		 
    Call ggoOper.LockField(Document, "Q")                                   '⊙: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet

	Call SetDefaultVal
    Call InitVariables		'⊙: Initializes local global variables
    
	Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
    
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
<!-- '#########################################################################################################
'       					6. Tag부 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>생산입고등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenOprRef()">공정내역</A> | <A href="vbscript:OpenProdRef()">실적내역</A> | <A href="vbscript:OpenRcptRef()">입고내역</A> | <A href="vbscript:OpenConsumRef()">부품소비내역</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>실적일</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/p4511ma1_I620418554_txtFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4511ma1_I921497106_txtToDt.js'></script>
									</TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14" ALT="품목명"></TD>								
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProdOrderNo() "></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value,0"></TD>								
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>입고창고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="입고창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=25 tag="14" ALT="입고창고명"></TD>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConWC()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="14" ALT="작업장명"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>입고일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4511ma1_I841741675_txtRcptDT.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>입고번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRcptNo" SIZE=18 MAXLENGTH=16 tag="25xxxU" ALT="입고번호"></TD>
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
								<TD HEIGHT="100%" colspan=4>
									<script language =javascript src='./js/p4511ma1_I395289701_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
							</TR>

							<TR>
								<TD WIDTH=66% colspan=2>
								<FIELDSET valign=top>
									<LEGEND>수량관련</LEGEND>
									<TABLE CLASS="TB2" CELLSPACING=0>
										<TR>
											<TD CLASS=TD5 NOWRAP>오더단위</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderUnit" SIZE=5 MAXLENGTH=3 tag="24" ALT="오더단위"></TD>										
											<TD CLASS=TD5 NOWRAP>기준단위</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBaseUnit" SIZE=5 MAXLENGTH=3 tag="24" ALT="기준단위"></TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>오더수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I522660672_txtOrderQty.js'></script>
											</TD>
											<TD CLASS=TD5 NOWRAP>오더수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_fpDoubleSingle1_txtOrderQty1.js'></script>
											</TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>실적수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I927203613_txtProdQty.js'></script>
											</TD>
											<TD CLASS=TD5 NOWRAP>실적수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I618573570_txtProdQty1.js'></script>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>양품수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I348217362_txtGoodQty.js'></script>
											</TD>
											<TD CLASS=TD5 NOWRAP>양품수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I191775000_txtGoodQty1.js'></script>
											</TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>입고수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I132385045_txtRcptQty.js'></script>
											</TD>
											<TD CLASS=TD5 NOWRAP>입고수량</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I490456032_txtRcptQty1.js'></script>
											</TD>
										</TR>
									</TABLE>	
								</FIELDSET>			
								</TD>
								<TD WIDTH=34% colspan=2>
								<FIELDSET valign=top>
									<LEGEND>일정관련</LEGEND>
									<TABLE CLASS="TB2" CELLSPACING=0>
										<TR>
											<TD CLASS=TD5 NOWRAP>착수예정일</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I354317295_txtPlanStratDt.js'></script>
											</TD>

										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>완료예정일</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I565412455_txtPlanEndDt.js'></script>
											</TD>

										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>작업지시일</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I479410910_txtReleaseDt.js'></script>
											</TD>
										</TR>	
										<TR>
											<TD CLASS=TD5 NOWRAP>실착수일</TD>
											<TD CLASS=TD6 NOWRAP>
												<script language =javascript src='./js/p4511ma1_I614313739_txtRealStratDt.js'></script>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD5 NOWRAP>지시상태</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderStatus" SIZE=10 tag="24" ALT="지시상태" ></TD>
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
    <TR HEIGHT="20">
      <TD WIDTH="100%">
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
			  <TD WIDTH=10>&nbsp;</TD>
			  <TD WIDTH="*" align="left"><a><button name="btnAutoSel" class="clsmbtn">전체선택</button></a></TD>
			  <TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
      </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hSlCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
