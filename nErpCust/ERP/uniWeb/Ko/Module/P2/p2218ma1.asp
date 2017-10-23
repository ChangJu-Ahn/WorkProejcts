<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p2218ma1
'*  4. Program Name			: MPS확정취소 
'*  5. Program Desc			:
'*  6. Business ASP List	: +p2218mb1.asp		'☆: List MPS
							  +p2218mb2.asp		'☆: cancel MPS(save)
							  +p2218mb3.asp 	'☆: Look up Plant
'*  7. Modified date(First)	:
'*  8. Modified date(Last)	:
'*  9. Modifier (First)		: Jung Yu Kyung
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				: MPS확정취소 
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'**********************************************************************************************-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "p2218ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,parent.gDateFormat)

<!-- #Include file="../../inc/lgVariables.inc" -->

'==========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
	Call SetDefaultVal
	Call InitVariables

	Call SetToolBar("11000000000011")

	frm1.txtPlndFromDt.Text = StartDate
	frm1.txtPlndToDt.Text = LastDate

	Call InitComboBox()
	
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		Call ExecMyBizASP(frm1, BIZ_PLANT_ID)	
		frm1.txtItemCd.focus
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MPS확정취소</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>계획일자</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/p2218ma1_fpDateTime3_txtPlndFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p2218ma1_fpDateTime3_txtPlndToDt.js'></script>
									</TD>																						
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo 0"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>MPS Status</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboMPSStatus" ALT="MPS Status" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
								<script language =javascript src='./js/p2218ma1_I973154413_vspdData.js'></script>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
						</TR>
						<TR>
							<TD WIDTH=55% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>Lot Sizing</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>최대오더수량</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p2218ma1_I242371708_txtMaxLotQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP>제조 L/T</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemLT" SIZE=10 tag="24" ALT="품목 L/T" STYLE="TEXT-ALIGN: right"></TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>최소오더수량</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p2218ma1_I550510692_txtMinLotQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP>구매 L/T</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAccumLT" SIZE=10 tag="24" ALT="누적 L/T" STYLE="TEXT-ALIGN: right"></TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>올림수</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p2218ma1_I620081776_txtRondQty.js'></script>
										</TD>
										<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
										<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									</TR>
								</TABLE>	
							</FIELDSET>			
							</TD>
							<TD WIDTH=45% colspan=2>
							<FIELDSET valign=top>
								<LEGEND>Time Fence</LEGEND>
								<TABLE CLASS="TB2" CELLSPACING=0>
									<TR>
										<TD CLASS=TD5 NOWRAP>DTF(Demand Time Fence)</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p2218ma1_I250055003_txtDTF.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>PTF(Planning Time Fence)</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p2218ma1_I680994743_txtPTF.js'></script>
										</TD>
									</TR>	
									<TR>
										<TD CLASS=TD5 NOWRAP>Planning Horizon</TD>
										<TD CLASS=TD6 NOWRAP>
											<script language =javascript src='./js/p2218ma1_I386777818_txtPH.js'></script>
										</TD>
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
					<TD><BUTTON NAME="btnAutoSel" CLASS="CLSMBTN">전체선택</BUTTON></TD></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlndFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hPlndToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hMPSOrigin" tag="24">
<INPUT TYPE=HIDDEN NAME="hMPSStatus" tag="24"><INPUT TYPE=HIDDEN NAME="txtMPSOpFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>