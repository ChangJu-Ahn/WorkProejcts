<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2345ma1
'*  4. Program Name         : MRP 일괄전환/승인취소 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Hyun Jae
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
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
<SCRIPT LANGUAGE = "VBScript" SRC = "p2345ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->

'==========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'==========================================================================================================
Sub LoadInfTB19029() 
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
End Sub 

'==========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 

Sub Form_Load()
	
	Call LoadInfTB19029()
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")

    Call SetToolbar("10000000000011")
    Call InitVariables    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		
		gLookUpEnable = True
		
		If DBQuery = False Then 
           Call RestoreToolBar()
           Exit Sub
        End If 
	Else
		frm1.txtPlantCd.focus 
    End If
    
    Set gActiveElement = document.activeElement	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP일괄전환/승인취소</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>	
					<TD WIDTH=* align=right><A href="vbscript:OpenErrorList()">ERROR내역리스트</A></TD>				
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="23XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>MRP 실행번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMRPHisNo" SIZE=18 MAXLENGTH=18 tag="24XXXU" ALT="MRP 실행번호"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>기준일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2345ma1_I215864616_txtFixExecFromDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>확정전개기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2345ma1_I500832748_txtFixExecToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>예시전개기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2345ma1_I256713265_txtPlanExecToDt.js'></script>
								</TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>가용재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg1 CHECKED><LABEL FOR=rdoAvailInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg2><LABEL FOR=rdoAvailInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>안전재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg1 CHECKED><LABEL FOR=rdoSafeInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg2><LABEL FOR=rdoSafeInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>승인일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2345ma1_I900466978_txtApproveDt.js'></script>
								</TD>		
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>승인자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApprover" SIZE=15 MAXLENGTH=13 tag="24" ALT="승인자"></TD>		
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계획오더번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStartOrderNo" SIZE=18 MAXLENGTH=18 tag="24" ALT="시작계획오더번호">&nbsp;~&nbsp;<INPUT TYPE=TEXT NAME="txtEndOrderNo" SIZE=18 MAXLENGTH=18 tag="24" ALT="종료계획오더번호">&nbsp;&nbsp;&nbsp</TD>		
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>전환오더수</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2345ma1_I303770047_txtConvQty.js'></script>
								</TD>		
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ERROR 수</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2345ma1_I283918213_txtErrQty.js'></script>
								</TD>		
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>Status</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtStatus" SIZE=15 MAXLENGTH=15 tag="24" ALT="Status"></TD>		
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
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1 onclick="BatchConfirmMRP()">MRP 일괄전환</BUTTON>&nbsp;<BUTTON NAME="btnApproveCancel" CLASS="CLSMBTN" Flag=1 onclick="ApproveCancelMRP()">MRP 승인취소</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
