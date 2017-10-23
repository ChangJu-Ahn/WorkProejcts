<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2210ma1
'*  4. Program Name         : MPS일괄생성 
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
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p2210ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'=========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

'=========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=========================================================================================================
Sub LoadInfTB19029() 
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
End Sub 


'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")

    Call SetToolBar("10000000000011")

    Call InitVariables    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm	
		frm1.txtPlanDt.focus 
	
		gLookUpEnable = True
	
        If DBQuery = False Then 
           Call RestoreToolBar()
           Exit Sub 
        End If 
	ELSE
		frm1.txtPlantCd.focus 
	End If    
	
    Set gActiveElement = document.activeElement
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MPS일괄생성</font></td>
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
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>							
							<TR>
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="23XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14">&nbsp;&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>MPS이력번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtMPSHistoryNo" SIZE=18 MAXLENGTH=18 tag="24XXXU" ALT="MPS이력번호"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>계획일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2210ma1_I662456655_txtPlanDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>가용재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoAvailInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg2 VALUE="N"><LABEL FOR=rdoAvailInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>안전재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoSafeInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg2 VALUE="N"><LABEL FOR=rdoSafeInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최대LOT감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMaxFlg ID=rdoMaxFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoMaxFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMaxFlg ID=rdoMaxFlg2 VALUE="N"><LABEL FOR=rdoMaxFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최소LOT감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMinFlg ID=rdoMinFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoMinFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMinFlg ID=rdoMinFlg2 VALUE="N"><LABEL FOR=rdoMinFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>올림수</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRoundFlg ID=rdoRoundFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoRoundFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRoundFlg ID=rdoRoundFlg2 VALUE="N"><LABEL FOR=rdoRoundFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>기준일자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoStartDtFlg ID=rdoStartDtFlg1 VALUE="Y"><LABEL FOR=rdoStartDtFlg1>DTF</LABEL>&nbsp;&nbsp;&nbsp;
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoStartDtFlg ID=rdoStartDtFlg2 VALUE="N" CHECKED><LABEL FOR=rdoStartDtFlg2>PTF</LABEL></TD>
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>DTF(Demand Time Fence)</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2210ma1_I567691153_txtDTF.js'></script>
								</TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>&nbsp;&nbsp;PTF(Planning Time Fence)</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2210ma1_I729984486_txtPTF.js'></script>
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
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" Flag=1 onclick="ExecuteMPS()">실행</BUTTON></TD>		
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
