<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 재고이동등록 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2003/05/13
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              :
'**********************************************************************************************-->
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="i1312ma1_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                       

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgMovType
Dim gMovTypeFlag
Dim StartDate
Dim Currentdate
	
Dim IsOpenPop					

'==========================================  2.1.1 InitVariables()  ======================================
Currentdate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(Currentdate, parent.gServerDateFormat, parent.gDateFormat)

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
	<%Call LoadBNumericFormatA("I","*","NOCOOKIE","MA") %>
End Sub


'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	
	call LoadInfTB19029
    
    If GetSetupMod(Parent.gSetupMod, "a") = "Y" Then
		txtOpenGL1Title.style.display = ""
		txtOpenGL2Title.style.display = ""
		txtCostTitle.style.display = ""
	Else
		frm1.txtCostCd1.tag = "25"
		ggoOper.SetReqAttr frm1.txtCostCd1, "Q"
	End if
	
	Call LockObjectField(frm1.txtYear,"R")
    Call LockObjectField(frm1.txtDocumentDt,"R")
    Call LockObjectField(frm1.txtPostingDt,"R")
    Call FormatDATEField(frm1.txtDocumentDt)
    Call FormatDATEField(frm1.txtPostingDt)
    Call FormatDATEField(frm1.txtYear)

    Call ggoOper.FormatDate(frm1.txtYear, Parent.gDateFormat, 3)
    Call InitSpreadSheet 
    Call InitVariables 
    Call SetDefaultVal

    Call SetToolBar("11101000000011")
    									
    gIsTab     = "Y" 
    gTabMaxCnt = 2  


End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_00%> >
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고이동등록</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<TD background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고이동내역</font></TD>
								<TD background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMoveInvRef1()">자재불출의뢰정보| </A>
					                        <A ID="txtOpenGL1Title" STYLE="DISPLAY: none" href="vbscript:OpenPopupGL()">회계전표정보 | </A>
					                        <A ID="txtOpenGL2Title" STYLE="DISPLAY: none" href="vbscript:OpenPopupGL2()">결의전표정보 | </A>
					                        <A href="vbscript:OpenMoveInvRef()">사내재고이동정보| </A>
					                        <A href="vbscript:OpenSubCtctRef()">사급품출고예정정보</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
		<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=no>
			<TABLE <%=LR_SPACE_TYPE_20%> >
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd1" CLASS=required STYLE="Text-Transform: uppercase" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant1()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm1" CLASS=protected readonly=true TABINDEX="-1" SIZE=28 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>이동번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDocumentNo1" CLASS=required STYLE="Text-Transform: uppercase" SIZE=20 MAXLENGTH=16  tag="12XXXU" ALT="이동번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocumentNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDocumentNo()"></TD>
									<TD CLASS="TD5" NOWRAP>년도</TD>
									<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/i1312ma1_ko441_fpDateTime1_txtYear.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>	
					<TD <%=HEIGHT_TYPE_03%> >
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=top HEIGHT=*>
					  <TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>이동일자</TD>
							<TD CLASS="TD656" NOWRAP>
								<script language =javascript src='./js/i1312ma1_ko441_fpDateTime2_txtDocumentDt.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>회계전표일자</TD>
							<TD CLASS="TD656" NOWRAP>
								<script language =javascript src='./js/i1312ma1_ko441_fpDateTime3_txtPostingDt.js'></script>
							</TD>							
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>수불유형</TD>
							<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtMovType"  CLASS=protected readonly=true  STYLE="Text-Transform: uppercase" SIZE=10 MAXLENGTH=3 tag="24XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMovType()">&nbsp;<INPUT TYPE=TEXT NAME="txtMovTypeNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 tag="24"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>창고</TD>
							<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd1" CLASS=required STYLE="Text-Transform: uppercase" SIZE=10 MAXLENGTH=7 tag="23XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL1()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm1" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 tag="24"></TD>
						</TR>
						<TR ID="txtCostTitle" STYLE="DISPLAY: none">
							<TD CLASS="TD5" NOWRAP>Cost Center</TD>
							<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd1" SIZE=10 MAXLENGTH=10 tag="25XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter1" align=top TYPE="BUTTON" tag="23" ONCLICK="vbscript:OpenCostCD1()">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm1" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=30 tag="24"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>이동번호</TD>
							<TD CLASS="TD656" NOWRAP><INPUT NAME="txtDocumentNo2" ALT="이동번호" STYLE="Text-Transform: uppercase" TYPE="Text" MAXLENGTH="16" SIZE=20 STYLE="" tag="25XXXU"></TD>
						</TR>
						<TR >
							<TD CLASS="TD5" NOWRAP>불출의뢰번호</TD>
							<TD CLASS="TD656" NOWRAP><INPUT NAME="txtDocumentText" ALT="불출의뢰번호" TYPE="Text" CLASS=protected readonly=true  MAXLENGTH="20" SIZE=20 tag="25"></TD>
						</TR>
						<% Call SubFillRemBodyTD656(12)%>
					  </TABLE>
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<DIV ID="TabDiv" STYLE="DISPLAY:none " SCROLL=no>
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%> >
			     <TR>
					<TD HEIGHT=* VALIGN=TOP WIDTH=100%>
			              <TABLE <%=LR_SPACE_TYPE_60%> > 
					        <TR> 
								<TD CLASS="TD5" NOWRAP><LABEL CLASS="normal" ID="txtPlantCd2Title" STYLE="DISPLAY: none">이동공장</LABEL></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd2" STYLE="DISPLAY: none" SIZE=10 MAXLENGTH=4 tag="25XXXU" ALT="이동공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" STYLE="DISPLAY: none" ONCLICK="vbscript:OpenPlant2()"><INPUT TYPE=TEXT NAME="txtPlantNm2" CLASS=protected readonly=true TABINDEX="-1" STYLE="DISPLAY: none" SIZE=20 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP><LABEL CLASS="normal" ID="txtTrackingNoTitle" STYLE="DISPLAY: none">Tracking No</LABEL></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo"  STYLE="DISPLAY: none" SIZE=15 MAXLENGTH=25 tag="25XXXU" ALT="Tracking No"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" STYLE="DISPLAY: none" onclick=vbscript:OpenTrackingNo1></TD>							
				           </TR>
						  <TR ID="txtSLCd2Title" STYLE="DISPLAY: none">
								<TD CLASS="TD5" NOWRAP>이동창고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd2" SIZE=10 MAXLENGTH=7 tag="25XXXU" ALT="이동창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL2()"><INPUT TYPE=TEXT NAME="txtSLNm2" CLASS=protected readonly=true TABINDEX="-1" SIZE=20 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
					      </TR>
						  <TR ID="txtCostCd2Title" STYLE="DISPLAY: none">
					      		<TD CLASS="TD5" NOWRAP>이동 Cost Center</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd2" SIZE=10 MAXLENGTH=10 tag="25XXXU" ALT="이동 Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCD2()"><INPUT TYPE=TEXT NAME="txtCostNm2" CLASS=protected readonly=true TABINDEX="-1" SIZE=20 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
						  </TR>
				          <TR>
					           <TD WIDTH=100% HEIGHT=100% COLSPAN=4> 
						    	<script language =javascript src='./js/i1312ma1_ko441_I963957387_vspdData.js'></script>
							  </TD>
						 </TR>
						</TABLE>	
					</TD>
				</TR>
			</TABLE>
		</DIV>
		</TD>
	</TR>
	<TR>
	    <TD <%=HEIGHT_TYPE_01%> >
	    </TD>
	</TR>
	<TR HEIGHT=20 >
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%> >
	        
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hGroupCount" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDocumentNo1" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hYear" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hGuiControlFlag3" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hGuiControlFlag2" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
