<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          :  기타입고등록 
'*  2. Function Name        :
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  7. Modified date(First) : 2002/10/31
'*  9. Modifier (First)     : Han Sung Gyu
'* 13. History              : 
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

<SCRIPT LANGUAGE="VBScript"   SRC="i1131ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                         

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgMovType
Dim IsOpenPop					
Dim gSelframeFlg
Dim StartDate
Dim Currentdate
Dim StrCompany

Currentdate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(Currentdate, parent.gServerDateFormat, parent.gDateFormat)

'**********************************Sub LoadInfTB19029()*****************************************
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
	<%Call LoadBNumericFormatA("I","*","NOCOOKIE","MA") %>
End Sub


'************************** Sub Form_Load() ****************************************************
Sub Form_Load()
	
	Call LoadInfTB19029
	
	If GetSetupMod(parent.gSetupMod, "a") = "Y" Then
		txtOpenGL1Title.style.display = ""
		txtOpenGL2Title.style.display = ""
		txtCostTitle.style.display = ""	
	Else
		frm1.txtCostCd.tag = "25"
		ggoOper.SetReqAttr frm1.txtCostCd, "Q"
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
    gTabMaxCnt =  2   
    
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타입고등록</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<TD background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타입고내역</font></TD>
								<TD background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
                   <TD WIDTH=* align=right><A ID="txtOpenGL1Title" STYLE="DISPLAY: none" href="vbscript:OpenPopupGL()">회계전표정보 |</A>&nbsp;<A ID="txtOpenGL2Title" STYLE="DISPLAY: none" href="vbscript:OpenPopupGL2()">결의전표정보</A></TD>
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
					<TD HEIGHT=30 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" CLASS=required STYLE="Text-Transform: uppercase" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=28 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>입고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDocumentNo1" CLASS=required STYLE="Text-Transform: uppercase" SIZE=20 MAXLENGTH=16 tag="12XXXU" ALT="입고번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocumentNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDocumentNo()"></TD>
									<TD CLASS="TD5" NOWRAP>년도</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/i1131ma1_fpDateTime1_txtYear.js'></script>
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
					<TD WIDTH=100% VALIGN=top>
					  <TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>입고일자</TD>
							<TD CLASS="TD656" NOWRAP>
								<script language =javascript src='./js/i1131ma1_fpDateTime2_txtDocumentDt.js'></script>
							</TD>
						</TR>						
						
						<TR>
							<TD CLASS="TD5" NOWRAP>회계전표일자</TD>
							<TD CLASS="TD656" NOWRAP>
								<script language =javascript src='./js/i1131ma1_fpDateTime3_txtPostingDt.js'></script>
							</TD>	
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>수불유형</TD>
							<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtMovType" CLASS=required STYLE="Text-Transform: uppercase" SIZE=10 MAXLENGTH=3 tag="23XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType2" align=top TYPE="BUTTON" tag="23" ONCLICK="vbscript:OpenMovType()">&nbsp;<INPUT TYPE=TEXT NAME="txtMovTypeNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=30 tag="24"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>창고</TD>
							<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" CLASS=required STYLE="Text-Transform: uppercase" SIZE=10 MAXLENGTH=7 tag="23XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" tag="23" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=30 tag="24" ALT="창고명"></TD>
						</TR>
						<TR ID="txtCostTitle" STYLE="DISPLAY: none">
							<TD CLASS="TD5" NOWRAP>Cost Center</TD>
							<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" CLASS=required STYLE="Text-Transform: uppercase" SIZE=10 MAXLENGTH=10 tag="23XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter" align=top TYPE="BUTTON" tag="23" ONCLICK="vbscript:OpenCostCD()">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=30 tag="24" ALT="Cost Center"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>입고번호</TD>
							<TD CLASS="TD656" NOWRAP><INPUT NAME="txtDocumentNo2" ALT="입고번호" STYLE="Text-Transform: uppercase" TYPE="Text" MAXLENGTH=16 SIZE=20 STYLE="" tag="25XXXU"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>비고</TD>
							<TD CLASS="TD656" NOWRAP><INPUT NAME="txtDocumentText" ALT="비고" TYPE="Text" MAXLENGTH="40" SIZE=50 tag="25"></TD>					
						</TR>
						<% Call SubFillRemBodyTD656(12)%>
					  </TABLE>

					</TD>
				</TR>
			</TABLE>
		</DIV>
		<DIV ID="TabDiv" STYLE="DISPLAY:none" SCROLL=no>
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%> >
				<TR>
					<TD WIDTH="100%" Height=100%><script language =javascript src='./js/i1131ma1_I243881094_vspdData.js'></script></TD>
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
<INPUT TYPE=HIDDEN NAME="hYear" tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

