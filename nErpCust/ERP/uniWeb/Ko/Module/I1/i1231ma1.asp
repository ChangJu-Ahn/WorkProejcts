<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'******************************************************************************
'*  1. Module Name          : 기타출고 등록
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2003/10/17
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="i1231ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        


<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgMovType
Dim IsOpenPop						
Dim StartDate
Dim Currentdate
Dim StrCompany

Currentdate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(Currentdate, parent.gServerDateFormat, parent.gDateFormat)

'****************************** Sub LoadInfTB19029() *****************************************
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
	<%Call LoadBNumericFormatA("I","*","NOCOOKIE","MA") %>
End Sub

'======================================== Form_Load() ===================================================
Sub Form_Load()

    Call LoadInfTB19029

	If GetSetupMod(Parent.gSetupMod, "a") = "Y" Then
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

	Call ggoOper.LockField(Document, "N")                                          
    Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call SetToolBar("11101000000011")										
    
    gIsTab     = "Y" 
    gTabMaxCnt =  2
	
End Sub

</SCRIPT>
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_00%>>
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
										<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타출고등록</font></td>
										<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
									</TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
									<TR>
										<TD background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<TD background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타출고내역</font></td>
										<TD background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100%>
									<FIELDSET CLASS="CLSFLD">
										<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>공장</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" CLASS=required STYLE="Text-Transform: uppercase" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=28 tag="14"></TD>
												<TD CLASS="TD5" NOWRAP></TD>
												<TD CLASS="TD6" NOWRAP></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>출고번호</TD>
												<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDocumentNo1" CLASS=required STYLE="Text-Transform: uppercase" SIZE=20 MAXLENGTH=16 tag="12XXXU" ALT="출고번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocumentNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDocumentNo()"></TD>
												<TD CLASS="TD5" NOWRAP>년도</TD>
												<TD CLASS="TD6" NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYY name=txtYear CLASSID=<%=gCLSIDFPDT%> tag="12X1" ALT="년도"> <PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
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
											<TD CLASS="TD5" NOWRAP>출고일자</TD>
											<TD CLASS="TD656" NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtDocumentDt CLASSID=<%=gCLSIDFPDT%> tag="23X1" ALT="출고일자"> <PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>회계전표일자</TD>
											<TD CLASS="TD656" NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime3 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPostingDt CLASSID=<%=gCLSIDFPDT%> tag="23X1" ALT="회계전표일자"> <PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
											</TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>수불유형</TD>
											<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtMovType" SIZE=10 MAXLENGTH=3 tag="23XXXU" ALT="수불유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMovType()">&nbsp;<INPUT TYPE=TEXT NAME="txtMovTypeNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=12 tag="24"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>창고</TD>
											<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="23XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=20 tag="24" ALT="창고명"></TD>
										</TR>
										<TR ID="txtCostTitle" STYLE="DISPLAY: none">
											<TD CLASS="TD5" NOWRAP>Cost Center</TD>
											<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="23XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCenter" align=top TYPE="BUTTON" onclick=vbscript:OpenCostCd()>&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 tag="24"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>작업장</LABEL></TD>
											<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtWCCd" SIZE=10 MAXLENGTH=7 tag="25XXXU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenWC()">&nbsp;<INPUT TYPE=TEXT NAME="txtWCNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 tag="24"></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>출고번호</TD>
											<TD CLASS="TD656" NOWRAP><INPUT NAME="txtDocumentNo2" ALT="출고번호" STYLE="Text-Transform: uppercase" TYPE="Text" MAXLENGTH="16" SIZE=20 STYLE="" tag="25XXXU"></TD>
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
					<DIV ID="TabDiv" STYLE="DISPLAY:none " SCROLL=no>	
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH=100% HEIGHT=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData width=100% height=100% TITLE="SPREAD" tag="2" id=OBJECT1> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>    
						</TABLE>
					</DIV>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_01%>>
				</TD>
			</TR>
			<TR HEIGHT=20>
				<TD>
					<TABLE <%=LR_SPACE_TYPE_30%>>
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
			<INPUT TYPE=HIDDEN NAME="hCostFlg" tag="14" TABINDEX="-1">
			<INPUT TYPE=HIDDEN NAME="hYear" tag="14" TABINDEX="-1">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>
