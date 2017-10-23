<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/05/06
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Ma JIn Ha
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"'****************************
'* 13. History              :
'*                            2000/05/08,2000/05/11
'********************************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   ***************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--==========================================  1.1.1 Style Sheet  =====================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--==========================================  1.1.2 공통 Include   ====================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="m5112ma1_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/JpQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Const BIZ_PGM_ID 		= "M5112mb1_KO441.asp"
Const BIZ_PGM_JUMP_ID 	= "M5111ma1"
Const BIZ_PGM_JUMP_ID2  = "M5113ma1"

'===============================  LoadInfTB19029()  ============================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call FormatDATEField(frm1.txtIvDt)
    Call FormatDoubleSingleField(frm1.txtivAmt)
    Call FormatDoubleSingleField(frm1.txtXchRt)
    Call FormatDoubleSingleField(frm1.txtnetAmt)
    Call FormatDoubleSingleField(frm1.txtvatAmt)
    
    Call LockHTMLField(frm1.txtIvNo, "R")
    Call LockObjectField(frm1.txtIvDt, "P")
    Call LockHTMLField(frm1.ChkPrepay, "P")
    Call LockObjectField(frm1.txtivAmt, "P")
    Call LockObjectField(frm1.txtXchRt, "P")
    Call LockObjectField(frm1.txtnetAmt, "P")
    Call LockObjectField(frm1.txtvatAmt, "P")
    
	Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	Call InitVariables                                                      '⊙: Initializes local global variables
    Call GetValue_ko441()
    Call SetDefaultVal
    Call CookiePage(0)
	    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입내역</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenGrRef()" >입출고참조</A>&nbsp;|&nbsp;
											<A href="vbscript:OpenExceptGrRef()">예외입고참조</A>&nbsp;|&nbsp;											
											<A href="vbscript:OpenPoRef()" >발주내역참조</A>&nbsp;|&nbsp;
											<A href="vbscript:OpenLLCRef()">LOCAL L/C내역참조</A>&nbsp;|&nbsp;
											<A href="vbscript:OpenRetRef()">예외반품출고참조</A></TD>
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
									<TD CLASS="TD5" nowrap>매입번호</TD>
									<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=32 MAXLENGTH=18 ALT="매입번호" CLASS=required STYLE="text-transform:uppercase" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvNo()">
														   <div STYLE="DISPLAY: none">
														   <INPUT NAME="txtSoNo1" STYLE="BORDER-RIGHT: 0px solid;BORDER-TOP: 0px solid;BORDER-LEFT: 0px solid;BORDER-BOTTOM: 0px solid" TYPE="Text" SIZE=1 DISABLED=TRUE Tag="11"></div></TD>
									<TD CLASS="TD6" nowrap></TD>
									<TD CLASS="TD6" nowrap></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>매입형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="매입형태" NAME="txtIvTypeCd" SIZE=10 CLASS=protected readonly=true tag="24X">
													   <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=20 ALT ="매입형태" CLASS=protected readonly=true tag="24X" ></TD>


								
								<TD CLASS="TD5" NOWRAP>매입등록일</TD>
								<TD CLASS="TD6" NOWRAP>
								    <Table cellpadding=0 cellspacing=0>
								        <TR>
								            <TD NOWRAP>
								                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=매입등록일 NAME="txtIvDt" style="HEIGHT: 20px; WIDTH: 80px" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" CLASS=protected readonly=true tag="24X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>&nbsp;
											</TD>
								            <TD NOWRAP> 
								                <INPUT TYPE=CHECKBOX CHECKED ID="ChkPrepay" CLASS=protected readonly=true tag="24" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid">
											</TD>
											<TD NOWRAP>
												선급금여부
											</TD>
								        </TR>
								   </TABLE>
								 </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSpplCd" SIZE=10 CLASS=protected readonly=true tag="24X">
													   <INPUT TYPE=TEXT NAME="txtSpplNm" SIZE=20 ALT ="공급처" CLASS=protected readonly=true tag="24X" ></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="구매그룹" NAME="txtGrpCd" SIZE=10 CLASS=protected readonly=true tag="24X">
													   <INPUT TYPE=TEXT NAME="txtGrpNm" SIZE=20 ALT ="구매그룹" CLASS=protected readonly=true tag="24X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>총매입금액</TD>
								<TD CLASS="TD6" NOWRAP>
								    <Table cellpadding=0 cellspacing=0>
								        <TR>
								            <TD NOWRAP>
								                <INPUT TYPE=TEXT ALT="화폐" NAME="txtCur" SIZE=10 CLASS=protected readonly=true tag="24X">&nbsp;</TD>
								            <TD NOWRAP> 
								                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=총금액 TYPE=TEXT NAME="txtivAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 150px" CLASS=protected readonly=true tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
		
								        </TR>
								   </TABLE>
								 </TD>
								<TD CLASS="TD5" NOWRAP>환율</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchRt" style="HEIGHT: 20px; WIDTH: 234px" CLASS=protected readonly=true tag="24X5" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
						
							<TR>
								<TD CLASS="TD5" NOWRAP>매입금액</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=매입금액 TYPE=TEXT NAME="txtnetAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 234px" CLASS=protected readonly=true tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
							
								<TD CLASS="TD5" NOWRAP>VAT금액</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT금액 TYPE=TEXT NAME="txtvatAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 234px" CLASS=protected readonly=true tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
							</TR>
							<TR>
								<TD HEIGHT="100%"  WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>			
						</TABLE>
		
					</TD>
				</TR>


			</TABLE>
		</TD>
	</TR>
    <tr>
      <td <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td>
						<BUTTON NAME="btnPosting" CLASS="CLSSBTN"  ONCLICK="Posting()">확정처리</BUTTON>&nbsp;
		         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;
					</td>	
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">매입세금계산서</a>|<a href="VBSCRIPT:CookiePage(2)">지급내역등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txthdnIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtQuerytype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPostingFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSppl" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvTypeNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMvmtType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPostDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnExceptflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnVatRt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoVatRt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnXch" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDiv" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnVatType" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnIssueType" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoTypeNm" tag="24">
<INPUT TYPE=HIDDEN NAME="hdvatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLcKind" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPayMeth" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>

