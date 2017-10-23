<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement 
'*  2. Function Name        : 
'*  3. Program ID           : U1122MA1
'*  4. Program Name         : 구매입고등록(procurement)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/03/02
'*  8. Modified date(Last)  : 2003/06/02
'*  9. Modifier (First)     : NHG
'* 10. Modifier (Last)      : NHG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   *****************************************!-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================!-->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ====================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="U1122MA1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit					

'==============================================================================================================================Const BIZ_PGM_ID = "m4111mb1.asp"									
Const BIZ_PGM_ID = "U1122MB1.asp"									
Const BIZ_PGM_JUMP_ID	= "M4131MA1"
'==============================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'==============================================================================================================================
Dim IsOpenPop          
Dim lblnWinEvent
Dim interface_Account

Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec   
Dim C_TrackingNo
Dim C_InspFlg	
Dim C_MvmtRcptLookUpBtn
Dim C_GrQty	
Dim C_StockQty
Dim C_GRUnit
Dim C_Cur	
Dim C_MvmtPrc
Dim C_DocAmt
Dim C_LocAmt
Dim C_SlCd	
Dim C_SlCdPop	
Dim C_SlNm		
Dim C_InspSts	
Dim C_GRMeth	
Dim C_LotNo		
Dim C_LotSeqNo	
Dim C_MakerLotNo
Dim C_MakerLotSeqNo
Dim C_GRNo
Dim C_GRSeqNo
Dim C_InspReqNo
Dim C_InspResultNo
Dim C_PoNo
Dim C_PoSeqNo
Dim C_CCNo
Dim C_CCSeqNo
Dim C_LLCNo
Dim C_LLCSeqNo
Dim C_Stateflg
Dim C_InspMethCd
Dim C_MvmtNo
Dim C_IvNo
Dim C_IvSeqNo
Dim C_RetOrdQty
Dim C_SplitSeqNo

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'=============================================================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                
    lgBlnFlgChgValue = False                 
    
    lgStrPrevKey = ""                        
    frm1.vspdData.MaxRows = 0
    Call ChangeTag(False)
    Call SetToolBar("1110000000001111")
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==============================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029                                                    
	
    call FormatDateField(frm1.txtGmDt)		
	call LockobjectField(frm1.txtGmDt,"R")
	
	
    Call InitSpreadSheet                                                   
    Call SetDefaultVal
    Call InitVariables
    Call CookiePage(0)
    
End Sub
'==============================================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구매입고등록(Procurement)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenFirmPORcpt()">입고예정참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>입고번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = REQUIRED STYLE = "text-transform:uppercase"  TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmtNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()"></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
								<TD CLASS="TD5" NOWRAP>입고형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS = REQUIRED STYLE = "text-transform:uppercase" TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="23NXXU" OnChange="VBScript:changeMvmtType()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMvmtType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>입고일</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="입고일" NAME="txtGmDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD Title="FPDATETIME" tag="23N1"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS = REQUIRED STYLE = "text-transform:uppercase" TYPE=TEXT NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="23XXXU" OnChange="VBScript:changeSpplCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT STYLE = "text-transform:uppercase" TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="25XXXU" OnChange="VBScript:changeGroupCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
													   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" SIZE=20 NAME="arrCond" tag="24X"></TD>								
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>입고번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT  STYLE = "text-transform:uppercase" TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo1" MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td>						
		         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>&nbsp;&nbsp;
		         		<BUTTON NAME="btnGlSel_1" CLASS="CLSMBTN"  ONCLICK="OpenBackFlushRef()">자품목Simulation</BUTTON>&nbsp;&nbsp;
					</td>					
					<td WIDTH="*" align=right><a href="VBSCRIPT:CookiePage(1)">검사결과등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"  TABINDEX=-1></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"  TABINDEX=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlNo" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnGlType" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hdnMvmtCur" tag="24" TabIndex="-1">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
                     