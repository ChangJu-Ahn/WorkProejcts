
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4313ma1
'*  4. Program Name			: 출고현황조회 
'*  5. Program Desc			:
'*  6. Comproxy List		: +B19029LookupNumericFormat
'*  7. Modified date(First)	: 2000/04/21
'*  8. Modified date(Last) 	: 2002/11/21
'*  9. Modifier (First) 	: Kim, GyoungDon
'* 10. Modifier (Last)		: Kang, HyoKu
'* 11. Comment				:
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="p4313ma1.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
Dim EndDate

	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	StartDate = UNIDateAdd("D",-7,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 처음 날짜 
	EndDate = UNIDateAdd("D", 7,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜 

<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>

End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1015", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboProdMgr, lgF0, lgF1, Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("I0004", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboInvMgr, lgF0, lgF1, Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboJobCd, lgF0, lgF1, Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1016", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboIssueMthd, lgF0, lgF1, Chr(11))
	
	frm1.cboProdMgr.value = ""
	frm1.cboInvMgr.value = ""	
	frm1.cboJobCd.value = ""	
	frm1.cboIssueMthd.value = ""
	
	
		
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) 

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
         
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
	
	Call InitComboBox
	Call InitSpreadComboBox	
			
    Call SetDefaultVal
    Call InitVariables		'⊙: Initializes local global variables

	Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
	
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtCompntCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출고현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenOprRef()">공정내역</A> | <A href="vbscript:OpenProdRef()">실적내역</A> | <A href="vbscript:OpenConsumRef()">부품소비내역</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>필요일</TD>
								    <TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtReqStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="12X1" id=fpDateTime1> </OBJECT>');</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtReqEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12X1" id=fpDateTime1> </OBJECT>');</SCRIPT>
								</TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부품</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCompntCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo1 frm1.txtCompntCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtCompntNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>완료여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueFlg" tag="11" ID="rdoIssueFlag1" VALUE="Y"><LABEL FOR="rdoIssueFlg">예</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueFlg" tag="11" ID="rdoIssueFlag2" VALUE="N"><LABEL FOR="rdoIssueFlg">아니오</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoIssueFlg" tag="11" CHECKED ID="rdoIssueFlag3" VALUE="A"><LABEL FOR="rdoIssueFlg">전체</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="TEXT" NAME="txtWCCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC()"> <INPUT TYPE=TEXT ID="txtWCNm" NAME="arrCond" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출고창고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="출고창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd frm1.txtSLCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>작업</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING= 0>
											<TR>
												<TD>
													<SELECT NAME="cboJobCd" ALT="작업" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
												</TD>
												<TD WIDTH="*">
													&nbsp;
												</TD>
												<TD  WIDTH="20" STYLE="TEXT-ALIGN: RIGHT" ><IMG SRC="../../../CShared/image/BigPlus.gif" Style="CURSOR: hand" ALT="DetailCondition" ALIGN= "TOP" ID = "IMG_DetailCondition" NAME="pop1" ONCLICK= 'vbscript:viewHidden "DetailCondition" ,3, 3' ></IMG></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR ID="DetailCondition1" style="display: none">
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value"></TD>	
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
								</TR>
								<TR ID="DetailCondition2" style="display: none">
									<TD CLASS=TD5 NOWRAP>재고담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboInvMgr" ALT="재고담당자" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>생산담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdMgr" ALT="생산담당자" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>	
								</TR>
								<TR ID="DetailCondition3" style="display: none">
									<TD CLASS=TD5 NOWRAP>출고방법</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboIssueMthd" ALT="출고방법" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
									<TD CLASS=TD5 NOWRAP>모품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="모품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo2 frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
								
							<TR HEIGHT="100%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=100% NAME=vspdData WIDTH=100% tag="2" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hCompntCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hSLCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdMgr" tag="24"><INPUT TYPE=HIDDEN NAME="hInvMgr" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="hJobCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hIssueFlag" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hIssueMthd" TABINDEX = "-1" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
