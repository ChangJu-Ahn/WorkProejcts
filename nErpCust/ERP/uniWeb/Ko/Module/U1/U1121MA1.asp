
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: SCM
'*  2. Function Name		: Firm Purchase Receipt Management
'*  3. Program ID			: Mscm16ma1.asp
'*  4. Program Name			: Manage Firm Purchase Receipt
'*  5. Program Desc			: Manage Firm Purchase Receipt
'*  6. Comproxy List		: 
'*  7. Modified date(First)	: 2004/07/12
'*  8. Modified date(Last)	: 
'*  9. Modifier (First)		: Park, BumSoo
'* 10. Modifier (Last)		: Park, BumSoo
'* 11. Comment	
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = " u1121ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit									'☜: indicates that All variables must be declared in advance

Dim LocSvrDate
Dim lgStrColorFlag
Dim lgStrPrevKey1
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

LocSvrDate = "<%=GetSvrDate%>"

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	'****************************
	'List Minor code(Order Type)
	'****************************
	<%
	Dim iData
    iData = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P3211' ")
	Response.write "Call SetCombo3(frm1.cboOrderType, """ &  iData & """) " & vbCrLf
	%>
	frm1.cboOrderType.value = "" 

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     				'⊙: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
    Call ggoOper.LockField(Document, "N")                                          			'⊙: Lock  Suitable  Field
	
	Call ggoOper.FormatNumber(frm1.txtDlvyTime,23,0,false)
    Call ggoOper.FormatNumber(frm1.txtDlvyTime2,59,0,false)

    Call InitSpreadSheet()                                                    				'⊙: Setup the Spread sheet

	Call SetDefaultVal
	
    Call InitVariables																		'⊙: Initializes local global variables

    Call InitSpreadComboBox()
'	Call SetToolBar("11000000000011")

		frm1.txtBpCd.focus	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>입고명세서조회/등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReqRef()">입고예정정보참조</A> </TD>
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
			 						<TD CLASS=TD5 NOWRAP>공급처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="12xxxU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDlvyno" align=top TYPE="BUTTON"ONCLICK="vbscript:Openbpcd()">&nbsp;
														 <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="공급처명"></TD>
									<TD CLASS=TD5 NOWRAP>명세서발행번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDlvyNo" ALT="명세서발행번호" TYPE="Text" MAXLENGTH=18 SiZE=18 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDlvyno" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDlvyNo()"></TD>
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
							<TR>
			 					<TD CLASS=TD5 NOWRAP>명세서발행번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDlvyNo2" SIZE=20 MAXLENGTH=18 tag="21" ALT="명세서발행번호"></TD>
								<TD CLASS=TD5 NOWRAP>납입시간</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtDlvyTime CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="납입시간"></OBJECT>');</SCRIPT>
											</TD>
											<TD>
											&nbsp;:&nbsp;
											</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtDlvyTime2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="납입시간"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>	
							</TR>
							<TR>
			 					<TD CLASS=TD5 NOWRAP>운송회사</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransCo" SIZE=30 MAXLENGTH=30 tag="21" ALT="운송회사"></TD>
								<TD CLASS=TD5 NOWRAP>차량번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 tag="21" ALT="차량번호"></TD></TD>
							</TR>
							<TR>
			 					<TD CLASS=TD5 NOWRAP>운전자명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDriver" SIZE=20 MAXLENGTH=20 tag="21" ALT="운전자명"></TD>
								<TD CLASS=TD5 NOWRAP>전화번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTelNo1" SIZE=20 MAXLENGTH=20 tag="21" ALT="전화번호1">&nbsp;<INPUT TYPE=TEXT NAME="txtTelNo2" SIZE=20 MAXLENGTH=20 tag="21" ALT="전화번호2"></TD>
							</TR>
							<TR>
			 					<TD CLASS=TD5 NOWRAP>입고장소</TD>
								<TD CLASS=TD6 COLSPAN = 3 NOWRAP><INPUT TYPE=TEXT NAME="txtDlvyPlace" SIZE=50 MAXLENGTH=50 tag="21" ALT="입고장소"></TD>
							</TR>
							<TR>
			 					<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 COLSPAN = 3 NOWRAP><INPUT TYPE=TEXT NAME="txtREMARK" SIZE=100 MAXLENGTH=100 tag="21" ALT="비고"></TD>
							</TR>
							<TR HEIGHT="100%">
								<TD WIDTH="100%" colspan = "4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData1 Id = "A" HEIGHT="100%" width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBPCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hDvFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hDvToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hrdoAppflg" tag="24"><INPUT TYPE=HIDDEN NAME="hPoFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hPoToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtcurr_dt" tag="24">
<P ID="divTextArea"></P>
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname">
	<input type="hidden" name="dbname">
	<input type="hidden" name="filename">
	<input type="hidden" name="condvar">
	<input type="hidden" name="date">                 
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>