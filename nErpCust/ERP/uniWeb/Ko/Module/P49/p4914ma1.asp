
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4914ma1.asp
'*  4. Program Name         : 작업일보 등록 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-01-17
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Yoon, Jeong Woo
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. History              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p4914ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim LocSvrDate
Dim StartDate
Dim EndDate

	LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
	StartDate = LocSvrDate	'UNIDateAdd("D",-10,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 처음 날짜 
	EndDate = UNIDateAdd("D", 20,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	<%
	Dim iData
    iData = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = 'P3211' ")
    Response.write "Call SetCombo3(frm1.cboOrderType, """ &  iData & """) " & vbCrLf
    %>
'	frm1.cboOrderType.value = ""

End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다.
'*********************************************************************************************************
'==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'===================================================================================================================
'----------------  ClickTab1(): Header Tab처리 부분 (Header Tab이 있는 경우만 사용)  ----------------------------
Function ClickTab1()

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		Call SetToolbar("11000000000111")
		Exit Function
    End If
	
	If gSelframeFlg = TAB1 Then Exit Function

'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field    
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call changeTabs(TAB1)	
	gSelframeFlg = TAB1
	lgIntFlgMode = parent.OPMD_CMODE
	
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function           
    End If 

End Function

Function ClickTab2()

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		Call SetToolbar("11000000000111")
		Exit Function
    End If

	If gSelframeFlg = TAB2 Then Exit Function

	If frm1.KeyProdtOrderNo2.value = "" Then
		Call DisplayMsgBox("800167", "X", "X", "X")
		Exit Function
	End If

'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field    

	ggoSpread.Source = frm1.vspdData4
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData5
    ggoSpread.ClearSpreadData
	Call changeTabs(TAB2)	
	gSelframeFlg = TAB2
	lgIntFlgMode = parent.OPMD_CMODE

	Call InitSpreadComboBox5()

'	frm1.txtProdOrderNo.Value = frm1.KeyProdtOrderNo2.value
	gMouseClickStatus = "SP4C"
	
    If DbQuery = False Then   
		Call RestoreToolBar()
		Exit Function           
    End If 
    	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리 
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
	Call InitSpreadSheet("*")															'⊙: Setup the Spread sheet

'	Call InitComboBox
	Call InitSpreadComboBox
	Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables

	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")

	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement
		Else
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement
		End If
	End If

    gTabMaxCnt = 2
    gIsTab = "Y"
	gSelframeFlg = TAB1
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>생산 및 유실 현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>간접공수/사고현황</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=300>&nbsp;</TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=7 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="라우팅명"></TD>
									<TD CLASS=TD5 NOWRAP>작업일자</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/p4914ma1_I604498142_txtprodDt.js'></script>
										<!--&nbsp;~&nbsp;-->
										<!--OBJECT classid=<%=gCLSIDFPDT%> name=txtprodToDt   CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="종료일"></OBJECT-->
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="12xxxU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConWC()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=25 tag="14" ALT="작업장명"></TD>
									<TD CLASS=TD5 NOWRAP><!--제조오더번호--></TD>
									<TD CLASS=TD6 NOWRAP><!--INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="제조오더번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProdOrderNo() "--></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<!-- style sheet 변경 -->
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<!-- 첫번째 탭 내용 -->
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD6Y6 NOWRAP>재적인원</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle1_txtDocAmt.js'></script>명</TD>
								<TD CLASS=TD6Y6 NOWRAP>지원인원(+)</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle2_txtDocAmt.js'></script>명</TD>
								<TD CLASS=TD6Y6 NOWRAP>잔업인원</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle3_txtDocAmt.js'></script>명</TD>
								<TD CLASS=TD6Y6 NOWRAP>작업공수</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle4_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD6Y6 NOWRAP>재적시간</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle5_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>지원시간(+)</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle6_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>잔업시간</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle7_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>총유실공수</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle8_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD6Y6 NOWRAP>휴업인원</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle9_txtDocAmt.js'></script>명</TD>
								<TD CLASS=TD6Y6 NOWRAP>지원인원(-)</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle10_txtDocAmt.js'></script>명</TD>
								<TD CLASS=TD6Y6 NOWRAP>잔업공수</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle11_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>&nbsp;</TD>
								<TD CLASS=TDT NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS=TD6Y6 NOWRAP>휴업시간</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle12_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>지원시간(-)</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle13_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>기타공수</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle14_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>
								<TD CLASS=TD6Y6 NOWRAP>실동공수</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle15_txtDocAmt.js'></script>&nbsp;&nbsp;</TD>

								<!--TD CLASS=TD6Y6 NOWRAP>실동공수</TD>
								<TD CLASS=TDT NOWRAP><script language =javascript src='./js/p4914ma1_fpDoubleSingle15_txtDocAmt.js'></script>&nbsp;&nbsp;</TD-->
							</TR>

							<TR>
								<TD WIDTH=100% HEIGHT=100% valign=top colspan=8>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR HEIGHT="50%">
											<TD WIDTH="100%" colspan=4>
												<script language =javascript src='./js/p4914ma1_A_vspdData1.js'></script>
											</TD>
										</TR>
										<TR HEIGHT="50%">
											<TD WIDTH="50%">
												<script language =javascript src='./js/p4914ma1_B_vspdData2.js'></script>
											</TD>
											<TD WIDTH="50%">
												<script language =javascript src='./js/p4914ma1_C_vspdData3.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>

						</TABLE>
						</DIV>

						<!-- 두번째 탭 내용 -->
						<DIV ID="TabDiv"  SCROLL="no" style="display:none">
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD WIDTH=100% HEIGHT=* valign=top>
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR HEIGHT="50%">
												<TD WIDTH="100%" colspan=4>
													<script language =javascript src='./js/p4914ma1_D_vspdData4.js'></script>
												</TD>
											</TR>
											<TR HEIGHT="50%">
												<TD WIDTH="100%" colspan=4>
													<script language =javascript src='./js/p4914ma1_E_vspdData5.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</DIV>

					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
    <TR HEIGHT="20">
      <TD WIDTH="100%">
		<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
			  <TD WIDTH=* align=right><A href="vbscript:OpenWorkDailyRef()">작업일보리스트</A> <!--| <A href="vbscript:ClickTab2()">간접공수/사고현황 등록</A--></TD>
	  		  <TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
      </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread1" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread3" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread4" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread5" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode0" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode1" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode2" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode3" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode4" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode5" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">

<INPUT TYPE=HIDDEN NAME="KeyProdtOrderNo2" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyOprNo2" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyProdtOrderNo3" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="KeyResourceCd3" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>