<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203ma1.asp
'*  4. Program Name         : Routing Information
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2002/12/17
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Hong Chang Ho
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p1203ma1_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim BaseDate
Dim StartDate

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

    Dim strCboCd 
    Dim strCboNm 
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    strCboCd = "" & vbTab
    strCboNm = "" & vbTab
  
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1006", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    ggoSpread.Source = frm1.vspdData
    lgF0 = "" & Chr(11) & lgF0
    lgF1 = "" & Chr(11) & lgF1
    ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_JobCd
    ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_JobNm
  
    '****************************
    'MileStone Flag Setting
    '****************************

    strCboCd = ""
    strCboCd = "Y" & vbTab & "N"
    
    ggoSpread.SetCombo strCboCd, C_MilestoneFlg
    
    '****************************
    'Insp Flag Setting
    '****************************

    strCboCd = ""
    strCboCd = "Y" & vbTab & "N"
    
    ggoSpread.SetCombo strCboCd,C_InspFlg
  
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
	Call InitSpreadSheet															'⊙: Setup the Spread sheet
	
	Call InitComboBox
    Call GetValue_ko441()
	Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolbar("11101101001011")
	Call ReadCookVal()
	
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			Call txtPlantCd_OnChange
			frm1.txtItemCd.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<!--'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>라우팅정보등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenRoutCopy()">라우팅 COPY</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="라우팅명"></TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value,0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>라우팅</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo" SIZE=15 MAXLENGTH=7 tag="12XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRoutingNo()">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="라우팅명"></TD>
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
								<TD CLASS=TD5 NOWRAP>품목</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="23XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd1.value,1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=50 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>라우팅</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRoutingNo1" SIZE=15 MAXLENGTH=7 tag="23XXXU" ALT="라우팅">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm1"  MAXLENGTH=40 SIZE=50 tag="21" ALT="라우팅명"></TD>
							<TR>
								<TD CLASS=TD5 NOWRAP>주라우팅</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoMajorRouting" ID="rdoMajorRouting1" Value="Y" CLASS="RADIO" tag="2X" CHECKED><LABEL FOR="rdoMajorRouting1">예</LABEL>
													 <INPUT TYPE="RADIO" NAME="rdoMajorRouting" ID="rdoMajorRouting2" Value="N" CLASS="RADIO" tag="2X"><LABEL op="rdoMajorRouting2">아니오</LABEL></TD>
								<TD CLASS=TD5 NOWRAP>유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME SIZE="10" MAXLENGTH="10" ALT="유효기간시작일" tag="23X1"> </OBJECT>');</SCRIPT>								
									&nbsp;~&nbsp; 
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22X1" ALT="유효기간종료일" MAXLENGTH="10" SIZE="10"> </OBJECT>');</SCRIPT>								
								</TD>
							</TR>					
							<TR>
								<TD CLASS=TD5 NOWRAP>작업지시 C/C</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=17 MAXLENGTH=10 tag="23XXXU" ALT="작업지시 C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCtr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCtr()">&nbsp;<INPUT NAME="txtCostNm" MAXLENGTH="20" SIZE=30 ALT ="코스트센타명" tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>라우팅 순서</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtALTRTVALUE CLASS=FPDS65 title=FPDOUBLESINGLE SIZE="3" MAXLENGTH="3" ALT="순서" tag="21X6Z" id=OBJECT1> </OBJECT>');</SCRIPT>
								</TD>							
							</TR>							
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=* Align=right><A href="vbscript:JumpAllocComp()">자품목투입정보 등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtBomNo" SIZE=10 MAXLENGTH=4 tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hRoutingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hOprCostFlag" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
