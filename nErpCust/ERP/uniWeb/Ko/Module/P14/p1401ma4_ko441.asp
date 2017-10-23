<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1401ma4.asp
'*  4. Program Name         : BOM등록(Multi)
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/10/15
'*  8. Modified date(Last)  : 2004/03/18
'*  9. Modifier (First)     : Jung Yu Kyung
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
<SCRIPT LANGUAGE = "VBScript" SRC = "p1401ma4_ko441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Dim i, iStrArr, iStrNmArr
    Dim strCbo  
    Dim strCboCd
    Dim strCboNm 
	'****************************
    'List Minor code(유무상구분)
    '****************************
    'strCboCd = "" & vbTab & ""
    'strCboNm = "" & vbTab 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2201", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iStrArr = Split(lgF0, Chr(11))
    iStrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iStrArr) - 1
        strCboCd = strCboCd & iStrArr(i) & vbTab
        strCboNm = strCboNm & iStrNmArr(i) & vbTab
	Next

	iStrFree = iStrNmArr(1)
	
    ggoSpread.SetCombo strCboCd, C_SupplyFlg 'parent.ggoSpread.SSGetColsIndex()              'Supply Flag setting
    ggoSpread.SetCombo strCboNm, C_SupplyFlgNm 'parent.ggoSpread.SSGetColsIndex()            'Supply Flag Nm Setting
    
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029  
	Call AppendNumberPlace("6", "5", "0")
	Call AppendNumberPlace("7", "2", "2")
	Call AppendNumberPlace("8", "11", "6")
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec, FALSE,, ggStrMinPart, ggStrMaxPart)
    
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet

    '----------  Coding part  -------------------------------------------------------------
    Call GetValue_ko441()
    Call SetDefaultVal

    Call InitVariables                                                      '⊙: Initializes local global variables
    Call InitComboBox
	
    Call SetToolbar("11100000000011")										'⊙: 버튼 툴바 제어 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement  
		Call txtPlantCd_OnChange()
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
    isClicked =  False

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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>BOM등록(Multi)</font></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="VBScript:OpenBOMCopy()">BOM 복사</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=90%>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtBaseDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="기준일" tag="12" VIEWASTEXT id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>
									</TD>
									<TD CLASS=TD5 ROWSPAN=2 NOWRAP>단계</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoStepType" ID="rdoSrchType1" CLASS="RADIO" tag="1X" Value="1"><LABEL FOR="rdoStepType1">단단계</LABEL></TD>													     
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>BOM Type</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBomNo" SIZE=5 MAXLENGTH=3 tag="12XXXU" ALT="BOM Type"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBomNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBomNo frm1.txtItemCd.value, frm1.txtBomNo.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtBOMNm" SIZE=15 tag="14"></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoStepType" ID="rdoSrchType2" CLASS="RADIO" tag="1X" Value="2" CHECKED><LABEL FOR="rdoStepType2">다단계</LABEL></TD>
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
								<TD CLASS=TD5 NOWRAP>품목계정</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=15 tag="24" ALT="품목계정"><INPUT TYPE=HIDDEN NAME="txtItemAcct" tag="24" ALT="품목계정"></TD>
								<TD CLASS=TD5 NOWRAP>기준단위</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBasicUnit" SIZE=5 MAXLENGTH=3 tag="24" ALT="기준단위"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>규격</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE=25 MAXLENGTH=40 tag="24" ALT="규격"></TD>
								<TD CLASS=TD5 NOWRAP>유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemFromDt CLASSID=<%=gCLSIDFPDT%> tag="24" ALT="시작일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemToDt CLASSID=<%=gCLSIDFPDT%> tag="24" ALT="종료일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>BOM 설명</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBOMDesc" SIZE=25 tag="24" ALT="BOM 설명"></TD>
								<TD CLASS=TD5 NOWRAP>설계변경번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="변경번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnECNNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNNo frm1.txtECNNo.value, 0"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>설계변경내용</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNDesc" SIZE=30 MAXLENGTH=100 tag="24" ALT="변경내용"></TD>
								<TD CLASS=TD5 NOWRAP>설계변경근거</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReasonCd" SIZE=4 MAXLENGTH=2 tag="24XXXU" ALT="변경근거"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReasonCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonCd frm1.txtReasonCd.value, 0">&nbsp;&nbsp<INPUT TYPE=TEXT NAME="txtReasonNm" SIZE=20 tag="24" ALT="변경근거명"></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" COLSPAN = 6>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBomType" tag="14">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHdrValidFromDt" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHdrValidToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHdrDrawingPath" tag="14">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
