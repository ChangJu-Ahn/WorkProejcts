<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1713ma1.asp
'*  4. Program Name         : 제조BOM작성(이관의뢰)
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/27
'*  8. Modified date(Last)  : 2005/01/27
'*  9. Modifier (First)     : Cho Yong Chill
'* 10. Modifier (Last)      : Cho Yong Chill
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
<SCRIPT LANGUAGE = "VBScript" SRC = "p1713ma1.vbs"></SCRIPT>
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
	'하단부 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo strCboCd, C_SupplyFlg              'Supply Flag setting
	ggoSpread.SetCombo strCboNm, C_SupplyFlgNm            'Supply Flag Nm Setting
	
	'상단부 
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.SetCombo strCboCd, C_SupplyFlg              'Supply Flag setting
	ggoSpread.SetCombo strCboNm, C_SupplyFlgNm            'Supply Flag Nm Setting
	
    
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
    
    Call SetDefaultVal

    Call InitVariables                                                      '⊙: Initializes local global variables
    Call InitComboBox
    Call SetToolbar("11101000000011")										'⊙: 버튼 툴바 제어 
    
    
    If parent.gPlant <> "" Then
		frm1.txtBasePlantCd.value = parent.gPlant
		'frm1.txtBasePlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement  
		Call txtBasePlantCd_OnChange()
	Else
		frm1.txtBasePlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
	
	Call SetCookieVal
	
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>제조BOM작성(이관의뢰)</font></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
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
								<TD CLASS=TD5 NOWRAP>기준공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBasePlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="기준공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBasePlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBasePlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtBasePlantNm" SIZE=25 tag="14"></TD>
								<TD CLASS=TD5 NOWRAP>모품목</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="모품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>대상공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDestPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="대상공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDestPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConDestPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtDestPlantNm" SIZE=25 tag="14"></TD>
								<TD CLASS=TD5 NOWRAP>이관의뢰번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReqTransNo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="이관의뢰번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReqTransNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReqTransNo()"></TD>
							</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
								    <TD CLASS=TD5 NOWRAP>이관의뢰번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReqTransNo2" SIZE=18 MAXLENGTH=18 tag="24XXXU" ALT="이관의뢰번호"></TD>
									<TD CLASS=TD5 NOWRAP>이관상태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtStatusNm" SIZE=20 MAXLENGTH=16 tag="24xxxU" ALT="이관상태"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=1 WIDTH=100%></TD>
				</TR>								
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData Width=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
    <TR HEIGHT="20">
	    <TD WIDTH="100%">
	    	<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="LEFT">
					    <BUTTON NAME="btnInit" CLASS="CLSMBTN" DISABLED="True">초기화</BUTTON>&nbsp;
					    <BUTTON NAME="btnRequest" CLASS="CLSMBTN" DISABLED="True">이관의뢰</BUTTON>&nbsp;
					    <BUTTON NAME="btnCancel" CLASS="CLSMBTN" DISABLED="True">이관의뢰취소</BUTTON>&nbsp;
					    <BUTTON NAME="btnCopy" CLASS="CLSMBTN" DISABLED="True" >복사</BUTTON>
					</TD>
				</TR>
	    	</TABLE>
	    </TD>
    </TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">

<INPUT TYPE=HIDDEN NAME="hStartDate" tag="24">
<INPUT TYPE=HIDDEN NAME="txtQueryType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtBaseDt" tag="14">

<INPUT TYPE=HIDDEN NAME="hBasePlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hDestPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqTransNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqTransDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hBomType" tag="24">
<INPUT TYPE=HIDDEN NAME="hMaxSeq" tag="14">	

<INPUT TYPE=HIDDEN NAME="hDescription" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemValidFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemValidToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hHdrValidFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hHdrValidToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDrawingPath" tag="24">
<INPUT TYPE=HIDDEN NAME="hStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hRequestingFlg" tag="24">

<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
