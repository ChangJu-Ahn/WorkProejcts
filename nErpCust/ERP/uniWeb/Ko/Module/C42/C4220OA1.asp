<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : C4220OA1
'*  4. Program Name         : 프로젝트별 제조원가명세서 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ig Sung, Cho
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================  -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>

<Script Language="VBScript">

Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "C4220OB1.asp"

Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgLngCurRows = 0
    
End Sub

Sub SetDefaultVal()

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)

	frm1.txtYyyyMmFr.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYyyyMmFr, Parent.gDateFormat, 2)
	frm1.txtYyyyMmTo.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYyyyMmTo, Parent.gDateFormat, 2)


End Sub

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("Q", "C", "NOCOOKIE", "BA")%>
End Sub

Function ExeReflect() 
	Dim strVal
	Dim strYyyymm
	Dim	strYear, strMonth, strDay
	
	If Not chkField(Document, "1") Then
		Exit Function
	End If

	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
	ExeReflect = False
    
	On Error Resume Next

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    
    strYYYYMM = strYear & strMonth	

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	strVal = strVal & "&txtYyyymm=" & strYyyymm

	Call RunMyBizASP(MyBizASP, strVal)

	frm1.txtSpread.value =  strVal

	
	ExeReflect = True
    
End Function

Function ExeReflectOk()
Dim IntRetCD

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	'window.status = "작업 완료"

End Function



Function SetPrintCond(StrEbrFile, strUrl)

	Dim strYyyymmFr, strYyyymmTo
	Dim	strYearFr, strMonthFr, strDayFr, strYearTo, strMonthTo, strDayTo
	Dim	strProjectCode, strTrackingNo

	SetPrintCond = False

	StrEbrFile = "c4220oa1"

	Call ExtractDateFrom(frm1.txtYyyyMmFr.Text,frm1.txtYyyyMmFr.UserDefinedFormat,Parent.gComDateType,strYearFr,strMonthFr,strDayFr)
	Call ExtractDateFrom(frm1.txtYyyyMmTo.Text,frm1.txtYyyyMmTo.UserDefinedFormat,Parent.gComDateType,strYearTo,strMonthTo,strDayTo)

	strYyyymmFr		= strYearFr & strMonthFr
	strYyyymmTo		= strYearTo & strMonthTo

	strProjectCode	= Trim(UCase(frm1.txtProjectCode.value))
	strTrackingNo	= Trim(UCase(frm1.txtTrackingNo.value))
	
	if strProjectCode = "" then		strProjectCode	= "%"
	if strTrackingNo = "" then		strTrackingNo	= "%"


	strUrl	= strUrl &	"strYyyyMmFr|"		& strYyyymmFr
	strUrl	= strUrl &	"|strYyyyMmTo|"		& strYyyymmTo
	StrUrl	= StrUrl &	"|strProjectCode|"	& strProjectCode
	StrUrl	= StrUrl &	"|strTrackingNo|"	& strTrackingNo
	
	SetPrintCond = True
End Function

Function FncBtnPreview() 
    
    Dim StrEbrFile, strUrl

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If

'	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
'	call FncEBCPreview(ObjName , strUrl)

	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPreview(ObjName , strUrl)

	
End Function



Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "프로젝트 팝업"
			arrParam(1) = "PMS_PROJECT"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "프로젝트"			
	
			arrField(0) = "PROJECT_CODE"
			arrField(1) = "PROJECT_NM"
			 
			arrHeader(0) = "프로젝트코드"
			arrHeader(1) = "프로젝트명"
		Case 1
			arrParam(0) = "Tracking No 팝업"
			arrParam(1) = "PMS_PRJ_PRODUCTS a(nolock) join PMS_PROJECT b(nolock) on a.project_code = b.project_code"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "Tracking No"
	
			arrField(0) = "a.tracking_no"
			arrField(1) = "b.PROJECT_CODE"
			arrField(2) = "b.PROJECT_NM"
			 
			arrHeader(0) = "Tracking No"
			arrHeader(1) = "프로젝트코드"
			arrHeader(2) = "프로젝트명"

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				frm1.txtProjectCode.focus
			Case 1
				frm1.txtTrackingNo.focus
			Case Else
		End Select
		Exit Function

	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

Function SetPopup(Byval arrRet, Byval iWhere)
	
	With frm1
	
		Select Case iWhere
			Case 0
				.txtProjectCode.focus
    			.txtProjectCode.value	= arrRet(0)
    			.txtProjectNm.value		= arrRet(1)
			Case 1
				.txtTrackingNo.focus
    			.txtTrackingNo.value	= arrRet(0)
			Case Else
				Exit Function
		End Select
	End With

End Function



Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables
	Call SetDefaultVal
	Call SetToolbar("10000000000000")

	frm1.txtYyyymm.focus
    
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub


Sub txtYyyyMmFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyyMmFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyyMmFr.focus
	End If
End Sub

Sub txtYyyyMmTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyyMmTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyyMmTo.focus
	End If
End Sub



Function FncQuery()

End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
	FncExit = True
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="작업년월" id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>								
							</TR>
							<TR>
									<TD CLASS="TD5" NOWRAP>출력년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyyMmFr" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="시작년월" id=fpDateTime2> </OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyyMmTo" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="종료년월" id=fpDateTime3> </OBJECT>');</SCRIPT>
									</TD>								
							</TR>
							<TR>
									<TD CLASS="TD5" NOWRAP>프로젝트</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtProjectCode" TYPE="Text" MAXLENGTH="20" tag="11XXXU" size="20" ALT="프로젝트코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProjectCode" align=top TYPE="BUTTON" onclick="vbscript:OpenPopup frm1.txtProjectCode.value, 0">
															<input NAME="txtProjectNm" TYPE="TEXT" tag="14XXX" size="30">
									</TD>
							</TR>
							<TR>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtTrackingNo" TYPE="Text" MAXLENGTH="20" tag="11XXXU" size="20" ALT="Tracking No"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" onclick="vbscript:OpenPopup frm1.txtTrackingNo.value, 1">
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
					<TD Width = 10> &nbsp </TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON>&nbsp;
						<BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="strUrl" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

