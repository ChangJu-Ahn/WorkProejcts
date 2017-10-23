<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : A_VAT
'*  3. Program ID		    : A6104BA
'*  4. Program Name         : 전자세금계산서일괄반영 
'*  5. Program Desc         : 전자세금계산서일괄반영 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/27
'*  8. Modified date(Last)  : 2002/08/28
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : hersheys
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit

'==========================================================================================================

Const BIZ_PGM_ID = "a6104bb1.asp"											 '☆: 비지니스 로직 ASP명 
 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgBlnFlgConChg				'☜: Condition 변경 Flag
Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode


'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt

Dim lgCurName()					'☆ : 개별 화면당 필요한 로칼 전역 변수 
'Dim cboOldVal          
Dim IsOpenPop          


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub

'=============================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A","NOCOOKIE","MA") %>
End Sub


'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim svrDate

	svrDate =  UNIDateClientFormat("<%=GetSvrDate%>")
	
	frm1.txtFromIssuedDt.text = UNIGetFirstDay(svrDate, Parent.gDateFormat)
	frm1.txtToIssuedDt.text   = svrDate
End Sub


'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: VAT 처리 함수 
'######################################################################################################### 
Function executeSP()
	Dim RetFlag
	Dim intRetCD

    Err.Clear                                                               '☜: Protect system from crashing
    
    Dim strVal
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If UniConvDateToYYYYMMDD(frm1.txtFromIssuedDt.text, Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToIssuedDt.text, Parent.gDateFormat,"")Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If

	RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
	'RetFlag = Msgbox("작업을 수행 하시겠습니까?", vbOKOnly + vbInformation, "정보")
	If RetFlag = VBNO Then
		Exit Function
	End IF

	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromIssuedDt.text)		'☆: 조회 조건 데이타 
    strVal = strVal & "&txtToDt="   & Trim(frm1.txtToIssuedDt.text)			'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
End Function


'===============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()


    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("10000000000011")    
    frm1.txtFromIssuedDt.focus 
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssuedDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtFromIssuedDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssuedDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssuedDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtToIssuedDt.Focus
    End If
End Sub


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>전자세금계산서일괄반영</font></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">발행일</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromIssuedDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="발행일" id=fpDateTime1></OBJECT>');</SCRIPT> ~ 
											    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToIssuedDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="발행일" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>
                                    <font color="red">* 주의<br>
                                                      &nbsp;&nbsp;본 화면은 다큐빌 적용 업체만 사용 가능합니다.<br>
                                                      &nbsp;&nbsp;다큐빌 미적용 업체는 적용이 불가 합니다.<br>
                                                      &nbsp;&nbsp;다큐빌을 통해 국세청 전송이 완료된 건에 대해서만 적용됩니다.
                                    </font>
                                </TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
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
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" OnClick="VBScript:Call executeSP()" Flag=1>실 행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tabindex="-1" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

