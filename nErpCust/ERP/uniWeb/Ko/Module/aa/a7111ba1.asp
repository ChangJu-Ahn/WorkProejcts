<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Asset Management
'*  3. Program ID           : a7111ma1.asp
'*  4. Program Name         : 감가상각계산 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             AS0051
'                             
'*  7. Modified date(First) : 2000/03/31
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->			<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>


<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'==========================================================================================================

Const BIZ_PGM_ID = "a7111bb1.asp"  

'========================================================================================================= 

Dim lgMpsFirmDate, lgLlcGivenDt

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim IsOpenPop          



Function OpenMasterRef(field_fg)

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		IsOpenPop = False
		Exit Function
	Else
		Call SetPoRef(arrRet,field_fg)
	End If	

	IsOpenPop = False

	'frm1.txtCondAsstNo.focus
	
End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(byval arrRet,byval field_fg)
       
	Select case field_fg
	case 0
		frm1.txtFrAsstCd.value	= arrRet(0)
		frm1.txtFrAsstNm.value	= arrRet(1)
		frm1.txtFrAsstCd.focus
	case 1
		frm1.txtToAsstCd.value	= arrRet(0)
		frm1.txtToAsstNm.value	= arrRet(1)
		frm1.txtToAsstCd.focus
	End select
	
	
		
End Sub




'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : Data Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(field_fg)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    
    Select Case field_fg
		Case 1  '계정코드 
			If IsOpenPop = True Or UCase(frm1.txtAcct.className) = "PROTECTED" Then Exit Function
		Case 2  '시작 자산코드 
			If IsOpenPop = True Or UCase(frm1.txtFrAsstCd.className) = "PROTECTED" Then Exit Function
		Case 3  '종료 자산코드 
			If IsOpenPop = True Or UCase(frm1.txtToAsstCd.className) = "PROTECTED" Then Exit Function
		Case 4  '시작 계정코드 
			If IsOpenPop = True Or UCase(frm1.txtFrAcctCd.className) = "PROTECTED" Then Exit Function
		Case 5  '종료 계정코드 
			If IsOpenPop = True Or UCase(frm1.txtToAcctCd.className) = "PROTECTED" Then Exit Function
	End Select 

	IsOpenPop = True
	Select Case  field_fg 
		Case 1
			arrParam(0) = "계정코드 팝업"	
			arrParam(1) = "A_ACCT"
			arrParam(2) = Trim(frm1.txtACCT.Value)
			arrParam(3) = Trim(frm1.txtACCTNm.Value)
			arrParam(4) = ""
			arrParam(5) = "계정코드"

			arrField(0) = "ACCT_CD"	
			arrField(1) = "ACCT_SH_NM"	

			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정명"
		Case 2, 3
			arrParam(0) = "자산코드 팝업"	
			arrParam(1) = "A_ASSET_MASTER"
			arrParam(2) = Trim(frm1.txtFrAsstCd.Value)
			arrParam(3) = Trim(frm1.txtFrAsstNm.Value)
			arrParam(4) = ""
			arrParam(5) = "자산코드"

			arrField(0) = "ASST_NO"
'			arrField(0) = "F2" & parent.gColSep & "ACQ_SEQ"
			arrField(1) = "ASST_NM"	

			arrHeader(0) = "자산코드"
			arrHeader(1) = "자산명"
		Case 4
			arrParam(0) = "계정코드 팝업"	
			arrParam(1) = "A_ASSET_ACCT A, A_ACCT B"
			arrParam(2) = Trim(frm1.txtFrAcctCd.Value)
			arrParam(3) = ""
			arrParam(4) = "A.ACCT_CD =B.ACCT_CD"
			arrParam(5) = "계정코드"

			arrField(0) = "A.ACCT_CD"
			arrField(1) = "B.ACCT_NM"

			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정명"
		Case 5
			arrParam(0) = "계정코드 팝업"	
			arrParam(1) = "A_ASSET_ACCT A, A_ACCT B"
			arrParam(2) = Trim(frm1.txtToAcctCd.Value)
			arrParam(3) = ""
			arrParam(4) = "A.ACCT_CD =B.ACCT_CD"
			arrParam(5) = "계정코드"

			arrField(0) = "A.ACCT_CD"
			arrField(1) = "B.ACCT_NM"

			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정명"

	End Select

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	PoupSetFocusVal(field_fg)
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,field_fg)
	End If	
End Function

'------------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(byval arrRet,byval field_fg)

	Select case field_fg
	case 1
		frm1.txtAcct.Value        = Trim(arrRet(0))
		frm1.txtAcctNm.Value      = arrRet(1)
	case 2
		frm1.txtFrAsstCd.Value    = Trim(arrRet(0))
		frm1.txtFrAsstNm.Value    = arrRet(1)
	case 3
		frm1.txtToAsstCd.Value    = Trim(arrRet(0))
		frm1.txtToAsstNm.Value    = arrRet(1)
	case 4
		frm1.txtFrAcctCd.Value    = Trim(arrRet(0))
		frm1.txtFrAcctNm.Value    = arrRet(1)
'		If Trim(frm1.txtToAcctCd.Value) = "" or Trim(frm1.txtFrAcctCd.Value) > Trim(frm1.txtToAcctCd.Value) Then
			frm1.txtToAcctCd.Value    = Trim(arrRet(0))
			frm1.txtToAcctNm.Value    = arrRet(1)
'		End If
	case 5
		frm1.txtToAcctCd.Value    = Trim(arrRet(0))
		frm1.txtToAcctNm.Value    = arrRet(1)
	End select
End Function

Function PoupSetFocusVal(byval field_fg)
	Select case field_fg
	case 1
		frm1.txtAcct.Focus
	case 2
		frm1.txtFrAsstCd.Focus
	case 3
		frm1.txtToAsstCd.Focus
	case 4
		frm1.txtFrAcctCd.Focus
	case 5
		frm1.txtToAcctCd.Focus
	End select
End Function
 '------------------------------------------  ExeReflect()  --------------------------------------------------
'	Name : ExeReflect()
'	Description : 실행 버튼 클릭 시 실행. 
'--------------------------------------------------------------------------------------------------------- 
Function ExeReflect()
    Dim strVal           
    Dim strFrdt
    Dim strTodt
    Dim strTarget
    Dim RetFlag
	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strYear1
	Dim strMonth1
	Dim strDay1
	
	ExeReflect = False  	
	
    '-----------------------
    'Check content area
    '-----------------------
    Err.Clear
    
    If Not chkField(Document, "2") Then        '⊙: Check contents area
		Exit Function
    End If


 	Call ExtractDateFrom(frm1.fpDateTime1.Text,frm1.fpDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
    strFrDt = strYear & strMonth

 	Call ExtractDateFrom(frm1.fpDateTime2.Text,frm1.fpDateTime2.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    strToDt = strYear1 & strMonth1

	 RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"X","X")   '☜ 바뀐부분 
	'RetFlag = Msgbox("작업을 수행 하시겠습니까?", vbOKOnly + vbInformation, "정보")

	If RetFlag = VBNO Then
		Exit Function
	End If   
	        
    Call LayerShowHide(1)
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002					'☜: 비지니스 처리 ASP의 상태    
       
    If frm1.Rb_WK1.checked = True Then
		strVal = strVal & "&txtRadio=" & "C"								' 기업회계기준 
    Else
		strVal = strVal & "&txtRadio=" & "T"								' 세법기준 
	End If    

    If frm1.Rb_CAL1.checked = True Then
		strVal = strVal & "&txtCAL=" & "C"									' 감가상각계산 
    Else
		strVal = strVal & "&txtCAL=" & "D"									' 계산취소 
	End If    

	strVal = strVal & "&txtFrAsstCd=" & frm1.txtFrAsstCd.value
	strVal = strVal & "&txtToAsstCd=" & frm1.txtToAsstCd.value
	strVal = strVal & "&txtFrAcctCd=" & frm1.txtFrAcctCd.value
	strVal = strVal & "&txtToAcctCd=" & frm1.txtToAcctCd.value
	
    strVal = strVal & "&txtFryymm=" & strFrDt & "&txtToyymm=" & strToDt   
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                       
End Function

Function ExeReflectOk()
    Dim IntRetCD 

	IntRetCD = DisplayMsgBox("990000","X","X","X")   '☜ 바뀐부분 
	
End function

Function ExeReflectNo()
	Dim IntRetCD 

    'Call DisplayMsgBox("","X","X","X") 				            '☆: 실행된 자료가 없습니다 

End Function
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
Function FncPrint() 
	Parent.fncPrint()    
End Function

Function FncQuery()
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "BA") %>
End Sub

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029	
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtFryymm, gDateFormat, 2)	
	
	frm1.fpDateTime1.Focus
	frm1.fpDateTime1.Text = UNIMonthClientFormat(parent.gFiscStart)	
	frm1.fpDateTime2.Text = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,gDateFormat)

    Call ggoOper.FormatDate(frm1.txtToyymm, gDateFormat, 2)    
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field    

    Call SetToolbar("10000000000011")										'⊙: 버튼 툴바 제어 
	
End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'======================================================================================================
'   Event Name : DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtFrYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrYYMM.Action = 7
	End If
End Sub

Sub txtToYYMM_DblClick(Button)
	If Button = 1 Then
		frm1.txtToYYMM.Action = 7
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0 >
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE  <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감가상각계산</font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>상각시작년월</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrYymm" CLASS=FPDTYYYYMM tag="22X1" Title="FPDATETIME" ALT=상각시작년월 id=fpDateTime1> </OBJECT>');</SCRIPT>
								</TD>
							</TR>					
							<TR>
								<TD CLASS="TD5" NOWRAP>상각종료년월</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToYymm" CLASS=FPDTYYYYMM tag="22X1" Title="FPDATETIME" ALT=상각종료년월 id=fpDateTime2> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>내용년수 구분</TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked><LABEL FOR=Rb_WK1>기업회계기준</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2><LABEL FOR=Rb_WK2>세법기준</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업 구분</TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio ID=Rb_CAL1 Checked><LABEL FOR=Rb_CAL1>감가상각계산</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO ID=Rb_CAL2><LABEL FOR=Rb_CAL2>계산취소</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 nowrap>시작계정코드</TD>
								<TD CLASS=TD6><INPUT NAME="txtFrAcctCd" ALT="시작계정코드" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag="2X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopUp(4)">&nbsp;
													 <INPUT NAME="txtFrAcctNm" ALT="계정코드명"       MAXLENGTH="30"  tag="24XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 nowrap>종료계정코드</TD>
								<TD CLASS=TD6><INPUT NAME="txtToAcctCd" ALT="종료계정코드" MAXLENGTH="20" STYLE="TEXT-ALIGN: left" tag="2X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopUp(5)">&nbsp;
													 <INPUT NAME="txtToAcctNm" ALT="계정코드명"       MAXLENGTH="30"   tag="24XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 nowrap>시작자산코드</TD>
								<TD CLASS=TD6><INPUT NAME="txtFrAsstCd" ALT="시작자산코드" MAXLENGTH="18" STYLE="TEXT-ALIGN: left" tag="2X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMasterRef(0)">&nbsp;
													 <INPUT NAME="txtFrAsstNm" ALT="자산명"       MAXLENGTH="7"  tag="24XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 nowrap>종료자산코드</TD>
								<TD CLASS=TD6><INPUT NAME="txtToAsstCd" ALT="종료자산코드" MAXLENGTH="18" STYLE="TEXT-ALIGN: left" tag="2X"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMasterRef(1)">&nbsp;
													 <INPUT NAME="txtToAsstNm" ALT="자산명"       MAXLENGTH="7"   tag="24XXXU"></TD>
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
					<TD><BUTTON NAME="btn배치" CLASS="CLSMBTN" OnClick="VBScript:Call ExeReflect()" Flag=1>실행</BUTTON> &nbsp</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" ></iframe>
</DIV>
</BODY>
</HTML>

