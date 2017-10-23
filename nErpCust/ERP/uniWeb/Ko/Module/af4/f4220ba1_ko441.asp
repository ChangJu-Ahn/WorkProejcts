<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4220ba1
'*  4. Program Name         : 차입금상환전개(posam)
'*  5. Program Desc         : Create of Loan Schedule
'*  6. Comproxy List        : PAFG400
'*  7. Modified date(First) : 2002.04.12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo min
'* 10. Modifier (Last)      : 2003.05.19
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'												1. 선 언 부 
'##############################################################################################################
'******************************************  1.1 Inc 선언   **************************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->		<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit  
                                                            '☜: indicates that All variables must be declared in advance 
 '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

 '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* 

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'<!-- #Include file="../../inc/lgvariables.inc" -->
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "f4220bb1_ko441.asp"  
 '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 

 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt
Dim lgIsOpenPop
Dim lgBlnFlgChgValue														   '☜: Variable is for Dirty flag
Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim IsOpenPop          
 '#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
'    lgIntFlgMode = OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
'    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'☆: 사용자 변수 초기화 
'    lgCboKeyPress = False
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
'<% Call LoadInfTB19029(gCurrency, "I", "*") %>
End Sub

 '******************************************  2.1 Pop-Up 함수   **********************************************
'	기능: Pop-Up 
'********************************************************************************************************* 

 '------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : Data Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopupLoan()

	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(3)	

	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("f4232ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f4232ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	

	IsOpenPop = True
    
	arrRet = window.showModalDialog(iCalledAspName & "?PGM=" & gStrRequestMenuID , Array(window.parent,arrParam), _
		     "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		frm1.txtLoanNo.focus
		Exit Function
	Else		
		frm1.txtLoanNo.value = arrRet(0)
		frm1.txtLoanNm.value = arrRet(1)
	End If
	
	frm1.txtLoanNo.focus

End Function

 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Sub SetDefaultVal()
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtDateFr.Text = frDt    
	frm1.txtDateTo.Text = toDt
End Sub

 '------------------------------------------  fnButtonExec()  --------------------------------------------------
'	Name : fnButtonExec()
'	Description : 실행 버튼 클릭 시 실행. 
'--------------------------------------------------------------------------------------------------------- 

Function fnButtonExec()
    Dim strVal           
    Dim strFrdt
    Dim strTodt
    Dim strTarget
	Dim IntRetCD
        
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
		Exit Function
    End If
	
	'--------------------------------------------------------------
	'차입일 조건 범위 Check
	'--------------------------------------------------------------

	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'⊙: GL Date Compare Common Function
			Exit Function
	End if
	
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
    Call LayerShowHide(1)
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태    
       
    strVal = strVal & "&txtDateFr=" & Trim(frm1.txtDateFr.text)  
    strVal = strVal & "&txtDateTo=" & Trim(frm1.txtDateTo.text)
    strVal = strVal & "&txtLoanNo=" & Trim(frm1.txtLoanNo.value)
    strVal = strVal & "&txtBizAreaCd=" & Trim(frm1.txtBizAreaCd.value)
    strVal = strVal & "&txtBizAreaCd1=" & Trim(frm1.txtBizAreaCd1.value)
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	    
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
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call InitVariables																'⊙: Initializes local global variables
    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.FormatDate(frm1.txtDateFr, parent.gDateFormat, 1)		'년월일 
    Call ggoOper.FormatDate(frm1.txtDateTo, parent.gDateFormat, 1)		'년월일 

    Call SetDefaultVal

    Call ggoOper.LockField(Document, "N")									'⊙: Lock  Suitable  Field
    Call SetToolbar("1000000000001111")
    
    frm1.fpDateFr.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txt1StPrRdpDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateFr.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDateFr.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDateTo_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateTo.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDateTo.Focus
    End If
    lgBlnFlgChgValue = True
End Sub

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd.Value	= arrRet(0)
		frm1.txtBizAreaNm.Value	= arrRet(1)
	End If
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd1.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		frm1.txtBizAreaCd1.Value = arrRet(0)
		frm1.txtBizAreaNm1.Value = arrRet(1)
	End If
End Function


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================
Function FncQuery()
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" ><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>차입금상환전개</font></td>
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
								<TD CLASS="TD5" NOWRAP>차입일자</TD>               
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f4220ba1_fpDateFr_txtDateFr.js'></script>&nbsp;~&nbsp;
													   <script language =javascript src='./js/f4220ba1_fpDateTo_txtDateTo.js'></script>
								</TD>
							</TR>					
							<TR>
								<TD CLASS="TD5" NOWRAP>차입금번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="차입금번호" SIZE=15 MAXLENGTH="18"  tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopUpLoan()">&nbsp;<INPUT NAME="txtLoanNm" ALT="차입금내역" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd()">
								                       <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd1()">
									                      <INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>처리건수</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCount" ALT="처리건수" SIZE=15 STYLE="TEXT-ALIGN:right" tag="14X"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" OnClick="VBScript:Call fnButtonExec()" Flag=1>실행</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>
