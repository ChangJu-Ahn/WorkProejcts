<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : 회계관리 
'*  2. Function Name        : 부가세관리 
'*  3. Program ID		    : A6111MA1
'*  4. Program Name         : 부가세누락분디스켓생성 
'*  5. Program Desc         : 부가세누락분디스켓생성 
'*  6. Component List       : +
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/09/11
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Hye young ,Lee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

'==========================================================================================================
Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
EndDate     =   "<%=GetSvrDate%>"

Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
StartDate   = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
EndDate     = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)


Const BIZ_PGM_ID = "a6111bb1.asp"											 '☆: 비지니스 로직 ASP명 
 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgBlnFlgConChg				'☜: Condition 변경 Flag
Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag
Dim lgIntGrpCount				'☜: Group View Size를 조사할 변수 
Dim lgIntFlgMode					'☜: Variable is for Operation Status

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""

Dim lgBlnStartFlag				' 메세지 관련하여 프로그램 시작시점 Check Flag

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt	 '☜: 비지니스 로직 ASP에서 참조하므로 

Dim  lgCurName()					'☆ : 개별 화면당 필요한 로칼 전역 변수 
Dim  cboOldVal          
Dim  IsOpenPop          



'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE   '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False    '⊙: Indicates that no value changed
    lgIntGrpCount = 0           '⊙: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False			'☆: 사용자 변수 초기화 
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	
	frm1.txtIssueDT1.Text = StartDate
	frm1.txtIssueDT2.Text = EndDate
	frm1.txtReportDt.Text = EndDate
	frm1.txtBizAreaCD.focus 
	
    'frm1.txtIssueDt1.focus
    'frm1.btnExecute.disabled = True
    
    'frm1.txtBizAreaCD.value	= parent.gBizArea

	lgBlnStartFlag = False
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
End Sub


'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "세금신고사업장 팝업"					' 팝업 명칭 
			arrParam(1) = "B_TAX_BIZ_AREA"	 			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "세금신고사업장코드"					' 조건필드의 라벨 명칭 

			arrField(0) = "TAX_BIZ_AREA_CD"				' Field명(0)
			arrField(1) = "TAX_BIZ_AREA_NM"				' Field명(0)
    
			arrHeader(0) = "세금신고사업장코드"					' Header명(0)
			arrHeader(1) = "세금신고사업장명"					' Header명(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCD.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' 사업장 
				.txtBizAreaCD.focus
				.txtBizAreaCD.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNM.value = arrRet(1)
		End Select
	End With
End Function

'========================================================================================================= 
Sub Form_Load()

    Call InitVariables							'⊙: Initializes local global variables
    Call LoadInfTB19029							'⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")		'⊙: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal

    Call SetToolbar("1000000000000001")										'⊙: 버튼 툴바 제어 
	frm1.txtBizAreaCD.focus 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtIssueDt1_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt1.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtIssueDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt1_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt1_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt2_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssueDt2.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtIssueDt2.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStartDt2_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtIssueDt2_Change()
    'lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReportDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtReportDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReportDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtReportDt_Change()
    'lgBlnFlgChgValue = True
End Sub

 '#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'######################################################################################################### 
Function subVatDisk() 
Dim RetFlag
Dim strVal
Dim IntRetCD
Dim intI, strFileName, intChrChk	'특수문자 Check

	'-----------------------
    'Check content area
    '-----------------------
    
    '화일명으로 사용할 수 없는 특수문자 \/:*?"<>|&. 포함여부 확인 
	strFileName = frm1.txtFileName.value
	
	For intI = 1 To Len(strFileName)
		intChrChk = ASC(Mid(strFileName, intI, 1))
		If intChrChk = ASC("\") Or intChrChk = ASC("/") Or intChrChk = ASC(":") Or intChrChk = ASC("*") Or _
			intChrChk = ASC("?") Or intChrChk = 34 Or intChrChk = ASC("<") Or intChrChk = ASC(">") Or _
			intChrChk = ASC("|") OR intChrChk = ASC("&") OR intChrChk = ASC(".") Then
				intRetCD =  DisplayMsgBox("970029","X" , frm1.txtFileName.Alt, frm1.txtIssueDt2.Alt)
				Exit Function
		End If
	Next
	
	' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then        '⊙: Check contents area
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtIssueDt1.text,frm1.txtIssueDt2.text,frm1.txtIssueDt1.Alt,frm1.txtIssueDt2.Alt, _
        	               "970025",frm1.txtIssueDt1.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtIssueDt1.focus
	   Exit Function
	End If
    

	RetFlag = DisplayMsgBox("900018", parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
	'RetFlag = Msgbox("작업을 수행 하시겠습니까?", vbOKOnly + vbInformation, "정보")
	If RetFlag = VBNO Then
		Exit Function
	End IF

    Err.Clear                                                               '☜: Protect system from crashing

    With frm1

		Call LayerShowHide(1)
	
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 비지니스 처리 ASP의 상태 
		strVal = strVal & "&txtIssueDt1=" & Trim(.txtIssueDt1.text)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtIssueDt2=" & Trim(.txtIssueDt2.text)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBizAreaCD=" & UCase(Trim(.txtBizAreaCD.value))	'☆: 조회 조건 데이타 
		strVal = strVal & "&txtReportDt=" & Trim(.txtReportDt.text)		'☆: 조회 조건 데이타 
		strVal = strVal & "&txtFileName=" & Trim(.txtFileName.value)			'☆: 조회 조건 데이타 

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
	
    End With
    
End Function

Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                               '☜: Protect system from crashing

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0002							'☜: 비지니스 처리 ASP의 상태 
	strVal = strVal & "&txtFileName=" & pFileName							'☆: 조회 조건 데이타 

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
End Function



'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
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
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()														'☆: 삭제 성공후 실행 로직 
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================

Function DbQueryOk()							'☆: 조회 성공후 실행로직 
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()			'☆: 저장 성공후 실행 로직 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>



<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1"  CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTABP"><font color=white>부가세누락분디스켓생성</font></td>
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
								<TD CLASS=TD5 NOWRAP>계산서발행일</TD>
								<TD CLASS=TD6><script language =javascript src='./js/a6111ba1_fpDateTime2_txtIssueDt1.js'></script>
											  &nbsp; ~ &nbsp;
											  <script language =javascript src='./js/a6111ba1_fpDateTime2_txtIssueDt2.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>
								<TR>
								<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" tag="12XXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
												<INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=30 MAXLENGTH=50 STYLE="TEXT-ALIGN: left" tag="14X" ALT="세금신고사업장"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>		
							<TR>
								<TD CLASS=TD5 NOWRAP>신고일자</TD>
								<TD CLASS=TD6><script language =javascript src='./js/a6111ba1_fpDateTime2_txtReportDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>		
							<TR>
								<TD CLASS=TD5 NOWRAP>화일명</TD>
								<TD CLASS=TD6><INPUT TYPE=TEXT ID="txtFileName" NAME="txtFileName" SIZE=30 MAXLENGTH=30 STYLE="TEXT-ALIGN: left" tag="12X" ALT="화일명"></TD>
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
					<TD><BUTTON NAME="btnExecute" CLASS="CLSMBTN" OnClick="VBScript:Call subVatDisk()" Flag=1>실 행</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>

