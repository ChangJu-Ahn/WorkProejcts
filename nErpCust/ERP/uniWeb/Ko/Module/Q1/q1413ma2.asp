<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MA2
'*  4. Program Name         : 전문가시스쳄 검사방식 선별형 적용화면	
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>T계수 선별형 샘플링 검사방식 적용</TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit 

Const BIZ_PGM_QRY_ID = "q1413Mb2.asp"							'☆: Query 비지니스 로직 ASP명 
Const PGM_JUMP_ID1 = "q1411ma1"
Const PGM_JUMP_ID2 = "Q1441MA1.asp"
Const PGM_JUMP_ID3 = "Q1442MA1.asp"

Dim lgQueryFlag				 '--- 1:New Query 0:Continuous Query 

Dim lgInspClassCd
Dim lgPlantCd
Dim lgItemCd
Dim lgInspReqNo
Dim lgBpCd

Dim hPlantCd
Dim hItemCd
Dim hInspReqNo
Dim hBpCd

Dim arrParam					 '--- First Parameter Group 
Dim arrReturn				 '--- Return Parameter Group 

Dim IsOpenPop          

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()

End Function

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	Dim strTypeNumber

	If ReadCookie("txtInsTypeNumber") <> "" Then
		strTypeNumber = ReadCookie("txtInsTypeNumber")
		WriteCookie "txtInsTypeNumber", ""
	End If
	
	With frm1
		.txtLotSize.AllowNull = True
		.txtLotSize.Text = ""
		.txtDefectRate.AllowNull = True
		.txtDefectRate.Text = ""
		
		If strTypeNumber = 0 then
			.rdoAssure.rdoAssure1.checked = true
			.txtTypeNumber.value = strTypeNumber
			.cboAOQL.value = ""
			Call LockAOQLLTPD("A")
			Call ClearDataAsLotQualityIndex("A")
		End if
	
		If strTypeNumber = 1 then
			.rdoAssure.rdoAssure2.checked = true
			.txtTypeNumber.value = strTypeNumber
			Call LockAOQLLTPD("L")
			Call ClearDataAsLotQualityIndex("L")
		End if
	End With
	
End Sub

'==========================================  2.2.2 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	Dim TmplgF0
	Dim TmplgF1
	Dim i

	TmplgF0 = split(lgF0,Chr(11))
	TmplgF1 = split(lgF1,Chr(11))	
	lgF0 = ""
	lgF1 = ""
	
	For i = 0 To UBound(TmplgF0) - 1
        lgF0 = lgF0 & uniConvNumAToB(TmplgF0(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next

	For i = 0 To UBound(TmplgF1) - 1
        lgF1 = lgF1 & uniConvNumAToB(TmplgF1(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next
	
    Call SetCombo2(frm1.cboAOQL ,lgF0  ,lgF1  ,Chr(11))
    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " major_cd=" & FilterVar("Q0021", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	
	TmplgF0 = split(lgF0,Chr(11))
	TmplgF1 = split(lgF1,Chr(11))	
	lgF0 = ""
	lgF1 = ""
	
	For i = 0 To UBound(TmplgF0) - 1
        lgF0 = lgF0 & uniConvNumAToB(TmplgF0(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next
	
	For i = 0 To UBound(TmplgF1) - 1
        lgF1 = lgF1 & uniConvNumAToB(TmplgF1(i),parent.gAPNum1000,parent.gAPNumDec,parent.gComNum1000,parent.gComNumDec,True,"X","X") & Chr(11)
	Next
	
    Call SetCombo2(frm1.cboLTPD ,lgF0  ,lgF1  ,Chr(11))  
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()

End Sub

'===========================================  2.3.1 ResultClick()  ==========================================
'=	Name : ResultClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function ResultClick()

	Dim strVal

	If Not chkField(Document, "2") Then  '⊙: Check contents area
       		Exit Function
    	End If

	IF frm1.rdoAssure.rdoAssure1.checked = true then
		frm1.txtTypeNumber.Value = "A"				'AOQL보증				'
	End IF

	IF frm1.rdoAssure.rdoAssure2.checked = true then
		frm1.txtTypeNumber.Value = "L"				'LTPD보증 
	End IF

	IF  Trim(frm1.txtTypeNumber.Value) = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'선택사항을 체크하십시오 
		Exit Function	
	End IF

	IF frm1.rdoAssure.rdoAssure1.checked = true and Trim(frm1.cboAOQL.Value) = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF

	IF frm1.rdoAssure.rdoAssure2.checked = true and Trim(frm1.cboLTPD.Value) = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF
	
	Call ggoOper.ClearField(Document, "1")										'⊙: Clear Contents  Field
	
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY_ID & "?txtLotSize=" & frm1.txtLotSize.Text   '☜: '☆: 조회 조건 데이타 
	strVal = strVal & "&txtAOQL=" & Trim(frm1.cboAOQL.Value)   
	strVal = strVal & "&txtDefectRate=" & frm1.txtDefectRate.Text
	strVal = strVal & "&txtLTPD=" & Trim(frm1.cboLTPD.Value)
	strVal = strVal & "&txtTypeNumber=" & Trim(frm1.txtTypeNumber.Value)		'AOQL or LTPD
	
	Call RunMyBizASP(MyBizASP, strVal)
		
	
End Function

'=========================================  2.3.2 ShowGraphClick()  ========================================
'=	Name : ShowGraphClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================

Function ShowGraphClick()

Dim strVal

	IF frm1.rdoDefect.rdoDefect1.checked = true then		
		frm1.txtReplaceMode.value = 0		'No replacement
	End IF
	
	IF frm1.rdoDefect.rdoDefect2.checked = true then		
		frm1.txtReplaceMode.value = 1		'Replacement
	End IF

	IF (frm1.rdoDefect.rdoDefect1.checked = False and frm1.rdoDefect.rdoDefect2.checked = False) then
		frm1.txtReplaceMode.value = 0
	End IF

	IF frm1.txtSampleSize.Text = "" then
		Call DisplayMsgBox("229920", "X", "X", "X") 		'결과항목이 없습니다 
		Exit Function	
	End IF

	strVal = PGM_JUMP_ID2 & "?txtLotSize=" & frm1.txtLotSize.Text
	strVal = strVal & "&txtSampleSize=" & frm1.txtSampleSize.Text
	strVal = strVal & "&txtAcceptSize=" & frm1.txtAcceptSize.Text
	strVal = strVal & "&txtDefectRate=" & frm1.txtDefectRate.Text	
	
	strVal = strVal & "&txtReplaceMode=" & Trim(frm1.txtReplaceMode.Value)	'Replacement
	'/* Issue: 검사방식 적용으로 Return - START */
	strVal = strVal & "&txtPageCode=" & "SA"
	'/* Issue: 검사방식 적용으로 Return - END */
	Location.href = strVal

End Function

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()
	PgmJump(PGM_JUMP_ID1)
End Function

'=======================================================================================================
'   Sub Name : LockAOQLLTPD()
'   Sub Desc : 
'=======================================================================================================
Sub LockAOQLLTPD(Byval vLotQualityIndex)
	With frm1
		Select Case vLotQualityIndex
			Case "A"
				Call ggoOper.SetReqAttr(.cboAOQL, "N")
				Call ggoOper.SetReqAttr(.cboLTPD, "Q")
			Case "L"
				Call ggoOper.SetReqAttr(.cboLTPD, "N")
				Call ggoOper.SetReqAttr(.cboAOQL, "Q")
			Case Else
				Call ggoOper.SetReqAttr(.cboLTPD, "Q")
				Call ggoOper.SetReqAttr(.cboAOQL, "Q")
		End Select 
	End With
End Sub

'=======================================================================================================
'   Event Name : ClearDataAsLotQualityIndex()
'   Event Desc : 
'=======================================================================================================
Sub ClearDataAsLotQualityIndex(Byval vLotQualityIndex)
		With frm1
		Select Case vLotQualityIndex
			Case "A"
				.cboLTPD.value=""
			Case "L"
				.cboAOQL.value=""
			Case Else
				.cboAOQL.value=""
				.cboLTPD.value=""
		End Select 
	End With
End Sub
'=========================================  2.3.4 Mouse Pointer 처리 함수 ===============================
'
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6", "3", "2")
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call InitVariables																'⊙: Initializes local global variables          
    Call InitComboBox
    Call SetDefaultVal
	Call SetToolbar("10000000000111")
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
						
End Sub

'=======================================================================================================
'   Event Name : rdoAssure1_Click()
'   Event Desc : 
'=======================================================================================================
Sub rdoAssure1_onClick()
	Call LockAOQLLTPD("A")
	Call ClearDataAsLotQualityIndex("A")
End Sub

'=======================================================================================================
'   Event Name : rdoAssure2_Click()
'   Event Desc : 
'=======================================================================================================
Sub rdoAssure2_onClick()
	Call LockAOQLLTPD("L")
	Call ClearDataAsLotQualityIndex("L")
End Sub

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then    				'⊙: Check contents area
       		Exit Function
    	End If

	Call ResultClick()
	   
End Function

'==========================================  3.2.1 Search_OnClick =======================================
'
'========================================================================================================
Sub Search_OnClick()    	
End Sub

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 Call parent.FncExport(Parent.C_SINGLE)					'☜: 화면 유형 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()

End Function

'********************************************  5.1 DbQueryOk()  *******************************************
' Function Name : DbQueryOk																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQueryOk()								'☆: 조회 성공후 실행로직 

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> BORDER=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>계수 선별형 샘플링 검사</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR height=*>
		<TD  VALIGN="TOP" WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD VALIGN="top"  WIDTH="100%" HEIGHT=*>
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
							<LEGEND>선택사항</LEGEND>
								<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
									<TR>
										<TD CLASS=TD5  HEIGHT=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>보증대상</TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAssure" TAG="2X" ID="rdoAssure1"><LABEL FOR="rdoAssure1">AOQL보증</LABEL></TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoAssure" TAG="2X" ID="rdoAssure2"><LABEL FOR="rdoAssure2">LTPD보증</LABEL></TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS=TD5   HEIGHT=15 NOWRAP>불량품 대체여부</TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDefect" TAG="2X" ID="rdoDefect1"><LABEL FOR="rdoDefect1">불량품 대체안함</LABEL></TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDefect" TAG="2X" ID="rdoDefect2"><LABEL FOR="rdoDefect2">불량품 대체함</LABEL></TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH=20% BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS=TD5  HEIGHT=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
										<TD WIDTH="20%" BGCOLOR="#F7F7F7" HEIGHT=5 CELLPADDING=5 NOWRAP></TD>
									</TR>
								</TABLE>
							</FIELDSET>							
							<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
							<LEGEND>입력사항</LEGEND>
								<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
									<TR>
										<TD CLASS="TD5" NOWRAP>로트크기</TD>
										<TD CLASS="TD6" NOWRAP>
											<script language =javascript src='./js/q1413ma2_txtLotSize_txtLotSize.js'></script>
										</TD>						
										<TD CLASS="TD5" NOWRAP></TD>
										<TD CLASS="TD6" NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>AOQL</TD>
										<TD CLASS="TD6" NOWRAP><SELECT NAME="cboAOQL" ALT="AOQL" STYLE="WIDTH: 80px" tag="22"></SELECT>&nbsp;%</TD>
										<TD CLASS="TD5" NOWRAP>LTPD</TD>
										<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLTPD" ALT="LTPD" STYLE="WIDTH: 80px" tag="22"></SELECT>&nbsp;%</TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>공정평균불량률</TD>
										<TD CLASS="TD6" NOWRAP>
											<script language =javascript src='./js/q1413ma2_fpDoubleSingle1_txtDefectRate.js'></script>&nbsp;%
										</TD>
										<td CLASS="TD5" NOWPAP HEIGHT=5></td>
										<td CLASS="TD6" NOWPAP HEIGHT=5></td>
									</TR>	
								</TABLE>
							</FIELDSET>							
							<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
							<LEGEND>결과</LEGEND>
								<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
									<TR>
										<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
										<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>샘플크기</TD>
										<TD CLASS="TD6" NOWRAP>
											<script language =javascript src='./js/q1413ma2_txtSampleSize_txtSampleSize.js'></script>
										</TD>
										<TD CLASS="TD5" NOWRAP>합격판정개수</TD>
										<TD CLASS="TD6" NOWRAP>
											<script language =javascript src='./js/q1413ma2_txtAcceptSize_txtAcceptSize.js'></script>
										</TD>
									</TR>
									<TR>
										<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
										<TD CLASS="TD5" HEIGHT=5 NOWRAP></TD>
										<TD CLASS="TD6" HEIGHT=5 NOWRAP></TD>
									</TR>
								</TABLE>
							</FIELDSET>
					</TD>	
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT="20">
      		<TD WIDTH="100%">
      			<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_30%>>
        				<TR>
        					<TD WIDTH=10>&nbsp;</TD>       			
	        				<TD><BUTTON NAME="btnResult" CLASS="CLSMBTN" ONCLICK="vbscript:ResultClick()">결과 보기</BUTTON></TD>
	        				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:ShowGraphClick">검사특성 그래프</A>&nbsp;|&nbsp;<A href="vbscript:ReturnClick()">전문가 시스템 질의</A></TD>
	        				<TD WIDTH=10>&nbsp;</TD>
	       			</TR>
      			</TABLE>
      		</TD>
    	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtTypeNumber" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtReplaceMode" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
