<%@ LANGUAGE="VBSCRIPT" %>  
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1413MA5
'*  4. Program Name         : 전문가시스템 검사방식 계량 조정형 적용화면	
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/07/30
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Koh Jae Woo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML> 
<HEAD> 
<TITLE>T계량 조정형 샘플링 검사방식 적용</TITLE>

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

Const BIZ_PGM_QRY_ID = "q1413Mb6.asp"							'☆: Query 비지니스 로직 ASP명 
Const PGM_JUMP_ID1 = "q1411ma1"
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
Dim arrReturn					 '--- Return Parameter Group 

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

'========================================== 2.2.1 SetDefaultVal() ======================================== 
' Name : SetDefaultVal() 
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting) 
'=========================================================================================================  
Sub SetDefaultVal() 
	'/* Issue: 표준편차에 초기값이 안들어가도록 수정 - START */
	With frm1
		.txtLotSize.AllowNull = True
		.txtLotSize.Text = ""
		.txtSD.AllowNull = True
		.txtSD.Text = ""
	End With
	'/* Issue: 표준편차에 초기값이 안들어가도록 수정 - END */
End Sub 

'=============================================  2.3.3()  ======================================
'=	Event Name : ReturnClick
'=	Event Desc :
'========================================================================================================
Function ReturnClick()
	PgmJump(PGM_JUMP_ID1)
End Function

'------------------------------------------  OpenAQL()  -------------------------------------------------
'	Name : OpenAQL()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAQL()
	OpenAQL = false
	
	Dim arrRet
	Dim arrParam1, arrParam2
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = frm1.txtAQL.value
	arrParam2 = "Q0012"
	
	iCalledAspName = AskPRAspName("q1211pa3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2), _
	              "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	frm1.txtAQL.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtAQL.value = arrRet(0)
		frm1.txtAQL.focus
	End If
	
	Set gActiveElement = document.activeElement
	OpenAQL = true
End Function

'------------------------------------------  OpenAQL()  -------------------------------------------------
'	Name : OpenAQL()
'	Description : AQL PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAQL1()
	OpenAQL = false

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "AQL팝업"					' 팝업 명칭 
	arrParam(1) = "B_MINOR"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtAQL.value)				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("Q0012", "''", "S") & ""				' Where Condition		
	arrParam(5) = "AQL"							' 조건필드의 라벨 명칭	
	arrField(0) = "MINOR_CD"						' Field명(0)
	arrField(1) = "MINOR_NM"						' Field명(0)
    arrHeader(0) = "코드"					' Header명(0)
    arrHeader(1) = "명"						' Header명(0)
    
   	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		If SetAQL(arrRet) = False Then	Exit Function
	End If	
	OpenAQL = true
End Function

'=======================================================================================================
'   Sub Name : ShowCondition()
'   Sub Desc : 
'=======================================================================================================
Function ShowCondition(Byval vSTDack)
	If vSTDack = "N" Then
		Q1.Style.display = ""	
	Else		' "Y", ""
		Q1.Style.display = "none"	
	End If
End Function
'/* Issue: 표준편차 선택에 따른 표준편차 대용 Show 관련 수정 - END */ 

'=========================================== 2.3.1 ResultClick() ========================================== 
'= Name : ResultClick()
'= Description : Return Array to Opener Window when OK button click = 
'======================================================================================================== 
Function ResultClick() 

	Dim strVal 

	
	If Not chkField(Document, "1") Then  '⊙: Check contents area
    	Exit Function
    End If
    	
	If frm1.rdoRigor.rdoRigor3.checked = true then
		frm1.txtRigor.value = "1"			
	Elseif frm1.rdoRigor.rdoRigor1.checked = true then
		frm1.txtRigor.value = "2"
	Elseif frm1.rdoRigor.rdoRigor2.checked = true then
		frm1.txtRigor.value = "3"
	Else
	End if		
	
	If frm1.rdoNormal.rdoNormal1.checked = true then
		frm1.txtDefectMode.value = "I"
	ElseIf frm1.rdoNormal.rdoNormal2.checked = true then
		frm1.txtDefectMode.value = "II"
	ElseIf frm1.rdoNormal.rdoNormal3.checked = true then
		frm1.txtDefectMode.value = "III"
	ElseIf frm1.rdoNormal.rdoNormal4.checked = true then
		frm1.txtDefectMode.value = "IV"
	ElseIf frm1.rdoNormal.rdoNormal5.checked = true then
		frm1.txtDefectMode.value = "V"
	Else	
	End if		
	
	IF frm1.rdoSTDack.rdoSTDack1.checked = true then
		frm1.txtSTDack.Value = "N"			'표준편차를 알고 있다.
	Else
		IF frm1.rdoSTDRep.rdoSTDRep1.checked = true then
			frm1.txtSTDack.Value = "S" 			'표준편차를 알지 못한다 -> s사용.
		ElseIf frm1.rdoSTDRep.rdoSTDRep2.checked = true then
			frm1.txtSTDack.Value = "R" 			'표준편차를 알지 못한다 -> R사용.
		Else
			Call DisplayMsgBox("229919", "X", "X", "X") 		'선택사항을 체크하십시오 
			Exit Function	
		End If
	End If

	IF frm1.rdoInsCri.rdoInsCri1.checked = true then
		frm1.txtInsCri.Value = 0			'하한규격 적용 
	ElseIF frm1.rdoInsCri.rdoInsCri2.checked = true then
		frm1.txtInsCri.Value = 1			'상한규격 적용 
	Else
		frm1.txtInsCri.Value = 2 			'양쪽규격 적용 
	End IF	
	
	If frm1.rdoForm.rdoForm1.checked = true then
		If frm1.txtInsCri.value = 2 then	'양쪽규격 적용이면 
			Call DisplayMsgBox("224306", "X", "X", "X") 		'형식-1은 한쪽규격일 때만 적용할 수 있습니다. 확인하십시오. 
			Exit Function
		End If
	End If
	
	IF frm1.rdoSTDack.rdoSTDack1.checked = true and frm1.txtSD.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri1.checked = true and frm1.txtLowerBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X")  		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri2.checked = true and frm1.txtUpperBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri3.checked = true and frm1.txtUpperBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF
	IF frm1.rdoInsCri.rdoInsCri3.checked = true and frm1.txtLowerBound.Text = "" then
		Call DisplayMsgBox("229918", "X", "X", "X") 		'선택사항과 입력사항이 일치하지 않습니다 
		Exit Function	
	End IF
	
	
	IF frm1.txtRigor.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'선택사항을 체크하십시오 
		Exit Function	
	ElseIF  frm1.txtDefectMode.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'선택사항을 체크하십시오 
		Exit Function	
	ElseIF  frm1.txtSTDack.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'선택사항을 체크하십시오 
		Exit Function	
	ElseIF  frm1.txtInsCri.Value = "" then
		Call DisplayMsgBox("229919", "X", "X", "X") 		'선택사항을 체크하십시오 
		Exit Function	
	End IF
	
	Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
	
	Call LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY_ID & "?txtLotSize=" & frm1.txtLotSize.Text   '☜: '☆: 조회 조건 데이타 
	strVal = strVal & "&txtAQL=" & frm1.txtAQL.Value
	strVal = strVal & "&txtSD=" & frm1.txtSD.Text   
	
	strVal = strVal & "&txtRigor=" & Trim(frm1.txtRigor.Value)	
	strVal = strVal & "&txtDefectMode=" & Trim(frm1.txtDefectMode.Value)	
	strVal = strVal & "&txtSTDack=" & Trim(frm1.txtSTDack.Value)	
	strVal = strVal & "&txtInsCri=" & Trim(frm1.txtInsCri.Value)	
	
	strVal = strVal & "&txtLowerBound=" & frm1.txtLowerBound.Text 
	strVal = strVal & "&txtUpperBound=" & frm1.txtUpperBound.Text 
	
	Call RunMyBizASP(MyBizASP, strVal) 
	
End Function 
 
'========================================= 2.3.2 ShowGraphClick() ======================================== 
'= Name : ShowGraphClick() = 
'= Description : Return Array to Opener Window for Cancel button click = 
'======================================================================================================== 
Function ShowGraphClick() 

	Dim strVal	
	
	IF frm1.txtSampleSize.Text = "" then
		Call DisplayMsgBox("229920", "X", "X", "X") 		'결과항목이 없습니다 
		Exit Function	
	End IF
			
	strVal = PGM_JUMP_ID3 & "?txtSampleSize=" & frm1.txtSampleSize.Text
	If frm1.rdoSTDack.rdoSTDack1.checked = true then
		strVal = strVal & "&txtSD=" & frm1.txtSD.Text
	Else
		Call DisplayMsgBox("229922", "X", "X", "X") 		'현재는 지원되지 않습니다 
		Exit Function	
		'strVal = strVal & "&txtSD=" & ""
	End If

	strVal = strVal & "&txtInsCri=" & Trim(frm1.txtInsCri.Value) 		'CheckBox값을 넘겨줍니다.

	strVal = strVal & "&txtLowerBound=" & frm1.txtLowerBound.Text
	strVal = strVal & "&txtUpperBound=" & frm1.txtUpperBound.Text
	
	'/* Issue: 검사방식 적용으로 Return - START */
	strVal = strVal & "&txtPageCode=" & "AV"
	'/* Issue: 검사방식 적용으로 Return - END */
	Location.href = strVal
End Function 

'/* Issue: 표준편차에 초기값이 안들어가는 것과 관련 수정 - START */
'=======================================================================================================
'   Sub Name : LockSD()
'   Sub Desc : 
'=======================================================================================================
Sub LockSD(Byval vSTDack)
	If vSTDack = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtSD, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtSD, "Q")
	End If
End Sub
'/* Issue: 표준편차에 초기값이 안들어가는 것과 관련 수정 - END */

'/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - START */
'=======================================================================================================
'   Sub Name : LockInspSpec()
'   Sub Desc : 
'=======================================================================================================
Sub LockInspSpec(Byval vSTrdoInsCri)
	SELECT CASE vSTrdoInsCri
		CASE "A"
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "N")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "N")
		CASE "U"
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "N")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "Q")
		CASE "L"
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "Q")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "N")
		CASE ELSE
			Call ggoOper.SetReqAttr(frm1.txtUpperBound, "Q")
			Call ggoOper.SetReqAttr(frm1.txtLowerBound, "Q")
	END SELECT
End Sub
'/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - END */

'========================================== 3.1.1 Form_Load() ====================================== 
' Name : Form_Load() 
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================  
Sub Form_Load() 
	Call LoadInfTB19029                                                     	'⊙: Load table , B_numeric_format
	'/* Issue: 수치 포맷 관련 수정 - START */
	Call AppendNumberPlace("6", "5", "3")
	Call AppendNumberPlace("7", "11", "4")
	'/* Issue: 수치 포맷 관련 수정 - END */
	Call ggoOper.LockField(Document, "N")                                   	'⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	'/* Issue: 표준편차에 초기값이 안들어가는 것과 관련 수정 - START */
	Call LockSD("")
	'/* Issue: 표준편차에 초기값이 안들어가는 것과 관련 수정 - END */
	'/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - START */
	Call LockInspSpec("")
	'/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - END */
	'/* Issue: 표준편차 선택에 따른 표준편차 대용 Show 관련 수정 - START */ 
	Call ShowCondition("")
	'/* Issue: 표준편차 선택에 따른 표준편차 대용 Show 관련 수정 - END */ 
	Call InitVariables																'⊙: Initializes local global variables          
	Call SetDefaultVal
	Call SetToolbar("10000000000111")
End Sub 
 
'========================================================================================== 
' Event Name : Form_QueryUnload 
' Event Desc : 
'========================================================================================== 
Sub Form_QueryUnload(Cancel , UnloadMode ) 

End Sub 

'/* Issue: 표준편차에 초기값이 안들어가는 것과 관련 수정 - START */
'========================================================================================== 
' Event Name : rdoSTDack1_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoSTDack1_onclick()
	Call LockSD("Y")
	Call ShowCondition("Y")
End Sub

'========================================================================================== 
' Event Name : rdoSTDack2_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoSTDack2_onclick()
	Call LockSD("N")
	Call ShowCondition("N")
End Sub
'/* Issue: 표준편차에 초기값이 안들어가는 것과 관련 수정 - END */

'/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - START */
'========================================================================================== 
' Event Name : rdoInsCri1_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoInsCri1_onclick()
	Call LockInspSpec("L")
End Sub

'========================================================================================== 
' Event Name : rdoInsCri2_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoInsCri2_onclick()
	Call LockInspSpec("U")
End Sub

'========================================================================================== 
' Event Name : rdoInsCri3_onclick 
' Event Desc : 
'========================================================================================== 
Sub rdoInsCri3_onclick()
	Call LockInspSpec("A")
End Sub
'/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - END */

'======================================================================================== 
' Function Name : FncQuery 
' Function Desc : This function is related to Query Button of Main ToolBar 
'======================================================================================== 
Function FncQuery() 
	FncQuery = False
End Function 
 
'======================================================================================== 
' Function Name : FncNew 
' Function Desc : This function is related to New Button of Main ToolBar 
'======================================================================================== 
Function FncNew() 
	FncNew = False
End Function 
 
'======================================================================================== 
' Function Name : Fnc 
' Function Desc : This function is related to Delete Button of Main ToolBar 
'======================================================================================== 
Function FncDelete() 
	FncDelete = False
End Function 
 
'======================================================================================== 
' Function Name : FncSave 
' Function Desc : This function is related to Delete Button of Main ToolBar 
'======================================================================================== 
Function FncSave() 
	FncSave = False
End Function 	
 
'======================================================================================== 
' Function Name : FncCopy 
' Function Desc : This function is related to Copy Button of Main ToolBar 
'======================================================================================== 
Function FncCopy() 
	FncCopy = False
End Function 
 
'======================================================================================== 
' Function Name : FncCancel 
' Function Desc : This function is related to Cancel Button of Main ToolBar 
'======================================================================================== 
Function FncCancel() 
	FncCancel = False
End Function 
 
'======================================================================================== 
' Function Name : FncInsertRow 
' Function Desc : This function is related to InsertRow Button of Main ToolBar 
'======================================================================================== 
Function FncInsertRow() 
	FncInsertRow = False
End Function 
 
'======================================================================================== 
' Function Name : FncDeleteRow 
' Function Desc : This function is related to DeleteRow Button of Main ToolBar 
'======================================================================================== 
Function FncDeleteRow() 
	FncDeleteRow = False
End Function 

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call Parent.FncPrint()
	FncPrint = True    
End Function
 
'======================================================================================== 
' Function Name : FncPrev 
' Function Desc : This function is related to Previous Button 
'======================================================================================== 
Function FncPrev() 
	FncPrev = False
End Function 
 
'======================================================================================== 
' Function Name : FncNext 
' Function Desc : This function is related to Next Button 
'======================================================================================== 
Function FncNext() 
	FncNext = False
End Function 

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
	Call parent.FncExport(Parent.C_SINGLE)					'☜: 화면 유형 
	FncExcel = True
End Function

Function FncExit() 
	FncExit = True
End Function 
 
</SCRIPT> 
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD> 
<BODY SCROLL=NO TABINDEX="-1"> 
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>계량 조정형 샘플링 검사</FONT></TD> 
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
					<TD VALIGN="top"  WIDTH="100%">
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
						<LEGEND>선택사항</LEGEND>
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>엄격도</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRigor" TAG="1X" ID="rdoRigor1"><LABEL FOR="rdoRigor1">보통검사</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRigor" TAG="1X" ID="rdoRigor2"><LABEL FOR="rdoRigor2">까다로운검사</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRigor" TAG="1X" ID="rdoRigor3"><LABEL FOR="rdoRigor3">수월한검사</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>검사수준</TD>
									<TD WIDTH="10%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="1X" ID="rdoNormal1"><LABEL FOR="rdoNormal1">Ⅰ</LABEL></TD>
									<TD WIDTH="10%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="1X" ID="rdoNormal2"><LABEL FOR="rdoNormal2">Ⅱ</LABEL></TD>
									<TD WIDTH="10%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="1X" ID="rdoNormal3"><LABEL FOR="rdoNormal3">Ⅲ</LABEL></TD>
									<TD WIDTH="10%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="1X" ID="rdoNormal4"><LABEL FOR="rdoNormal4">IV</LABEL></TD>
									<TD WIDTH="10%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoNormal" TAG="1X" ID="rdoNormal5"><LABEL FOR="rdoNormal5">V</LABEL></TD>
								</TR>				
								<!-- '/* Issue: 표준편차 선택에 따른 표준편차 대용 Show 관련 수정 - START */ -->
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>표준편차 인지 여부</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSTDack" TAG="1X" ID="rdoSTDack1"><LABEL FOR="rdoSTDack1">표준편차를 알고 있슴</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSTDack" TAG="1X" ID="rdoSTDack2"><LABEL FOR="rdoSTDack2">표준편차를 알지 못함</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<!-- '/* Issue: 표준편차 선택에 따른 표준편차 대용 Show 관련 수정 - END */ -->
								<TR ID=Q1>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>표준편차 대체항목</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSTDRep" TAG="1X" ID="rdoSTDRep1"><LABEL FOR="rdoSTDRep1">샘플 표준편차(S)</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSTDRep" TAG="1X" ID="rdoSTDRep2"><LABEL FOR="rdoSTDRep2">샘플 범위(R)</LABEL></TD>									
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>형식</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoForm" TAG="1X" ID="rdoForm1"><LABEL FOR="rdoForm1">형식-1</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoForm" TAG="1X" ID="rdoForm2"><LABEL FOR="rdoForm2">형식-2</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5   HEIGHT=15 NOWRAP>검사규격 선택</TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInsCri" TAG="1X" ID="rdoInsCri1"><LABEL FOR="rdoInsCri1">한쪽(하한)규격</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInsCri" TAG="1X" ID="rdoInsCri2"><LABEL FOR="rdoInsCri2">한쪽(상한)규격</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInsCri" TAG="1X" ID="rdoInsCri3"><LABEL FOR="rdoInsCri3">양쪽규격</LABEL></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
									<TD WIDTH="20%" BGCOLOR="#F7F7F7" ALIGN="LEFT" CELLPADDING=5 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;"> 
						<LEGEND>입력사항</LEGEND> 
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%> 
								<TR>
									<TD CLASS="TD5" NOWRAP>로트크기</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtLotSize_txtLotSize.js'></script>
									</TD>								
									<TD CLASS="TD5" NOWRAP>AQL</TD>
									
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtAQL_txtAQL.js'></script>
										<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAQL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenAQL()">
									</TD>									
									
								</TR>
								<!-- '/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - START */ -->
								<TR> 
									<TD CLASS="TD5" NOWRAP>상한규격</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtUpperBound_txtUpperBound.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP>하한규격</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtLowerBound_txtLowerBound.js'></script>
									</TD>
								</TR> 
								<!-- '/* Issue: 검사규격 선택에 따른 해당 필드 Enable/Disable - END */ -->
								<!-- '/* Issue: 수치 포맷 관련 수정 - START */ -->
								<TR> 
									<TD CLASS="TD5" NOWRAP>표준편차</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtSD_txtSD.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD> 
									<TD CLASS="TD6" NOWRAP></TD> 
								</TR> 
								<!-- '/* Issue: 수치 포맷 관련 수정 - END */ -->
							</TABLE> 
						</FIELDSET> 		
						<FIELDSET STYLE="margin-left:10px; margin-right:10px;"> 
						<LEGEND>결과</LEGEND> 
							<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%> 
								<TR> 
									<TD CLASS="TD5" NOWRAP>샘플크기</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtSampleSize_txtSampleSize.js'></script>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD> 
									<TD CLASS="TD6" NOWRAP></TD> 
								</TR>
								<TR> 
									<TD CLASS="TD5" NOWRAP>합격판정계수</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtAcceptSize_txtAcceptSize.js'></script>
									</TD> 
									<TD CLASS="TD5" NOWRAP>최대허용불량률</TD> 
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/q1413ma6_txtMaxTolPercent_txtMaxTolPercent.js'></script>
									</TD> 
								</TR>  
							</TABLE> 
						</FIELDSET> 
					</TD>	
				</TR>
			</TABLE>
		</DIV>
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
<INPUT TYPE=HIDDEN NAME="txtRigor" tag="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtDefectMode" tag="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtSTDack" tag="14" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtInsCri" tag="14" tabindex=-1>
</FORM> 
<DIV ID="MousePT" NAME="MousePT"> 
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe> 
</DIV> 
</BODY> 
</HTML>

