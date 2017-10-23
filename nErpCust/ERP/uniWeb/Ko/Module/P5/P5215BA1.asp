<%@	LANGUAGE="VBSCRIPT"	%>
<!--
======================================================================================================
'**********************************************************************************************
'*	1. Module	Name					:	Facility
'*	2. Function	Name				:
'*	3. Program ID						:	Fb102ma1
'*	4. Program Name					:	설비점검계획수립
'*	5. Program Desc					:	설비점검계획수립
'*	6. Component List				:
'*	7. Modified	date(First)	:	2005/01/19
'*	8. Modified	date(Last)	:	2005/01/21
'*	9. Modifier	(First)			:	Lee	chang-je
'* 10. Modifier	(Last)			:	Lee	chang-je
'* 11. Comment							:
'* 12. Common	Coding Guide	:	this mark(☜)	means	that "Do not change"
'*														this mark(⊙)	Means	that "may	 change"
'*														this mark(☆)	Means	that "must change"
'* 13. History							:
'**********************************************************************************************
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include	file="../../inc/IncSvrCcm.inc" -->
<!-- #Include	file="../../inc/incSvrHTML.inc"	-->

<LINK	REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script	Language="VBScript">
Option Explicit

'========================================================================================================
'=											 4.2 Constant	variables
'========================================================================================================
Const	BIZ_PGM_ID	=	"P5215bb1.asp"
'========================================================================================================
'=											 4.3 Common	variables
'========================================================================================================
<!-- #Include	file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=											 4.4 User-defind Variables
<%'========================================================================================================%>
Dim	IsOpenPop

'========================================================================================================
'	Name : InitVariables()
'	Desc : Initialize	value
'========================================================================================================
Sub	InitVariables()
		lgIntFlgMode = Parent.OPMD_CMODE									 'Indicates	that current mode	is Create	mode
		lgBlnFlgChgValue = False										'Indicates that	no value changed
		lgIntGrpCount	=	0														'initializes Group View	Size

		lgStrPrevKey = ""														'initializes Previous	Key
		lgLngCurRows = 0														'initializes Deleted Rows	Count
End	Sub

'========================================================================================================
'	Name : SetDefaultVal()
'	Desc : Set default value
'========================================================================================================

Sub	SetDefaultVal()
	Dim	strYear
	Dim	strMonth
	Dim	strDay

	frm1.txtprov_dt.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat	,	Parent.gServerDateType ,strYear,strMonth,strDay)

	frm1.txtprov_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtprov_dt.Month	=	strMonth
	frm1.txtprov_dt.Day	=	strDay
End	Sub

'========================================================================================================
'	Name : LoadInfTB19029()
'	Desc : Set System	Number format
'========================================================================================================
Sub	LoadInfTB19029()
	<!-- #Include	file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call	loadInfTB19029A("Q", "H",	"NOCOOKIE",	"BA")	%>
End	Sub

'========================================================================================================
'	Name : CookiePage()
'	Description	:	Item Popup에서 Return되는	값 setting
'========================================================================================================
<%'========================================================================================================%>

Function CookiePage(ByVal	flgs)
		Const	CookieSplit	=	4877

	Dim	strTemp

	If flgs	=	0	Then																			 '☜:	h6012ma1.asp 의	쿠기값을 받고	있음.
		strTemp	=	ReadCookie("PROV_DT")											 '				 절대수정금지	요망........!
		If strTemp = ""	then Exit	Function

				frm1.txtprov_dt.text = ReadCookie("PROV_DT")
		frm1.txthFacility_Cd.value = ReadCookie("PROV_TYPE")
		frm1.txthFacility_Nm.value = ReadCookie("PROV_TYPE_NM")


		MainQuery()
		WriteCookie	"PROV_DT"	,	""
			WriteCookie	"PROV_TYPE"	,	""
				WriteCookie	"TRANS_DT"	,	""

	ElseIf flgs	=	1	Then
				WriteCookie	"PROV_DT"	,	frm1.txtprov_dt.text
	End	IF
End	Function

'========================================================================================================
'	Function Name	:	MakeKeyStream
'	Function Desc	:	This method	set	focus	to pos of	err
'========================================================================================================
Sub	MakeKeyStream(pOpt)
	 With	frm1

			lgKeyStream	=	.txtprov_dt.Text & Parent.gColSep
			lgKeyStream	=	lgKeyStream	&	.txthFacility_Cd.Text	&	Parent.gColSep

	 End With
End	Sub

'========================================================================================================
'	Name : Form_Load
'	Desc : developer describe	this line	Called by	Window_OnLoad()	evnt
'========================================================================================================
Sub	Form_Load()

		Err.Clear																																				'☜: Clear err status
	Call LoadInfTB19029																															'☜: Load	table	,	B_numeric_format

	Call ggoOper.LockField(Document, "N")																		'⊙: Lock	 Suitable	 Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call InitVariables																										 '⊙:	Setup	the	Spread sheet
		Call ggoOper.FormatDate(frm1.txtprov_dt, Parent.gDateFormat, 1)

	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼	툴바 제어 
	Call CookiePage(0)																														 '☜:	Check	Cookie

End	Sub

'========================================================================================================
'	Name : Form_QueryUnload
'	Desc : developer describe	this line	Called by	Window_OnUnLoad()	evnt
'========================================================================================================
Sub	Form_QueryUnload(Cancel, UnloadMode)

End	Sub

'========================================================================================================
'	Name : FncDelete
'	Desc : developer describe	this line	Called by	MainDelete in	Common.vbs
'========================================================================================================
Function FncDelete()
		Dim	intRetCD

		FncDelete	=	False																														 '☜:	Processing is	NG
		Err.Clear																																		 '☜:	Clear	err	status

		FncDelete	=	True																														 '☜:	Processing is	OK
End	Function

'========================================================================================================
'	Name : FncCancel
'	Desc : developer describe	this line	Called by	MainCancel in	Common.vbs
'========================================================================================================
Function FncCancel()
	On Error Resume	Next																												'☜: Protect system	from crashing
End	Function

'========================================================================================================
'	Name : FncPrint
'	Desc : developer describe	this line	Called by	MainDeleteRow	in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()																											'☜: Protect system	from crashing
End	Function

'========================================================================================================
'	Name : FncExcel
'	Desc : developer describe	this line	Called by	MainExcel	in Common.vbs
'========================================================================================================
Function FncExcel()
	Call Parent.FncExport(Parent.C_SINGLE)
End	Function

'========================================================================================================
'	Name : FncFind
'	Desc : developer describe	this line	Called by	MainFind in	Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind(Parent.C_SINGLE, False)
End	Function

'========================================================================================================
'	Name : FncExit
'	Desc : developer describe	this line	Called by	MainExit in	Common.vbs
'========================================================================================================
Function FncExit()
	Dim	IntRetCD

	FncExit	=	False

	FncExit	=	True
End	Function

'========================================================================================================
'	Name : DbQuery
'	Desc : This	function is	called by	FncQuery
'========================================================================================================
Function DbQuery()
		Dim	strVal
		Err.Clear																																		 '☜:	Clear	err	status

		DbQuery	=	True																															 '☜:	Processing is	NG
End	Function
'========================================================================================================
'	Name : DbSave
'	Desc : This	function is	called by	FncSave
'========================================================================================================
Function DbSave()
	Dim	strVal
		Err.Clear																																		 '☜:	Clear	err	status

	DbSave = False																		 '☜:	Processing is	NG

		DbSave	=	True																															 '☜:	Processing is	NG
End	Function
'========================================================================================================
'	Name : DbDelete
'	Desc : This	function is	called by	FncDelete
'========================================================================================================
Function DbDelete()
	Dim	strVal
		Err.Clear																																		 '☜:	Clear	err	status

	DbDelete = False																											 '☜:	Processing is	NG

	DbDelete = True																															 '⊙:	Processing is	NG
End	Function
'========================================================================================================
'	Function Name	:	DbQueryOk
'	Function Desc	:	Called by	MB Area	when query operation is	successful
'========================================================================================================
Function DbQueryOk()


End	Function

'========================================================================================================
'	Function Name	:	DbSaveOk
'	Function Desc	:	Called by	MB Area	when save	operation	is successful
'========================================================================================================
Function DbSaveOk()

End	Function

'========================================================================================================
'	Function Name	:	DbDeleteOk
'	Function Desc	:	Called by	MB Area	when delete	operation	is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
End	Function

'========================================================================================================
'	Name : OpenFacility_Popup()
'	Desc : developer describe	this line
'========================================================================================================
Function OpenFacility_Popup(Byval	iWhere)
	Dim	arrRet
	Dim	arrParam(5), arrField(6),	arrHeader(6)

	If IsOpenPop = True	 Then
		Exit Function
	End	If

	IsOpenPop	=	True
	Select Case	iWhere
		Case "1"
			arrParam(0)	=	"설비코드 팝업"
			arrParam(1)	=	"Y_FACILITY"
			arrParam(2)	=	frm1.txthFacility_Cd.value
			arrParam(3)	=	""												'	Name Cindition
			arrParam(4)	=	""																		'	Where	Condition
			arrParam(5)	=	"설비코드"												'	TextBox	명칭 

			arrField(0)	=	"Facility_cd"											'	Field명(0)
			arrField(1)	=	"Facility_Nm"											'	Field명(1)

			arrHeader(0) = "설비코드"									'	Header명(0)
			arrHeader(1) = "설비코드명"									'	Header명(1)
	End	Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",	Array(arrParam,	arrField,	arrHeader),	_
	"dialogWidth=420px;	dialogHeight=450px;	center:	Yes; help: No; resizable:	No;	status:	No;")

	IsOpenPop	=	False
	If arrRet(0) = ""	Then
			 Frm1.txthFacility_Cd.focus
		 Exit	Function
	Else
		 Call	SetCondArea(arrRet,iWhere)
	End	If

End	Function

'======================================================================================================
'	Name : SetCondArea()
'	Description	:	Item Popup에서 Return되는	값 setting
'=======================================================================================================
Sub	SetCondArea(Byval	arrRet,	Byval	iWhere)
	With Frm1
		Select Case	iWhere
			Case "1"
					.txthFacility_Cd.value = arrRet(0)
					.txthFacility_Nm.value = arrRet(1)
		End	Select
	End	With
End	Sub


'========================================================================================================
'		Event	Name : txthFacility_Cd_Onchange()						 '<==코드만	입력해도 앤터키,탭키를 치면	코드명을 불러준다 
'		Event	Desc :
'========================================================================================================
Function txthFacility_Cd_Onchange()
	Dim	iDx
	Dim	IntRetCd
	IF frm1.txthFacility_Cd.value	=	"" THEN
		frm1.txthFacility_Nm.value = ""
	ELSE
		IntRetCd = CommonQueryRs(" facility_nm "," y_facility	","	 facility_cd =	"	&	FilterVar(frm1.txthFacility_Cd.value , "''", "S")	&	"" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	'unicode
		If IntRetCd	=	false	then
'			Call DisplayMsgBox("970000","X","설비코드","X")
			frm1.txthFacility_Nm.value = ""
			frm1.txthFacility_Cd.focus
		ELSE
			frm1.txthFacility_Nm.value = Trim(Replace(lgF0,Chr(11),""))
		END	IF
	END	IF
End	Function

'======================================================================================================
'	Function Name	:	ExeReflect
'	Function Desc	:
'=======================================================================================================
Function ExeReflect()
	Dim	strVal
	Dim	strprov_dt,	stracct_dt
	Dim	IntRetCD
	Dim strWhere
	ExeReflect = False																													'⊙: Processing	is NG

	On Error Resume	Next																									 '☜:	Protect	system from	crashing

	If Not chkField(Document,	"1") Then
		Call	BtnDisabled(0)
		Exit	Function																		 '☜:	This function	check	required field
	End	If

	IF frm1.txthFacility_Cd.value	=	"" THEN
		frm1.txthFacility_Nm.value = ""
	ELSE
		IntRetCd = CommonQueryRs(" facility_nm "," y_facility	","	 facility_cd =	"	&	FilterVar(frm1.txthFacility_Cd.value , "''", "S")	&	"" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	'unicode
		If IntRetCd	=	false	then
			Call DisplayMsgBox("970000","X","설비코드","X")
			frm1.txthFacility_Nm.value = ""
			frm1.txthFacility_Cd.focus
			Exit Function
		ELSE
			frm1.txthFacility_Nm.value = Trim(Replace(lgF0,Chr(11),""))
		END	IF
	END	IF

	if frm1.txthFacility_Cd.value = "" then
		strWhere = " GUBUN_CD ='10' AND PLAN_GUBUN='1' AND INSP_FLAG = 'N' AND WORK_DT <=  " & FilterVar(frm1.txtprov_dt.text,"''","S")
	else
		strWhere = " GUBUN_CD ='10' AND PLAN_GUBUN='1' AND INSP_FLAG = 'N' AND WORK_DT <=  " & FilterVar(frm1.txtprov_dt.text,"''","S") & " AND FAC_CAST_CD = " & FilterVar(frm1.txthFacility_Cd.value,"''","S")
	end if
	IF CommonQueryRs( "Count(*) mycount " , "Y_FAC_CAST_PLAN" , strWhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
		IF ABS(Cdbl(Trim(lgF0))) > 0 THEN
			IntRetCD = DisplayMsgBox("800358",Parent.VB_YES_NO,"X","X")
		
			If IntRetCD	=	vbNo Then
				Call BtnDisabled(0)
				Exit Function
			End	If		
		ELSE
			IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
		
			If IntRetCD	=	vbNo Then
				Call BtnDisabled(0)
				Exit Function
			End	If
		END IF 
	End IF	



	If	 LayerShowHide(1)	=	False	Then
			 Call	BtnDisabled(0)
			 Exit	Function
	End	If
	Call BtnDisabled(1)

'	strprov_dt = UniConvDateToYYYYMMDD(frm1.txtprov_dt.text, Parent.gDateFormat, Parent.gComDateType)
	strprov_dt = frm1.txtprov_dt.Year & Right("0" & frm1.txtprov_dt.Month,2) & Right("0" & frm1.txtprov_dt.Day,2)

	strVal = BIZ_PGM_ID	&	"?txtMode="	&	Parent.UID_M0006
	strVal = strVal	&	"&txtWork_dt=" & strprov_dt
	strVal = strVal	&	"&txtFacility_Cd=" & frm1.txthFacility_Cd.value


	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스	ASP	를 가동 

	ExeReflect = True																														'⊙: Processing	is NG
	Call BtnDisabled(0)
End	Function

'======================================================================================================
'	Function Name	:	ExeReflectOk
'	Function Desc	:	ExeReflect가 성공적일	경우 MyBizASP	에서 호출되는	Function,	현재 FncSave에 있는것을	옮김 
'=======================================================================================================
Function ExeReflectOk()										'☆: 저장	성공후 실행	로직 
	Dim	IntRetCD

	IntRetCD =DisplayMsgBox("990000","X","X","X")

End	Function

Function ExeReflectNo()										'☆: 실행된	자료가 없습니다 
End	Function

Sub	txtprov_dt_DblClick(Button)
	If Button	=	1	Then
		Call SetFocusToDocument("M")
		frm1.txtprov_dt.Action = 7
		frm1.txtprov_dt.focus
	End	If
End	Sub




</SCRIPT>
<!-- #Include	file="../../inc/UNI2KCM.inc" -->
</HEAD>


<BODY	TABINDEX="-1"	SCROLL="NO">
<FORM	NAME=frm1	TARGET="MyBizASP"	METHOD="POST">
<TABLE CLASS="BatchTB1"	CELLSPACING=0	CELLPADDING=0>
	<TR>
		<TD	<%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR	HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD	>
						<TABLE ID="MyTab"	CELLSPACING=0	CELLPADDING=0>
							<TR>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif"><img	src="../../../CShared/image/table/seltab_up_left.gif"	width="9"	height="23"></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="center"	CLASS="CLSMTAB"><font	color=white>설비점검계획수립</font></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img	src="../../../CShared/image/table/seltab_up_right.gif" width="10"	height="23"></td>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR	HEIGHT=*>
		<TD	CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD	HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD	CLASS="TD5"	NOWRAP>기간</TD>
								<TD	CLASS="TD6"	NOWRAP>
									<table cellpadding=0 cellspacing=0>
										<tr>
											<td>설비별 최종점검일	~	&nbsp; <td>
											<td>
												<script language =javascript src='./js/p5215ba1_txtprov_dt_txtprov_dt.js'></script>
											</td>
										</tr>
									</table>
								</TD>

							</TR>
							<TR>
								<TD	CLASS="TD5"	NOWRAP>설비코드</TD>
								<TD	CLASS="TD6"	NOWRAP><INPUT	NAME="txthFacility_Cd" MAXLENGTH="18"	 SIZE="18" ALT ="설비코드" TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname"	ALIGN=top	TYPE="BUTTON"	ONCLICK="VBScript: OpenFacility_Popup('1')">
																			 <INPUT	NAME="txthFacility_Nm" SIZE="40"	ALT	="설비명"	tag="14"></TD>

													</TR>
							</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD	<%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR	HEIGHT=20>
		<TD>
				<TABLE <%=LR_SPACE_TYPE_30%>>
						<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD>
								 <BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()"	Flag=1>실행</BUTTON>
								</TD>
					<TD	WIDTH=*	ALIGN="right">&nbsp;</TD>
					<TD	WIDTH=10>&nbsp;</TD>
						</TR>
				</TABLE>
		</TD>
	</TR>
	<TR>
		<TD	HEIGHT=20><IFRAME	NAME="MyBizASP"	SRC	=	"../../blank.htm"	WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no	noresize framespacing=0  TABINDEX="-1"></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"	tag="24"><INPUT	TYPE=HIDDEN	NAME="txtFlgMode"	tag="24">
</FORM>
<DIV ID="MousePT"	NAME="MousePT">
<iframe	name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	width=220	height=41	src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


