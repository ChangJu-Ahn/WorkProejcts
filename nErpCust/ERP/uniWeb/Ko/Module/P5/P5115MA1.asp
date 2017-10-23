<%@	LANGUAGE="VBSCRIPT"	%>
<!--
======================================================================================================
*	 1.	Module Name					 : Production
*	 2.	Function Name				 :
*	 3.	Program	ID					 : b1b02ma1
*	 4.	Program	Name				 : Item	Image	Management
*	 5.	Program	Desc				 :
*	 6.	Component	List			 :
*	 7.	Modified date(First) : 2001/06/29
*	 8.	Modified date(Last)	 : 2003/01/15
*	 9.	Modifier (First)		 : Im	Hyun Soo
*	10.	Modifier (Last)			 : Hong	Chang	Ho
*	11.	Comment							 :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include	file="../../Inc/IncSvrCcm.inc" -->
<!-- #Include	file="../../Inc/incSvrHTML.inc"	-->

<LINK	REL="stylesheet" TYPE="Text/css" HREF="../../Inc/SheetStyle.css">

<SCRIPT	LANGUAGE="VBScript"		SRC="../../Inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../Inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../Inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../Inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../Inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript"		SRC="../../Inc/Cookie.vbs"></SCRIPT>
<SCRIPT	LANGUAGE="VBScript">

Option Explicit

Const	BIZ_PGM_ID	=	"P5115mb1.asp"
Const	BIZ_PGM_ID1	=	"P5115mb2.asp"
Const	IMG_LOAD_PATH	=	"../../ComAsp/imgTemp.asp?src="

Const	DIR_INIT_FILE	=	"../../../CShared/image/unierp20logo.gif"

<!-- #Include	file="../../inc/lgVariables.inc" -->

Dim	IsOpenPop

'========================================================================================================
'	Name : InitVariables()
'	Desc : Initialize	value
'========================================================================================================
<%'========================================================================================================%>

Sub	InitVariables()
	lgIntFlgMode		=	parent.OPMD_CMODE										'⊙: Indicates that	current	mode is	Create mode
	lgBlnFlgChgValue	=	False										'⊙: Indicates that	no value changed
	lgIntGrpCount		=	0										'⊙: Initializes Group View	Size
	lgStrPrevKey		=	""																			'⊙: initializes Previous	Key
	lgSortKey			=	1								'⊙: initializes sort	direction
End	Sub

'========================================================================================================
'	Name : CookiePage()
'	Description	:	Item Popup에서 Return되는	값 setting
'========================================================================================================
Function CookiePage(ByVal	flgs)
	On Error Resume	Next

	Const	CookieSplit	=	4877
	Dim	strTemp

	If flgs	=	1	Then
		WriteCookie	CookieSplit	,	frm1.txtFacility_Cd.Value
	ElseIf flgs	=	0	Then
		frm1.txtFacility_Cd.Value	=	ReadCookie("txtFacility_Cd")
		frm1.txtFacility_Nm.value	=	ReadCookie("txtFacility_Nm")

		WriteCookie	"txtFacility_Cd",	""
		WriteCookie	"txtFacility_Nm",	""

		If frm1.txtFacility_Cd.value <>	"" Then
			Call MainQuery()
		End	If

	End	If

End	Function

'========================================================================================================
'	Function Name	:	MakeKeyStream
'	Function Desc	:	This method	set	focus	to pos of	err
'========================================================================================================
Sub	MakeKeyStream(pOpt)
	If pOpt	=	"Q"	Then
		lgKeyStream	=	Frm1.txtFacility_Cd.Value	&	parent.gColSep
	Else
		lgKeyStream	=	Frm1.txtFacility_Cd.Value	&	parent.gColSep
	End	If
End	Sub

Sub	Form_Load()

	Err.Clear								'☜: Clear err status

	Call ggoOper.FormatField(Document, "2",	ggStrIntegeralPart,	ggStrDeciPointPart,parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock	Field

	Call SetToolbar("1110100000000111")												'⊙: Set ToolBar
	frm1.txtFileExist.Value = "False"
	Call InitVariables

	frm1.txtFacility_Cd.focus
	Set	gActiveElement = document.ActiveElement

	Call CookiePage	(0)															'☜: Check Cookie

End	Sub

'========================================================================================================
'	Name : FncQuery
'	Desc : developer describe	this line	Called by	MainQuery	in Common.vbs
'========================================================================================================
Function FncQuery()
	Dim	IntRetCD

	FncQuery = False															 '☜:	Processing is	NG
	Err.Clear									'☜:	Clear	err	status

	If lgBlnFlgChgValue	=	True Then
	IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X",	"X")					 '☜:	Data is	changed.	Do you want	to display it?
	If IntRetCD	=	vbNo Then
		Exit Function
	End	If
	End	If

	Call ggoOper.ClearField(Document,	"2")										 '☜:	Clear	Contents	Field

	If Not chkField(Document,	"1") Then													 '☜:	This function	check	required field
		 Exit	Function
	End	If
	
	If ChkKeyField() =false then 
		exit function 
	End If

	Call InitVariables								'⊙:	Initializes	local	global variables
	Call MakeKeyStream("Q")

	If GetFacilityCd = False Then
		Exit Function
	End	If
	'Call	DbQuery																	'☜: Query db	data

	FncQuery = True								'☜:	Processing is	OK

End	Function

'========================================================================================================
'	Name : FncNew
'	Desc : developer describe	this line	Called by	MainNew	in Common.vbs
'========================================================================================================
Function FncNew()
	Dim	IntRetCD

	FncNew = False																 '☜:	Processing is	NG
	Err.Clear									'☜:	Clear	err	status

	If lgBlnFlgChgValue	=	True Then
		 IntRetCD	=	DisplayMsgBox("900015",	parent.VB_YES_NO,	"X", "X")					 '☜:	Data is	changed.	Do you want	to make	it new?
		 If	IntRetCD = vbNo	Then
			Exit Function
		 End If
	End	If

	Call ggoOper.ClearField(Document,	"A")										'☜:	Clear	Condition	Field
	Call ggoOper.LockField(Document	,	"N")										'☜:	Lock	Field

	document.all.ImgItemImage.src	=	DIR_INIT_FILE

	Call SetToolbar("11101000000001")
	Call InitVariables												'⊙: Initializes local global	variables

	frm1.txtFacility_Cd.focus
	Set	gActiveElement = document.ActiveElement

	FncNew = True																 '☜:	Processing is	OK
End	Function

'========================================================================================================
'	Name : FncDelete
'	Desc : developer describe	this line	Called by	MainDelete in	Common.vbs
'========================================================================================================
Function FncDelete()
	Dim	intRetCD

	FncDelete	=	False									'☜:	Processing is	NG
	Err.Clear									'☜:	Clear	err	status

	If lgIntFlgMode	<> parent.OPMD_UMODE Then								'☜: Please	do Display first.
		Call DisplayMsgBox("900002", "X",	"X", "X")
		Exit Function
	End	If

	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X",	"X")												'☜: Do	you	want to	delete?
	If IntRetCD	=	vbNo Then
		Exit Function
	End	If

		Call MakeKeyStream("D")

	If DbDelete	=	False	Then
		Exit Function
	End	If

	Set	gActiveElement = document.ActiveElement
	FncDelete	=	True									'☜:	Processing is	OK

End	Function

'========================================================================================================
'	Name : FncSave
'	Desc : developer describe	this line	Called by	MainSave in	Common.vbs
'========================================================================================================
Function FncSave()
	Dim	IntRetCD,	iStrFileType

	FncSave	=	False								'☜:	Processing is	NG

	Err.Clear									'☜:	Clear	err	status

	If lgBlnFlgChgValue	=	False	Then
		IntRetCD = DisplayMsgBox("900001", "X",	"X", "X")					'☜:There	is no	changed	data.
		Exit Function
	End	If

	If Not chkField(Document,	"1") Then															'⊙: Check contents	area
		 Exit	Function
	End	If

	If Not chkField(Document,	"2") Then											 '☜:	Check	contents area
		 Exit	Function
	End	If

	If frm1.txtFileExist.value = "True" then
		If Not ggoSaveFile.FileExists(frm1.txtPath.value)	=	0	Then
			Call DisplayMsgBox("115191", "X",	"X", "X")
			Exit Function
		End	If
	
		iStrFileType = Right(Trim(UCase(frm1.txtPath.value)),	3)
	
		If Not (iStrFileType = "BMP" Or	iStrFileType = "GIF" Or	iStrFileType = "JPG")	Then
			Call DisplayMsgBox("122904", "X",	"X", "X")
			Exit Function
		End	If
	End if


	Call MakeKeyStream("S")

	If DbSave	=	False	Then
		Exit Function
	End	If

	FncSave	=	True									'☜: Processing	is OK
End	Function

'========================================================================================================
'	Name : FncPrint
'	Desc : developer describe	this line	Called by	MainDeleteRow	in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()										'☜: Protect system	from crashing
End	Function

'========================================================================================================
'	Name : FncPrev
'	Desc : developer describe	this line	Called by	MainPrev in	Common.vbs
'========================================================================================================
Function FncPrev()

	Dim	strVal
	Dim	IntRetCD

	FncPrev	=	False								'☜:	Processing is	OK
	Err.Clear									'☜:	Clear	err	status

	If lgIntFlgMode	<> parent.OPMD_UMODE Then								'☜: Please	do Display first.
			Call DisplayMsgBox("900002", "X",	"X", "X")
			Exit Function
	End	If

	If lgBlnFlgChgValue	=	True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X",	"X")					 '☜:	Will you destory previous	data
		If IntRetCD	=	vbNo Then
			Exit Function
		End	If
	End	If

	Call MakeKeyStream("P")
	Call ggoOper.ClearField(Document,	"2")										 '⊙:	Clear	Contents Area

	Call InitVariables											'⊙:	Initializes	local	global variables

	LayerShowHide(1)

	strVal = BIZ_PGM_ID	&	"?txtMode="					 & parent.UID_M0001												'☜: Query
	strVal = strVal			&	"&txtKeyStream="		 & lgKeyStream										 '☜:	Query	Key
	strVal = strVal			&	"&txtPrevNext="			 & "P"															 '☜:	Direction

	Call RunMyBizASP(MyBizASP, strVal)												'☜:	 Run biz logic
	Set	gActiveElement = document.ActiveElement

	FncPrev	=	True								'☜:	Processing is	OK

End	Function

'========================================================================================================
'	Name : FncNext
'	Desc : developer describe	this line	Called by	MainNext in	Common.vbs
'========================================================================================================
Function FncNext()

	Dim	strVal
	Dim	IntRetCD

	FncNext	=	False								'☜:	Processing is	OK
	Err.Clear									'☜:	Clear	err	status

	If lgIntFlgMode	<> parent.OPMD_UMODE Then
			Call DisplayMsgBox("900002", "X",	"X", "X")
			Exit Function
	End	If

	If lgBlnFlgChgValue	=	True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X",	"X")					 '☜:	Will you destory previous	data
		If IntRetCD	=	vbNo Then
			Exit Function
		End	If
	End	If

	Call MakeKeyStream("N")
	Call ggoOper.ClearField(Document,	"2")										 '⊙:	Clear	Contents Area

	Call InitVariables											'⊙:	Initializes	local	global variables

	LayerShowHide(1)

	strVal = BIZ_PGM_ID	&	"?txtMode="					 & parent.UID_M0001												'☜: Query
	strVal = strVal			&	"&txtKeyStream="		 & lgKeyStream										 '☜:	Query	Key
	strVal = strVal			&	"&txtPrevNext="			 & "N"															 '☜:	Direction

	Call RunMyBizASP(MyBizASP, strVal)												'☜:	 Run biz logic
	Set	gActiveElement = document.ActiveElement

	FncNext	=	True								'☜:	Processing is	OK

End	Function

'========================================================================================================
'	Name : FncExcel
'	Desc : developer describe	this line	Called by	MainExcel	in Common.vbs
'========================================================================================================
Function FncExcel()
	Call Parent.FncExport(parent.C_SINGLE)
End	Function

'========================================================================================================
'	Name : FncFind
'	Desc : developer describe	this line	Called by	MainFind in	Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind(parent.C_SINGLE, True)
End	Function

'========================================================================================================
'	Name : FncExit
'	Desc : developer describe	this line	Called by	MainExit in	Common.vbs
'========================================================================================================
Function FncExit()
	Dim	IntRetCD

	FncExit	=	False
	If lgBlnFlgChgValue	=	True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X",	"X")			'⊙: Data	is changed.	 Do	you	want to	exit?
		If IntRetCD	=	vbNo Then
			Exit Function
		End	If
	End	If

	FncExit	=	True
End	Function

'========================================================================================================
'	Name : DbQuery
'	Desc : This	function is	called by	FncQuery
'========================================================================================================
Function DbQuery(KeyItemVal)
	Dim	strVal
	Err.Clear									'☜:	Clear	err	status

	On Error Resume	Next

	DbQuery	=	False								'☜:	Processing is	NG

'	LayerShowHide(1)

	'------	Developer	Coding part	(Start)	 --------------------------------------------------------------
	If frm1.txtPrevNext.value	=	"" Then
		If CommonQueryRs(" FAC_CAST_CD "," Y_FAC_CAST_IMAGE	","	FAC_CAST_CD	=	"	&	FilterVar(KeyItemVal,	"''",	"S"),	lgF0,	lgF1,	lgF2,	lgF3,	lgF4,	lgF5,	lgF6)	=	False	Then
			Call DisplayMsgBox("Y52900", "X",	"X", "X")
			document.all.ImgItemImage.src= DIR_INIT_FILE
			frm1.txtPath.focus
			Set	gActiveElement = document.ActiveElement
			Exit Function
		End	If
	End	If

	strVal = "../../ComAsp/CPictRead.asp"	&	"?txtKeyValue="	&	KeyItemVal			'☜: query key
	strVal = strVal			&	"&txtDKeyValue=" & "default"														'☜: default value
	strVal = strVal			&	"&txtTable="		 & "Y_FAC_CAST_IMAGE"												'☜: Table Name
	strVal = strVal			&	"&txtField="		 & "PIC_IMAGE"												'☜: Field
	strVal = strVal			&	"&txtKey="			 & "FAC_CAST_CD"												'☜: Key

	document.all.ImgItemImage.src	=	ValueEscape(strVal)
	
	if strVal <> "" then
		frm1.txtFileExist.Value = "False"
	end if
	lgIntFlgMode = parent.OPMD_UMODE

	Call SetToolbar("11111000110001")

	DbQuery	=	True								'☜:	Processing is	NG

End	Function

'========================================================================================================
'	Name : DbSave
'	Desc : This	function is	called by	FncSave
'========================================================================================================
Function DbSave()
	Dim	strVal
	Err.Clear									'☜:	Clear	err	status

	On Error Resume	Next

	DbSave = False									'☜:	Processing is	NG

	LayerShowHide(1)

	'------	Developer	Coding part	(Start)	 --------------------------------------------------------------
	lgIntFlgMode = parent.OPMD_UMODE
	With Frm1
		.txtMode.value				=	parent.UID_M0002								'☜: Delete
		.txtFlgMode.value			=	lgIntFlgMode
		.txtKeyStream.Value		=	lgKeyStream										'☜:	Save Key
	End	With

	Call ExecMyBizASP(frm1,	BIZ_PGM_ID1)

	DbSave	=	True								'☜:	Processing is	NG
End	Function

'========================================================================================================
'	Name : DbDelete
'	Desc : This	function is	called by	FncDelete
'========================================================================================================
Function DbDelete()
	Dim	strVal
	Err.Clear									'☜:	Clear	err	status

	DbDelete = False																		'☜:	Processing is	NG

	LayerShowHide(1)

	strVal = BIZ_PGM_ID	&	"?txtMode="					 & parent.UID_M0003												'☜: Query
	strVal = strVal			&	"&txtKeyStream="		 & lgKeyStream										 '☜:	Query	Key

	Call RunMyBizASP(MyBizASP, strVal)												'☜:	 Run biz logic
	Set	gActiveElement = document.ActiveElement

	DbDelete = True								'⊙:	Processing is	NG
End	Function

'========================================================================================================
'	Function Name	:	DbQueryOk
'	Function Desc	:	Called by	MB Area	when query operation is	successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode			=	parent.OPMD_UMODE

	Frm1.txtFacility_Cd.focus

	Call SetToolbar("11111000111001")

	Call ggoOper.LockField(Document, "Q")
	Set	gActiveElement = document.ActiveElement

End	Function

'========================================================================================================
'	Function Name	:	DbSaveOk
'	Function Desc	:	Called by	MB Area	when save	operation	is successful
'========================================================================================================
Function DbSaveOk()

	Call InitVariables
	frm1.txtPath.value = NULL

	Call MainQuery()
End	Function

'========================================================================================================
'	Function Name	:	DbDeleteOk
'	Function Desc	:	Called by	MB Area	when delete	operation	is successful
'========================================================================================================
Function DbDeleteOk()

	Call InitVariables()
	Call FncNew()
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
			arrParam(2)	=	frm1.txtFacility_Cd.value
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
		 Frm1.txtFacility_Cd.focus
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
				.txtFacility_Cd.value	=	arrRet(0)
				.txtFacility_Nm.value	=	arrRet(1)
		 		.txtFacility_Cd.focus

		End	Select
	End	With
End	Sub


'------------------------------------------	 GetFacilityCd()	--------------------------------------------------
'	Name : GetFacilityCd()
'	Description	:
'---------------------------------------------------------------------------------------------------------
Function GetFacilityCd()

	Dim	strVal

	On Error Resume	Next																		'☜: Clear err status

	Err.Clear

	LayerShowHide(1)

	strVal = BIZ_PGM_ID	&	"?txtMode="					 & parent.UID_M0001												'☜: Query
	strVal = strVal			&	"&txtKeyStream="		 & lgKeyStream										 '☜:	Query	Key
	strVal = strVal			&	"&txtPrevNext="			 & ""															 '☜:	Direction

	Call RunMyBizASP(MyBizASP, strVal)												'☜:	 Run biz logic
	Set	gActiveElement = document.ActiveElement

End	Function



Sub	txtPath_OnChange()
	Dim	iStrFileType

	lgBlnFlgChgValue = True

	If Not ggoSaveFile.FileExists(frm1.txtPath.value)	=	0	Then
		frm1.txtFileExist.Value = "False"
		Exit Sub
	End	If

	iStrFileType = Right(Trim(UCase(frm1.txtPath.value)),	3)

	If Not (iStrFileType = "BMP" Or	iStrFileType = "GIF" Or	iStrFileType = "JPG")	Then
		Call DisplayMsgBox("122904", "X",	"X", "X")
		Exit Sub
	End	If
	document.all.ImgItemImage.src= ValueEscape(IMG_LOAD_PATH & frm1.txtPath.value)

	frm1.txtFileExist.Value = "True"
End	sub


'==========================================  2.2.6 ChkKeyField()  =======================================
'	Name : ChkKeyField()
'	Description : 
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		

	If Trim(frm1.txtFacility_Cd.value) <> "" Then
		strWhere = " Facility_cd =  " & FilterVar(frm1.txtFacility_Cd.value, "''", "S") & "  "

		Call CommonQueryRs(" Facility_Nm "," Y_FACILITY ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X","설비코드","X")
			frm1.txtFacility_Cd.focus 
			frm1.txtFacility_nm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtFacility_nm.value = strDataNm(0)
	Else
		frm1.txtFacility_nm.value = ""
	End If

End Function

</SCRIPT>
<!-- #Include	file="../../inc/uni2kcm.inc" -->
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->

<BODY	TABINDEX="-1"	SCROLL="NO">
<FORM	NAME=frm1	TARGET="MyBizASP"	METHOD="POST"	ENCTYPE="MULTIPART/FORM-DATA">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD	<%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR	HEIGHT=23>
		<TD	WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD	CLASS="CLSMTABP">
						<TABLE ID="MyTab"	CELLSPACING=0	CELLPADDING=0>
							<TR>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif"><img	src="../../../CShared/image/table/seltab_up_left.gif"	width="9"	height="23"></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="center"	CLASS="CLSMTAB"><font	color=white>설비사진등록</font></td>
								<td	background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img	src="../../../CShared/image/table/seltab_up_right.gif" width="10"	height="23"></td>
								</TR>
						</TABLE>
					</TD>
					<TD	WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR	HEIGHT=*>
		<TD	WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD	<%=HEIGHT_TYPE_02%>	WIDTH=100%></TD>
				</TR>
				<TR>
					<TD	HEIGHT=20	WIDTH=100%>
						<FIELDSET	CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR>
									<TD	CLASS="TD5"	NOWRAP>설비코드</TD>
									<TD	CLASS="TD656"	NOWRAP><INPUT	ID=txtFacility_Cd	NAME="txtFacility_Cd"	ALT="설비코드" TYPE="Text" SiZE="18" MAXLENGTH="18"	tag="12XXXU"><IMG	SRC="../../../CShared/image/btnPopup.gif"	NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript:	OpenFacility_Popup('1')">
											<INPUT	ID=txtFacility_Nm	NAME="txtFacility_Nm"	ALT="설비코드명" TYPE="Text" SiZE="40" MAXLENGTH="40"	tag="14XXXU"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD	<%=HEIGHT_TYPE_03%>	WIDTH=100%></TD>
				</TR>
				<TR>
					<TD	WIDTH=100% HEIGHT=*	valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD	CLASS=TD5	NOWRAP>비고</TD>
								<TD	CLASS=TD656	NOWRAP><INPUT	NAME="txtMemo" ALT="비고"	TYPE="Text"	SiZE=40	MAXLENGTH=40	tag="21XXXX"></TD>
							</TR>
							<TR>
								<TD	CLASS=TD5	NOWRAP>경로</TD>
								<TD	CLASS=TD656	NOWRAP><INPUT	TYPE=FILE	NAME="txtPath" SIZE=40 MAXLENGTH=100 tag=21	ALT="경로"></TD>
							</TR>
							<TR>
								<TD	CLASS=TD5	NOWRAP>&nbsp;</TD>
								<TD	CLASS=TD656	NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD	CLASS=TD5	NOWRAP>사진</TD>
								<TD	CLASS=TD656	NOWRAP colspan=3>
												<IFRAME	NAME="ImgItemImage"	SRC="../../../CShared/image/unierp20logo.gif"	marginwidth=0	marginheight=0 WIDTH=100%	HEIGHT=365 FRAMEBORDER=1 FRAMESPACING=0></IFRAME>
								</TD>
							</TR>
							<TR>
								<TD	CLASS=TD5	NOWRAP>&nbsp;</TD>
								<TD	CLASS=TD656	NOWRAP>&nbsp;</TD>
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
	<TR>
		<TD	WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"	WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	tabindex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"		TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"	TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"	TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFileExist"	TAG="24">

</FORM>
<DIV ID="MousePT"	NAME="MousePT">
	<IFRAME	NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO	noresize framespacing=0	WIDTH=220	HEIGHT=41	SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
