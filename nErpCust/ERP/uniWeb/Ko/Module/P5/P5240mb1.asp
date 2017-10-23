<%@	LANGUAGE=VBSCript%>
<%Option Explicit		 %>
<!-- #Include	file="../../inc/IncSvrMain.asp"	-->
<!-- #Include	file="../../inc/IncSvrNumber.inc"	-->
<!-- #Include	file="../../inc/IncSvrDate.inc"	-->
<!-- #Include	file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include	file="../../inc/IncSvrDBAgentVariables.inc"	-->
<!-- #Include	file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q",	"P", "NOCOOKIE","MB")

'**********************************************************************************************
'*	1. Module	Name					:	Procurement
'*	2. Function	Name				:
'*	3. Program ID						:	m9111ma1
'*	4. Program Name					:
'*	5. Program Desc					:
'*	6. Comproxy	List				:	PM9G111(Maint)
'								PM9G112(확정)
'*	7. Modified	date(First)	:	2002/12/06
'*	8. Modified	date(Last)	:
'*	9. Modifier	(First)			:	Oh Chang Won
'* 10. Modifier	(Last)			:
'* 11. Comment							:
'* 12. Common	Coding Guide	:	this mark(☜)	means	that "Do not change"
'*														this mark(⊙)	Means	that "may	 change"
'*														this mark(☆)	Means	that "must change"
'* 13. History							:
'*
'*
'*
'*
'* 14. Business	Logic	of m9111ma1(재고이동요청)
'**********************************************************************************************
Dim	lgOpModeCRUD

Dim	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0									'☜	:	DBAgent	Parameter	선언 
Dim	rs1, rs2,	rs3, rs4,rs5
Dim	istrData
Dim	iStrPoNo
Dim	StrNextKey		'	다음 값 
Dim	lgStrPrevKey	'	이전 값 
Dim	iLngMaxRow		'	현재 그리드의	최대Row
Dim	iLngRow
Dim	GroupCount
Dim	lgCurrency
Dim	index,Count			'	저장 후	Return 해줄	값을 넣을때	쓴는 변수 
Dim	lgDataExist
Dim	lgPageNo
Dim	lgMaxCount
Dim	strFlag

	Const	C_SHEETMAXROWS_D	=	100

		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status

		Call HideStatusWnd																															 '☜:	Hide Processing	message
	'------	Developer	Coding part	(Start ) ------------------------------------------------------------------

	'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------

		lgOpModeCRUD	=	Request("txtMode")

																																	'☜: Read	Operation	Mode (CRUD)
		Select Case	lgOpModeCRUD
				Case CStr(UID_M0001)																												 '☜:	Query
						 Call	 SubBizQueryMulti()
				Case CStr(UID_M0002)																												 '☜:	Save,Update
						 Call	SubBizSaveMulti()
				Case CStr(UID_M0003)
						 Call	SubBizSaveMulti()
		End	Select

'============================================================================================================
'	Name : SubBizQuery
'	Desc : Query Data	from Db
'============================================================================================================
Sub	SubBizQuery()
		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status

End	Sub
'============================================================================================================
'	Name : SubBizSave
'	Desc : Save	Data
'============================================================================================================
Sub	SubBizSave()
		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status
End	Sub
'============================================================================================================
'	Name : SubBizDelete
'	Desc : Delete	DB data
'============================================================================================================
Sub	SubBizDelete()
		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status
End	Sub

'============================================================================================================
'	Name : SubBizQuery
'	Desc : Query Data	from Db
'============================================================================================================
Sub	SubBizQueryMulti()

		On Error Resume	Next

	iStrPoNo = Trim(Request("txtPoNo"))
	lgPageNo			 = UNICInt(Trim(Request("lgPageNo")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
		lgMaxCount		 = C_SHEETMAXROWS_D														'☜	:	한번에 가져올수	있는 데이타	건수 
	lgDataExist			=	"No"
	iLngMaxRow		 = CDbl(lgMaxCount)	*	CDbl(lgPageNo) + 1

	lgStrPrevKey = Request("lgStrPrevKey")


	Call FixUNISQLData()
	Call QueryData()

	'====================
	'Call	PO_DTL List
	'====================

	'-----------------------
	'Result	data display area
	'-----------------------
	Response.Write "<Script	Language=vbscript>"	&	vbCr
	Response.Write "With parent" & vbCr

	 Response.Write	"	If .frm1.vspdData.MaxRows	<	1	then"						&	vbCr
	 Response.Write	"	End	if"							&	vbCr


		Response.Write "	.ggoSpread.Source				=	.frm1.vspdData "			&	vbCr
		Response.Write "	.ggoSpread.SSShowData			"""	&	istrData	 & """"	&	vbCr
		Response.Write "	.lgPageNo	 = """ & lgPageNo		&	"""" & vbCr

		Response.Write " .DbQueryOk	"	&	vbCr
		Response.Write "End	With"		&	vbCr
		Response.Write "</Script>"		&	vbCr

End	Sub

'----------------------------------------------------------------------------------------------------------
'	Set	DB Agent arg
'----------------------------------------------------------------------------------------------------------
'	Query하기	전에	DB Agent 배열을	이용하여 Query문을 만드는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub	FixUNISQLData()

	Dim	strItemCd
	Dim	strYear, strMonth
	Dim	strFacility_Accnt
	Dim	strFacility_Cd


	Redim	UNISqlId(2)																											'☜: SQL ID	저장을 위한	영역확보 
	Redim	UNIValue(2,	3)



	UNISqlId(0)	=	"160902saa"
	UNISqlId(1)	=	"P5110P580"



	IF Request("txtItemCd")	=	"" Then
		 strItemCd = "|"
	ELSE
		 strItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	END	IF
	IF Request("txtYear")	=	"" Then
		 strYear = "|"
	ELSE
		 strYear = FilterVar(Ucase(Trim(Request("txtYear"))),"''","S")
	END	IF
	IF Request("txtMonth") = ""	Then
		 strMonth	=	"|"
	ELSE
		 strMonth	=	FilterVar(Ucase(Trim(Request("txtMonth"))),"''","S")
	END	IF



	UNIValue(0,	0) = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")

	UNIValue(1,	0) = "^"
	UNIValue(1,	1) = strItemCd
	UNIValue(1,	2) = strYear
	UNIValue(1,	3) = strMonth

	UNILock	=	DISCONNREAD	:	UNIFlag	=	"1"


End	Sub

'----------------------------------------------------------------------------------------------------------
'	Query	Data
'	ADO의	Record Set이용하여 Query를 하고	Record Set을 넘겨서	MakeSpreadSheetData()으로	Spreadsheet에	데이터를 
'	뿌림 
'	ADO	객체를 생성할때	prjPublic.dll파일을	이용한다.(상세내용은 vb로	작성된 prjPublic.dll 소스	참조)
'----------------------------------------------------------------------------------------------------------
Sub	QueryData()
		Dim	lgstrRetMsg																							'☜	:	Record Set Return	Message	변수선언 
		Dim	lgADF																										'☜	:	ActiveX	Data Factory 지정	변수선언 
		Dim	iStr

		Set	lgADF		=	Server.CreateObject("prjPublic.cCtlTake")

		lgstrRetMsg	=	lgADF.QryRs(gDsnNo,	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0, rs1)

	Set	lgADF		=	Nothing

		iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <>	"0"	Then
				Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
		End	If


		If	rs0.EOF	And	rs0.BOF	 Then
		strFlag	=	"ERROR_PLANT"
				Response.Write "<Script	Language=vbscript>"	&	vbCr
		Response.Write "parent.frm1.txtItemCd.value	=	"""	&	"" & """"	&	vbCr
		Response.Write "parent.frm1.txtItemNm.value	=	"""	&	"" & """"	&	vbCr
				Response.Write "</Script>"		&	vbCr
	Else
				Response.Write "<Script	Language=vbscript>"	&	vbCr
		Response.Write "parent.frm1.txtItemNm.value	=	"""	&	ConvSPChars(rs0("Item_nm"))	&	"""" & vbCr
				Response.Write "</Script>"		&	vbCr
	End	If

		rs0.Close
		Set	rs0	=	Nothing


		If	rs1.EOF	And	rs1.BOF	 Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
				rs1.Close
				Set	rs1	=	Nothing
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "</Script>"		&	vbCr
				Response.end
		Else
				Call	MakeSpreadSheetData()
		End	If

'			Call DisplayMsgBox("x",	vbInformation, "이상하넹", "FASDFADS1111", I_MKSCRIPT)


End	Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서	Query가	되면 MakeSpreadSheetData()에 의해서	데이터를 스프레드시트에	뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub	MakeSpreadSheetData()
		Dim	iLoopCount
		Dim	iRowStr
		Dim	ColCnt

		lgDataExist		 = "Yes"
		If CLng(lgPageNo)	>	0	Then
			 rs1.Move			=	CLng(lgMaxCount) * CLng(lgPageNo)									 'lgMaxCount:Max Fetched Count at	once , lgStrPrevKeyIndex : Previous	PageNo
		End	If




	iLoopCount = 0
	Do while Not (rs1.EOF	Or rs1.BOF)
		iLoopCount =	iLoopCount + 1
		iRowStr	=	""
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("PLANT_CD"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("PLANT_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("ITEM_CD"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("ITEM_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("SPEC"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("BASIC_UNIT"))
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs1("PREV_GOOD_QTY"),ggExchRate.DecPoint,0)	'16
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs1("GOOD_ON_HAND_QTY"),ggExchRate.DecPoint,0)	'16

		iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow + iLoopCount



		If iLoopCount	-	1	<	lgMaxCount	Then
			istrData = istrData	&	iRowStr	&	Chr(11)	&	Chr(12)
		Else
			lgPageNo = lgPageNo	+	1
			Exit Do
		End	If
		rs1.MoveNext
	Loop



		If iLoopCount	<= lgMaxCount	Then																			'☜: Check if	next data	exists
			 lgPageNo	=	""
		End	If
		rs1.Close																												'☜: Close recordset object
		Set	rs1	=	Nothing																							'☜: Release ADF
End	Sub

'============================================================================================================
'	Name : SubBizSaveMulti
'	Desc : Save	Data
'============================================================================================================
Sub	SubBizSaveMulti()

End	Sub


'============================================================================================================
'	Name : SubBizSaveCreate
'	Desc : Query Data	from Db
'============================================================================================================
Sub	SubBizSaveMultiCreate(arrColVal)
On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

'----------	Developer	Coding part	(Start)	---------------------------------------------------------------
'A developer must	define field to	create record
'--------------------------------------------------------------------------------------------------------

'----------	Developer	Coding part	(End	)	---------------------------------------------------------------
End	Sub
'============================================================================================================
'	Name : SubBizSaveMultiUpdate
'	Desc : Update	Data from	Db
'============================================================================================================
Sub	SubBizSaveMultiUpdate(arrColVal)

On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

'----------	Developer	Coding part	(Start)	---------------------------------------------------------------
'A developer must	define field to	update record
'--------------------------------------------------------------------------------------------------------

'----------	Developer	Coding part	(End	)	---------------------------------------------------------------
End	Sub
'============================================================================================================
'	Name : SubBizSaveMultiDelete
'	Desc : Delete	Data from	Db
'============================================================================================================
Sub	SubBizSaveMultiDelete(arrColVal)

On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

'----------	Developer	Coding part	(Start)	---------------------------------------------------------------
'A developer must	define field to	update record
'--------------------------------------------------------------------------------------------------------

'----------	Developer	Coding part	(End	)	---------------------------------------------------------------
End	Sub
'============================================================================================================
'	Name : SubMakeSQLStatements
'	Desc : Make	SQL	statements
'============================================================================================================
Sub	SubMakeSQLStatements(pDataType,arrColVal)
Dim	iSelCount

On Error Resume	Next

'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End	Sub
'============================================================================================================
'	Name : CommonOnTransactionCommit
'	Desc : This	Sub	is called	by OnTransactionCommit Error handler
'============================================================================================================
Sub	CommonOnTransactionCommit()
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End	Sub

'============================================================================================================
'	Name : CommonOnTransactionAbort
'	Desc : This	Sub	is called	by OnTransactionAbort	Error	handler
'============================================================================================================
Sub	CommonOnTransactionAbort()
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End	Sub

'============================================================================================================
'	Name : SetErrorStatus
'	Desc : This	Sub	set	error	status
'============================================================================================================
Sub	SetErrorStatus()
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End	Sub
'============================================================================================================
'	Name : SubHandleError
'	Desc : This	Sub	handle error
'============================================================================================================
Sub	SubHandleError(pOpCode,pConn,pRs,pErr)
On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status


End	Sub
'==============================================================================
'	Function : SheetFocus
'	Description	:	에러발생시 Spread	Sheet에	포커스줌 
'==============================================================================
Function SheetFocus(Byval	lRow,	Byval	lCol,	Byval	iLoc)


End	Function

%>
