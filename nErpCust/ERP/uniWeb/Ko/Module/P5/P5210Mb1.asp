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
Dim	istrData1
Dim	istrData2
Dim	iStrPoNo
Dim	StrNextKey		'	다음 값 
Dim	lgStrPrevKey	'	이전 값 
Dim	iLngMaxRow1		'	현재 그리드의	최대Row
Dim	iLngMaxRow2		'	현재 그리드의	최대Row
Dim	iLngMaxRow3		'	현재 그리드의	최대Row
Dim	iLngRow
Dim	GroupCount
Dim	lgCurrency
Dim	index,Count			'	저장 후	Return 해줄	값을 넣을때	쓴는 변수 
Dim	lgDataExist
Dim	lgPageNo_A
Dim	lgPageNo_B
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
	lgPageNo_A			 = UNICInt(Trim(Request("lgPageNo_A")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_B			 = UNICInt(Trim(Request("lgPageNo_B")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
		lgMaxCount			 = C_SHEETMAXROWS_D														'☜	:	한번에 가져올수	있는 데이타	건수 
	lgDataExist			 = "No"
	iLngMaxRow1			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_A) + 1
	iLngMaxRow2			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_B) + 1

	lgStrPrevKey = Request("lgStrPrevKey")


	Call FixUNISQLData()
	Call QueryData()

	'====================
	'Call	PO_DTL List
	'====================

	'-----------------------
	'Result	data display area
	'-----------------------
	if Request("txtType")	=	"A"	Then							'☜	:	디테일 검색 

		Response.Write "<Script	Language=vbscript>"	&	vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData1.MaxRows < 1 then"						&	vbCr
		Response.Write "	End	if"							&	vbCr


		Response.Write "	.ggoSpread.Source				=	.frm1.vspdData1	"			&	vbCr
		Response.Write "	.ggoSpread.SSShowData			"""	&	istrData1	 & """"	&	vbCr
		Response.Write "	.lgPageNo_A	 = """ & lgPageNo_A		&	"""" & vbCr

		Response.Write " .DbQueryOk	"	&	vbCr
		Response.Write "End	With"		&	vbCr
		Response.Write "</Script>"		&	vbCr
	Else
		Response.Write "<Script	Language=vbscript>"	&	vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData2.MaxRows < 1 then"						&	vbCr
		Response.Write "	End	if"							&	vbCr


		Response.Write "	.ggoSpread.Source				=	.frm1.vspdData2	"			&	vbCr
		Response.Write "	.ggoSpread.SSShowData			"""	&	istrData2	 & """"	&	vbCr
		Response.Write "	.lgPageNo_A	 = """ & lgPageNo_B		&	"""" & vbCr

		Response.Write " .DbDtlQueryOk "	&	vbCr
		Response.Write "End	With"		&	vbCr
		Response.Write "</Script>"		&	vbCr
	End	if
End	Sub

'----------------------------------------------------------------------------------------------------------
'	Set	DB Agent arg
'----------------------------------------------------------------------------------------------------------
'	Query하기	전에	DB Agent 배열을	이용하여 Query문을 만드는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub	FixUNISQLData()

	Dim	strFacility_Accnt
	Dim	strFACILITY_LVL1
	Dim	strFACILITY_LVL2


	Redim	UNISqlId(4)																											'☜: SQL ID	저장을 위한	영역확보 
	Redim	UNIValue(4,	3)

	UNISqlId(0)	=	"I2241QA2A5"
	UNISqlId(1)	=	"I2241QA2A5"
	UNISqlId(2)	=	"P5110P510"
	UNISqlId(3)	=	"P5110P520"


	IF Request("txtFacility_Accnt")	=	"" Then
		 strFacility_Accnt = "|"
	ELSE
		 strFacility_Accnt = FilterVar(Ucase(Trim(Request("txtFacility_Accnt"))),"''","S")
	END	IF

	IF Request("txtItemGroupCd1")	=	"" Then
		 strFACILITY_LVL1	=	"|"
	ELSE
		 strFACILITY_LVL1	=	FilterVar(Ucase(Trim(Request("txtItemGroupCd1"))),"''","S")
	END	IF

	IF Request("txtItemGroupCd2")	=	"" Then
		 strFACILITY_LVL2	=	"|"
	ELSE
		 strFACILITY_LVL2	=	FilterVar(Ucase(Trim(Request("txtItemGroupCd2"))),"''","S")
	END	IF

	UNIValue(0,	0) = FilterVar(Ucase(Trim(Request("txtItemGroupCd1"))),"''","S")

	UNIValue(1,	0) = FilterVar(Ucase(Trim(Request("txtItemGroupCd2"))),"''","S")

	UNIValue(2,	0) = "^"
	UNIValue(2,	1) = strFacility_Accnt
	UNIValue(2,	2) = strFACILITY_LVL1
	UNIValue(2,	3) = strFACILITY_LVL2


	UNIValue(3,	0) = "^"
	UNIValue(3,	1) = strFacility_Accnt
	UNIValue(3,	2) = FilterVar(Ucase(Trim(Request("txtItemGroupCd1"))),"''","S")
	UNIValue(3,	3) = FilterVar(Ucase(Trim(Request("txtItemGroupCd2"))),"''","S")


	UNILock	=	DISCONNREAD	:	UNIFlag	=	"1"


End	Sub

'----------------------------------------------------------------------------------------------------------
'	Query	Data
'	ADO의	Record Set이용하여 Query를 하고	Record Set을 넘겨서	MakeSpreadSheetData1()으로 Spreadsheet에 데이터를 
'	뿌림 
'	ADO	객체를 생성할때	prjPublic.dll파일을	이용한다.(상세내용은 vb로	작성된 prjPublic.dll 소스	참조)
'----------------------------------------------------------------------------------------------------------
Sub	QueryData()
		Dim	lgstrRetMsg																							'☜	:	Record Set Return	Message	변수선언 
		Dim	lgADF																										'☜	:	ActiveX	Data Factory 지정	변수선언 
		Dim	iStr

		Set	lgADF		=	Server.CreateObject("prjPublic.cCtlTake")

		lgstrRetMsg	=	lgADF.QryRs(gDsnNo,	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0, rs1,	rs2, rs3)

	Set	lgADF		=	Nothing

		iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <>	"0"	Then
				Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
		End	If

	if Request("txtType")	<> "B" Then							'☜	:	디테일 검색 


			rs0.Close
			Set	rs0	=	Nothing
			rs1.Close
			Set	rs1	=	Nothing
			rs3.Close
			Set	rs3	=	Nothing

			If	rs2.EOF	And	rs2.BOF	 Then
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
					Response.Write "<Script	Language=vbscript>"	&	vbCr
					Response.Write "</Script>"		&	vbCr
					Response.end
			Else
					Call	MakeSpreadSheetData1()
			End	If
	Else
			rs0.Close
			Set	rs0	=	Nothing
			rs1.Close
			Set	rs1	=	Nothing
			rs2.Close
			Set	rs2	=	Nothing

			If	rs3.EOF	And	rs3.BOF	 Then
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
					Response.Write "<Script	Language=vbscript>"	&	vbCr
					Response.Write "</Script>"		&	vbCr
					Response.end
			Else
					Call	MakeSpreadSheetData2()
			End	If
	End	If


'			Call DisplayMsgBox("x",	vbInformation, "이상하넹", "FASDFADS1111", I_MKSCRIPT)


End	Sub

Sub	MakeHeaderData()

	Dim	strPoType				'이동유형 
	Dim	strPoTypeNm			'이동유형명 
		Dim	strPoDt						'등록일 
		Dim	strSupplierCd			'공급창고 
		Dim	strSupplierNm			'공급창고명 
		Dim	strGroupCd				'구매그룹 
	Dim	strGroupNm				'구매그룹명 
		Dim	strSuppPrsn				'공급처담당 
		Dim	strTel						'긴급연락처 
		Dim	strRemark					'비고 
		Dim	strReleaseflg			'확정여부 
	Dim	strClsflg					'마감여부 
	Dim	strStosono				'수주번호 

	strPoType			=	rs0(1)
	strPoTypeNm		=	rs0(2)
		strSupplierCd	=	rs0(3)
		strSupplierNm	=	rs0(4)
		strGroupCd		=	rs0(5)
	strGroupNm		=	rs0(6)
		strPoDt				=	rs0(7)
		strSuppPrsn		=	rs0(8)
		strTel				=	rs0(9)
		strReleaseflg	=	rs0(10)
		strClsflg			=	rs0(11)
		strRemark			=	rs0(12)
		strStosono		=	rs0(13)

	Response.Write "<Script	Language=vbscript>"	&	vbCr
	Response.Write "With parent" & vbCr
	Response.Write "if .frm1.vspdData1.MaxRows = 0 then	"	&	vbCr
	Response.Write "	.frm1.txtSupplierCd.value	=	"""	&	Trim(UCase(ConvSPChars(strSupplierCd)))							 	&	"""" & vbCr
	Response.Write "	.frm1.txtSupplierNm.value	=	"""	&	Trim(UCase(ConvSPChars(strSupplierNm)))							 	&	"""" & vbCr
	Response.Write "	.frm1.txtGroupCd.value		=	"""	&	Trim(UCase(ConvSPChars(strGroupCd)))									&	"""" & vbCr
	Response.Write "	.frm1.txtGroupNm.value		=	"""	&	Trim(UCase(ConvSPChars(strGroupNm)))							&	"""" & vbCr
	Response.Write "	.frm1.txtPoTypeCd.value		=	"""	&	Trim(UCase(ConvSPChars(strPoType)))				&	"""" & vbCr
	Response.Write "	.frm1.txtPoTypeCdNm.value		=	"""	&	Trim(UCase(ConvSPChars(strPoTypeNm)))				&	"""" & vbCr
	Response.Write "	.frm1.txtPoNo1.value		=	"""	&	Trim(UCase(ConvSPChars(iStrPoNo)))							 & """"	&	vbCr
	Response.Write "	.frm1.txtPoNo.value				=	"""	&	Trim(UCase(ConvSPChars(iStrPoNo)))							 & """"	&	vbCr
	Response.Write "	.frm1.txtPoDt.text			 = """ & UNIDateClientFormat(strPoDt)					&	"""" & vbCr


	Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE	Then parent.lgIntFlgMode = parent.parent.OPMD_UMODE	"	&	vbCr

	Response.Write " end if		"	&	vbCr
	Response.Write " End With	"	&	vbCr
		Response.Write "</Script>"	&	vbCr

		rs0.Close																												'☜: Close recordset object
		Set	rs0	=	Nothing
End	Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서	Query가	되면 MakeSpreadSheetData1()에	의해서 데이터를	스프레드시트에 뿌려주는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub	MakeSpreadSheetData1()
		Dim	iLoopCount
		Dim	iRowStr
		Dim	ColCnt

		lgDataExist		 = "Yes"
		If CLng(lgPageNo_A)	>	0	Then
			 rs2.Move			=	CLng(lgMaxCount) * CLng(lgPageNo_A)									 'lgMaxCount:Max Fetched Count at	once , lgStrPrevKeyIndex : Previous	PageNo
		End	If

	iLoopCount = 0
	Do while Not (rs2.EOF	Or rs2.BOF)
				iLoopCount =	iLoopCount + 1
				iRowStr	=	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("Facility_Accnt"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("Facility_Accnt_Nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("Facility_Lvl1"))
				iRowStr	=	iRowStr	&	Chr(11)	&	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("FACILITY_LVL1_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("Facility_Lvl2"))
				iRowStr	=	iRowStr	&	Chr(11)	&	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("FACILITY_LVL2_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	iloopcount
				iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow1	+	iLoopCount

				If iLoopCount	-	1	<	lgMaxCount Then
					 istrData1 = istrData1 & iRowStr & Chr(11) & Chr(12)
				Else
					 lgPageNo_A	=	lgPageNo_A + 1
					 Exit	Do
				End	If
				rs2.MoveNext
	Loop

		If iLoopCount	<= lgMaxCount	Then																			'☜: Check if	next data	exists
			 lgPageNo_A	=	""
		End	If
		rs2.Close																												'☜: Close recordset object
		Set	rs2	=	Nothing																							'☜: Release ADF

End	Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서	Query가	되면 MakeSpreadSheetData2()에	의해서 데이터를	스프레드시트에 뿌려주는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub	MakeSpreadSheetData2()
		Dim	iLoopCount
		Dim	iRowStr
		Dim	ColCnt

		lgDataExist		 = "Yes"
		If CLng(lgPageNo_B)	>	0	Then
			 rs3.Move			=	CLng(lgMaxCount) * CLng(lgPageNo_B)									 'lgMaxCount:Max Fetched Count at	once , lgStrPrevKeyIndex : Previous	PageNo
		End	If

	iLoopCount = 0
	Do while Not (rs3.EOF	Or rs3.BOF)
				iLoopCount =	iLoopCount + 1
				iRowStr	=	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("SEQ"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("ZINSP_PART"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("ZINSP_PART_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("INSP_PART"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("INSP_PART_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("INSP_METH"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("INSP_METH_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("INSP_DECISION"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("INSP_DECISION_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("ST_GO_GUBUN"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs3("ST_GO_GUBUN_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	request("hChkFlag")
				iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow2	+	iLoopCount

				If iLoopCount	-	1	<	lgMaxCount Then
					 istrData2 = istrData2 & iRowStr & Chr(11) & Chr(12)
				Else
					 lgPageNo_B	=	lgPageNo_B + 1
					 Exit	Do
				End	If
				rs3.MoveNext
	Loop

		If iLoopCount	<= lgMaxCount	Then																			'☜: Check if	next data	exists
			 lgPageNo_B	=	""
		End	If
		rs3.Close																												'☜: Close recordset object
		Set	rs3	=	Nothing																							'☜: Release ADF
End	Sub
'============================================================================================================
'	Name : SubBizSaveMulti
'	Desc : Save	Data
'============================================================================================================
Sub	SubBizSaveMulti()

	On Error Resume	Next
	Err.Clear

	Dim	pPY5G210		'구	pS13111
	Dim	iErrorPosition

	On Error Resume	Next																																 '☜:	Protect	system from	crashing
	Err.Clear																			 '☜:	Clear	Error	status

	Set	pPY5G210 = Server.CreateObject("PY5G210.CsF_Chk_LISTMultiSvr")

	If CheckSYSTEMError(Err,True)	=	true then
		Exit Sub
	End	If

	Dim	reqtxtSpread
	reqtxtSpread = Request("txtSpread")
	Call pPY5G210.PY5_MAINT_FA_CHK_LIST_MULTI_SVR(gStrGlobalCollection,	trim(reqtxtSpread),	iErrorPosition)

	If CheckSYSTEMError2(Err,	True,	iErrorPosition & "행","","","","") = True	Then
		 Set pPY5G210	=	Nothing
		 Exit	Sub
	End	If

	Set	pPY5G210 = Nothing

	Response.Write "<Script	Language=vbscript>"	&	vbCr
	Response.Write "Parent.DBSaveOK	"						&	vbCr
	Response.Write "</Script>"									&	vbCr
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

If Trim(lRow)	=	"" Then	Exit Function
If iLoc	=	I_INSCRIPT Then
	strHTML	=	"parent.frm1.vspdData1.focus"	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.Row = " & lRow	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.Col = " & lCol	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.Action	=	0" & vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.SelStart	=	0	"	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.SelLength = len(parent.frm1.vspdData1.Text) " & vbCrLf
	Response.Write strHTML
ElseIf iLoc	=	I_MKSCRIPT Then
	strHTML	=	"<"	&	"Script	LANGUAGE=VBScript" & ">" & vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.focus"	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.Row = " & lRow	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.Col = " & lCol	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.Action	=	0" & vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.SelStart	=	0	"	&	vbCrLf
	strHTML	=	strHTML	&	"parent.frm1.vspdData1.SelLength = len(parent.frm1.vspdData1.Text) " & vbCrLf
	strHTML	=	strHTML	&	"</" & "Script"	&	">"	&	vbCrLf
	Response.Write strHTML
End	If
End	Function

%>
