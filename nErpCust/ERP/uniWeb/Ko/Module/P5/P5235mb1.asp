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
'								PM9G112(Ȯ��)
'*	7. Modified	date(First)	:	2002/12/06
'*	8. Modified	date(Last)	:
'*	9. Modifier	(First)			:	Oh Chang Won
'* 10. Modifier	(Last)			:
'* 11. Comment							:
'* 12. Common	Coding Guide	:	this mark(��)	means	that "Do not change"
'*														this mark(��)	Means	that "may	 change"
'*														this mark(��)	Means	that "must change"
'* 13. History							:
'*
'*
'*
'*
'* 14. Business	Logic	of m9111ma1(����̵���û)
'**********************************************************************************************
Dim	lgOpModeCRUD

Dim	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0									'��	:	DBAgent	Parameter	���� 
Dim	rs1, rs2,	rs3, rs4,rs5
Dim	istrData
Dim	iStrPoNo
Dim	StrNextKey		'	���� �� 
Dim	lgStrPrevKey	'	���� �� 
Dim	iLngMaxRow		'	���� �׸�����	�ִ�Row
Dim	iLngRow
Dim	GroupCount
Dim	lgCurrency
Dim	index,Count			'	���� ��	Return ����	���� ������	���� ���� 
Dim	lgDataExist
Dim	lgPageNo
Dim	lgMaxCount
Dim	strFlag

	Const	C_SHEETMAXROWS_D	=	100

		On Error Resume	Next																														 '��:	Protect	system from	crashing
		Err.Clear																																				 '��:	Clear	Error	status

		Call HideStatusWnd																															 '��:	Hide Processing	message
	'------	Developer	Coding part	(Start ) ------------------------------------------------------------------

	'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------

		lgOpModeCRUD	=	Request("txtMode")

																																	'��: Read	Operation	Mode (CRUD)
		Select Case	lgOpModeCRUD
				Case CStr(UID_M0001)																												 '��:	Query
						 Call	 SubBizQueryMulti()
				Case CStr(UID_M0002)																												 '��:	Save,Update
						 Call	SubBizSaveMulti()
				Case CStr(UID_M0003)
						 Call	SubBizSaveMulti()
		End	Select

'============================================================================================================
'	Name : SubBizQuery
'	Desc : Query Data	from Db
'============================================================================================================
Sub	SubBizQuery()
		On Error Resume	Next																														 '��:	Protect	system from	crashing
		Err.Clear																																				 '��:	Clear	Error	status

End	Sub
'============================================================================================================
'	Name : SubBizSave
'	Desc : Save	Data
'============================================================================================================
Sub	SubBizSave()
		On Error Resume	Next																														 '��:	Protect	system from	crashing
		Err.Clear																																				 '��:	Clear	Error	status
End	Sub
'============================================================================================================
'	Name : SubBizDelete
'	Desc : Delete	DB data
'============================================================================================================
Sub	SubBizDelete()
		On Error Resume	Next																														 '��:	Protect	system from	crashing
		Err.Clear																																				 '��:	Clear	Error	status
End	Sub

'============================================================================================================
'	Name : SubBizQuery
'	Desc : Query Data	from Db
'============================================================================================================
Sub	SubBizQueryMulti()

		On Error Resume	Next

	iStrPoNo = Trim(Request("txtPoNo"))
	lgPageNo			 = UNICInt(Trim(Request("lgPageNo")),0)		 '��:	"0"(First),"1"(Second),"2"(Third),"3"(...)
		lgMaxCount		 = C_SHEETMAXROWS_D														'��	:	�ѹ��� �����ü�	�ִ� ����Ÿ	�Ǽ� 
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
'	Query�ϱ�	����	DB Agent �迭��	�̿��Ͽ� Query���� �����	���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub	FixUNISQLData()

	Dim	strItemCd
	Dim	strAppFrDt,	strAppToDt
	Dim	strFacility_Accnt
	Dim	strFacility_Cd


	Redim	UNISqlId(2)																											'��: SQL ID	������ ����	����Ȯ�� 
	Redim	UNIValue(2,	3)



	UNISqlId(0)	=	"160902saa"
	UNISqlId(1)	=	"P5110P570"



	IF Request("txtItemCd")	=	"" Then
		 strItemCd = "|"
	ELSE
		 strItemCd = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	END	IF
	IF Request("txtAppFrDt") = ""	Then
		 strAppFrDt	=	"|"
	ELSE
		 strAppFrDt	=	FilterVar(Ucase(Trim(Request("txtAppFrDt"))),"''","S")
	END	IF
	IF Request("txtAppToDt") = ""	Then
		 strAppToDt	=	"|"
	ELSE
		 strAppToDt	=	FilterVar(Ucase(Trim(Request("txtAppToDt"))),"''","S")
	END	IF



	UNIValue(0,	0) = FilterVar(Ucase(Trim(Request("txtItemCd"))),"''","S")
	UNIValue(1,	0) = "^"

	UNIValue(1,	1) = strItemCd
	UNIValue(1,	2) = strAppFrDt
	UNIValue(1,	3) = strAppToDt


	UNILock	=	DISCONNREAD	:	UNIFlag	=	"1"

End	Sub

'----------------------------------------------------------------------------------------------------------
'	Query	Data
'	ADO��	Record Set�̿��Ͽ� Query�� �ϰ�	Record Set�� �Ѱܼ�	MakeSpreadSheetData()����	Spreadsheet��	�����͸� 
'	�Ѹ� 
'	ADO	��ü�� �����Ҷ�	prjPublic.dll������	�̿��Ѵ�.(�󼼳����� vb��	�ۼ��� prjPublic.dll �ҽ�	����)
'----------------------------------------------------------------------------------------------------------
Sub	QueryData()
		Dim	lgstrRetMsg																							'��	:	Record Set Return	Message	�������� 
		Dim	lgADF																										'��	:	ActiveX	Data Factory ����	�������� 
		Dim	iStr

		Set	lgADF		=	Server.CreateObject("prjPublic.cCtlTake")

		lgstrRetMsg	=	lgADF.QryRs(gDsnNo,	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0, rs1)

	Set	lgADF		=	Nothing

		iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <>	"0"	Then
				Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
		End	If


		If	rs0.EOF	And	rs0.BOF	 Then
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
				rs0.Close
				rs1.Close
				Set	rs0	=	Nothing
				Set	rs1	=	Nothing
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "</Script>"		&	vbCr
				Response.end
		Else

'					Call	MakeHeaderData()
				Call	MakeSpreadSheetData()
		End	If

'			Call DisplayMsgBox("x",	vbInformation, "�̻��ϳ�", "FASDFADS1111", I_MKSCRIPT)


End	Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ�	Query��	�Ǹ� MakeSpreadSheetData()�� ���ؼ�	�����͸� ���������Ʈ��	�ѷ��ִ� ���ν��� 
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
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("MVMT_NO"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("PLANT_CD"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("PLANT_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("MVMT_SL_CD"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("SL_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("ITEM_CD"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("ITEM_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("BP_CD"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("BP_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs1("MVMT_BASE_UNIT"))
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs1("MVMT_BASE_QTY"),ggExchRate.DecPoint,0)	'16
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs1("MVMT_PRC"),ggExchRate.DecPoint,0)	'16
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs1("MVMT_DOC_AMT"),ggExchRate.DecPoint,0)	'16

				iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow + iLoopCount

				If iLoopCount	-	1	<	lgMaxCount Then
					 istrData	=	istrData & iRowStr & Chr(11) & Chr(12)
				Else
					 lgPageNo	=	lgPageNo + 1
					 Exit	Do
				End	If
				rs1.MoveNext
	Loop



		If iLoopCount	<= lgMaxCount	Then																			'��: Check if	next data	exists
			 lgPageNo	=	""
		End	If
		rs1.Close																												'��: Close recordset object
		Set	rs1	=	Nothing																							'��: Release ADF
End	Sub

'============================================================================================================
'	Name : SubBizSaveMulti
'	Desc : Save	Data
'============================================================================================================
Sub	SubBizSaveMulti()

	On Error Resume	Next
	Err.Clear

	Dim	pPY5G220		'��	pS13111
	Dim	iErrorPosition

	On Error Resume	Next																																 '��:	Protect	system from	crashing
	Err.Clear																			 '��:	Clear	Error	status

	Set	pPY5G220 = Server.CreateObject("PY5G220.CsF_Cast_PlanMultiSvr")

	If CheckSYSTEMError(Err,True)	=	true then
		Exit Sub
	End	If

	Dim	reqtxtSpread
	reqtxtSpread = Request("txtSpread")
	Call pPY5G220.PY5_MAINT_Y_FAC_CAST_PLAN_MULTI_SVR(gStrGlobalCollection,	trim(reqtxtSpread),	iErrorPosition)

	If CheckSYSTEMError2(Err,	True,	iErrorPosition & "��","","","","") = True	Then
		 Set pPY5G220	=	Nothing
		 Exit	Sub
	End	If

	Set	pPY5G220 = Nothing

	Response.Write "<Script	Language=vbscript>"	&	vbCr
	Response.Write "Parent.DBSaveOK	"						&	vbCr
	Response.Write "</Script>"									&	vbCr
End	Sub


'============================================================================================================
'	Name : SubBizSaveCreate
'	Desc : Query Data	from Db
'============================================================================================================
Sub	SubBizSaveMultiCreate(arrColVal)
On Error Resume	Next																														 '��:	Protect	system from	crashing
Err.Clear																																				 '��:	Clear	Error	status

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

On Error Resume	Next																														 '��:	Protect	system from	crashing
Err.Clear																																				 '��:	Clear	Error	status

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

On Error Resume	Next																														 '��:	Protect	system from	crashing
Err.Clear																																				 '��:	Clear	Error	status

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
On Error Resume	Next																														 '��:	Protect	system from	crashing
Err.Clear																																				 '��:	Clear	Error	status


End	Sub
'==============================================================================
'	Function : SheetFocus
'	Description	:	�����߻��� Spread	Sheet��	��Ŀ���� 
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