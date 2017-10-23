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

Call LoadInfTB19029B("Q",	"P", "NOCOOKIE","RB")

'**********************************************************************************************
'*	1. Module	Name		:	Procurement
'*	2. Function	Name		:
'*	3. Program ID			:	s3112rb51
'*	4. Program Name			:
'*	5. Program Desc			:
'*	6. Comproxy	List		:	
'*	7. Modified	date(First)	:	2005/11/29
'*	8. Modified	date(Last)	:
'*	9. Modifier	(First)		:	nhg
'* 10. Modifier	(Last)		:
'* 11. Comment				:
'* 12. Common
'*     Coding Guide			:	this mark(☜)	means	that "Do not change"
'*								this mark(⊙)	Means	that "may	 change"
'*								this mark(☆)	Means	that "must change"
'* 13. History				:
'**********************************************************************************************
Dim	lgOpModeCRUD

Dim	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0									'☜	:	DBAgent	Parameter	선언 
Dim	rs1, rs2,	rs3, rs4,rs5
Dim	istrData1
Dim	istrData2

Dim	istrData3

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

Dim	lgPageNo_C

Dim	lgMaxCount

Dim	strFlag


Const	C_SHEETMAXROWS_D	=	100

On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

Call HideStatusWnd																															 '☜:	Hide Processing	message
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------

'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
lgOpModeCRUD	=	Request("txtMode")

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
Sub SubBizQuery()
		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status

End Sub
'============================================================================================================
'	Name : SubBizSave
'	Desc : Save	Data
'============================================================================================================
Sub SubBizSave()
		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status
End Sub
'============================================================================================================
'	Name : SubBizDelete
'	Desc : Delete	DB data
'============================================================================================================
Sub SubBizDelete()
		On Error Resume	Next																														 '☜:	Protect	system from	crashing
		Err.Clear																																				 '☜:	Clear	Error	status
End Sub


'============================================================================================================
'	Name : SubBizQuery
'	Desc : Query Data	from Db
'============================================================================================================
Sub SubBizQueryMulti()

	On Error Resume	Next

	lgPageNo_A			 = UNICInt(Trim(Request("lgPageNo_A")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_B			 = UNICInt(Trim(Request("lgPageNo_B")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_C			 = UNICInt(Trim(Request("lgPageNo_C")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount			 = C_SHEETMAXROWS_D														'☜	:	한번에 가져올수	있는 데이타	건수 
	lgDataExist			 = "No"
	iLngMaxRow1			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_A)
	iLngMaxRow2			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_B)
	iLngMaxRow2			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_C)

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
	Response.Write "	If .frm1.vspddata.MaxRows < 1 then"						&	vbCr
	Response.Write "	End If"							&	vbCr
	Response.Write "	.ggoSpread.Source				=	.frm1.vspddata "			&	vbCr
	Response.Write "	.ggoSpread.SSShowData			"""	&	istrData1	 & """"	&	vbCr
	Response.Write "	.lgPageNo_A	 = """ & lgPageNo_A		&	"""" & vbCr
	Response.Write " .DbDtlQuery1Ok	"	&	vbCr
	Response.Write "End	With"		&	vbCr
	Response.Write "</Script>"		&	vbCr

End Sub

'----------------------------------------------------------------------------------------------------------
'	Set	DB Agent arg
'----------------------------------------------------------------------------------------------------------
'	Query하기	전에	DB Agent 배열을	이용하여 Query문을 만드는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim	UNISqlId(0)																											'☜: SQL ID	저장을 위한	영역확보 
	Redim	UNIValue(0,	0)

	UNISqlId(0)	=	"S3112RA51"

' ********************************************************
' 헤더와 디테일이 검색조건이 달라서 부득히 변수 하나 더 씀 
' ********************************************************

	UNIValue(0,	0) = FilterVar(Ucase(Trim(Request("txtSoNo"))),"''","S")

	UNILock	=	DISCONNREAD	:	UNIFlag	=	"1"

End Sub

'----------------------------------------------------------------------------------------------------------
'	Query	Data
'	ADO의	Record Set이용하여 Query를 하고	Record Set을 넘겨서	MakeSpreadSheetData1()으로 Spreadsheet에 데이터를 
'	뿌림 
'	ADO	객체를 생성할때	prjPublic.dll파일을	이용한다.(상세내용은 vb로	작성된 prjPublic.dll 소스	참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim	lgstrRetMsg																							'☜	:	Record Set Return	Message	변수선언 
	Dim	lgADF																										'☜	:	ActiveX	Data Factory 지정	변수선언 
	Dim	iStr

	Set	lgADF = Server.CreateObject("prjPublic.cCtlTake")
	lgstrRetMsg	= lgADF.QryRs(gDsnNo,	UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	Set	lgADF = Nothing

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <>	"0"	Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
	End If

	If	rs0.EOF	And	rs0.BOF	 Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	ELSE
		Call MakeSpreadSheetData1()
	END IF
End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서	Query가	되면 MakeSpreadSheetData1()에	의해서 데이터를	스프레드시트에 뿌려주는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData1()

	Dim	iLoopCount
	Dim	iRowStr
	Dim  i, j
	lgDataExist		 = "Yes"

	If CLng(lgPageNo_A)	>	0	Then
		 rs0.Move			=	CLng(lgMaxCount) * CLng(lgPageNo_A)
	End If

	iLoopCount = 0
	Do while Not (rs0.EOF	Or rs0.BOF)
		iLoopCount =	iLoopCount + 1
		iRowStr	= ""


		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(0))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(1))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(2))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(3))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(4))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(5))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(6))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(7))
		
		If ConvSPChars(rs0(7)) = "Y" Then
			iRowStr	= iRowStr & Chr(11) & "진단가"
		Else
			iRowStr	= iRowStr & Chr(11) & "가단가"
		End If	
		
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(8))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(9))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(10))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(11))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(12))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(13))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(14))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(15))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(16))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(17))
		iRowStr	= iRowStr & Chr(11) & ConvSPChars(rs0(18))
		
		If ConvSPChars(rs0(18)) = "1" Then
			iRowStr	= iRowStr & Chr(11) & "포함"
		Else
			iRowStr	= iRowStr & Chr(11) & "별도"
		End If	
		
		iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow1	+	iLoopCount

		If iLoopCount	-	1	<	lgMaxCount Then
			 istrData1 = istrData1 & iRowStr & Chr(11) & Chr(12)
		Else
			 lgPageNo_A	=	lgPageNo_A + 1
			 Exit	Do
		End If
		rs0.MoveNext
	Loop

	If iLoopCount	<= lgMaxCount	Then
		 lgPageNo_A	=	""
	End If
	rs0.Close
	Set	rs0	=	Nothing

End Sub

'============================================================================================================
'	Name : SubBizSaveMulti
'	Desc : Save	Data
'============================================================================================================
Sub SubBizSaveMulti()


End Sub


'============================================================================================================
'	Name : SubBizSaveCreate
'	Desc : Query Data	from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

'----------	Developer	Coding part	(Start)	---------------------------------------------------------------
'A developer must	define field to	create record
'--------------------------------------------------------------------------------------------------------

'----------	Developer	Coding part	(End	)	---------------------------------------------------------------
End Sub
'============================================================================================================
'	Name : SubBizSaveMultiUpdate
'	Desc : Update	Data from	Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

'----------	Developer	Coding part	(Start)	---------------------------------------------------------------
'A developer must	define field to	update record
'--------------------------------------------------------------------------------------------------------

'----------	Developer	Coding part	(End	)	---------------------------------------------------------------
End Sub
'============================================================================================================
'	Name : SubBizSaveMultiDelete
'	Desc : Delete	Data from	Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status

'----------	Developer	Coding part	(Start)	---------------------------------------------------------------
'A developer must	define field to	update record
'--------------------------------------------------------------------------------------------------------

'----------	Developer	Coding part	(End	)	---------------------------------------------------------------
End Sub

'============================================================================================================
'	Name : CommonOnTransactionCommit
'	Desc : This	Sub is called	by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End Sub

'============================================================================================================
'	Name : CommonOnTransactionAbort
'	Desc : This	Sub is called	by OnTransactionAbort	Error	handler
'============================================================================================================
Sub CommonOnTransactionAbort()
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End Sub

'============================================================================================================
'	Name : SetErrorStatus
'	Desc : This	Sub set	error	status
'============================================================================================================
Sub SetErrorStatus()
'------	Developer	Coding part	(Start ) ------------------------------------------------------------------
'------	Developer	Coding part	(End	 ) ------------------------------------------------------------------
End Sub
'============================================================================================================
'	Name : SubHandleError
'	Desc : This	Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
On Error Resume	Next																														 '☜:	Protect	system from	crashing
Err.Clear																																				 '☜:	Clear	Error	status


End Sub



%>
