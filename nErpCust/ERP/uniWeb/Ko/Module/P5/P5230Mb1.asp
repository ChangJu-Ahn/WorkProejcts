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
Call LoadInfTB19029B("I",	"*", "NOCOOKIE","MB")

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

	lgPageNo_A			 = UNICInt(Trim(Request("lgPageNo_A")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_B			 = UNICInt(Trim(Request("lgPageNo_B")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_C			 = UNICInt(Trim(Request("lgPageNo_C")),0)		 '☜:	"0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount			 = C_SHEETMAXROWS_D														'☜	:	한번에 가져올수	있는 데이타	건수 
	lgDataExist			 = "No"
	iLngMaxRow1			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_A) + 1
	iLngMaxRow2			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_B) + 1
	iLngMaxRow3			 = CDbl(lgMaxCount)	*	CDbl(lgPageNo_C) + 1

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
		Response.Write "	If .frm1.vspdData.MaxRows	<	1	then"						&	vbCr
		Response.Write "	End	if"							&	vbCr


		Response.Write "	.ggoSpread.Source				=	.frm1.vspdData "			&	vbCr
		Response.Write "	.ggoSpread.SSShowData			"""	&	istrData1	 & """"	&	vbCr
		Response.Write "	.lgPageNo_A	 = """ & lgPageNo_A		&	"""" & vbCr

		Response.Write " .DbQueryOk	"	&	vbCr
		Response.Write "End	With"		&	vbCr
		Response.Write "</Script>"		&	vbCr
	Elseif Request("txtType")	=	"B"	Then
		Response.Write "<Script	Language=vbscript>"	&	vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData1.MaxRows < 1 then"						&	vbCr
		Response.Write "	End	if"							&	vbCr
		Response.Write "	.ggoSpread.Source				=	.frm1.vspdData1	"			&	vbCr
		Response.Write "	.ggoSpread.SSShowData			"""	&	istrData2	 & """"	&	vbCr
		Response.Write "	.lgPageNo_B	 = """ & lgPageNo_B		&	"""" & vbCr
		Response.Write " .DbDtlQuery2	"	&	vbCr
		Response.Write "End	With"		&	vbCr
		Response.Write "</Script>"		&	vbCr
	Elseif Request("txtType")	=	"C"	Then
		Response.Write "<Script	Language=vbscript>"	&	vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData2.MaxRows < 1 then"						&	vbCr
		Response.Write "	End	if"							&	vbCr
		Response.Write "	.ggoSpread.Source				=	.frm1.vspdData2	"			&	vbCr
		Response.Write "	.ggoSpread.SSShowData			"""	&	istrData3	 & """"	&	vbCr
		Response.Write "	.lgPageNo_C	 = """ & lgPageNo_C		&	"""" & vbCr

'			Response.Write " .DbDtlQueryOk2	"	&	vbCr
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

	Dim	strFacility_Accnt, strFacility_Cd
	Dim	strWork_Dt,	strPlantCd

	Redim	UNISqlId(3)																											'☜: SQL ID	저장을 위한	영역확보 
	Redim	UNIValue(3,	4)

	UNISqlId(0)	=	"I2241QA2A4"
	UNISqlId(1)	=	"P5110P5AB"
	UNISqlId(2)	=	"P5110P541"

	IF Request("txtWork_Dt") = ""	Then
		 strWork_Dt	=	"|"
	ELSE
		 strWork_Dt	=	FilterVar(Ucase(Trim(Request("txtWork_Dt"))),"''","S")
	END	IF
	IF Request("txtFacility_Cd") = ""	Then
		 strFacility_Cd	=	"|"
	ELSE
		 strFacility_Cd	=	FilterVar(Ucase(Trim(Request("txtFacility_Cd"))),"''","S")
	END	IF
	IF Request("txtPlantCd") = ""	Then
		 strPlantCd	=	"|"
	ELSE
		 strPlantCd	=	FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	END	IF
	IF Request("CboFacility_Accnt")	=	"" Then
		 strFacility_Accnt = "|"
	ELSE
		 strFacility_Accnt = FilterVar(Ucase(Trim(Request("CboFacility_Accnt"))),"''","S")
	END	IF


	UNIValue(0,	0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")

	UNIValue(1,	0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(1,	1) = FilterVar(Ucase(Trim(Request("txtFacility_Cd"))),"''","S")

	UNIValue(2,	0) = "^"
	UNIValue(2,	1) = strWork_Dt
	UNIValue(2,	2) = strPlantCd
	UNIValue(2,	3) = strFacility_Accnt
	UNIValue(2,	4) = strFacility_Cd

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
	Dim	strFacility_Accnt, strFacility_Cd
	Dim	strWork_Dt,	strPlantCd


	IF Request("txtWork_Dt") = ""	Then
		 strWork_Dt	=	"|"
	ELSE
		 strWork_Dt	=	FilterVar(Ucase(Trim(Request("txtWork_Dt"))),"''","S")
	END	IF
	IF Request("txtFacility_Cd") = ""	Then
		 strFacility_Cd	=	"|"
	ELSE
		 strFacility_Cd	=	FilterVar(Ucase(Trim(Request("txtFacility_Cd"))),"''","S")
	END	IF
	IF Request("txtPlantCd") = ""	Then
		 strPlantCd	=	"|"
	ELSE
		 strPlantCd	=	FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	END	IF
	IF Request("CboFacility_Accnt")	=	"" Then
		 strFacility_Accnt = "|"
	ELSE
		 strFacility_Accnt = FilterVar(Ucase(Trim(Request("CboFacility_Accnt"))),"''","S")
	END	IF


	if Request("txtType")	=	"A"	Then							'☜	:	디테일 검색 

			Set	lgADF		=	Server.CreateObject("prjPublic.cCtlTake")
			lgstrRetMsg	=	lgADF.QryRs(gDsnNo,	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0, rs1,	rs2)

		Set	lgADF		=	Nothing

			iStr = Split(lgstrRetMsg,gColSep)

		If iStr(0) <>	"0"	Then
			Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
		End	If


		If Trim(Request("txtPlantCd")) <> "" Then
			If	rs0.EOF	And	rs0.BOF	 Then
				Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "parent.frm1.txtPlantNm.value = """ & ""	&	"""" & vbCr
				Response.Write "</Script>"		&	vbCr
				Response.end
			Else
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0("Plant_Nm"))	&	"""" & vbCr
				Response.Write "</Script>"		&	vbCr
			End	If
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		End If

		If Trim(Request("txtFacility_cd")) <> "" Then
			If	rs1.EOF	And	rs1.BOF	 Then
				Call DisplayMsgBox("Y50050", vbOKOnly, "", "", I_MKSCRIPT)
				strFlag	=	"ERROR_FACILITY"
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "parent.frm1.txtFacility_Nm.value = """ & ""	&	"""" & vbCr
				Response.Write "</Script>"		&	vbCr
				Response.end
			Else
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "parent.frm1.txtFacility_Nm.value = """ & ConvSPChars(rs1("FACILITY_NM")) & """"	&	vbCr
				Response.Write "</Script>"		&	vbCr
			End	If
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtFacility_Nm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		End If



		rs0.Close
		Set	rs0	=	Nothing
		rs1.Close
		Set	rs1	=	Nothing

		If	rs2.EOF	And	rs2.BOF	 Then
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
				Response.Write "<Script	Language=vbscript>"	&	vbCr
				Response.Write "</Script>"		&	vbCr
				Response.end
		Else
				Call	MakeSpreadSheetData1()
		End	If
	Elseif Request("txtType")	=	"B"	Then

		Redim	UNISqlId(1)																											'☜: SQL ID	저장을 위한	영역확보 
		Redim	UNIValue(1,	4)
		UNIValue(0,	0) = "^"
		UNIValue(0,	1) = strWork_Dt
		UNIValue(0,	2) = strPlantCd
		UNIValue(0,	3) = strFacility_Accnt
		UNIValue(0,	4) = FilterVar(Ucase(Trim(Request("txtFacility_Cd"))),"''","S")

		UNISqlId(0)	=	"P5110P551"

			Set	lgADF		=	Server.CreateObject("prjPublic.cCtlTake")
			lgstrRetMsg	=	lgADF.QryRs(gDsnNo,	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0)

		Set	lgADF		=	Nothing

			iStr = Split(lgstrRetMsg,gColSep)

		If iStr(0) <>	"0"	Then
					Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
			End	If

			If	rs0.EOF	And	rs0.BOF	 Then
'				Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'						Response.Write "<Script	Language=vbscript>"	&	vbCr
'						Response.Write "</Script>"		&	vbCr
'						Response.end
			Else
					Call	MakeSpreadSheetData2()
			End	If
	Elseif Request("txtType")	=	"C"	Then

		Redim	UNISqlId(1)																											'☜: SQL ID	저장을 위한	영역확보 
		Redim	UNIValue(1,	4)
		UNIValue(0,	0) = "^"
		UNIValue(0,	1) = strWork_Dt
		UNIValue(0,	2) = strPlantCd
		UNIValue(0,	3) = strFacility_Accnt
		UNIValue(0,	4) = FilterVar(Ucase(Trim(Request("txtFacility_Cd"))),"''","S")

		UNISqlId(0)	=	"P5110P561"

		Set	lgADF		=	Server.CreateObject("prjPublic.cCtlTake")
		lgstrRetMsg	=	lgADF.QryRs(gDsnNo,	UNISqlId,	UNIValue,	UNILock, UNIFlag,	rs0)

		Set	lgADF		=	Nothing
		iStr = Split(lgstrRetMsg,gColSep)

		If iStr(0) <>	"0"	Then
			Call ServerMesgBox(lgstrRetMsg , vbInformation,	I_MKSCRIPT)
		End	If

		If	rs0.EOF	And	rs0.BOF	 Then

		Else
				Call	MakeSpreadSheetData3()
		End	If
	End	If

'			Call DisplayMsgBox("x",	vbInformation, "이상하넹", "FASDFADS1111", I_MKSCRIPT)


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
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("FAC_CAST_CD"))
		iRowStr	=	iRowStr	&	Chr(11)	&	""
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("FACILITY_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("SET_PLANT"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("PLANT_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("FACILITY_ACCNT_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	UNIDateClientFormat(rs2("WORK_DT"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("INSP_TEXT"))
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs2("INSP_HOUR"),ggQty.DecPoint,0)
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs2("INSP_MIN"),ggQty.DecPoint,0)
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("REQ_DEPT"))
		iRowStr	=	iRowStr	&	Chr(11)	&	""
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("REQ_DEPT_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("INSP_DEPT"))
		iRowStr	=	iRowStr	&	Chr(11)	&	""
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("INSP_DEPT_NM"))
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs2("INSP_EMP_QTY"),ggQty.DecPoint,0)
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs2("PAYROLL"),ggQty.DecPoint,0)
		iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs2("MATL_COST"),ggQty.DecPoint,0)
		iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs2("INSP_FLAG"))
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
		 rs0.Move			=	CLng(lgMaxCount) * CLng(lgPageNo_B)									 'lgMaxCount:Max Fetched Count at	once , lgStrPrevKeyIndex : Previous	PageNo
	End	If

	iLoopCount = 0
	Do while Not (rs0.EOF	Or rs0.BOF)
				iLoopCount =	iLoopCount + 1
				iRowStr	=	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("SEQ"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("ZINSP_PART"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("ZINSP_PART_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_PART"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_PART_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_METH"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_METH_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_DECISION"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_DECISION_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("ST_GO_GUBUN"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("ST_GO_GUBUN_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("SURY_ASSY"))
				iRowStr	=	iRowStr	&	Chr(11)	&	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("ITEM_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs0("S_QTY"),ggQty.DecPoint,0)
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs0("PRICE"),ggQty.DecPoint,0)
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs0("SURY_AMT"),ggQty.DecPoint,0)
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("CURRENCY"))
				iRowStr	=	iRowStr	&	Chr(11)	&	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("SURI_TYPE"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("SURI_TYPE_NM"))

				iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow2	+	iLoopCount

				If iLoopCount	-	1	<	lgMaxCount Then
					 istrData2 = istrData2 & iRowStr & Chr(11) & Chr(12)
				Else
					 lgPageNo_B	=	lgPageNo_B + 1
					 Exit	Do
				End	If
				rs0.MoveNext
	Loop
		If iLoopCount	<= lgMaxCount	Then																			'☜: Check if	next data	exists
			 lgPageNo_B	=	""
		End	If
		rs0.Close																												'☜: Close recordset object
		Set	rs0	=	Nothing																							'☜: Release ADF
End	Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서	Query가	되면 MakeSpreadSheetData3()에	의해서 데이터를	스프레드시트에 뿌려주는	프로시져 
'----------------------------------------------------------------------------------------------------------
Sub	MakeSpreadSheetData3()
		Dim	iLoopCount
		Dim	iRowStr
		Dim	ColCnt

		lgDataExist		 = "Yes"


		If CLng(lgPageNo_C)	>	0	Then
			 rs0.Move			=	CLng(lgMaxCount) * CLng(lgPageNo_C)									 'lgMaxCount:Max Fetched Count at	once , lgStrPrevKeyIndex : Previous	PageNo
		End	If

	iLoopCount = 0
	Do while Not (rs0.EOF	Or rs0.BOF)
				iLoopCount =	iLoopCount + 1
				iRowStr	=	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("SEQ"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_EMP_GB"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_EMP_GB_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("INSP_EMP_CD"))
				iRowStr	=	iRowStr	&	Chr(11)	&	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("Emp_nm"))
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("CUST_CD"))
				iRowStr	=	iRowStr	&	Chr(11)	&	""
				iRowStr	=	iRowStr	&	Chr(11)	&	ConvSPChars(rs0("CUST_NM"))
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs0("INSP_HOUR"),ggQty.DecPoint,0)
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs0("INSP_MIN"),ggQty.DecPoint,0)
				iRowStr	=	iRowStr	&	Chr(11)	&	UNINumClientFormat(rs0("PAYROLL"),ggQty.DecPoint,0)

				iRowStr	=	iRowStr	&	Chr(11)	&	iLngMaxRow3	+	iLoopCount

				If iLoopCount	-	1	<	lgMaxCount Then
					 istrData3 = istrData3 & iRowStr & Chr(11) & Chr(12)
				Else
					 lgPageNo_C	=	lgPageNo_C + 1
					 Exit	Do
				End	If
				rs0.MoveNext
	Loop

		If iLoopCount	<= lgMaxCount	Then																			'☜: Check if	next data	exists
			 lgPageNo_C	=	""
		End	If
		rs0.Close																												'☜: Close recordset object
		Set	rs0	=	Nothing																							'☜: Release ADF
End	Sub

'============================================================================================================
'	Name : SubBizSaveMulti
'	Desc : Save	Data
'============================================================================================================
Sub	SubBizSaveMulti()

	On Error Resume	Next
	Err.Clear

	Dim	pPY5G230		'구	pS13111
	Dim	iErrorPosition

	On Error Resume	Next																																 '☜:	Protect	system from	crashing
	Err.Clear																			 '☜:	Clear	Error	status

	Set	pPY5G230 = Server.CreateObject("PY5G230.CsF_Cast_PlanMultiSvr")

	If CheckSYSTEMError(Err,True)	=	true then
		Exit Sub
	End	If

	Dim	reqtxtSpread
	reqtxtSpread = Request("txtSpread")
	Call pPY5G230.PY5_MAINT_Y_FAC_CAST_MULTI_SVR(gStrGlobalCollection, trim(reqtxtSpread), iErrorPosition)

	If CheckSYSTEMError2(Err,	True,	iErrorPosition & "행","","","","") = True	Then
		 Set pPY5G230	=	Nothing
		 Exit	Sub
	End	If

	Set	pPY5G230 = Nothing

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
