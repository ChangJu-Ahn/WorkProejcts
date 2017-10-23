<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : 금형관리 
'*  2. Function Name        : 
'*  3. Program ID           : P6110Mb1.asp
'*  4. Program Name         : 금형제원정보조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005-01-25
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Lee Sang Ho
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter 선언 
Dim strQryMode
Dim lgPageNo, lgMaxCount
Dim lgDataExist
Dim GroupCount
Dim istrData
Dim iLngMaxRow
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Err.Clear																	'☜: Protect system from crashing

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 2)

Dim strCastCd
Dim strCarKind
Dim strSetPlantCd

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)
lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
lgDataExist     = "No"
iLngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

If IsNull(Trim(Request("txtCastCd"))) Or Trim(Request("txtCastCd")) = "" Then
	strCastCd = "%"
Else
	strCastCd = Trim(Request("txtCastCd"))
End If

If IsNull(Trim(Request("txtCarKind"))) Or Trim(Request("txtCarKind")) = "" Then
	strCarKind = "%"
Else
	strCarKind = Trim(Request("txtCarKind"))
End If

If IsNull(Trim(Request("txtSetPlantCd"))) Or Trim(Request("txtSetPlantCd")) = "" Then
	strSetPlantCd = "%"
Else
	strSetPlantCd = Trim(Request("txtSetPlantCd"))
End If


UNISqlId(0) = "Y6110MB101"
UNIValue(0, 0) = FilterVar(Ucase(strSetPlantCd),"''","S")
UNIValue(0, 1) = FilterVar(Ucase(strCarKind),"''","S")
UNIValue(0, 2) = FilterVar(Ucase(strCastCd),"''","S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

Set ADF = Nothing

If  rs0.EOF And rs0.BOF  Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "</Script>"		& vbCr
    Response.end
End If
Call MakeSpreadSheetData()

'-----------------------
'Result data display area
'-----------------------
if GroupCount > 0 then
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "	With parent " & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
	Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr
	Response.Write "	.lgPageNo  = """ & lgPageNo   & """" & vbCr
	
	Response.Write " 	.DbQueryOk "	& vbCr
	Response.Write "	End With "		& vbCr
	Response.Write "</Script>"		& vbCr
End if
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

	Dim iLoopCount
	Dim iRowStr

	lgDataExist    = "Yes"
	If CLng(lgPageNo) > 0 Then
	   rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If
	
	iLoopCount = 0
	
	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAST_CD"))	
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAST_NM"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SET_PLANT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SET_PLANT_NM"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAR_KIND"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAR_KIND_NM"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("MAKE_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("STR_TYPE"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("CHECK_END_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_CD_1"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("CLOSE_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("PIC_FLAG"))
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

	If iLoopCount - 1 < lgMaxCount Then
		istrData = istrData & iRowStr & Chr(11) & Chr(12)
	Else
		lgPageNo = lgPageNo + 1
		Exit Do
	End If
	rs0.MoveNext
Loop

If iLoopCount <= lgMaxCount Then                                      '☜: Check if next data exists
	lgPageNo = ""
End If
GroupCount = iLoopCount
rs0.Close                                                       '☜: Close recordset object
Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
