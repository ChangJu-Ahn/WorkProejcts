<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111qb1
'*  4. Program Name         : 멀티컴퍼니수주조회 
'*  5. Program Desc         : 멀티컴퍼니수주조회-멀티 
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/24
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
Dim lgDataExist
Dim lgPageNo


Dim SupplierNM			'☜ : 발주법인 

Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows

Dim istr
Dim istrYN

intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Dim strSpread																'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 



Call HideStatusWnd                                                               '☜: Hide Processing message

lgOpModeCRUD  = Request("txtMode")

'response.write lgOpModeCRUD & lgOpModeCRUD &"<br>"
'response.write UID_M0002 & UID_M0002 &"<br>"

'response.end


Select Case lgOpModeCRUD
	Case CStr(UID_M0001)                                                         '☜: Query
		Call  SubBizQueryMulti()
End Select

Sub SubBizQueryMulti()
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set

	Call QueryData()			'☜ : DB-Agent를 통한 ADO query

	'-----------------------
	'Result data display area
	'-----------------------
%>
	<Script Language=vbscript>
		With parent
			.frm1.txtSupplierNM.Value	= "<%=SupplierNM%>"

			.frm1.hdnItem.value = "<%=ConvSPChars(Request("txtitemcd"))%>"

			.frm1.txtPO_NO.focus
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then

				'Show multi spreadsheet data from this line

				.ggoSpread.Source    = .frm1.vspdData
				.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data

				.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag

				.DbQueryOk <%=intARows%>,<%=intTRows%>

			End If
		End with
	</Script>
<%
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear
	Dim iErrorPosition
	Dim LngMaxRow
	Dim arrTemp
	Dim arrVal
	Dim lGrpCnt
	Dim LngRow
	Dim iRow_cnt

	Dim iPO_COMPANY
	Dim iSO_COMPANY
	Dim iPO_NO
	Dim ObjPSMG111


	LngMaxRow = CInt(Request("txtMaxRows"))								'☜: 최대 업데이트된 갯수 
	arrTemp = Split(Request("txtSpread"), gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 

	lGrpCnt = 0

	Set ObjPSMG111 = Server.CreateObject ("PSMG111.CMaintMcCustPoSoSvr")

	If CheckSYSTEMError(Err,True) = true then
		Set ObjPSMG111 = Nothing
		Exit Sub
	End If

    For LngRow = 1 To LngMaxRow
		Err.Clear
		lGrpCnt = lGrpCnt 														'☜: Group Count

		arrVal = Split(arrTemp(LngRow-1), gColSep)

		iPO_COMPANY	= arrVal(2)
		iSO_COMPANY	= arrVal(3)
		iPO_NO 		= arrVal(4)

		Call ObjPSMG111.S_UPDATE_MC_PO_STS_SOMK(gStrGlobalCollection,iPO_COMPANY,iSO_COMPANY,iPO_NO,iErrorPosition)

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If CheckSYSTEMError2(Err, True, LngRow & "행:", "", "", "", "") = True Then
		    	Err.Clear
			'처리가 완료된것은 Check Box 가 풀림.
			Response.Write "<Script language=vbscript> "		& vbCr
			Response.Write "	Dim iBln "				& vbCr
			Response.Write "            iBln = MsgBox (""계속진행하시겠습니까?"", vbYesNo, """") "				& vbCr
			Response.Write "            If iBln = vbNo Then   "				& vbCr
			Response.Write "	       Parent.DbSaveOk    "				& vbCr
			Response.Write "	    End If"						& vbCr
			Response.Write "</Script> "
		Else
			'처리가 완료된것은 Check Box 가 풀림.
			Response.Write "<Script language=vbscript> "		& vbCr
			Response.Write "On error resume Next"				& vbCr
			Response.Write "	with Parent.frm1.vspdData"      & vbCr
			Response.Write "		Dim iIndex, iRowNo	"		& vbCr
			Response.Write "		for iIndex = 1 to .MaxRows	"      & vbCr
			Response.Write "			.Col = Parent.C_PO_NO	"      & vbCr
			Response.Write "			.Row = iIndex	"		& vbCr
			Response.Write "			If Trim(.text) = """	&  iPO_NO & """ then "     & vbCr
			Response.Write "				iRowNo = iIndex	"   & vbCr
			Response.Write "			End if	"				& vbCr
			Response.Write "		Next	"					& vbCr
			Response.Write "		.Col = parent.C_CfmFlg	"   & vbCr
			Response.Write "		.Row = iRowNo "				& vbCr
			Response.Write "		.Text = 0 "					& vbCr
			Response.Write "	end with "						& vbCr
			Response.Write "</Script> "

		End If

	Next

	If NOT(ObjPSMG111 is Nothing) Then
		Set ObjPSMG111 = Nothing
	End If

    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "


End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt

	Const C_SOLD_TO_PARTY		= 0	'발주법인 
	Const C_BP_FULL_NM		= 1	'발주법인명 
	Const C_CUST_PO_NO		= 2	'고객발주번호 
	Const C_SO_NO			= 3	'수주번호 
	Const C_EXPORT_FLAG		= 4	'내외자구분 
	Const C_CFM_FLAG		= 5	'수주확정여부 
	Const C_SO_DT			= 6	'수주일 
	Const C_SALES_GRP		= 7	'영업그룹 
	Const C_SALES_GRP_FULL_NM	= 8	'영업그룹명 
	Const C_CUR			= 9		'화페 
	Const C_NET_AMT			= 10		'수주금액 
	Const C_VAT_AMT			= 11		'부가세금액 
	Const C_NET_VAT_TOTAMT		= 12		'수주총금액 
	Const C_VAT_TYPE		= 13		'부가세유형 
	Const C_VAT_TYPE_NM		= 14		'부가세유형명 
	Const C_VAT_RATE		= 15		'부가세율 
	Const C_PAY_METH		= 16		'결제방법 
	Const C_PAY_METH_NM		= 17		'결제방법명 
	Const C_INCOTERMS		= 18		'가격조건 
	Const C_INCOTERMS_NM		= 19		'가격조건명 

	lgDataExist    = "Yes"

	If CLng(lgPageNo) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		intTRows	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
	End If

	'----- 레코드셋 칼럼 순서 ----------
	'-----------------------------------
	iLoopCount = 0

    	ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		If ConvSPChars(rs0(C_EXPORT_FLAG))="N" Then
			istr = "내자"
		Else
			istr = "외자"
		End If

		If ConvSPChars(rs0(C_CFM_FLAG))="Y" Then
			istrYN = "확정"
		Else
			istrYN = "미확정"
		End If

		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SOLD_TO_PARTY))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_FULL_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_CUST_PO_NO))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_NO))
		iRowStr = iRowStr & Chr(11) & istr
		iRowStr = iRowStr & Chr(11) & istrYN
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(C_SO_DT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SALES_GRP))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SALES_GRP_FULL_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_CUR))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_NET_AMT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_AMT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_NET_VAT_TOTAMT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_TYPE))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_TYPE_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_RATE))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PAY_METH))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PAY_METH_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_INCOTERMS))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_INCOTERMS_NM))

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)

        	Else
	   	   istrData = ""
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop


	intARows = iLoopCount
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
	  lgPageNo = ""
	End If

	rs0.Close                                                       '☜: Close recordset object
	Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next

    SetConditionData = false

	If Not(rs1.EOF Or rs1.BOF) Then
		SupplierNM = rs1("BP_FULL_NM")
		Set rs1 = Nothing

	Else
		Set rs1 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "발주법인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		    Exit Function
		End If
	End If

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
   	Dim strVal
	ReDim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Redim UNIValue(1,7)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
	strVal = ""
	UNISqlId(0) = "SM111QA1"		'상의Splead Query
	UNISqlId(1) = "SM111MA101"		'발주법인 PopUp

	UNIValue(1,0) = "'zzzz'"

	'//UNIValue(0,0) = "^"

	'발주법인 
	If Trim(Request("txtSupplierCd")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
		UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,0) = "|"

	End If

	'고객발주일(From)
	If Trim(Request("txtFrDt")) <> "" Then
		UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("txtFrDt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,1) = "|"
	End If

	'고객발주일(To)
	If Trim(Request("txtToDt")) <> "" Then
		UNIValue(0,2) = " '"& FilterVar(Trim(UCase(Request("txtToDt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,2) = "|"
	End If

	'수주일(From)
	If Trim(Request("txtSo_Frdt")) <> "" Then
		UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("txtSo_Frdt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,3) = "|"
	End If

	'수주일(To)
	If Trim(Request("txtSo_Todt")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("txtSo_Todt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,4) = "|"
	End If

	'수주확정여부 
	If Trim(Request("rdoCfmFlag")) <> "" Then
		UNIValue(0,5) = " '"& FilterVar(Trim(UCase(Request("rdoCfmFlag"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,5) = "|"
	End If

	'고객발주번호 
	If Trim(Request("txtPO_NO")) <> "" Then
		UNIValue(0,6) = " '"& FilterVar(Trim(UCase(Request("txtPO_NO"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,6) = "|"
	End If

	'수주번호 
	If Trim(Request("txtSO_NO")) <> "" Then
		UNIValue(0,7) = " '"& FilterVar(Trim(UCase(Request("txtSO_NO"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,7) = "|"
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	On Error Resume Next
	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
	Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	'팝업필드 체크 
	If Setconditiondata = False Then Exit Sub

	If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
	Else

		Call  MakeSpreadSheetData()

	End If
End Sub



%>
