<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

'**********************************************************************************************
'*  1. Module Name          : MOLD management 
'*  2. Function Name        :
'*  3. Program ID           : P6310MB1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/02/21
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Lee Sang-Ho
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*
'*
'*
'*
'* 14. Business Logic of P6310MA1(금형별실적조회)
'**********************************************************************************************
Dim lgOpModeCRUD

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData1
Dim istrData2
Dim istrData3
Dim iStrPoNo
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim iLngMaxRow1		' 현재 그리드의 최대Row
Dim iLngMaxRow2		' 현재 그리드의 최대Row
Dim iLngMaxRow3		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
Dim lgDataExist
Dim lgPageNo_A
Dim lgPageNo_B
Dim lgMaxCount
Dim strFlag

	Const C_SHEETMAXROWS_D  = 100

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    lgOpModeCRUD  = Request("txtMode")

										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	
    lgMaxCount       = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
	lgDataExist      = "No"

	lgStrPrevKey = Request("lgStrPrevKey")

	Call FixUNISQLData()
	
	If Request("txtType") = "A" Then									
		lgPageNo_A       = UNICInt(Trim(Request("lgPageNo_A")),0)   
		iLngMaxRow1	     = CDbl(lgMaxCount) * CDbl(lgPageNo_A) + 1
		Call QueryData()
		
		Response.Write "<Script Language=vbscript>"									& vbCr
		Response.Write "With parent"												& vbCr
		Response.Write "	If .frm1.vspdData.MaxRows < 1 then"						& vbCr
		Response.Write "	End if"													& vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData "				& vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & istrData1	 & """"		& vbCr
		Response.Write "	.lgPageNo_A  = """ & lgPageNo_A   & """"				& vbCr
		Response.Write "	.DbQueryOk "											& vbCr
		Response.Write "End With"													& vbCr
		Response.Write "</Script>"													& vbCr
	Elseif Request("txtType") = "B" Then
		lgPageNo_B       = UNICInt(Trim(Request("lgPageNo_B")),0) 
		iLngMaxRow1	     = CDbl(lgMaxCount) * CDbl(lgPageNo_B) + 1
		Call QueryData()
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData1.MaxRows < 1 then"					& vbCr
		Response.Write "	End if"													& vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData1 "				& vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & istrData2	 & """"		& vbCr
		Response.Write "	.lgPageNo_B  = """ & lgPageNo_B   & """"				& vbCr
		Response.Write "	.DbDtlQueryOk1 "										& vbCr
		Response.Write "End With"													& vbCr
		Response.Write "</Script>"													& vbCr
	End if
	
	
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData

	If Request("txtType") = "A" Then
	
		Dim strCast_Cd, strProd_Item_Cd
		Dim strWork_Dt, strPlantCd
		Dim strProd_Dt_Fr, strProd_Dt_To
		Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
		Redim UNIValue(0, 2)

		UNISqlId(0) = "Y6310MB01"
	
		StrPlantCd = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
		
		IF Request("txtCast_Cd") = "" Then
		   strCast_Cd = FilterVar("%", "''", "S")
		ELSE
		   strCast_Cd = FilterVar(Ucase(Trim(Request("txtCast_Cd"))),"''","S")
		END IF

		IF Request("txtProd_Item_Cd") = "" Then
		   strProd_Item_Cd = FilterVar("%", "''", "S")
		ELSE
		   strProd_Item_Cd = FilterVar(Ucase(Trim(Request("txtProd_Item_Cd"))),"''","S")
		END IF

		UNIValue(0, 0) = strPlantCd
		UNIValue(0, 1) = strCast_Cd
		UNIValue(0, 2) = strProd_Item_Cd

	Else
	
		Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
		Redim UNIValue(0, 2)

		
		UNISqlId(0) = "Y6310MB02"
		
		UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtCastCd1"))),"''","S")
		
		If Trim(Request("txtProd_Dt_Fr")) = "" Then
			UNIValue(0, 1) = FilterVar("1900-01-01", "''", "S")
		Else
			UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtProd_Dt_Fr"))),"''","S")
		End If
		
		If Trim(Request("txtProd_Dt_To")) = "" Then
			UNIValue(0, 2) = FilterVar("2999-12-31", "''", "S")
		Else
			UNIValue(0, 2) = FilterVar(Ucase(Trim(Request("txtProd_Dt_To"))),"''","S")
		End If
		
	End If
	UNILock = DISCONNREAD :	UNIFlag = "1"

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData1()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr


	if Request("txtType") = "A" Then							'☜ : 디테일 검색 
	    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

		Set lgADF   = Nothing
	
	    iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If
	    
	    If  rs0.EOF And rs0.BOF  Then
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	        Response.Write "<Script Language=vbscript>" & vbCr
	        Response.Write "</Script>"		& vbCr
	        Response.end
	    Else
	        Call  MakeSpreadSheetData1()
	    End If
	    
	Elseif Request("txtType") = "B" Then

	    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

		Set lgADF   = Nothing
	
	    iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If

	    If  rs0.EOF And rs0.BOF  Then
 			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	        Response.Write "<Script Language=vbscript>" & vbCr
	        Response.Write "</Script>"		& vbCr
	        Response.end
	    Else
	        Call  MakeSpreadSheetData2()
	    End If

	End If

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData1()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData1()

    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt
    lgDataExist    = "Yes"
    If CLng(lgPageNo_A) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo_A)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
	
	iLoopCount = 0
	Do while Not (rs0.EOF Or rs0.BOF)
		
        iLoopCount =  iLoopCount + 1
        
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAST_CD"))
	    iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAST_NM"))
	    iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SET_PLACE"))
	    iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SET_PLACE_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAR_KIND"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CAR_KIND_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_CD_1"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CUR_ACCNT"))        
        iRowStr = iRowStr & Chr(11) & iLngMaxRow1 + iLoopCount
		
        If iLoopCount - 1 < lgMaxCount Then
           istrData1 = istrData1 & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo_A = lgPageNo_A + 1
           Exit Do
        End If
        rs0.MoveNext
	Loop

    If iLoopCount <= lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo_A = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData2()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData2()

    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt

    lgDataExist    = "Yes"
    If CLng(lgPageNo_B) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo_B)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

	iLoopCount = 0
	Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
   
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("REPORT_DT"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("PRODT_ORDER_NO"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("OPR_NO"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SEQ"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_CD"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
        iRowStr = iRowStr & Chr(11) & UniConvNumberDBToCompany(rs0("CAST_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("REPORT_TYPE"))
        iRowStr = iRowStr & Chr(11) & UniConvNumberDBToCompany(rs0("PROD_QTY_IN_ORDER_UNIT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("PRODT_ORDER_UNIT"))
        iRowStr = iRowStr & Chr(11) & iLngMaxRow2 + iLoopCount

        If iLoopCount - 1 < lgMaxCount Then
           istrData2 = istrData2 & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo_B = lgPageNo_B + 1
           Exit Do
        End If

        rs0.MoveNext
	Loop
    If iLoopCount <= lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo_B = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

'---------- Developer Coding part (Start) ---------------------------------------------------------------
'A developer must define field to create record
'--------------------------------------------------------------------------------------------------------

'---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

'---------- Developer Coding part (Start) ---------------------------------------------------------------
'A developer must define field to update record
'--------------------------------------------------------------------------------------------------------

'---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)

End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()

End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()

End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()

End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

End Sub
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)

End Function

%>