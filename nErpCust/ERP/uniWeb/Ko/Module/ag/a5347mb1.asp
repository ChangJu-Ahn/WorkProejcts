<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call loadInfTB19029B("Q", "*","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "*", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
    Dim lgstrData                                                              '☜ : data for spreadsheet data
    Dim lgStrPrevKey                                                           '☜ : 이전 값 
    Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgDataExist
    Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
    Dim iPrevEndRow
    Dim iEndRow
    Dim strSql
    
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
 
    Call HideStatusWnd


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
Sub TrimData()
	Dim strtxtFromReqDt
	Dim strtxtToReqDt
	Dim strtxtBizArea
	Dim strtxtTransType
	Dim strFlag
	Dim strSqlTemp,strSqlTemp2 ,strSqlTemp3
	Dim strtxtGlInputType
	Dim strtxtfrBatchNo	
	Dim strtxtToBatchNo	
	Dim strtxtfrRefNo	
	Dim strtxtToRefNo
	Dim rdoDiff

	strFlag	= Trim(Request("RADIO"))

	strtxtFromReqDt	= UNIConvDate( Request("txtFromReqDt"))

	strtxtToReqDt	= UNIConvDate(Request("txtToReqDt"))
	strtxtBizArea	= FilterVar(Request("txtBizCd"),"","S")
	strtxtTransType	= FilterVar(Request("txtTransType"),"","S")
	strtxtGlInputType	= FilterVar(Request("txtGlInputType"),"","S")
	strtxtfrBatchNo	= FilterVar(Request("txtfrBatchNo"),"","S")
	strtxtToBatchNo	= FilterVar(Request("txtToBatchNo"),"","S")

	strtxtfrRefNo	= FilterVar(Request("txtfrRefNo"),"","S")
	strtxtToRefNo	= FilterVar(Request("txtToRefNo"),"","S")

	strFlag	= Trim(Request("RdoDispType"))
	strtxtMaxRows		= Request("txtMaxRows")
	rdoDiff = Request("rdoDiff")
	
	strSql =  " AND  A.GL_DT >=  " & FilterVar(strtxtFromReqDt , "''", "S") & " AND A.GL_DT <=  " & FilterVar(strtxtToReqDt , "''", "S") & "" & vbcr
	IF strtxtBizArea <> "" then			strSql  = strSql & "  AND  A.BIZ_AREA_CD  =  " & FilterVar(strtxtBizArea , "''", "S") & " " & vbcr
	IF strtxtTransType <> "" then		strSql  = strSql & "  AND  DBO.UFN_A_GETTRANSTYPE(A.BATCH_NO, A.BIZ_AREA_CD)=  " & FilterVar(strtxtTransType , "''", "S") & " " & vbcr
	IF strtxtGlInputType <> "" then		strSql  = strSql & " AND A.GL_INPUT_TYPE =  " & FilterVar(strtxtGlInputType , "''", "S") & " " & vbcr
	IF strtxtfrBatchNo <> "" then		strSql  = strSql & " AND A.BATCH_NO >=  " & FilterVar(strtxtfrBatchNo , "''", "S") & " " & vbcr
	IF strtxtToBatchNo <> "" then		strSql  = strSql & " AND A.BATCH_NO <=  " & FilterVar(strtxtToBatchNo , "''", "S") & " " & vbcr
	IF strtxtfrRefNo <> "" then			strSql  = strSql & " AND A.REF_NO >=  " & FilterVar(strtxtfrRefNo , "''", "S") & " " & vbcr
	IF strtxtToRefNo <> "" then			strSql  = strSql & " AND A.REF_NO <=  " & FilterVar(strtxtfrRefNo , "''", "S") & " " & vbcr

	If rdoDiff	 = "True" Then
		strSql  = strSql & " AND A.CHAIN_NO IN ( " & vbcr
		strSql  = strSql & " SELECT UN.CHAIN_NO FROM ( " & vbcr
		strSql  = strSql & " 	SELECT A.CHAIN_NO, A.ACCT_CD, A.DR_CR_FG, (A.ITEM_LOC_AMT) BATCHAMT, 0 GLAMT FROM AV_BATCH_TOT_POST A , A_ACCT B " & vbcr
		strSql  = strSql & " 	WHERE A.ACCT_CD = B.ACCT_CD  AND B.HQ_BRCH_FG <> " & FilterVar("Y", "''", "S") & "  AND B.ACCT_TYPE NOT IN (" & FilterVar("XP", "''", "S") & "," & FilterVar("XL", "''", "S") & ") " & vbcr
		strSql  = strSql & " 	UNION ALL " & vbcr
		strSql  = strSql & " 	SELECT A.REF_NO AS CHAIN_NO, A.ACCT_CD, A.DR_CR_FG, 0, A.ITEM_LOC_AMT FROM AV_GL_TEMP A, A_ACCT B " & vbcr
		strSql  = strSql & " 	WHERE A.ACCT_CD = B.ACCT_CD  AND B.HQ_BRCH_FG <> " & FilterVar("Y", "''", "S") & "  AND B.ACCT_TYPE NOT IN (" & FilterVar("XP", "''", "S") & "," & FilterVar("XL", "''", "S") & ") " & vbcr
		strSql  = strSql & " 	) UN " & vbcr
		strSql  = strSql & "  GROUP BY UN.CHAIN_NO, UN.ACCT_CD, UN.DR_CR_FG " & vbcr
		strSql  = strSql & " HAVING SUM(UN.BATCHAMT) <> SUM(UN.GLAMT) " & vbcr
		strSql  = strSql & " ) " & vbcr
	End If
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 30     
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo

    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1

            Exit Do
        End If
        rs0.MoveNext
	Loop

	
    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,3)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A5347MA101"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList          
'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)
                                '☜: Select list
  	UNIValue(0,1)  = strSql	'UNIConvDate(Request("txtFromGlDt") )
'	UNIValue(0,2)  = strSql2
'	UNIValue(0,3)  = strSql3

'    UNIValue(0,0) = strSql
'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)

'  	UNIValue(0,1)  = strSql2	'UNIConvDate(Request("txtFromGlDt") )
'	UNIValue(0,2)  = strSql
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtBizCd.Value      = Parent.Frm1.txtBizCd.Value                  'For Next Search
'          Parent.Frm1.hRadio.Value         = Parent.Frm1.RdoDiff.Value
          Parent.Frm1.htxtTransType.Value  = Parent.Frm1.txtTransType.value
          Parent.Frm1.hFromReqDt.Value    = Parent.Frm1.txtFromReqDt.Text
          Parent.Frm1.hToReqDt.Value      = Parent.Frm1.txtToReqDt.Value                  'For Next Search
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                  '☜ : Display data       
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   

</Script>	

