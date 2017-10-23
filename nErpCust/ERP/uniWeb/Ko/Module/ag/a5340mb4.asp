<%Option Explicit%>
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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
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
    Dim strSql2
    Dim strSql3
	Dim lgtxtTransDt
	Dim lgtxtBizArea
	Dim lgtxtTransType
	Dim lgJnlCd
	Dim lgEventCd
	Dim lgFlag
	Dim bAbNomal
	Dim lgtxtFromReqDt
	Dim lgtxtToReqDt
	Dim lgtxtMaxRows

    
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
'   Call SvrMsgBox("bb" , vbInformation, I_MKSCRIPT)
 
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
	Dim strSqlTemp,strSqlTemp2 

	lgFlag	= Trim(Request("RADIO"))

	lgtxtFromReqDt		= UNIConvDate( Request("txtFromReqDt"))
	lgtxtToReqDt			= UNIConvDate(Request("txtToReqDt"))
	
	lgtxtBizArea		= FilterVar(Trim(Request("txtBizCd")),"","S")
	lgJnlCd	= FilterVar(Trim(Request("JnlCd")),"","S")
	lgtxtTransType	= FilterVar(Trim(Request("txtTransType")),"","S")
	lgEventCd	= FilterVar(Trim(Request("EventCd")),"","S")

	lgFlag	= Trim(Request("Rdodt"))
	lgtxtMaxRows		= Request("txtMaxRows")
	
	IF  Trim(lgEventCd) = "" and Trim(lgJnlCd) = "" Then
		bAbNomal = "Y"
	end if
	
	strSql  =  " AND  A.GL_DT >=   " & FilterVar(lgtxtFromReqDt, "''", "S") & " AND A.GL_DT <=  " & FilterVar( lgtxtToReqDt, "''", "S") & ""
	strSql2 =  " AND  TRANS_DT >=   " & FilterVar(lgtxtFromReqDt, "''", "S") & " AND TRANS_DT <=  " & FilterVar( lgtxtToReqDt, "''", "S") & ""
	strSql3  =  " AND  A.GL_DT >=   " & FilterVar(lgtxtFromReqDt, "''", "S") & " AND A.GL_DT <=  " & FilterVar( lgtxtToReqDt, "''", "S") & ""
	
	strSqlTemp  =  strSqlTemp & "  AND A.BIZ_AREA_CD  =  " & FilterVar(lgtxtBizArea , "''", "S") & " "
	strSqlTemp  = strSqlTemp & " AND TRANS_TYPE =  " & FilterVar(lgtxtTransType , "''", "S") & " "
	
	If bAbNomal <> "Y" Then
		strSqlTemp  =  strSqlTemp & "  AND A.JNL_CD  =  " & FilterVar(lgJnlCd , "''", "S") & " "
		strSqlTemp  =  strSqlTemp & "  AND A.EVENT_CD  =  " & FilterVar(lgEventCd , "''", "S") & " "
	End If
	
'	Call SvrMsgBox(lgEventCd , vbInformation, I_MKSCRIPT)
'	Call SvrMsgBox(lgJnlCd , vbInformation, I_MKSCRIPT)

'	Call SvrMsgBox(bAbNomal , vbInformation, I_MKSCRIPT)
	strSql = strSql & strSqlTemp
	strSql2 = strSql2 & strSqlTemp
	strSql3 =  strSql3 & "  AND DBO.UFN_A_GETTRANSTYPE(A.BATCH_NO, A.BIZ_AREA_CD)= " & FilterVar(lgtxtTransType, "''", "S") & " "
	strSql3 = strSql3 & "  AND A.BIZ_AREA_CD  =  " & FilterVar(lgtxtBizArea , "''", "S") & " "
	
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 100     
    
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

	IF  bAbNomal = "Y"  Then
	    UNISqlId(0) = "A5340MA106"
	else
	    UNISqlId(0) = "A5340MA105"
	end if
	

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList          
'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)
                                '☜: Select list
	IF  bAbNomal = "Y"  Then
		UNIValue(0,2)  = strSql3	
		UNIValue(0,1)  = strSql2	

	Else
	  	UNIValue(0,1)  = strSql	'UNIConvDate(Request("txtFromGlDt") )
	
	End if
'	UNIValue(0,2)  = strSql2
'	UNIValue(0,3)  = strSql3

'    UNIValue(0,0) = strSql          
'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)
                                '☜: Select list
'  	UNIValue(0,1)  = strSql2	'UNIConvDate(Request("txtFromGlDt") )
'	UNIValue(0,2)  = strSql
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
 '   UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
 
 
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
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData1
       Parent.frm1.vspdData1.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                  '☜ : Display data       
       Parent.lgPageNo_B      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.frm1.vspdData1.Redraw = True
    End If   

</Script>	

