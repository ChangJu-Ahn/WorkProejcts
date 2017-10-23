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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3                              '☜ : DBAgent Parameter 선언 
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
   	Dim lgtxtBizArea
	Dim lgtxtTransType
	Dim lgtxtGlInputType
	Dim lgtxtMaxRows

    Dim BIZ_AREA_NM
    Dim TRANS_NM
    Dim GlInputType_Nm

    
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
	Dim lgtxtFromReqDt
	Dim lgtxtToReqDt
	Dim lgFlag
	Dim strSqlTemp

	lgFlag	= Trim(Request("RADIO"))

	lgtxtFromReqDt		= UNIConvDate( Request("txtFromReqDt"))

	lgtxtToReqDt			= UNIConvDate(Request("txtToReqDt"))
	lgtxtBizArea		= FilterVar(Trim(Request("txtBizCd")),"","S")
	lgtxtTransType	= FilterVar(Trim(Request("txtTransType")),"","S")
	lgtxtGlInputType	= FilterVar(UCase(Trim(Request("txtGlInputType"))),"","S")

	lgFlag	= Trim(Request("Rdodt"))
	lgtxtMaxRows		= Request("txtMaxRows")
	
	strSqlTemp = " "
	strSql  =  " AND  A.GL_DT >=   " & FilterVar(lgtxtFromReqDt, "''", "S") & " AND A.GL_DT <=  " & FilterVar( lgtxtToReqDt, "''", "S") & ""
	strSql2 =  " AND  TRANS_DT >=   " & FilterVar(lgtxtFromReqDt, "''", "S") & " AND TRANS_DT <=  " & FilterVar( lgtxtToReqDt, "''", "S") & ""
	strSql3 =  " AND  A.GL_DT >=   " & FilterVar(lgtxtFromReqDt, "''", "S") & " AND A.GL_DT <=  " & FilterVar( lgtxtToReqDt, "''", "S") & ""

	IF lgtxtBizArea <> "" then		 
		strSql  =  strSql & " AND A.BIZ_AREA_CD  =  " & FilterVar(lgtxtBizArea , "''", "S") & " "
		strSql3 = strSql3 & " AND A.BIZ_AREA_CD  =  " & FilterVar(lgtxtBizArea , "''", "S") & " "
	end if	  

	If lgtxtGlInputType<> "" Then
		strSql  =  strSql  & " AND A.GL_INPUT_TYPE  =  " & FilterVar(lgtxtGlInputType , "''", "S") & " "
		strSql3 =  strSql3 & " AND A.GL_INPUT_TYPE  =  " & FilterVar(lgtxtGlInputType , "''", "S") & " "
	end if

	IF lgtxtTransType <> "" then		 
		strSql  =  strSql  & " AND A.TRANS_TYPE =  " & FilterVar(lgtxtTransType , "''", "S") & " "
		strSql2 =  strSql2 & " AND TRANS_TYPE =  " & FilterVar(lgtxtTransType , "''", "S") & " "
		strSql3 =  strSql3 & " AND  DBO.UFN_A_GETTRANSTYPE(A.BATCH_NO, A.BIZ_AREA_CD)=  " & FilterVar(lgtxtTransType , "''", "S") & " "
	end if

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

    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,4)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A5340MA102"
    UNISqlId(1) = "COMMONQRY"					'BizNm
    UNISqlId(2) = "COMMONQRY"					'TransTypeNm
    UNISqlId(3) = "COMMONQRY"					'GlInputTypeNm

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList          
'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)
                                '☜: Select list
  	UNIValue(0,1)  = strSql	'UNIConvDate(Request("txtFromGlDt") )
	UNIValue(0,2)  = strSql2
	UNIValue(0,3)  = strSql3

'    UNIValue(0,0) = strSql          
'    Call SvrMsgBox(strSql , vbInformation, I_MKSCRIPT)
                                '☜: Select list
'  	UNIValue(0,1)  = strSql2	'UNIConvDate(Request("txtFromGlDt") )
'	UNIValue(0,2)  = strSql

	UNIValue(1,0)  = "SELECT BIZ_AREA_NM FROM B_BIZ_AREA WHERE BIZ_AREA_CD =  " & FilterVar( lgtxtBizArea , "''", "S") & ""		'rs1에 대한 Value값 setting
	UNIValue(2,0)  = "SELECT TRANS_NM FROM A_ACCT_TRANS_TYPE WHERE TRANS_TYPE =  " & FilterVar( lgtxtTransType , "''", "S") & ""		'rs1에 대한 Value값 setting
	UNIValue(3,0)  = "SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  and MINOR_CD =  " & FilterVar( lgtxtGlInputType , "''", "S") & ""		
	
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If  
    
   IF NOT (rs1.EOF or rs1.BOF) then
	    BIZ_AREA_NM = rs1("BIZ_AREA_NM")
	ELSE
		BIZ_AREA_NM=""
    END IF
    rs1.Close
    Set rs1 = Nothing
    
    IF NOT (rs2.EOF or rs2.BOF) then
	    TRANS_NM = rs2("TRANS_NM")
	ELSE
		TRANS_NM=""
    END IF
    rs2.Close
    Set rs2 = Nothing
    
    IF NOT (rs3.EOF or rs3.BOF) then
	    GlInputType_Nm = rs3("MINOR_NM")
	ELSE
		GlInputType_Nm=""
    END IF
    rs3.Close
    Set rs3 = Nothing

      
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
          Parent.Frm1.hBizCd.Value      = Parent.Frm1.txtBizCd.Value                  'For Next Search
          Parent.Frm1.hRdoDt.Value         = Parent.Frm1.RdoDt.Value
          Parent.Frm1.hTransType.Value  = Parent.Frm1.txtTransType.value
          Parent.Frm1.hFromReqDt.Value    = Parent.Frm1.txtFromReqDt.Text
          Parent.Frm1.hToReqDt.Value      = Parent.Frm1.txtToReqDt.Text
          parent.Frm1.hGlInputType.value  = Parent.Frm1.txtGlInputType.Value
       End If

       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
'redraw true, false 반드시 해야함              
       Parent.frm1.vspdData.Redraw = False
' showdata 마지막 인자 "F"를 반드시 줘야함       
       Parent.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                  '☜ : Display data
' 시작 row넘버와 마지막 row넘버를 반드시 지정해야 함 
       
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
'redraw true, false 반드시 해야함       
       Parent.frm1.vspdData.Redraw = True
    End If
    parent.frm1.txtBizNm.value = "<%=ConvSPChars(BIZ_AREA_NM)%>"  
	parent.frm1.txtTransTypeNm.value = "<%=ConvSPChars(TRANS_NM)%>"  
	parent.frm1.txtGlInputTypeNm.value = "<%=ConvSPChars(GlInputType_Nm)%>"     

</Script>	

