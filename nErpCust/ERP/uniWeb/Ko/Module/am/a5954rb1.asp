<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 

On Error Resume Next
Err.Clear

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0	                           '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strDB
Dim strCond
Dim strStdDt, strStdYYMM
Dim strBankCd
Dim strPgmId
Dim strMsgCd, strMsg1, strMsg2
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")

Const C_SHEETMAXROWS_D  = 100

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
lgDataExist    = "No"
strStdDt	= UniConvDate(Request("txtStdDt"))	
strStdYYMM	= Request("txtStdYYMM")    
strPgmId	= Request("txtPgmId")
    
Call TrimData()
Call FixUNISQLData()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Trim Data
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
'on Error Resume Next

Dim strStdDtYear, strStdDtMonth
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strCond = ""
	If strPgmId = "F5954RA101" Then 
		strCond = strCond & " apprl_yrmnth = " & FilterVar(strStdYYMM, "''", "S") 
		strDB = "B_MONTHLY_EXCHANGE_RATE A"
	ElseIF strPgmId = "F5954RA102" Then 
		strCond = strCond & " apprl_dt = " & FilterVar(UNIConvDate(strStdDt), "''", "S") 
		strDB = "B_DAILY_EXCHANGE_RATE A"
	End If	
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,3)

    UNISqlId(0) = "F5954RA101"
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strDB
    UNIValue(0,2) = strCond   
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode


End Sub    

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim lgADF
    Dim iStr
    Dim lgstrRetMsg   
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(iStr(1), vbInformation, I_MKSCRIPT)
    End If    

	If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' MakeSpreadSheetData
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim	 YYYYMM
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
     
    'rs0에 대한 결과 
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1		
		    if ColCnt = 2 And Len(Trim(rs0(ColCnt))) = 6 Then      
				YYYYMM = UNIConvYYYYMMDDToDate(gServerDateFormat,Mid(rs0(ColCnt),1,4),Mid(rs0(ColCnt),5,2),"01")
				iRowStr = iRowStr & Chr(11) & UNIMonthClientFormat(YYYYMM)
		    else
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		    end if
		    
		Next
'		Response.Write  iRowStr
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
                    '☜: ActiveX Data Factory Object Nothing
End Sub

%>

<Script Language=vbscript>
    With parent
		If "<%=lgDataExist%>" = "Yes" Then
			 If "<%=lgPageNo%>" = "1" Then          
				.frm1.hStdDt.value = "<%=strStdt%>"
				.frm1.hStdYYMM.value = "<%=strStYYMM%>"
			 End if
				.ggoSpread.Source  = Parent.frm1.vspdData
         		.ggoSpread.SSShowData "<%=lgstrData%>"                           '☜: Display data 
			    .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
				.DbQueryOk
		End If	
	End with
</Script>
