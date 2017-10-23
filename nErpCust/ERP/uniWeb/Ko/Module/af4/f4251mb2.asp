<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4107mb2
'*  4. Program Name         : 차입금상환내역조회 
'*  5. Program Desc         : Query of Loan Repay
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002.04.12
'*  8. Modified date(Last)  : 2003.05.05
'*  9. Modifier (First)     : Hwang Eun Hee
'* 10. Modifier (Last)      : Ahn do hyun
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  차입금번호 오류 Check
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                         '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const C_SHEETMAXROWS_D = 30

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strLoanNo
Dim strIntDtFr, strIntDtTo
Dim cboConfFg, cboApSts
Dim strWhere
Dim TotalLoanAmt, TotalLoanLocAmt
Dim strMsgCd

Dim  iLoopCount
Dim  LngMaxRow

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    strLoanNo		= Request("txtLoanNo")	
    strIntDtFr		= Request("txtIntDtFr")	
    strIntDtTo		= Request("txtIntDtTo")	
	cboConfFg	= Trim(Request("cboConfFg"))
	cboApSts	= Trim(Request("cboApSts"))

    lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount		= C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList		= Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist		= "No"
	LngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1
    
    Call FixUNISQLData()
    Call QueryData()
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iRowStr

    lgstrData = ""
    lgDataExist    = "Yes"

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(lgMaxCount) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
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

    Redim UNIValue(0,8)

    UNISqlId(0) = "F4251mb02"

	strWhere = ""
	If cboConfFg	= "C" Then strWhere = strWhere & " and conf_fg IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " ) " 
	If cboConfFg	= "U" Then	strWhere = strWhere & " and conf_fg   =  " & FilterVar(cboConfFg , "''", "S") & " " 
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList  
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = Filtervar(UniConvDate(strIntDtFr)	, "", "S")
    UNIValue(0,2) = Filtervar(UniConvDate(strIntDtTo)	, "", "S") 
    UNIValue(0,3) = strWhere
    UNIValue(0,4) = Filtervar(strLoanNo	, "", "S")
    UNIValue(0,5) = Filtervar(UniConvDate(strIntDtFr)	, "", "S") 
    UNIValue(0,6) = Filtervar(UniConvDate(strIntDtTo)	, "", "S") 
    UNIValue(0,7) = Filtervar(strLoanNo	, "", "S")

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim lgstrRetMsg   

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set lgADF = Nothing 
    
    iStr = Split(lgstrRetMsg,gColSep)
    If iStr(0) <> "0" Then		
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If     

    If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Exit sub
	Else   
        Call  MakeSpreadSheetData()
    End If
																'☜: ActiveX Data Factory Object Nothing
End Sub

%>

<Script Language=vbscript>
If "<%=lgDataExist%>" = "Yes" Then
	With parent
		.ggoSpread.Source    = .frm1.vspdData2 
		.frm1.vspdData2.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo_B      =  "<%=lgPageNo%>"								'☜: set next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("B",1),.GetKeyPos("B",2),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("B",1),.GetKeyPos("B",3),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("B",1),.GetKeyPos("B",4),   "A" ,"I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData2,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,.GetKeyPos("B",1),.GetKeyPos("B",5),   "A" ,"I","X","X")
		.DbQueryOk("2")
		.frm1.vspdData2.Redraw = True
	End with
End if
</Script>	
