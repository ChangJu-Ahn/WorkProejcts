
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim iPrevEndRow
Dim iEndRow	

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
	Dim strFrApDt	                                                           
	Dim strToApDt
	Dim strFrApNo	                                                           
	Dim strToApNo
	Dim strdeptcd
	Dim strBpCd
	Dim strCond

	Dim strMsgCd
	Dim strMsg1
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
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
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End if
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(1,2)

    UNISqlId(0) = "A7101RA1"
    UNISqlId(1) = "COMMONQRY"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    
	UNIValue(1,0) = "SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD =  " & FilterVar(strBpCd, "''", "S") & " "
    'UNIValue(1,0) = UCase("'" & FilterVar(Trim(strBpCd),"","S") & "'" )
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim strMsg
    Dim strMsgCd
    
    strMsg = Trim(Request("txtBpcd_Alt"))
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    Set lgADF = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    IF NOT (rs1.EOF or rs1.BOF) then
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = "<%=rs1(0)%>"   
		End With
		</Script>
<%			
	ELSE
		if Trim(Request("txtBpCd")) <> "" Then
			strMsgCd = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%	
		Else 
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%			
		End if
	End if
    rs1.Close
    Set rs1 = Nothing 
    
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strFrPrNo	   = Request("txtFrPrNo")                                                          
     strToPrNo     = Request("txtToPrNo")     
     strFrPrDt     = UNIConvDate(Request("txtFrPrDt"))
     strToPrDt     = UNIConvDate(Request("txtToPrDt"))
     strBpCd	   = Request("txtBpCd")     

     If strFrPrNo <> "" Then
		strCond = strCond & " and A.PRRCPT_NO >=  " & FilterVar(strFrPrNo, "''", "S") & " "	 
     End If
     
     If strToPrNo <> "" Then
		strCond = strCond & " and A.PRRCPT_NO <=  " & FilterVar(strToPrNo, "''", "S") & " "
     End If
         
     If strToPrDt <> "" Then
		strCond = strCond & " and A.PRRCPT_DT <=  " & FilterVar(strToPrDt , "''", "S") & ""
     End If
     
     If strFrPrDt <> "" Then
		strCond = strCond & " and A.PRRCPT_DT >=  " & FilterVar(strFrPrDt , "''", "S") & ""
     End If  
     
     If strBpCd <> "" Then
		strCond = strCond & " and A.BP_CD =  " & FilterVar(strBpCd, "''", "S") & " "
     End If
     
     strCond = strCond & " and A.PRRCPT_FG = " & FilterVar("CT", "''", "S") & " "
     
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtFrPrDt.Value      = Parent.Frm1.txtFrPrDt.Text
          Parent.Frm1.htxtToPrDt.Value         = Parent.Frm1.txtToPrDt.Text
          Parent.Frm1.htxtFrPrNo.Value  = Parent.Frm1.txtFrPrNo.Value
          Parent.Frm1.htxtToPrNo.Value    = Parent.Frm1.txtToPrNo.Value
          Parent.Frm1.htxtBpCd.Value    = Parent.Frm1.txtBpCd.Value
       End If
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
		Parent.frm1.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",4),"A", "Q" ,"X","X")
		Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   
</Script>	



