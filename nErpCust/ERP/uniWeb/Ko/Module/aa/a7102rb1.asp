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
Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 


Dim lgADF            
Dim lgPID                                                           '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strFrPrDt
Dim strToPrDt
Dim strFrPrNo
Dim strToPrNo
Dim strDeptCd
Dim strFrAsstNo
Dim strToAsstNo


Dim strCond

Dim iPrevEndRow
Dim iEndRow
Dim lgDataExist

' 권한관리 추가 
Dim lgAuthBizAreaCd	' 사업장 
Dim lgInternalCd	' 내부부서 
Dim lgSubInternalCd	' 내부부서(하위포함)
Dim lgAuthUsrID		' 개인 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    lgPID          = UCase(Request("PID"))
    lgPageNo   = Cint(Request("lgStrPrevKey"))                               '☜ : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgDataExist    = "No"
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

	iPrevEndRow = 0
    iEndRow = 0

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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
    lgstrData = ""

    iPrevEndRow = 0

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
            lgstrData	=	lgstrData      & iRowStr & Chr(11) & Chr(12)
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

    Redim UNIValue(0,2)

    UNISqlId(0) = "A7102RA1"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    'UNIValue(0,2) = UCase(Trim(strToPrDt)) A7101RA1
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

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
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strFrPrNo	   = UCase(Trim(Request("txtFrPrNo")))
     strToPrNo     = UCase(Trim(Request("txtToPrNo")))
     strFrPrDt     = UniConvDate(Request("txtFrPrDt"))
     strToPrDt     = UniConvDate(Request("txtToPrDt"))
     strDeptCd	   = UCase(Trim(Request("txtDeptCd")))
     strFrAsstNo   = UCase(Trim(Request("txtFrAsstNo")))
     strToAsstNo   = UCase(Trim(Request("txtToAsstNo")))

     If strFrAsstNo <> "" Then
		strCond = strCond & " and C.ASST_NO >=  " & FilterVar( strFrAsstNo, "''", "S") & " "	 
     End If
     
     If strToAsstNo <> "" Then
		strCond = strCond & " and C.ASST_NO <=  " & FilterVar(strToAsstNo, "''", "S") & " "
     End If

	 If strFrPrNo <> "" Then
		strCond = strCond & " and A.ACQ_NO >=  " & FilterVar(strFrPrNo, "''", "S") & " "	 
     End If
     
     If strToPrNo <> "" Then
		strCond = strCond & " and A.ACQ_NO <=  " & FilterVar(strToPrNo, "''", "S") & " "
     End If
         
     If Trim(Request("txtToPrDt")) <> "" Then
		strCond = strCond & " and A.ACQ_DT <=  " & FilterVar(strToPrDt , "''", "S") & ""
     End If
     
     If Trim(Request("txtFrPrDt")) <> "" Then
		strCond = strCond & " and A.ACQ_DT >=  " & FilterVar(strFrPrDt , "''", "S") & ""
     End If  
     
     If strDeptCd <> "" Then
		strCond = strCond & " and A.DEPT_CD =  " & FilterVar(strDeptCd, "''", "S") & " "
     End If
     
   
     IF lgPID = "A7122MA1" or lgPID = "A7122MA1_KO441" Then	
     strCond = strCond & " and A.ACQ_FG = " & FilterVar("03", "''", "S") & " "
	 else
     strCond = strCond & " and A.ACQ_FG <> " & FilterVar("03", "''", "S") & " "
	 END IF

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If  	 		
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub
%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
       Parent.lgStrPrevKey      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
</Script>	

