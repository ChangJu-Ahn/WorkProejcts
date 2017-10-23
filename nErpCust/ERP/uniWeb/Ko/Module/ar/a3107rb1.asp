<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0  ,rs1, rs2                   '☜ : DBAgent Parameter 선언 
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

dim strRcptDt
dim strToRcptDt
dim strBizCd
dim strBpCd
dim strDocCur
Dim strAllcDt
	                                                           '⊙ : 발주일 
Dim strCond
Dim strMsgCd
Dim strMsg1


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","A","NOCOOKIE","RB")
    Call HideStatusWnd 


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
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A3107RB110"
	UNISqlId(1) = "ABPNM"
    UNISqlId(2) = "ABizNM"
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,1)  = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )
    UNIValue(2,0) = UCase(" " & FilterVar(strBizCd, "''", "S") & " " )

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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBpCd_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.txtBpNm.value = ""
		End With
		</Script>
<%
	Else
%>
		<Script Language=vbScript>
		With parent
			.txtBpCd.value = "<%=Trim(ConvSPChars(rs1(0)))%>"
			.txtBpNm.value = "<%=Trim(ConvSPChars(rs1(1)))%>"
		End With
		</Script>
<%
    End If
	Set rs1 = Nothing 


    If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strBizCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizCd_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.txtBizNm.value = ""
		End With
		</Script>
<%
    Else
%>
		<Script Language=vbScript>
		With parent
			.txtBizCd.value = "<%=Trim(ConvSPChars(rs2(0)))%>"
			.txtBizNm.value = "<%=Trim(ConvSPChars(rs2(1)))%>"
		End With
		</Script>
<%
    End If
	Set rs2 = Nothing 
		
    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Resonse.End
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

'----------------------------------------------------------------------------------------------------------
' Trim Data
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

	strRcptDt		= UCase(Trim(UNIConvDate(Request("txtRcptDt"))))
	strToRcptDt		= UCase(Trim(UNIConvDate(Request("txtToRcptDt"))))
	strBizCd		= UCase(Trim(Request("txtBizcd")))
	strBpCd			= UCase(Trim(Request("txtBpcd")))
	strDocCur		= UCase(Trim(Request("txtDocCur")))
	strAllcDt		= UNIConvDate(Trim(Request("txtAllcDt")))
	strRcptNo		= UCase(Trim(Request("txtRcptNo")))
	strRemark		= UCase(Trim(Request("txtRemark")))
		
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))		 
		
	strCond = "AND ar.RCPT_STS = " & FilterVar("O", "''", "S") & "  AND ar.BAL_AMT <> 0 AND ar.RCPT_FG = " & FilterVar("S", "''", "S") & "  AND ar.CONF_FG = " & FilterVar("C", "''", "S") & "  AND ar.GL_NO <> ''"
	strCond = strCond & " AND ar.RCPT_DT <=  " & FilterVar(strAllcDt , "''", "S") & ""
		
	If strRcptDt <> "" Then
		strCond = strCond & " AND ar.RCPT_DT  >=  " & FilterVar(strRcptDt , "''", "S") & ""
	End If

	If strToRcptDt <> "" Then
		strCond = strCond & " AND ar.RCPT_DT  <=  " & FilterVar(strToRcptDt , "''", "S") & ""
	End If

	IF strBpCd <> "" Then
		strCond = strCond & " AND ar.BP_CD  =  " & FilterVar(strBpCd, "''", "S") & " "
	End IF	

	IF strDocCur <> "" Then
		strCond = strCond & " AND ar.DOC_CUR  =  " & FilterVar(strDocCur, "''", "S") & " "
	END IF	

	IF strBizCd <> "" Then
		strCond = strCond & " AND ar.BIZ_AREA_CD  =  " & FilterVar(strBizCd, "''", "S") & " "
	ENd IF
		
	' 가수금번호,비고 조건추가 
	IF strRcptNo <> "" Then
		strCond = strCond & " AND ar.RCPT_NO  =  " & FilterVar(strRcptNo, "''", "S") & " "
	ENd IF
		
	IF strRemark <> "" Then
		strCond = strCond & " AND ar.RCPT_DESC  like  " & FilterVar("%" & strRemark & "%", "''", "S") & " "
	ENd IF
	
	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND AR.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND AR.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND AR.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND AR.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If	

	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

End Sub 

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
'			Parent.htxtToPrDt.Value		= Parent.txtToPrDt.Text
			Parent.htxtRcptDt.Value		= Parent.txtRcptDt.text
			Parent.htxtToRcptDt.Value		= Parent.txtToRcptDt.Text			
			Parent.htxtBizCd.Value		= Parent.txtBizCd.Value
			Parent.htxtBpcd.Value		= Parent.txtBpcd.Value
			Parent.htxtDocCur.Value		= Parent.txtDocCur.Value
       End If
			Parent.ggoSpread.Source  = Parent.vspdData
			Parent.vspdData.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",7),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",8),"A", "I" ,"X","X")
			 Parent.vspdData.Redraw = True
			Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag

			Parent.DbQueryOk
    End If   

</Script>	

