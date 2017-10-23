<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3104mb1
'*  4. Program Name         : 예적금입출내역조회 
'*  5. Program Desc         : Query of Deposit Income/Outgo
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.12
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  사업장코드, 은행코드 오류 Check
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4          '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strBizAreaCd															'⊙ : 사업장 
Dim strBizAreaCd1															'⊙ : 사업장1
Dim strDpstFg																'⊙ : 예적금구분 
Dim strDpstType																'⊙ : 예적금유형 
Dim strBankCd																'⊙ : 은행 
Dim strTransSts																'⊙ : 거래상태 
Dim strDocCur																'⊙ : 통화 
Dim strWhere																'⊙ : Where 조건 
Dim strMsgCd, strMsg1, strMsg2

Dim  iLoopCount
Dim  LngMaxRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")
    Call HideStatusWnd 

    lgPageNo			= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgStrPrevKey		= Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount			= 100
    lgSelectList		= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList			= Request("lgTailList")                                 '☜ : Orderby value
	LngMaxRow			= CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

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
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

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
            lgStrPrevKey = rs0(2)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                   '☜: Check if next data exists
        lgPageNo = ""              
        lgStrPrevKey = ""                                   '☜: 다음 데이타 없다.
    End If
  	
'	rs0.Close
'    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(4,2)

    UNISqlId(0) = "F3104MA101"
    UNISqlId(1) = "F3104MA105"	'사업장코드 
    UNISqlId(2) = "F3104MA106"	'은행코드 
    UNISqlId(3) = "F3104MA107"	'은행코드 
    UNISqlId(4) = "F3104MA105"	'사업장코드1
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strWhere))
    UNIValue(1,0) = FilterVar(strBizAreaCd , "''", "S")
    UNIValue(2,0) = FilterVar(strBankCd , "''", "S") 
    UNIValue(3,0) = FilterVar(strDocCur , "''", "S")
    UNIValue(4,0) = FilterVar(strBizAreaCd1 , "''", "S")
    
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBizAreaCd <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBizAreaCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%		
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strBankCd <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBankCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBankCd.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			.txtBankNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
		End With
		</Script>
<%
	End If
	
	rs2.Close
	Set rs2 = Nothing
	
	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strDocCur <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtDocCur_Alt")
		End If
	End If
	
	rs3.Close
	Set rs3 = Nothing
	
	If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" And strBizAreaCd1 <> "" Then 
			strMsgCd = "970000"
			strMsg1 = Request("txtBizAreaCd_Alt1")
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBizAreaCd1.value = "<%=ConvSPChars(Trim(rs4(0)))%>"
			.txtBizAreaNm1.value = "<%=ConvSPChars(Trim(rs4(1)))%>"
		End With
		</Script>
<%		
	End If
	
	rs4.Close
	Set rs4 = Nothing
	
    If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
	
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strBizAreaCd	= UCase(Trim(Request("txtBizAreaCd")))
	strBizAreaCd1	= UCase(Trim(Request("txtBizAreaCd1")))
	
	strBankCd		= UCase(Trim(Request("txtBankCd")))
	strDpstType		= UCase(Trim(Request("cboDpstType")))
	strTransSts		= UCase(Trim(Request("cboTransSts")))
	strDocCur		= UCase(Trim(Request("txtDocCur")))
	
	strWhere = ""
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " and A.biz_area_cd >=  " & FilterVar(strBizAreaCd , "''", "S") & ""
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " and A.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""
	end if
	
	If strBankCd   <> "" Then strWhere = strWhere & " and A.bank_cd   =  " & FilterVar(strBankCd , "''", "S") & " "
	If strDpstType <> "" Then strWhere = strWhere & " and A.dpst_type =  " & FilterVar(strDpstType , "''", "S") & " "
	If strTransSts <> "" Then strWhere = strWhere & " and A.trans_sts =  " & FilterVar(strTransSts , "''", "S") & " "
	If strDocCur   <> "" Then strWhere = strWhere & " and A.doc_cur   =  " & FilterVar(strDocCur , "''", "S") & " "
	
	' 권한관리 추가 
'	If lgAuthBizAreaCd <> "" Then
'		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
'	End If
'	
'	If lgInternalCd <> "" Then
'		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
'	End If
'	
'	If lgSubInternalCd <> "" Then
'		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
'	End If
'	
'	If lgAuthUsrID <> "" Then
'		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
'	End If
'	
'	' 권한관리 추가 
'	strWhere	= strWhere & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	
	
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    With parent
        .ggoSpread.Source     = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=lgstrData%>" , "F"                 '☜ : Display data
		.lgPageNo_A           =  "<%=lgPageNo%>"               '☜ : Next next data tag
		.lgStrPrevKey_A       = "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
		Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=LngMaxRow%>" , "<%=LngMaxRow + iLoopCount%>" ,parent.GetKeyPos("A",3),parent.GetKeyPos("A",4),   "A" ,"I","X","X")
		.DbQueryOk()
		.frm1.vspdData.Redraw = True
	End with
</Script>	

