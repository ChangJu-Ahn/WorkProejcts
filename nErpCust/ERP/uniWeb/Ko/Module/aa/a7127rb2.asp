<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                        '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo


'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim LngRow
Dim GroupCount    
Dim strVal

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

Dim strSoldyyyymm
Dim strFrAcqDt
Dim strToAcqDt
Dim strDeptCd
Dim strFrAsstNo
Dim strToAsstNo
Dim strAcctCd
Dim strOrgChangeId

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

Dim strCond
	
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Dim DEPT_NM
Dim ACCT_NM 
Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D 'CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()


Const ConDate = "1899/12/30"

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

  
  
   Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "A7127RA201"
	UNISqlId(1) = "COMMONQRY"
    UNISqlId(2) = "COMMONQRY"

    Redim UNIValue(2,6)

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	UNIValue(0,1) = FilterVar(strSoldyyyymm, "''", "S")
	UNIValue(0,2) = FilterVar(strSoldyyyymm, "''", "S")
	UNIValue(0,3) = FilterVar(strSoldyyyymm, "''", "S")
	'UNIValue(0,4) = FilterVar(strSoldyyyymm, "''", "S")
	UNIValue(0,5) = strCond

	UNIValue(1,0) = " SELECT DEPT_NM FROM B_ACCT_DEPT WHERE DEPT_CD =  " & FilterVar(UCase(Request("txtDeptCd")), "''", "S") & "" & _
					" AND ORG_CHANGE_ID =  " & FilterVar(UCase(Request("txtOrgChangeId")), "''", "S") & ""
    UNIValue(2,0) = "SELECT ACCT_NM FROM A_ACCT WHERE ACCT_CD =  " & FilterVar(UCase(Request("txtAcctCd")), "''", "S") & ""

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

'	UNIValue(0,0) = strWhere
'    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim strMsg
    Dim strMsg1
    Dim strMsgCd
    Dim strMsgCd1
    
    strMsg = Trim(Request("txtDeptCd_Alt"))
    strMsg1 = Trim(Request("txtAcctCd_Alt"))
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	IF NOT (rs1.EOF or rs1.BOF) then
		DEPT_NM = rs1(0)
%>
		<Script Language=vbScript>
			With parent
				.Frm1.txtDeptNm.Value  = "<%=DEPT_NM%>"
			End With
		</Script>
<%			
	ELSE
		if Trim(Request("txtDeptCd")) <> "" Then
			strMsgCd = "970000"
%>
			<Script Language=vbScript>
				With parent
					.Frm1.txtDeptNm.Value  = ""
				End With
			</Script>
<%	
		Else
%>
			<Script Language=vbScript>
				With parent
					.Frm1.txtDeptNm.Value  = ""
				End With
			</Script>
<%			
		End if
	End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2에 대한 결과
    IF NOT (rs2.EOF or rs2.BOF) then
	    ACCT_NM = rs2(0)
%>
		<Script Language=vbScript>
			With parent
				.Frm1.txtAcctNm.Value = "<%=ACCT_NM%>"   
			End With
		</Script>
<%			    
	ELSE
		if Trim(Request("txtAcctCd")) <> "" Then
			strMsgCd1 = "970000"
%>
			<Script Language=vbScript>
				With parent
					.Frm1.txtAcctNm.Value = ""   
				End With
			</Script>
<%	
		Else
%>
			<Script Language=vbScript>
				With parent
					.Frm1.txtAcctNm.Value = ""   
				End With
			</Script>
<%		
		End if
	END IF
	
	rs2.Close
	
	Set rs2 = Nothing
		
	If  Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
	    Response.End													'☜: 비지니스 로직 처리를 종료함
	End If
	    
	If  Trim(strMsgCd1) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
	    Response.End													'☜: 비지니스 로직 처리를 종료함
	End If
	
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Response.End
	Else
		Call  MakeSpreadSheetData()
	End If
	    
	Set rs0 = Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
	
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strCond = ""
	strSoldyyyymm	= Trim(Request("txtSoldyyyymm"))
	strFrAcqDt	= UCase(Trim(UNIConvDate(Request("txtFrAcqDt"))))
	strToAcqDt	= UCase(Trim(UNIConvDate(Request("txtToAcqDt"))))
	strDeptCd	= UCase(Trim(Request("txtDeptCd")))
	strOrgChangeId	= UCase(Trim(Request("txtOrgChangeId")))
	strFrAsstNo	= UCase(Trim(Request("txtFrAsstNo")))
	strToAsstNo	= UCase(Trim(Request("txtToAsstNo")))
	strAcctCd	= UCase(Trim(Request("txtAcctCd")))

	' 권한관리 추가
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

	If Trim(Request("txtFrAcqDt")) <> "" Then
	   strCond = strCond & " AND A.REG_DT >=  " & FilterVar(strFrAcqDt , "''", "S") & ""
	End If
	     
	If Trim(Request("txtToAcqDt")) <> "" Then
	   strCond = strCond & " AND A.REG_DT <=  " & FilterVar(strToAcqDt , "''", "S") & ""
	End If

	If strDeptCd <> "" Then
		strCond = strCond & " AND D.DEPT_CD =  " & FilterVar(strDeptCd , "''", "S") & ""
		strCond = strCond & " AND D.ORG_CHANGE_ID =  " & FilterVar(strOrgChangeId , "''", "S") & ""
	End If
	     
	If strFrAsstNo <> "" Then
	   strCond = strCond & " AND A.ASST_NO >=  " & FilterVar(strFrAsstNo , "''", "S") & ""
	End If
	     
	If strToAsstNo <> "" Then
	   strCond = strCond & " AND A.ASST_NO <=  " & FilterVar(strToAsstNo , "''", "S") & ""
	End If   
	     
	If strAcctCd <> "" Then
		strCond = strCond & " AND A.ACCT_CD =  " & FilterVar(strAcctCd , "''", "S") & "" 
	End If

	' 권한관리 추가
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	' 권한관리 추가
	strCond	= strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub
'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------
'Function FilterVar(Byval str,Byval strALT)
'     Dim strL
'     strL = UCASE(Trim(str))
'     If Len(strL) Then
'        FilterVar = "'" & strL  & "'"
'     Else
'        FilterVar = strALT   
'     End If
'End Function


%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          parent.Frm1.hSoldyyyymm.Value	= Parent.Frm1.txtSoldyyyymm.Text
          parent.Frm1.hFrAcqDt.Value	= Parent.Frm1.txtFrAcqDt.Text
          Parent.Frm1.hToAcqDt.Value    = Parent.Frm1.txtToAcqDt.Text
          Parent.Frm1.hFrAsstNo.Value	= Parent.Frm1.txtFrAsstNo.Value
          Parent.Frm1.hToAsstNo.Value	= Parent.Frm1.txtToAsstNo.Value
          Parent.Frm1.hAcctCd.Value		= Parent.Frm1.txtAcctCd.Value
		  Parent.Frm1.hDeptCd.Value		= Parent.Frm1.txtDeptCd.value
		  
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   
</Script>

