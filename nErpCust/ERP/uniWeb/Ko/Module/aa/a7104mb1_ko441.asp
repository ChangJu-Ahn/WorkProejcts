<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<% 

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3 , rs4, rs5, rs6    '☜ : DBAgent Parameter 선언 
Dim lgstrData																	'☜ : data for spreadsheet data
Dim lgStrPrevKey																'☜ : 이전 값 
Dim lgTailList																	'☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtDeptCd
Dim txtDeptNm
Dim txtAcctCd
Dim txtAcctNm
Dim txtCondAsstNo
Dim txtCondAsstNm
Dim strBizAreaCd															'⊙ : 시작사업장 
Dim strBizAreaNm
Dim strBizAreaCd1															'⊙ : 종료사업장 
Dim strBizAreaNm1
Dim strMsgCd, strMsg1, strMsg2

Dim iPrevEndRow
Dim iEndRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call LoadBasisGlobalInf()    

    Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")   
    Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB") 

    Call HideStatusWnd 

    lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList		= Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist		= "No"
    
	txtDeptCd		= Trim(Request("txtDeptCd"))
	txtAcctCd		= Trim(Request("txtAcctCd"))
	txtCondAsstNo	= Trim(Request("txtCondAsstNo"))
	
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))					'사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))					'사업장To

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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

	Const C_SHEETMAXROWS_D  = 100                                          '☆: Fetch max count at once
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    iPrevEndRow = 0

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(C_SHEETMAXROWS_D) * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                   'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData		=	lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
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

    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Dim strWhere

    Redim UNIValue(6,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A7104MA101KO441"		'>> air
    UNISQLID(1) = "commonqry"
    UNISQLID(2) = "commonqry"
    UNISQLID(3) = "commonqry"
	UNISqlId(4) = "A7104MA102KO441"		'>> air
	UNISqlId(5) = "A_GETBIZ"
    UNISqlId(6) = "A_GETBIZ"
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    
    strWhere = ""
    strWhere = strWhere & " AND D.MAJOR_CD = " & FilterVar("A2004", "''", "S") & " AND E.MAJOR_CD = " & FilterVar("A2004", "''", "S") & " "
    If txtDeptCd <> "" Then
		strWhere = strWhere & " AND A.DEPT_CD = " & FilterVar(txtDeptCd ,"''"	,"S")		'관리부서 
	End If

	strWhere = strWhere & " AND A.ASST_NO >= " & FilterVar(txtCondAsstNo ,"''"	,"S")	'자산번호 

	If txtAcctCd <> "" Then
		strWhere = strWhere & " AND A.ACCT_CD = " & FilterVar(txtAcctCd ,"''"	,"S")		'계정명 
	ENd If
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " AND a.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere = strWhere & " AND a.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " AND a.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere = strWhere & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & ""
	End if

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
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strWhere	= strWhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	
	
	UNIValue(0,1)  = strWhere
	UNIValue(4,0)  = strWhere
	UNIValue(1,0) = "select DEPT_NM from B_ACCT_DEPT Where dept_cd= " & FilterVar(txtDeptCd ,"''"	,"S")
	UNIValue(2,0) = "select acct_nm from A_ACCT Where acct_cd = " & FilterVar(txtAcctCd ,"''"	,"S")
	UNIValue(3,0) = "select asst_nm from A_ASSET_MASTER Where asst_no = " & FilterVar(txtCondAsstNo ,"''"	,"S")
	UNIValue(5,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(6,0)  = FilterVar(strBizAreaCd1, "''", "S")

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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    'rs1
    If txtDeptCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
'		Response.Write "txtDeptCd yes" & "<br>"
			txtDeptNm = Trim(rs1("Dept_Nm"))
		Else
'		Response.Write "txtDeptCd no" & "<br>"
			txtDeptNm = ""
			Call DisplayMsgBox("127800", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		    rs1.Close
		    Set rs1 = Nothing 
		    Exit sub
		End IF
		rs1.Close
		Set rs1 = Nothing
	End If

    'rs2
    If txtAcctCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
			txtAcctNm = Trim(rs2("acct_nm"))
'		Response.Write "txtAcctCd yes" & "<br>"
		Else
'		Response.Write "txtAcctCd no" & "<br>"
			txtAcctNm = ""
			Call DisplayMsgBox("110100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		    rs2.Close
		    Set rs2 = Nothing 
		    Exit sub
		End IF
		rs2.Close
		Set rs2 = Nothing
	End If
    
    'rs3
    If txtCondAsstNo <> "" Then
		If Not (rs3.EOF OR rs3.BOF) Then
			txtCondAsstNm = Trim(rs3("asst_nm"))
'		Response.Write "txtCondAsstNo yes" & "<br>"
		Else
'		Response.Write "txtCondAsstNo no" & "<br>"
			txtCondAsstNm = ""
		End IF
		rs3.Close
		Set rs3 = Nothing
    End If

If (rs5.EOF And rs5.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs5(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs5(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs5.Close
	Set rs5 = Nothing   
    
    
If (rs6.EOF And rs6.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs6(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs6(1))%>"
	End With
	</Script>
<%
    End If
    rs6.Close
	Set rs6 = Nothing 
	
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("117400", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
    
    If Not (rs4.EOF OR rs4.BOF) Then
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtSum1.Text  = "<%=UNINumClientFormat(Trim(rs4(0)), ggAmtOfMoney.DecPoint, 0)%>"
	End With
	</Script>
<%
	End IF

	rs4.Close
	Set rs4 = Nothing
    
End Sub

%>

<Script Language=vbscript>
With Parent
	If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.Frm1.htxtDeptCd.Value		= .Frm1.txtDeptCd.Value                  'For Next Search
			.Frm1.htxtAcctCd.Value		= .Frm1.txtAcctCd.Value
			.Frm1.htxtCondAsstNo.Value	= .Frm1.txtCondAsstNo.Value
			.frm1.htxtBizAreaCd.value	= .frm1.txtBizAreaCd.value
			.frm1.htxtBizAreaCd1.value	= .frm1.txtBizAreaCd1.value
				
		End If

		Parent.ggoSpread.Source  = Parent.frm1.vspdData
		Parent.frm1.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
		Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		Parent.DbQueryOk
		Parent.frm1.vspdData.Redraw = True
    End If

	.frm1.txtDeptNm.value = "<%=ConvSPChars(txtDeptNm)%>"			'rs1 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	.frm1.txtAcctNm.value = "<%=ConvSPChars(txtAcctNm)%>"			'rs2 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	.frm1.txtCondAsstNm.value = "<%=ConvSPChars(txtCondAsstNm)%>"	'rs3 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 

	.frm1.txtDeptCd.focus
End With
</Script>
