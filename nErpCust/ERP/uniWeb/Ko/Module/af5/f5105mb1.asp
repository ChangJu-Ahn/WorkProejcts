
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
Err.Clear 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3, rs4         '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim NOSumAmt,STSumAmt
Dim strMsgCd
Dim lgtxtMaxRows
Dim strFromIssueDt 
Dim strToIssueDt
Dim strNoteFg
Dim strBankCd
Dim strBpCd
Dim strStsCd
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

Dim strWhere0
Dim strWhere1
Dim strWhere2

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
    
    lgPageNo       = Request("lgPageNo")                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
    Const C_SHEETMAXROWS_D = 100

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
    lgstrData      = ""

    If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then                 
          lgPageNo = CInt(lgPageNo)
       End If   
    Else   
       lgPageNo = 0
    End If   

       'rs0에 대한 결과 
    rs0.PageSize     = C_SHEETMAXROWS_D                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1

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
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If

    
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()


    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "F5105MA101"	'어음(f_note)조회 
	UNISqlId(1) = "F5105MA102"	'어음금액 합계 
	UNISqlId(2) = "F5105MA103"	'결제금액 합계 
	UNISqlId(3) = "A_GETBIZ"
    UNISqlId(4) = "A_GETBIZ"

	Redim UNIValue(4,2)
	
    UNIValue(0,0) = Trim(lgSelectList)                                          '☜: Select list
    UNIValue(0,1) = UCase(Trim(strWhere0))										'where0조건절 list
	UNIValue(1,0) = UCase(Trim(strWhere1))										'where1조건절 list
	UNIValue(2,0) = UCase(Trim(strWhere2))										'where2조건절 list
	
	UNIValue(3,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(4,0)  = FilterVar(strBizAreaCd1, "''", "S")

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
	Dim txtBankNm, txtBpNm, txtStsNm
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1 ,rs2, rs3, rs4)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	'은행명 
	If Trim(request("txtBankCd")) <> "" Then				
		Call CommonQueryRs(" BANK_NM "," B_BANK "," BANK_CD =  " & FilterVar(strBankCd, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtBankNm = ""
		    Call DisplayMsgBox("120800", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
			Call SetErrorStatus()
			exit sub
		Else   
		  txtBankNm = ConvSPChars(Trim(Replace(lgF0,Chr(11),"")))
		End if    	    
	Else 
		txtBankNm = ""
	End If
	
	'거래처명 
	If Trim(request("txtBpCd")) <> "" Then				
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER "," BP_CD =  " & FilterVar(strBpCd, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtBpNm = ""
		    Call DisplayMsgBox("126100", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
			Call SetErrorStatus()
			exit sub
		Else   
		  txtBpNm = ConvSPChars(Trim(Replace(lgF0,Chr(11),"")))
		End if    	    
	Else 
		txtBpNm = ""
	End If
	
	If Trim(request("txtStsCd")) <> "" Then				
		Call CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1008", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strStsCd, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
		  txtStsNm = ""
		    Call DisplayMsgBox("206139", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
			Call SetErrorStatus()
			exit sub
		Else   
		  txtStsNm = ConvSPChars(Trim(Replace(lgF0,Chr(11),"")))
		End if    	    
	Else 
		txtStsNm = ""
	End If
	
	
%>
<Script Language=vbscript>
	With Parent.Frm1
	 .txtBankNm.Value				= "<%=txtBankNm%>"
	 .txtBpNm.Value					= "<%=txtBpNm%>"
	 .txtStsNm.Value				= "<%=txtStsNm%>"
	End With
	 
</Script>       
<%  


If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs3(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs3(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs3.Close
	Set rs3 = Nothing   
    
    
If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs4(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs4(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs4.Close
	Set rs4 = Nothing 


	If Not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False Then NOSumAmt = rs1(0)
	End If
%>
<Script Language=vbscript>
		With parent.frm1
			.txtNoteSum.Text = "<%=UNINumClientFormat(NOSumAmt, ggAmtOfMoney.DecPoint, 0)%>"	'어음금액합계 
		End With 
</script>
<%
	rs1.Close
	Set rs1 = Nothing	

	If Not(rs2.EOF And rs2.BOF) Then
		If IsNull(rs2(0)) = False Then STSumAmt = rs2(0)
	End If
%>
<Script Language=vbscript>
		With parent.frm1
			.txtSttlSum.Text = "<%=UNINumClientFormat(STSumAmt, ggAmtOfMoney.DecPoint, 0)%>"	'결제금액합계 
		End With 
</script>
<%
	rs2.Close
	Set rs2 = Nothing	
	
	
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		rs0.Close
        Set rs0 = Nothing
        Exit Sub													'☜: 비지니스 로직 처리를 종료함 
	End If
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
		rs0.close
		Set rs0 = nothing
    End If
    

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strFromIssueDt	= uniconvdate(Request("txtFromIssueDt"))
	strToIssueDt	= uniconvdate(Request("txtToIssueDt"))
	strNoteFg		= UCase(request("cboNoteFg"))							'어음구분 
	strBankCd		= UCase(request("txtBankCd"))							'은행코드 
	strBpCd			= UCase(request("txtBpCd"))								'거래처코드 
	strStsCd		= UCase(request("txtStsCd"))							'어음상태	
	lgtxtMaxRows	= Request("txtMaxRows")
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))					'사업장From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))					'사업장To

	strWhere0 = ""
	strWhere0 = strWhere0 & " a.bank_cd = c.bank_cd "
	strWhere0 = strWhere0 & " and a.bp_cd = b.bp_cd"
	strWhere0 = strWhere0 & " and a.issue_Dt between  " & FilterVar(strFromIssueDt, "''", "S") & " and  " & FilterVar(strToIssueDt, "''", "S") & " "
	strWhere0 = strWhere0 & " and a.note_fg =  " & FilterVar(strNoteFg , "''", "S") & ""
	strWhere0 = strWhere0 & " and d.major_cd = " & FilterVar("F1008", "''", "S") & " "
	strWhere0 = strWhere0 & " and d.minor_cd = a.note_sts "
	

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



	
	If strBankCd <> "" Then
	strWhere0 = strWhere0 & " and a.bank_cd =  " & FilterVar(strBankCd , "''", "S") & " "
	End If 
	
	If strBpCd <> "" Then
	strWhere0 = strWhere0 & " and a.bp_cd =  " & FilterVar(strBpCd , "''", "S") & " "
	End If
	
	If strStsCd <> "" Then
	strWhere0 = strWhere0 & " and a.note_sts =  " & FilterVar(strStsCd , "''", "S") & " "
	End If
	
	if strBizAreaCd <> "" then
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD >= " & FilterVar(" ", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if
	

	' 권한관리 추가 
	strWhere0	= strWhere0	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

	strWhere0 = strWhere0 & " order by a.note_no"




	strWhere1 = ""
	strWhere1 = strWhere1 & " a.bank_cd = c.bank_cd "
	strWhere1 = strWhere1 & " and a.bp_cd = b.bp_cd"
	strWhere1 = strWhere1 & " and a.issue_Dt between  " & FilterVar(strFromIssueDt, "''", "S") & " and  " & FilterVar(strToIssueDt, "''", "S") & " "
	strWhere1 = strWhere1 & " and a.note_fg =  " & FilterVar(strNoteFg , "''", "S") & ""

	If strBankCd <> "" Then
	strWhere1 = strWhere1 & " and a.bank_cd =  " & FilterVar(strBankCd , "''", "S") & " "
	End If 
	
	If strBpCd <> "" Then
	strWhere1 = strWhere1 & " and a.bp_cd =  " & FilterVar(strBpCd , "''", "S") & " "
	End If
	
	If strStsCd <> "" Then
	strWhere1 = strWhere1 & " and a.note_sts =  " & FilterVar(strStsCd , "''", "S") & ""
	End If
	
	if strBizAreaCd <> "" then
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD >= " & FilterVar(" ", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if
	


	strWhere2 = "select ISNULL(SUM(A.note_amt),0) as sttl_amt_sum from f_note a, b_biz_partner b, b_bank c where"
	strWhere2 = strWhere2 & " a.bank_cd = c.bank_cd "
	strWhere2 = strWhere2 & " and a.bp_cd = b.bp_cd "
	strWhere2 = strWhere2 & " and a.issue_Dt between  " & FilterVar(strFromIssueDt, "''", "S") & " and  " & FilterVar(strToIssueDt, "''", "S") & " "
	strWhere2 = strWhere2 & " and a.note_fg =  " & FilterVar(strNoteFg , "''", "S") & ""

	if strStsCd = "" then
	strWhere2 = strWhere2 & " and a.note_sts = " & FilterVar("SM", "''", "S") & " "
	elseif strStsCd <> "" and strStsCd = "SM" then
	strWhere2 = strWhere2 & " and a.note_sts =  " & FilterVar(strStsCd , "''", "S") & ""
	else 
	strWhere2 = strWhere2 & " and a.note_sts = ''"
	End if 
	
	If strBankCd <> "" Then
	strWhere2 = strWhere2 & " and a.bank_cd =  " & FilterVar(strBankCd , "''", "S") & " "
	End If 
	
	If strBpCd <> "" Then
	strWhere2 = strWhere2 & " and a.bp_cd =  " & FilterVar(strBpCd , "''", "S") & " "
	End If
	
	if strBizAreaCd <> "" then
		strWhere2 = strWhere2 & " AND a.BIZ_AREA_CD >= " & FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere2 = strWhere2 & " AND a.BIZ_AREA_CD >= " & FilterVar(" ", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere2 = strWhere2 & " AND a.BIZ_AREA_CD <= " & FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere2 = strWhere2 & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if


	
	' 권한관리 추가 
	strWhere1	= strWhere1	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	strWhere2	= strWhere2	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

	
End Sub
%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
  			Parent.frm1.htxtFromIssueDt.value	= "<%=strFromIssueDt%>"
			Parent.frm1.htxtToIssueDt.value		= "<%=strToIssueDt%>"  		
			Parent.frm1.hcboNoteFg.value		= "<%=strNoteFg%>"
			Parent.frm1.htxtBankCd.value		= "<%=strBankCd%>"
			Parent.frm1.htxtBpCd.value			= "<%=strBpCd%>"
			Parent.frm1.htxtStsCd.value			= "<%=strStsCd%>"		
      End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
       Parent.lgPageNo			=  "<%=lgPageNo%>"   
       
       Parent.DbQueryOk
    End If   

</Script>	

