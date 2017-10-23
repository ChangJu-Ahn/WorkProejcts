<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%'======================================================================================================
'*  1. Module Name          : 구매 
'*  2. Function Name        : L/C관리 
'*  3. Program ID           : m3211rb2
'*  4. Program Name         : Local L/C참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : M32118ListLcHdrForAmendSvr
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2003/05/20
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kang Su-hwan	
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date 표준적용 
'*							  2002/04/25 ADO 변환 
'=======================================================================================================

                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3 '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iTotstrData

Dim strBeneficiary											  ' 거래처명 
Dim strPurGrp												  ' 구매그룹명 
Dim strPayTerms											      ' 결제방법명 

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiary =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtBeneficiary")) Then
			Call DisplayMsgBox("970000", vbInformation, "수혜자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			Exit Function
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrp =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtPurGrp")) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			Exit Function
		End If			
    End If   	
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strPayTerms =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtPayTerms")) Then
			Call DisplayMsgBox("970000", vbInformation, "결제방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			Exit Function
		End If				
    End If      
    
    SetConditionData = TRUE
    
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)

    UNISqlId(0) = "M3211QA003"  										' main query(spread sheet에 뿌려지는 query statement)
	UNISqlId(1) = "s0000qa002"  										' 거래처코드/명 
	UNISqlId(2) = "s0000qa019"  										' 구매그룹코드/명 
	UNISqlId(3) = "M3211QA103"  										' 결제방법코드/명 

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
	
    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
	strVal = " "	

	If Len(Request("txtBeneficiary")) Then		' 거래처 
		strVal = "AND A.BENEFICIARY = " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "	
	End If	
	arrVal(0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")

	If Len(Request("txtPurGrp")) Then			'구매그룹 
		strVal = strVal & " AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S") & " "		
	End If
	arrVal(1) = FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S")
		   
	If Len(Request("txtPayTerms")) Then			'결제방법 
		strVal = strVal & " AND A.PAY_METHOD = " & FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S") & " "		
	End If
	arrVal(2) = FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S")
    
 	If Len(Request("txtOpenFrDt")) Then			'개설일 
		strVal = strVal & " AND A.OPEN_DT >= " & FilterVar(UNIConvDate(Request("txtOpenFrDt")), "''", "S") & " "		
	End If	    
	
    If Len(Request("txtOpenToDt")) Then		
		strVal = strVal & " AND A.OPEN_DT <= " & FilterVar(UNIConvDate(Request("txtOpenToDt")), "''", "S") & " "			
	End If		
	
    UNIValue(0,1) = strVal   
    UNIValue(1,0) = arrVal(0)						' 거래처코드/명 
    UNIValue(2,0) = arrVal(1)					    ' 구매그룹코드/명 
    UNIValue(3,0) = arrVal(2)			' 결제방법코드/명 
    
'    UNIValue(0,UBound(UNIValue,2)) = "ORDER BY A.LC_NO DESC"
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    IF SetConditionData() = FALSE THEN EXIT SUB
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtBeneficiaryNm.value 	= "<%=ConvSPChars(strBeneficiary)%>" 
		.frm1.txtPurGrpNm.value 		= "<%=ConvSPChars(strPurGrp)%>" 
        .frm1.txtPayTermsNm.value 		= "<%=ConvSPChars(strPayTerms)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHBeneficiary.value		= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
				.frm1.txtHPurGrp.value			= "<%=ConvSPChars(Request("txtPurGrp"))%>"
				.frm1.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"
				.frm1.txtHOpenFrDt.value		= "<%=Request("txtOpenFrDt")%>"
				.frm1.txtHOpenToDt.value		= "<%=Request("txtOpenToDt")%>"
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                            '☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			parent.DbQueryOk
		End If
	End with
</Script>	
