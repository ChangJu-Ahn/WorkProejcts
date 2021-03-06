<!--
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 발주참조 Popup
'*  3. Program ID           : M3111RB3
'*  4. Program Name         : P/O Reference ASP
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/17
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/04 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/17 : ADO변환 
'**************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next													   '실행 오류가 발생할 때 오류가 발생한 문장 바로 다음에 실행이 계속될 수 있는 문으로 컨트롤을 옮길 수 있도록 지정합니다.				

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iTotstrData

Dim strBeneficiaryNm
Dim strPurGrpNm
Dim strPoTypeNm
Dim strPaymeth
	
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
	
	Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(4,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 
	
	UNISqlId(0) = "M3111RA301"												' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(1) = "S0000QA002"	'수출자 
    UNISqlId(2) = "S0000QA022"	'구매그룹 
    UNISqlId(3) = "S0000QA020"	'발주형태 
    UNISqlId(4) = "S0000QA000"	'결제방법 
    
    '--- 2004-08-19 by Byun Jee Hyun for UNICODE
    strVal = ""
	If Len(Request("txtBeneficiary")) Then
		strVal = strVal & " AND mhdr.BP_CD =  " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
	end if
	
	If Len(Trim(Request("txtGroup"))) Then
		strVal = strVal & " AND mhdr.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S") & " "
	End If
	
	If Len(Request("txtPOType")) Then
		strVal = strVal & " AND mhdr.PO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtPOType"))), " " , "S") & " "
	end if
	
	If Len(Request("txtFrPODt")) Then
		strVal =  strVal & " AND mhdr.PO_DT >=  " & FilterVar(UNIConvDate(Request("txtFrPODt")), "''", "S") & ""
	End If	
	
	If Len(Request("txtToPODt")) Then
		strVal =  strVal & " AND mhdr.PO_DT <=  " & FilterVar(UNIConvDate(Request("txtToPODt")), "''", "S") & ""
	End If	
		
	If Len(Trim(Request("txtPayTerms"))) Then
		strVal = strVal & " AND mhdr.PAY_METH =  " & FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("gBizArea"))) Then
		strVal = strVal & " AND mhdr.PUR_BIZ_AREA =  " & FilterVar(Trim(UCase(Request("gBizArea"))), " " , "S") & " "
	End If

	If Len(Trim(Request("gPurGrp"))) Then
		strVal = strVal & " AND mhdr.PUR_GRP =  " & FilterVar(Trim(UCase(Request("gPurGrp"))), " " , "S") & " "
	End If

	If Len(Trim(Request("gPurOrg"))) Then
		strVal = strVal & " AND mhdr.PUR_ORG =  " & FilterVar(Trim(UCase(Request("gPurOrg"))), " " , "S") & " "
	End If

	UNIValue(0,0) = lgSelectList                                    '☜: Select list
	UNIValue(0,1) = strVal											'	UNISqlId(0)의 두번째 ?에 입력됨	
	
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")  				'수혜자 
	UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S")  						'구매그룹 
	UNIValue(3,0) = FilterVar(Trim(UCase(Request("txtPOType"))), " " , "S")  						'발주형태 
	UNIValue(4,0) = FilterVar("B9004", " " , "S")  																			'결제방법(Major_cd)
	UNIValue(4,1) = FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S") 					'결제방법 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
'    UNIValue(0,UBound(UNIValue,2)) = " ORDER BY mhdr.PO_NO DESC "
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
	
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

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
    
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
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
        strBeneficiaryNm = rs1("Bp_Nm")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtBeneficiary"))) Then
			Call DisplayMsgBox("970000", vbInformation, "수혜자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			EXIT FUNCTION
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrpNm = rs2("Pur_Grp_Nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			EXIT FUNCTION			
		End If
	End If  
	
	If Not(rs3.EOF Or rs3.BOF) Then
        strPoTypeNm = rs3("Po_Type_Nm")
   		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtPOType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "발주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			EXIT FUNCTION			
		End If
	End If  
	
	If Not(rs4.EOF Or rs4.BOF) Then
       strPaymeth = rs4("Minor_Nm")
   		Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Trim(Request("txtPayTerms"))) Then
			Call DisplayMsgBox("970000", vbInformation, "결제방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			EXIT FUNCTION			
		End If
	End If  
	
	SetConditionData = TRUE
	
End Function

%>

<Script Language=vbscript>
    With parent
       
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.txtHBeneficiary.value	= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
				.txtHPOType.value		= "<%=ConvSPChars(Request("txtPOType"))%>"
				.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"
				.txtHFrPODt.Value		= "<%=Request("txtFrPODt")%>"
				.txtHToPODt.Value		= "<%=Request("txtToPODt")%>"
				.txtHGrp.Value			= "<%=ConvSPChars(Request("txtGroup"))%>" 
			End If    
			'Show multi spreadsheet data from this line
				       
			.ggoSpread.Source    = .vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                            '☜: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
