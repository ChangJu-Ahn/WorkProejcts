<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5211rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : B/L Reference ASP															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/04/30																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'********************************************************************************************************

%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
                                                                         
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3,rs4			   '�� : DBAgent Parameter ���� 
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData
	
	Dim strPayTerms												  '�� : ������� 
	Dim strIncoterms											  '�� : �������� 
	Dim strPurGrp												  '�� : ���ű׷� 
	Dim strForwarderNm											  '�� : ����ȸ��� 
		
	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query
    
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

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
function SetConditionData()
    'On Error Resume Next
    
    SetConditionData= true
  
    If Not(rs1.EOF Or rs1.BOF) Then
        strForwarderNm = rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtBeneficiary")) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = false
			exit function        
		End If
	End If     
    
    If Not(rs2.EOF Or rs2.BOF) Then
		strPurGrp = rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtPurGrp")) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = false
			exit function        
		End If
	End If   	
    
     
	If Not(rs3.EOF Or rs3.BOF) Then
        strPayTerms = rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtPayTerms")) Then
			Call DisplayMsgBox("970000", vbInformation, "�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = false
			exit function        
		End If			
    End If   	
    
    If Not(rs4.EOF Or rs4.BOF) Then
        strIncoterms = rs4(1)
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtIncoterms")) Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = false
			exit function        
		End If				
    End If      
  
End function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	Dim arrVal(4)														  '��: ȭ�鿡�� �˾��Ͽ� query
	
	Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(4,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
   
    UNISqlId(0) = "M5211RA101"  ' main query(spread sheet�� �ѷ����� query statement)   
	UNISqlId(1) = "s0000qa002"	'����ȸ��� 
	UNISqlId(2) = "s0000qa019"  '���ű׷�                                                
	UNISqlId(3) = "s0000qa000"  '�������                                                
	UNISqlId(4) = "s0000qa000"	'��������                                                  			                                                     
    
    '--- 2004-08-20 by Byun Jee Hyun for UNICODE
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
	strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
	strVal = " "	
	
	If Len(Request("txtIssueFromDt")) Then 
		strVal = strVal & " AND A.BL_ISSUE_DT >=  " & FilterVar(UNIConvDate(Request("txtIssueFromDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtIssueToDt")) Then
		strVal = strVal & " AND A.BL_ISSUE_DT <=  " & FilterVar(UNIConvDate(Request("txtIssueToDt")), "''", "S") & " "
	End If
	
 
	If Len(Request("txtBLDocNo")) Then
		strVal = strVal & " AND A.BL_DOC_NO =  " & FilterVar(Trim(UCase(Request("txtBLDocNo"))), " " , "S") & " "
	ELSE
	   strVal = strVal & " AND A.BL_DOC_NO <=  " & FilterVar(strEnd , "''", "S") & " "
	end if
	

	If Len(Trim(Request("txtBeneficiary"))) Then
		strVal = strVal & " AND A.BENEFICIARY =  " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
	End If
	arrVal(1) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")
	
	If Len(Trim(Request("txtPurGrp"))) Then
		strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S") & " "
	End If
	arrVal(2) =  FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S")
	
	If Len(Trim(Request("txtPayTerms"))) Then
		strVal = strVal & " AND A.PAY_METHOD =  " & FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S") & " "
	End If
	arrVal(3) =  FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S")
	
	If Len(Trim(Request("txtIncoterms"))) Then
		strVal = strVal & " AND A.INCOTERMS =  " & FilterVar(Trim(UCase(Request("txtIncoterms"))), " " , "S") & " "
	End If

     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND A.BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND A.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND A.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
	
	
	arrVal(4) =  FilterVar(Trim(UCase(Request("txtIncoterms"))), " " , "S")
	
	UNIValue(0,1) = strVal    '	UNISqlId(0)�� �ι�° ?�� �Էµ�	
	UNIValue(1,0) = arrVal(1)
	UNIValue(2,0) = arrVal(2)	

	UNIValue(3,0) = FilterVar("B9004", " " , "S")			
	UNIValue(3,1) = arrVal(3)
	
	UNIValue(4,0) = FilterVar("B9006", " " , "S")
	UNIValue(4,1) = arrVal(4)
	
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3,rs4)

	Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If  SetConditionData = false then exit sub

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
		    
		.frm1.txtPayTermsNm.value = "<%=ConvSPChars(strPayTerms)%>"
		.frm1.txtIncotermsNm.value = "<%=ConvSPChars(strIncoterms)%>"		       
		.frm1.txtPurGrpNm.value = "<%=ConvSPChars(strPurGrp)%>"		       
		.frm1.txtBeneficiaryNm.value = "<%=ConvSPChars(strForwarderNm)%>"		       								       
       
       If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.hdnBeneficiary.value = "<%=ConvSPChars(Request(txtBeneficiary))%>"
				.frm1.hdnIssueFromDt.value = "<%=ConvSPChars(Request("txtIssueFromDt"))%>"
				.frm1.hdnIssueToDt.value = "<%=ConvSPChars(Request("txtIssueToDt"))%>"
			End If    
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                            '��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
		End If
		
		
       
   	End with
</Script>	
 	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>


