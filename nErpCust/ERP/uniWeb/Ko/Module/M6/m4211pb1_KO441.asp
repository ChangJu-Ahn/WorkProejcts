
<%

'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4211pb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : ������� PopUp��� Transaction ó���� ASP									*
'*  7. Modified date(First) : 2000/05/12																*
'*  8. Modified date(Last)  : 2000/05/12																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
	'On Error Resume Next
	                                                                         
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2      '�� : DBAgent Parameter ���� 
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim SortNo													  ' Sort ���� 
	Dim iTotstrData

	Dim strBeneficiary											  ' �ŷ�ó 


    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

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

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	Dim arrVal(2)														  '��: ȭ�鿡�� �˾��Ͽ� query
	
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(1,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
   
    UNISqlId(0) = "m4211pa101"												' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(1) = "s0000qa002"												' 2��° query(spread sheet�� �ѷ����� query statement)
    
    '--- 2004-08-20 by Byun Jee Hyun for UNICODE	
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
	
	strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
	strVal = " "
	
	If Len(Request("txtFromDt")) Then 
		strVal = strVal & " AND A.ID_DT >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtToDt")) Then
		strVal = strVal & " AND A.ID_DT <=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "
	End If
	
		
	If Len(Trim(Request("txtBeneficiary"))) Then
		strVal = strVal & " AND A.BENEFICIARY =  " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
	ELSE	
		strVal = strVal & " AND A.BENEFICIARY  <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "
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
	
	UNIValue(0,1) = strVal    '	UNISqlId(0)�� �ι�° ?�� �Էµ�	
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")				  		  '	UNISqlId(1)�� ù��° ?�� �Էµ�		
    

    'UNIValue(0,UBound(UNIValue,2)) = "order by A.CC_NO DESC"
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                       '��: set ADO read mode
 
 
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
    
    if SetConditionData = false then Exit Sub

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub
'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData= true
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiaryNm = rs1("BP_NM")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtBeneficiary"))) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = false
			Exit Function
		End If
		
	End If   	
    SetConditionData = TRUE 
    
End Function    
'-----------------------------

%>

<Script Language=vbscript>
    With parent
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnBeneficiary.value = "<%=Request("txtBeneficiary")%>"
				.frm1.hdnFromDt.value = "<%=Request("txtFromDt")%>"
				.frm1.hdnToDt.value = "<%=Request("txtToDt")%>" 
			End If    
			'Show multi spreadsheet data from this line
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                  '��: Display data 
																	'0: ����None 1 :��������  2: ��������					
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
