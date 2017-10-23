<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3211pb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open L/C PopUp��� Transaction ó���� ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/22																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************

'On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort ���� 
Dim strItemNm
Dim iTotstrData

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
   
    UNISqlId(0) = "m3211pa101"												' main query(spread sheet�� �ѷ����� query statement)
    UNIValue(0,0) = lgSelectList
    UNISqlId(1) = "s0000qa024"                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    
   
	strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
	strVal = " WHERE A.LC_KIND = " & FilterVar("L", "''", "S") & "  "
	
	If Len(Trim(Request("txtFrDt"))) Then					
		strVal = strVal & " AND A.REQ_DT >= " & FilterVar(UNIConvDate(Request("txtFrDt")), "''", "S") & " "
	End If

	If Len(Trim(Request("txtToDt"))) Then					
		strVal = strVal & " AND A.REQ_DT <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "
	End If			
	
	If Len(Trim(Request("txtBeneficiary"))) Then
		strVal = strVal & " AND A.BENEFICIARY =  " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
	ELSE	
		strVal = strVal & " AND A.BENEFICIARY  <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "
	End If
  
	UNIValue(0,1) = strVal    '	UNISqlId(0)�� �ι�° ?�� �Էµ�	
 
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")		
End Sub

Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strItemNm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtBeneficiary")) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			Exit Function
		End If
	End If   	

	SetConditionData = TRUE
	
End Function

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

		.frm1.txtBeneficiaryNm.Value	= "<%=ConvSPChars(strItemNm)%>"
		
        If "<%=lgDataExist%>" = "Yes" Then

			.frm1.hdnBeneficiary.value = "<%=ConvSPChars(Request(txtBeneficiary))%>"
			.frm1.hdnFrDt.value = "<%=Request("txtFrDt")%>"
			.frm1.hdnToDt.value = "<%=Request("txtToDt")%>"
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"     
			.lgPageNo			 =  "<%=lgPageNo%>"		
			.DbQueryOk
		End If
   	End with
</Script>	
 	
<%
    'Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>




    
