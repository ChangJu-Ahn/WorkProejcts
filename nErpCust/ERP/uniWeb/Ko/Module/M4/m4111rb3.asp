<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3111rb3.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : �������� PopUp ASP														*
'*  7. Modified date(First) : 2000/04/17																*
'*  8. Modified date(Last)  : 2002/05/11																*
'*  9. Modifier (First)     : 																			*
'* 10. Modifier (Last)      : Kim Jin Ha																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2002/05/11 : ADO Conv.												*
'********************************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iTotstrData

Dim strBeneficiaryNm
Dim strPurGrpNm
Dim strPaymeth
Dim strPoTypeNm
Dim strIOTypeNm
	
	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist    = "No"
  
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	
	Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(5,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
	
	UNISqlId(0) = "M4111RA301"												' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(1) = "S0000QA002"	'������ 
    UNISqlId(2) = "S0000QA022"	'���ű׷�(���Դ��)
    UNISqlId(3) = "S0000QA000"	'������� 
    UNISqlId(4) = "s0000qa020"	'�������� 
	UNISqlId(5) = "s0000qa023"	'�԰����� 
	
	'--- 2004-08-19 by Byun Jee Hyun for UNI Code
    strVal = " "
	
	If Len(Request("txtMVFrDt")) Then
		strVal =  strVal & " AND mvmt.MVMT_DT >=  " & FilterVar(UNIConvDate(Request("txtMVFrDt")), "''", "S") & ""
	End If	
	
	If Len(Request("txtMVToDt")) Then
		strVal =  strVal & " AND mvmt.MVMT_DT <=  " & FilterVar(UNIConvDate(Request("txtMVToDt")), "''", "S") & ""
	End If
	
	If Len(Request("txtBeneficiary")) Then
		strVal = strVal & " AND mhdr.BP_CD =  " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
	end if
	
	If Len(Trim(Request("txtPurGrp"))) Then
		strVal = strVal & " AND mhdr.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("txtPayTerms"))) Then
		strVal = strVal & " AND mhdr.PAY_METH =  " & FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("txtPOType"))) Then
		strVal = strVal & " AND mhdr.PO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtPOType"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("txtMvmtType"))) Then
		strVal = strVal & " AND mvmt.IO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S") & " "
	End If
	
	UNIValue(0,0) = lgSelectList                                    '��: Select list
	UNIValue(0,1) = strVal											'	UNISqlId(0)�� �ι�° ?�� �Էµ�	
	
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") 				'������ 
	UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S") 					'���ű׷� 
	UNIValue(3,0) = FilterVar("B9004""))), " " , "S")																		'�������(Major_cd)
	UNIValue(3,1) = FilterVar(Trim(UCase(Request("txtPayTerms"))), " " , "S")					'������� 
	UNIValue(4,0) = FilterVar(Trim(UCase(Request("txtPOType"))), " " , "S") 					'�������� 
	UNIValue(5,0) = FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S") 				'�԰����� 
 
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
   ' UNIValue(0,UBound(UNIValue,2)) = " ORDER BY mvmt.MVMT_NO DESC "
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

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
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
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
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiaryNm = rs1("Bp_Nm")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtBeneficiary"))) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			EXIT FUNCTION
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrpNm = rs2("Pur_Grp_Nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtPurGrp"))) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			EXIT FUNCTION			
		End If
	End If  
	
	If Not(rs3.EOF Or rs3.BOF) Then
       strPaymeth = rs3("Minor_Nm")
   		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtPayTerms"))) Then
			Call DisplayMsgBox("970000", vbInformation, "�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			EXIT FUNCTION			
		End If
	End If  
	
	If Not(rs4.EOF Or rs4.BOF) Then
       strPoTypeNm = rs4("Po_TYPE_Nm")
   		Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Trim(Request("txtPOType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			EXIT FUNCTION			
		End If
	End If  	
	
	If Not(rs5.EOF Or rs5.BOF) Then
       strIOTypeNm = rs5("IO_TYPE_NM")
   		Set rs5 = Nothing
    Else
		Set rs5 = Nothing
		If Len(Trim(Request("txtMvmtType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "�԰�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			EXIT FUNCTION
		End If
	End If  	
	
	SetConditionData = TRUE
	
End Function

%>

<Script Language=vbscript>
    With parent
		.frm1.txtBeneficiaryNm.value	= "<%=ConvSPChars(strBeneficiaryNm)%>" 
		.frm1.txtPurGrpNm.value			= "<%=ConvSPChars(strPurGrpNm)%>" 
        .frm1.txtPayTermsNm.value		= "<%=ConvSPChars(strPaymeth)%>"
        .frm1.txtPOTypeNm.value			= "<%=ConvSPChars(strPoTypeNm)%>"
        .frm1.txtMvmtTypeNm.value		= "<%=ConvSPChars(strIOTypeNm)%>"
        
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHBeneficiary.value		= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
				.frm1.txtHPurGrp.value			= "<%=ConvSPChars(Request("txtPurGrp"))%>"
				.frm1.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"
				.frm1.txtHPOType.value			= "<%=ConvSPChars(Request("txtPOType"))%>"
				.frm1.txtHMvmtType.value		= "<%=ConvSPChars(Request("txtMvmtType"))%>"
				.frm1.txtHMVFrDt.Value			= "<%=Request("txtMVFrDt")%>"
				.frm1.txtHMVToDt.Value			= "<%=Request("txtMVToDt")%>"
			End If    
			'Show multi spreadsheet data from this line
				       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                            '��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
