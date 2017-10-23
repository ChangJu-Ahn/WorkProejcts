<!--
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M9211PB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/17
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : KO MYOUNG JIN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/04 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'*                            -2002/04/17 : ADO��ȯ 
'**************************************************************************************
%>
-->
<!-- #Include file="../../inc/IncServer.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				
Const SortNo = 2

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
dim istrData

Dim strBeneficiaryNm
Dim strPurGrpNm
Dim strPaymeth
Dim strIncoterms
Dim strMvmtTypenm	
	Call HideStatusWnd 

	

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"

    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	
	Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(3,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
	
	
    UNISqlId(0) = "M9211PA101"											' main query(spread sheet�� �ѷ����� query statement)
    'UNISqlId(1) = "S0000QA002"	'����ó 
    'UNISqlId(2) = "S0000QA022"	'���ű׷� 
    UNISqlId(1) = "s0000qa024"	'����ó 
    UNISqlId(2) = "s0000qa019"	'���ű׷� 
    UNISqlId(3) = "s0000qa023"	'�԰����� 
   
	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
	strVal = " "

	If Len(Request("txtFrRcptDt")) AND  Len(Request("txtToRcptDt")) Then
		strVal = " AND Z.MVMT_RCPT_DT >=  " & FilterVar(UNIConvDate(Request("txtFrRcptDt")), "''", "S") & " AND Z.MVMT_RCPT_DT <=  " & FilterVar(UNIConvDate(Request("txtToRcptDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtMvmtType")) Then
	  strVal = strVal & " AND Z.IO_TYPE_CD =  " & FilterVar(Request("txtMvmtType"), "''", "S") & " "
	ELSE	
	   strVal = strVal & " AND Z.IO_TYPE_CD <=  " & FilterVar(strEnd , "''", "S") & " "
	end if	    
    		
	If Len(Trim(Request("txtSupplier"))) Then
		strVal = strVal & " AND Z.BP_CD =  " & FilterVar(Request("txtSupplier"), "''", "S") & " "
	ELSE	
		strVal = strVal & " AND Z.BP_CD <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "
	End If
	
	IF LEN(Trim(Request("txtGroup"))) THEN
		strVal = strVal & " AND Z.PUR_GRP =  " & FilterVar(Request("txtGroup"), "''", "S") & " "
	ELSE
	    strVal = strVal & " AND Z.PUR_GRP <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "  
	END IF	 
	
	strVal = strVal & "AND B.RCPT_FLG = " & FilterVar("Y", "''", "S") & "  " 
	strVal = strVal & "AND B.RET_FLG = " & FilterVar("N", "''", "S") & "  " 
	strVal = strVal & "order by Z.MVMT_RCPT_NO desc " 	  
   
        
	UNIValue(0,0) = lgSelectList                                    '��: Select list
	UNIValue(0,1) = strVal											'	UNISqlId(0)�� �ι�° ?�� �Էµ�	
	
	UNIValue(1,0) = FilterVar(Trim(Request("txtSupplier"))," ","S") 			    	'����ó 
	UNIValue(2,0) = FilterVar(Trim(Request("txtGroup"))," ","S") 						'���ű׷� 
    UNIValue(3,0) = FilterVar(Trim(Request("txtMvmtType"))," ","S") 						'�԰����� 
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
  
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
	
	
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If SetConditionData = false then Exit Sub
         
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
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

   iLoopCount = 0
   Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""

 		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(0))           '�԰��ȣ 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(1))		    '�԰����� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(2))		    '�԰����¸� 
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(3))   '�԰����� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(4))		    '����ó 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(5))		    '����ó�� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(6))	        '���ű׷�                                'ǰ��԰� '8	
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(7))	        '���ű׷�� 
		'iRowStr = iRowStr & Chr(11) & ""							'14								'27
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             
      
        If iLoopCount - 1 < lgMaxCount Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
  Loop
    
    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
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
   
    SetConditionData = false
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiaryNm = rs1("Bp_Nm")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtSupplier"))) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrpNm = rs2("Pur_Grp_Nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
	If Not(rs3.EOF Or rs3.BOF) Then
        strMvmtTypenm = rs3(1)
   		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtMvmtType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "�԰�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
	
	SetConditionData = true
	
End Function

%>

<Script Language=vbscript>
    With parent
		.frm1.txtSupplierNm.value = "<%=ConvSPChars(strBeneficiaryNm)%>" 
		.frm1.txtGroupNm.value = "<%=ConvSPChars(strPurGrpNm)%>" 
      	.frm1.txtMvmtTypeNm.value = "<%=ConvSPChars(strMvmtTypenm)%>" 

      	If "<%=lgDataExist%>" = "Yes" Then
			
			'Set condition data to hidden area			
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			    .frm1.hdnMvmtType.Value	 	= "<%=ConvSPChars(Request("txtMvmtType"))%>"				
				.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnFrRcptDt.Value 	= "<%=Request("txtFrRcptDt")%>"
				.frm1.hdnToRcptDt.Value 	= "<%=Request("txtToRcptDt")%>"
				.frm1.hdnInspFlag.value     = "<%=ConvSPChars(Request("txtInspFlag"))%>"
				.frm1.hdnGroup.Value 	    = "<%=ConvSPChars(Request("txtGroup"))%>" 				
			End If    
			
			'Show multi spreadsheet data from this line
				       
			.ggoSpread.Source    = .frm1.vspdData 
						
			.ggoSpread.SSShowData "<%=istrData%>"                            '��: Display data 
			
			'.vspdSort .C_PoNo, "<%=SortNo%>"							'��: Sort Bp_cd 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
			
		End If
	End with
</Script>	
 	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

