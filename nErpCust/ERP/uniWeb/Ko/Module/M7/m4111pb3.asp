<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : �԰��ȣ Popup
'*  3. Program ID           : M4111PB1
'*  4. Program Name         : 
'*  5. Program Desc         : �����԰���ȭ���� �԰��ȣ �˾� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
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
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%							                           '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	'On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				
	
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
	Dim rs1, rs2, rs3, rs4
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim iTotstrData
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

	Dim strBeneficiaryNm
	Dim strPurGrpNm
	Dim strMvmtTypenm	
	
	Dim iPrevEndRow
    Dim iEndRow
	
	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
	

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	
	iPrevEndRow = 0
    iEndRow = 0
    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	Dim strEnd
	
	Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(3,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
	
	
    UNISqlId(0) = "M4111PA301"											' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(1) = "s0000qa024"	'����ó 
    UNISqlId(2) = "s0000qa019"	'���ű׷� 
    UNISqlId(3) = "M4111PA302"	'�԰�����(STO������ ���� ����(KJH:03-01-06)
    
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	
	strVal = " "

	If Len(Request("txtFrRcptDt")) Then 
		strVal = strVal & " AND A.MVMT_RCPT_DT >=  " & FilterVar(UNIConvDate(Request("txtFrRcptDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtToRcptDt")) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <=  " & FilterVar(UNIConvDate(Request("txtToRcptDt")), "''", "S") & " "
	End If
			
	
	If Len(Request("txtMvmtType")) Then
	  strVal = strVal & " AND B.IO_TYPE_CD =  " & FilterVar(Request("txtMvmtType"), "''", "S") & " "
	ELSE	
	   strVal = strVal & " AND B.IO_TYPE_CD <=  " & FilterVar(strEnd , "''", "S") & " "
	end if	    
    		
	If Len(Trim(Request("txtSupplier"))) Then
		strVal = strVal & " AND A.BP_CD =  " & FilterVar(Request("txtSupplier"), "''", "S") & " "
	ELSE	
		strVal = strVal & " AND A.BP_CD <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "
	End If
	
	IF LEN(Trim(Request("txtGroup"))) THEN
		strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Request("txtGroup"), "''", "S") & " "
	ELSE
	    strVal = strVal & " AND A.PUR_GRP <=  " & FilterVar(LEFT(strEnd,10), "''", "S") & " "  
	END IF	 
 
    strVal = strVal & " AND (A.DLVY_ORD_FLG <> " & FilterVar("Y", "''", "S") & "  OR A.DLVY_ORD_FLG is Null) "
	
	UNIValue(0,0) = lgSelectList                                    '��: Select list
	UNIValue(0,1) = strVal											'	UNISqlId(0)�� �ι�° ?�� �Էµ�	
'	call svrmsgbox(strVal, vbinformation, i_mkscript)
	UNIValue(1,0) = FilterVar(Trim(Request("txtSupplier"))," ","S") 			    	'����ó 
	UNIValue(1,1) = " AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "											'��ܰŷ�ó�� 
	UNIValue(2,0) = FilterVar(Trim(Request("txtGroup"))," ","S") 						'���ű׷� 
    UNIValue(3,0) = " AND A.IO_TYPE_CD =  " & FilterVar(Request("txtMvmtType"), "''", "S") & "  "	'�԰����� 
    
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
    Dim FalsechkFlg
    
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
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim PvArr
    
    Const C_SHEETMAXROWS_D = 100     
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 
	End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
            PvArr(iLoopCount) = lgstrData	
		    lgstrData = ""
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
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
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
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
			
			If "<%=lgPageNo%>" = "1" Then  
			    .frm1.hdnMvmtType.Value	 	= "<%=ConvSPChars(Request("txtMvmtType"))%>"				
				.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnFrRcptDt.Value 	= "<%=Request("txtFrRcptDt")%>"
				.frm1.hdnToRcptDt.Value 	= "<%=Request("txtToRcptDt")%>"
				.frm1.hdnInspFlag.value     = "<%=ConvSPChars(Request("txtInspFlag"))%>"
				.frm1.hdnGroup.Value 	    = "<%=ConvSPChars(Request("txtGroup"))%>" 				
			End If    
			
			.ggoSpread.Source    = .frm1.vspdData 
			.frm1.vspdData.Redraw = false
		    .ggoSpread.SSShowData "<%=iTotstrData%>"                            '��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
 	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

