<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : �������� Popup
'*  3. Program ID           : m4111pb4
'*  4. Program Name         : �԰�����˾� 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/02/18
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee, Eun Hee
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            
'**************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	On Error Resume Next
	Err.Clear
														   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
	Dim rs1, rs2, rs3, rs4
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim iTotstrData
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

	Dim strMvmtType
	Dim strMvmtTypenm	
	Dim strBeneficiaryNm
	
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
	On Error Resume Next			
    Err.Clear
    			
    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	Dim strEnd
	
	Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(2,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
	
	UNISqlId(0) = "M4111PA501"											' main query(spread sheet�� �ѷ����� query statement)
	UNISqlId(1) = "M4111PA302"											'�԰�����(STO������ ���� ����(KJH:03-01-06)
	UNISqlId(2) = "s0000qa024"	'����ó 
	
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    strEnd   = "ZZZZZZZZZZZZZZZZZZ"
	strVal = " "
	
	If Len(Trim(Request("txtSupplierCd"))) Then
		strVal = strVal & " AND A.BP_CD =  " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("txtPlantCd"))) Then
		strVal = strVal & " AND A.PLANT_CD =  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("txtItemCd"))) Then
		strVal = strVal & " AND A.ITEM_CD =  " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
	End If
	
	If Len(Request("txtMvmtType")) Then
		strVal = strVal & " AND A.IO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S") & " "
	Else	
		strVal = strVal & " AND A.IO_TYPE_CD <=  " & FilterVar(strEnd , "''", "S") & " "
	End If
	
	If Len(Request("txtPoNo")) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "S") & " "		
	Else	
		strVal = strVal & " AND A.PO_NO <=  " & FilterVar(strEnd , "''", "S") & " "
	End If	
	
	If Len(Trim(Request("txtFrRcptDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT >= " & FilterVar(UNIConvDate(Request("txtFrRcptDt")), "''", "S") & ""
	Else
		strVal = strVal & " AND A.MVMT_RCPT_DT >=" & "" & FilterVar("1900/01/01", "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToRcptDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <= " & FilterVar(UNIConvDate(Request("txtToRcptDt")), "''", "S") & ""		
	Else
		strVal = strVal & " AND A.MVMT_RCPT_DT <=" & "" & FilterVar("2999/12/30", "''", "S") & ""		
	End If
	
	UNIValue(0,0) = lgSelectList                                    '��: Select list
	UNIValue(0,1) = strVal											'	UNISqlId(0)�� �ι�° ?�� �Էµ�	
	
	UNIValue(1,0) = " AND A.IO_TYPE_CD =  " & FilterVar(FilterVar(Request("txtMvmtType")),"","SNM", "''", "S") & " "	'�԰����� 
    
    UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtSupplierCd"))), "''" , "S") 			    	'����ó 
	UNIValue(2,1) = " AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "											'��ܰŷ�ó�� 
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
    
    On Error Resume Next			
    Err.Clear			
    
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
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim  PvArr
    
    On Error Resume Next			
    Err.Clear
    
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
		strMvmtType = rs1(0)
        strMvmtTypenm = rs1(1)
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtMvmtType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "�԰�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strBeneficiaryNm = rs2("Bp_Nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtSupplierCd"))) Then
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
	SetConditionData = true
End Function

%>

<Script Language=vbscript>
    With parent
		.frm1.txtMvmtType.value = "<%=ConvSPChars(strMvmtType)%>" 
      	.frm1.txtMvmtTypeNm.value = "<%=ConvSPChars(strMvmtTypenm)%>" 
      	.frm1.txtSupplierNm.value = "<%=ConvSPChars(strBeneficiaryNm)%>" 

      	If "<%=lgDataExist%>" = "Yes" Then
			
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			    .frm1.hdnPlantCd.Value	 	= "<%=ConvSPChars(Request("txtPlantCd"))%>"				
				.frm1.hdnItemCd.Value 		= "<%=ConvSPChars(Request("txtItemCd"))%>"
				.frm1.hdnSupplierCd.Value 	= "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.frm1.hdnFrRcptDt.Value 	= "<%=Request("txtFrRcptDt")%>"
				.frm1.hdnToRcptDt.value     = "<%=Request("txtToRcptDt")%>"
				.frm1.hdnMvmtType.Value 	= "<%=ConvSPChars(Request("txtMvmtType"))%>" 				
				.frm1.hdnPoNo.Value 	    = "<%=ConvSPChars(Request("txtPoNo"))%>"
			End If    
			
			.ggoSpread.Source    = .frm1.vspdData 
			.frm1.vspdData.Redraw = false
			.ggoSpread.SSShowData "<%=iTotstrData%>","F"          '�� : Display data
       
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",21),.GetKeyPos("A",18),"C","I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",21),.GetKeyPos("A",19),"A","I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",21),.GetKeyPos("A",20),"A","I","X","X")
		
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = true
		End If
	End with
</Script>	
 	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

