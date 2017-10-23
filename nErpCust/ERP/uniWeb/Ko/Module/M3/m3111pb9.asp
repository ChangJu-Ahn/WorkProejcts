<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111pb9
'*  4. Program Name         : ���ֹ�ȣ 
'*  5. Program Desc         : ���ֹ�ȣ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/20		
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Min, HJ	
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgPageNo                                                           '�� : ���� �� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strRcpt
Dim strIv
Dim strSubcontra														   '�� : ���ְ������� 
'----------------------- �߰��� �κ� ----------------------------------------------------------------------
Dim arrRsVal(10)								'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
Dim iFrPoint
iFrPoint=0
'----------------------------------------------------------------------------------------------------------
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)              '�� : Next key flag
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call TrimData()

	strRcpt = Request("txtRcptFlg")
	strIv = Request("txtIvFlg")
	strSubcontra = Request("txtSubcontraFlg")

	Err.Clear												'��: Protect system from crashing

	If Len(Trim(Request("txtFrPoDt"))) Then
		If UNIConvDate(Request("txtFrPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrPoDt", 0, I_MKSCRIPT)
		    Response.End	
		End If
	End If
		
	If Len(Trim(Request("txtToPoDt"))) Then
		If UNIConvDate(Request("txtToPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToPoDt", 0, I_MKSCRIPT)
		    Response.End	
		End If
	End If
	
	Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    
    Dim strVal
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,2)
    
	If Trim(strRcpt) = "N" And Trim(strIv) = "N" Then		'��ǰ ��� (RI) 
	    UNISqlId(0) = "M3111pa901"									'* : ������ ��ȸ�� ���� SQL�� ���� 
	End If
	If Trim(strRcpt) = "Y" And Trim(strIv) = "N" Then		'��ǰ�԰�(RR)
		UNISqlId(0) = "M3111pa902"									'* : ������ ��ȸ�� ���� SQL�� ���� 
	End If
	If Trim(strRcpt) = "N" And Trim(strIv) = "Y" Then		'��ǰ ��� (RF) - �������� 
		UNISqlId(0) = "M3111pa903"									'* : ������ ��ȸ�� ���� SQL�� ���� 
	End If
	
	UNISqlId(1) = "M3111QA103"								              '�������¸� 
    UNISqlId(2) = "M3111QA102"											  '����ó��   
    UNISqlId(3) = "M3111QA104"											  '���ű׷�� 
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     
	strVal = " "
	If Trim(strSubcontra) <> "" Then	' ���ְ������� �߰� 
	    strVal = strVal & " AND A.SUBCONTRA_FLG = " & FilterVar(Trim(strSubcontra), " " , "S") & " "
	End If
	
	If Trim(Request("txtPotype")) <> "" Then
		strVal = strVal & " AND A.PO_TYPE_CD >= " & FilterVar(Trim(UCase(Request("txtPotype"))), " " , "S") & "  AND A.PO_TYPE_CD <=  " & FilterVar(Trim(UCase(Request("txtPotype"))), " " , "S") & " "
	Else
		strVal = strVal & " AND A.PO_TYPE_CD >='' AND A.PO_TYPE_CD <= " & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If
	
    If Trim(Request("txtSupplier")) <> "" Then
		strVal = strVal & " AND A.BP_CD >= " & FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S") & "  AND A.BP_CD <=  " & FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S") & " "
	Else
		strVal = strVal & " AND A.BP_CD >='' AND A.BP_CD <= " & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If
    
    If Len(Trim(Request("txtFrPoDt"))) Then
    	strVal = strVal & " AND A.PO_DT >= " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & ""		
	Else
    	strVal = strVal & " AND A.PO_DT >=" & FilterVar("1900/01/01", "''", "S") & ""
    End If

	If Len(Trim(Request("txtToPoDt"))) Then
		strVal = strVal & " AND A.PO_DT <= " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & ""		
	Else
    	strVal = strVal & " AND A.PO_DT <=" & FilterVar("2999/12/30", "''", "S") & ""
    End If

	If Trim(Request("txtGroup")) <> "" Then
		strVal = strVal & " AND A.PUR_GRP >= " & FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S") & "  AND A.PUR_GRP <=  " & FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S") & " "
	Else
		strVal = strVal & " AND A.PUR_GRP >='' AND A.PUR_GRP <= " & FilterVar("zzzzzzzzz", "''", "S") & ""		
	End If
	
    UNIValue(0,1) = strVal   
    UNIValue(1,0)  = " " & FilterVar(Trim(UCase(Request("txtPotype"))), " " , "S") & " "
    UNIValue(2,0)  = " " & FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S") & " "
    UNIValue(3,0)  = " " & FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S") & " "      
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim iStr
   
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
     
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
   
    IF SetConditionData() = FALSE THEN EXIT SUB
   
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    
    Else
		Call  MakeSpreadSheetData()        
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
	Dim PvArr
    
    lgstrData      = ""
  
  	If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgPageNoIndex : Previous PageNo
		iFrPoint    = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
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
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        PvArr(iLoopCount) = lgstrData
        lgstrData=""
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '��: Check if next data exists
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

    '============================= �߰��� �κ� =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtPotype")) Then
           Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       Exit Function
		End If
    Else    
		arrRsVal(1) = rs1(0)
		arrRsVal(2) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtSupplier")) Then
		   Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       Exit Function
		End If
    Else    
		arrRsVal(3) = rs2(0)
		arrRsVal(4) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtGroup")) Then
		   Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       Exit Function
		End If
    Else    
		arrRsVal(5) = rs3(0)
		arrRsVal(6) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

End Sub


%>

<Script Language=vbscript>
    With parent
        
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowData "<%=lgstrData%>","F"                            '��: Display data 

	    Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",7), Parent.GetKeyPos("A",6),"A", "I" ,"X","X")	'���ֱݾ� 

  		.frm1.hdnFrDt.value  =  "<%=ConvSPChars(Request("txtFrPoDt"))%>" 	
  		.frm1.hdnToDt.value	 =  "<%=ConvSPChars(Request("txtToPoDt"))%>" 	
  		.frm1.hdnRcptFlg.Value 		= "<%=strRcpt%>"
		.frm1.hdnIvFlg.Value 		= "<%=strIv%>"
				
  		.frm1.txtPotypeNm.Value 	= "<%=ConvSPChars(arrRsVal(2))%>"
		.frm1.txtSupplierNm.Value	= "<%=ConvSPChars(arrRsVal(4))%>"
		.frm1.txtGroupNm.Value 		= "<%=ConvSPChars(arrRsVal(6))%>"
		.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag

		.frm1.hdnPotype.Value	 	= "<%=ConvSPChars(Request("txtPotype"))%>"
		.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplier"))%>"
		.frm1.hdnFrDt.Value 		= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
		.frm1.hdnToDt.Value 		= "<%=ConvSPChars(Request("txtToPoDt"))%>"
		.frm1.hdnGroup.Value 		= "<%=ConvSPChars(Request("txtGroup"))%>"
		.frm1.hdnRcptFlg.Value 		= "<%=ConvSPChars(Request("txtRcptFlg"))%>"
		.frm1.hdnIvFlg.Value 		= "<%=ConvSPChars(Request("txtIvFlg"))%>"
		.DbQueryOk
	End with
</Script>	

<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

