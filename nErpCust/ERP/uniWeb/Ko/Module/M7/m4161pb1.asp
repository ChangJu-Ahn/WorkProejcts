<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4161pb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : ���ó����ȣPopup															*
'*  7. Modified date(First) :																			*
'*  8. Modified date(Last)  : 2003/05/23																*
'*  9. Modifier (First)     : 																			*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	On Error Resume Next
	                                                                         
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1,rs2,rs3   '�� : DBAgent Parameter ���� 
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim iTotstrData
	Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim SortNo													  ' Sort ���� 
	
	Dim iPrevEndRow
    Dim iEndRow
    
	DIM strMvmtType
	DIM	strSupplier
	Dim strGroup 

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
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim  PvArr
    
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
        strMvmtType =  rs1("IO_TYPE_NM")
        Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtMvmtType")) Then
			Call DisplayMsgBox("970000", vbInformation, "�԰�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
    If Not(rs2.EOF Or rs2.BOF) Then
        strSupplier =  rs2("BP_NM")
        Set rs2 = Nothing
	Else
		Set rs2 = Nothing
		If Len(Request("txtSupplier")) Then
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
    If Not(rs3.EOF Or rs3.BOF) Then
        strGroup =  rs3("PUR_GRP_NM")
        Set rs3 = Nothing
	Else
		Set rs3 = Nothing
		If Len(Request("txtGroup")) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  		
	
	
	SetConditionData = true
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(3,2)

    UNISqlId(0) = "M4161PA101"
    UNISqlId(1) = "S0000QA023"		'// �԰����� 
    UNISqlId(2) = "S0000QA024"       '// ����ó 
    UNISqlId(3) = "M3111QA104"		'// ���ű׷� 

    UNIValue(0,0) = Trim(lgSelectList)    
    
	strVal = " "

	If Len(Request("txtMvmtType")) Then
		strVal = strVal & " AND D.IO_TYPE_CD = " & FilterVar(UCase(Request("txtMvmtType")), "''", "S") & "  "	
	End If
	arrVal(0) = FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S")

	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND C.BP_CD = " & FilterVar(UCase(Request("txtSupplier")), "''", "S") & "  "	
	End If
	arrVal(1) = FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S")

	If Len(Request("txtFrRcptDt")) > 0 Then
		strVal = strVal & " AND A.MVMT_DT >= " & FilterVar(UniConvDate(Request("txtFrRcptDt")), "''", "S") & " "	
	End If
	
	If Len(Request("txtToRcptDt")) > 0 Then
		strVal = strVal & " AND A.MVMT_DT <= " & FilterVar(UniConvDate(Request("txtToRcptDt")), "''", "S") & " "	
	End If	
	
	If Len(Request("txtGroup")) Then
		strVal = strVal & " AND B.PUR_GRP = " & FilterVar(UCase(Request("txtGroup")), "''", "S") & "  "	
	End If
	arrVal(2) = " " & FilterVar(UCase(Request("txtGroup")), "''", "S") & " "
	
	UNIValue(0,1) = strVal 
	UNIValue(1,0) = arrVal(0)									'�԰����� 
    UNIValue(2,0) = arrVal(1)									'����ó 
    UNIValue(2,1) = " AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "					'��ܰŷ�ó�� 
    UNIValue(3,0) = arrVal(2)									'���ű׷� 
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2,rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If SetConditionData = false then Exit Sub
         
    If  rs0.EOF And rs0.BOF Then  '// And FalsechkFlg =  False
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
		.frm1.txtMoveTypeNm.value	= "<%=ConvSPChars(strMvmtType)%>" 
		.frm1.txtSupplierNm.value	= "<%=ConvSPChars(strSupplier)%>" 
		.frm1.txtGroupNm.value	= "<%=ConvSPChars(strGroup)%>" 

		If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then           
				.frm1.hdnMvmtType.Value	 	= "<%=ConvSPChars(Request("txtMvmtType"))%>"
				.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnFrDt.Value 		= "<%=Request("txtFrRcptDt")%>"
				.frm1.hdnToDt.Value 		= "<%=Request("txtToRcptDt")%>"
				.frm1.hdnGroup.Value 		= "<%=ConvSPChars(Request("txtGroup"))%>"	
			End If    
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.frm1.vspdData.Redraw = false
			.ggoSpread.SSShowData "<%=iTotstrData%>"                  '��: Display data 
																	'0: ����None 1 :��������  2: ��������					
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = true
		End If
	End with
</Script>	
