<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : ��������� ���� 
'*  3. Program ID           : M4141PB1 
'*  4. Program Name         : Receipt No Popup Biz.
'*  5. Program Desc         : ���Ź�ǰ�� ������ȣ �˾� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2003/05/28
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'*							  ADO Conv. 	
'**************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1, rs2, rs3	'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim iTotstrData
Dim lgpageNo	                                            '�� : ���� �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim iPrevEndRow
Dim iEndRow

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim strIOtypeNm, strSupplierNm, strGroupNm

Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
     
lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)              '�� : Next key flag
lgSelectList    = Request("lgSelectList")
lgTailList      = Request("lgTailList")
lgSelectListDT  = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgDataExist     = "No"
iPrevEndRow = 0
iEndRow = 0	 

Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	Dim sTemp
	
	Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(4,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
    UNISqlId(0) = "U1123PA101"
    UNISqlId(1) = "s0000qa025"	'��������� 
    UNISqlId(2) = "s0000qa024"	'����ó 
    UNISqlId(3) = "s0000qa019"	'���ű׷� 
    
    '--- 2004-08-20 by Byun Jee Hyun for UNICODE	
    
    '������� 
    If Len(Trim(Request("txtFrRcptDt"))) Then
		strVal = strVal & " AND gdmvmt.MVMT_RCPT_DT >=  " & FilterVar(UniConvDate(Request("txtFrRcptDt")), "''", "S") & " "	
	End If
			
    If Len(Trim(Request("txtToRcptDt"))) Then
		strVal = strVal & " AND gdmvmt.MVMT_RCPT_DT <=  " & FilterVar(UniConvDate(Request("txtToRcptDt")), "''", "S") & " "	
	End If
	
    '��������� 
    If Len(Trim(Request("txtMvmtType"))) Then
		strVal = strVal & " AND mvtype.IO_TYPE_CD =  " & FilterVar(UCase(Request("txtMvmtType")), "''", "S") & "  "	
	End If
    
    '����ó 
    If Len(Trim(Request("txtSupplier"))) Then
		strVal = strVal & " AND bizpart.BP_CD =  " & FilterVar(UCase(Request("txtSupplier")), "''", "S") & "  "	
	End If
    
	'���ű��� 
	If Len(Trim(Request("txtGroup"))) Then
		strVal = strVal & " AND pgrp.PUR_GRP =  " & FilterVar(UCase(Request("txtGroup")), "''", "S") & "  "	
	End If
    
	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select �� �ʵ� 
	UNIValue(0,1) = strVal													'---WHERE �� 
	
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtMvmtType"))), " " , "S")
	UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S")
	UNIValue(2,1) = " AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "											'��ܰŷ�ó�� 
	UNIValue(3,0) = FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S")
    
    UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)					
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
     
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
   
    If SetConditionData = False Then Exit Sub
   
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
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
        strIOtypeNm = rs1("IO_TYPE_NM")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtMvmtType"))) Then
			Call DisplayMsgBox("970000", vbInformation, "���������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If   	
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSupplierNm = rs2("BP_NM")
		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtSupplier"))) Then
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If			
    End If     
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strGroupNm = rs3("PUR_GRP_NM")
		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If			
    End If  
     
    SetConditionData = True 
     
End Function 


%>

<Script Language=vbscript>
    With parent

		.frm1.txtMoveTypeNm.value	= "<%=ConvSPChars(strIOtypeNm)%>"
		.frm1.txtSupplierNm.value	= "<%=ConvSPChars(strSupplierNm)%>"
		.frm1.txtGroupNm.value		= "<%=ConvSPChars(strGroupNm)%>"

		If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				
				.frm1.hdnMvmtType.Value 	= "<%=ConvSPChars(Request("txtMvmtType"))%>"
				.frm1.hdnSupplier.Value 	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnFrDt.Value 		= "<%=Request("txtFrRcptDt")%>"
				.frm1.hdnToDt.Value 		= "<%=Request("txtToRcptDt")%>"
				.frm1.hdnGroup.Value 		= "<%=ConvSPChars(Request("txtGroup"))%>"
				.frm1.hdnRcptFlg.Value 		= "<%=ConvSPChars(Request("txtRcptFlg"))%>"			
			End If    
			.ggoSpread.Source = .frm1.vspdData 
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=iTotstrData%>"					'��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
