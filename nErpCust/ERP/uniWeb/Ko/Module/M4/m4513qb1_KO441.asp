<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M4513QA1
'*  4. Program Name         : 
'*  5. Program Desc         : ����������Ȳ ��ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Eun Hee
'* 10. Modifier (Last)      : 
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
<%	
	Call HideStatusWnd
	call LoadBasisGlobalInf()
	call LoadInfTB19029B("Q", "M","NOCOOKIE","QB") 
	call LoadBNumericFormatB("Q","M","NOCOOKIE","QB")

    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2 ,rs3                '�� : DBAgent Parameter ���� 
    Dim lgTailList
    Dim lgPageNo
    Dim istrData
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim iLngRow
                                                                       
	Dim iTotstrData
	
	Dim strPurGrpNm
	Dim strBeneficiaryNm
 
   ' On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    'On Error Resume Next
	Err.Clear
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)  
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	
	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim iStrSQL
    Dim strVal, strVal1
	Redim UNISqlId(3)														'��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(3,1)														'��: DB-Agent�� ���۵� parameter�� ���� ���� 
																			'parameter�� ���� ���� ������ 
    UNISqlId(0) = "M4513QA101"												'Detial 											
	UNISqlId(1) = "s0000qa024"	'������ 
    UNISqlId(2) = "S0000QA022"	'���ű׷� 
    
	strVal = ""
	
	If Len(Trim(Request("txtPurGrpCd"))) Then
		strVal = strVal & " AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPurGrpCd"))), " " , "S") & "  "
	End If
	
	IF Len(Trim(Request("txtBpCd"))) Then
		strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "S") & "  "
	End If
	
	If Len(Request("txtPoFrDt")) Then
		strVal = strVal & " AND A.PO_DT >=  " & FilterVar(UNIConvDate(Request("txtPoFrDt")), "''", "S") & " "
	End IF
	
	If Len(Request("txtPoToDt")) Then
		strVal = strVal & " AND A.PO_DT <=  " & FilterVar(UNIConvDate(Request("txtPoToDt")), "''", "S") & " "
	End If
	
	IF Len(Trim(Request("txtPoNo"))) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "S") & "  "
	End If

     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND A.PUR_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND A.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND A.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
	
	UNIValue(0,0) = strVal
	UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtBpCd"))), "''" , "S") 				'������ 
	UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtPurGrpCd"))), "''" , "S")					'���ű׷�	
		
	
	UNIValue(0,UBound(UNIValue,2)) = " ORDER BY A.PO_NO DESC "
	UNILock = DISCONNREAD :	UNIFlag = "1"                                
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2,rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
		    
	If Not(rs1.EOF Or rs1.BOF) Then
        strBeneficiaryNm = rs1("Bp_Nm")
   		rs1.Close
   		Set rs1 = Nothing
    Else
		rs1.Close
		Set rs1 = Nothing
		If Len(Trim(Request("txtBeneficiary"))) Then
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			exit Sub
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrpNm = rs2("Pur_Grp_Nm")
   		rs2.Close
   		Set rs2 = Nothing
    Else
		rs2.Close
		Set rs2 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			exit Sub			
		End If
	End If 
    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"				& vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & iTotstrData	    & """ " & vbCr
   	
	Response.Write "	.lgPageNo  = """ & lgPageNo   & """"			& vbCr 
    Response.Write "	.frm1.hdnPoFrDt.value		= """ & Trim(ConvSPChars(Request("txtPoFrDt")))           & """" & vbCr
    Response.Write "	.frm1.hdnPoToDt.value		= """ & Trim(ConvSPChars(Request("txtPoToDt")))           & """" & vbCr
    Response.Write "	.frm1.hdnPoNo.value			= """ & Trim(ConvSPChars(Request("txtPoNo")))             & """" & vbCr
    Response.Write "	.frm1.hdnPurGrpCd.value		= """ & Trim(ConvSPChars(Request("txtPurGrpCd")))             & """" & vbCr
    Response.Write "	.frm1.hdnBpCd.value			= """ & Trim(ConvSPChars(Request("txtBpCd")))           & """" & vbCr
	
	Response.Write "	.frm1.txtPurGrpNm.value		= """ & ConvSPChars(strPurGrpNm)              				& """" & vbCr
	Response.Write "	.frm1.txtBpNm.value			= """ & ConvSPChars(strBeneficiaryNm)              				& """" & vbCr
		
	Response.Write "    .DbQueryOk"										& vbCr
	Response.Write "End With"											& vbCr
    Response.Write "</Script>"											& vbCr        

End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Dim iLoopCount  
	Dim TmpAmt
	
	Const C_SHEETMAXROWS_D  = 100
    
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

   iLoopCount = -1
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
   
   Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
				
        iRowStr  = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("PO_NO"))	                       
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("PAY_METH"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("PO_CUR"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("PO_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("BP_CD"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("LC_NO"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("LC_DOC_NO"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("LC_AMEND_SEQ"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("OPEN_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("LC_TYPE"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("BL_NO"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("BL_DOC_NO"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("LOADING_DT"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("BL_ISSUE_DT"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("SETLMNT_DT"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("DISCHGE_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CC_NO"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ID_NO"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("ID_DT"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("MVMT_RCPT_NO"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("MVMT_RCPT_DT"))
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount - 1

        If iLoopCount < C_SHEETMAXROWS_D Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = istrData	
		   istrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
	    rs0.MoveNext
	Loop
	
	
	iLngRow = iLoopCount
	iTotstrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

%>
