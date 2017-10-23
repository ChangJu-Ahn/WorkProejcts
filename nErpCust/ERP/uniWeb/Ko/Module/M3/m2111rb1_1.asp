<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m2111rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Purchase Order Detail ���� PopUp ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2002/04/23																*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Kim Jae Soon																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************
%>
	
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0       		   '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgStrData_1
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strPtnBpNm												  ' ��ǰó�� 
Dim strDNTypeNm												  ' �������¸� 
Dim strSOTypeNm											      ' ����Ÿ�Ը� 
Dim gridNum													  ' �׸��� ���� Ȯ�� 

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	'�̼��� �߰� 
	lgPageNo_1       = UNICInt(Trim(Request("lgPageNo_1")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	
	gridNum			= Request("txtGridNum")

	Call FixUNISQLData(gridNum)									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query
 
 '----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount 
    '�̼��� �߰� 
    Dim iLoopCount_1                                                                    
    Dim iRowStr,iRowStr_1
    Dim ColCnt
    
    lgDataExist    = "Yes"

	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    If gridNum = "B" then
    
	    iLoopCount_1 = -1
        
		lgstrData_1	   = ""
		
		Do while Not (rs0.EOF Or rs0.BOF)
   
		     iLoopCount_1 =  iLoopCount_1 + 1
		     iRowStr_1 = ""
		     
				For ColCnt = 0 To UBound(lgSelectListDT) - 1 
					iRowStr_1 = iRowStr_1 & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
				Next
 
'		     If iLoopCount < lgMaxCount Then
			 If iLoopCount_1 < lgMaxCount Then
		        lgstrData_1 = lgstrData_1 & iRowStr_1 & Chr(11) & Chr(12)
		     Else
		        lgPageNo_1 = lgPageNo_1 + 1
		        Exit Do
		     End If
		     
		     rs0.MoveNext
		Loop
	Else
		
		iLoopCount = -1

		lgstrData      = ""
		Do while Not (rs0.EOF Or rs0.BOF)
   
		     iLoopCount =  iLoopCount + 1
		     iRowStr = ""
		     
				For ColCnt = 0 To UBound(lgSelectListDT) - 1 
		         iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
				Next
 
		     If iLoopCount < lgMaxCount Then
		        lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
		     Else
		        'call svrmsgbox(lgPageNo , vbinformation, i_mkscript)
		        lgPageNo = lgPageNo + 1
		        Exit Do
		     End If
		     
		     rs0.MoveNext
		Loop
	End if
	
    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    If iLoopCount_1 < lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo_1 = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub
   
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData(byVal gridNum)
	
    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,2)
    
	strVal = " "
	If Len(Request("txtSoNo")) Then
		strVal = strVal & "AND A.SO_NO like " & FilterVar(Trim(UCase(Request("txtSoNo"))), " " , "S") & ""	
	End If

	If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO like " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & " "		
	End If		
	
	If Len(Request("txtSupplier")) Then
		If Request("txtSTOflg") = "Y" then					
 			strVal = strVal & " AND g.SPPL_CD like "		'2002-12-16(LJT)
 		Else
 			strVal = strVal & " AND g.SPPL_CD like " & FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S") & " "		 			
 		End If
	End If	
    
 	If Len(Request("txtGroup")) Then
 			strVal = strVal & " AND g.PUR_GRP like " & FilterVar(Trim(UCase(Request("txtGroup"))), " " , "S") & " "		
	End If	    
 	
 	If Len(Request("txtProcure")) Then
			strVal = strVal & " AND a.procure_type like " & FilterVar(Trim(UCase(Request("txtProcure"))), " " , "S") & " "		
	End If	    
	
    If Len(Request("txtFrPoDt")) Then
		strVal = strVal & " AND g.PUR_PLAN_DT >= '" & UNIConvDate(Request("txtFrPoDt")) & "' "			
	else
	    strVal = strVal & " AND g.PUR_PLAN_DT >= '1900-01-01' "			
	End If		
	
	If Len(Request("txtToPoDt")) Then
		strVal = strVal & " AND g.PUR_PLAN_DT <= '" & UNIConvDate(Request("txtToPoDt")) & "' "		
	else
	    strVal = strVal & " AND g.PUR_PLAN_DT <= '2999-12-31' "		
	End If
	
	If Len(Request("txtFrDlvyDt")) Then
		strVal = strVal & " AND A.DLVY_DT >= '" & UNIConvDate(Request("txtFrDlvyDt")) & "' "		
	End If		
	
	If Len(Request("txtToDlvyDt")) Then
		strVal = strVal & " AND A.DLVY_DT <= '" & UNIConvDate(Request("txtToDlvyDt")) & "' "		
	End If	
	
	'�̼��� �߰� PlantCD
	If Len(Request("txtPlantCd")) Then
		strVal = strVal & " AND A.PLANT_CD = " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "		
	End If	
	
	if gridNum = "B" then
		UNISqlId(0) = "M2111RA1_1_2"
		
		UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
		UNIValue(0,1) = strVal & "Group by b.pur_grp,B.PUR_GRP_NM,g.SPPL_CD,E.BP_NM, a.procure_type " & UCase(Trim(lgTailList)) 
    else

		UNISqlId(0) = "M2111RA1_1_1"
	
		UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
		UNIValue(0,1) = strVal & UCase(Trim(lgTailList)) 
		
	End if
	
	
'--------------- ������ coding part(�������,End)------------------------------------------------------
   
    
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
        
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
    iStr = Split(lgstrRetMsg,gColSep)

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
    
    Set lgADF   = Nothing
    
End Sub



%>

<Script Language=vbscript>
    With parent
		
		If "<%= gridNum %>" = "B" Then
			If "<%=lgDataExist%>" = "Yes" Then
				.frm1.hdnFrDt.Value 		= "<%=Request("txtFrPoDt")%>"
				.frm1.hdnToDt.Value 		= "<%=Request("txtToPoDt")%>"
				.frm1.hdnFrDt2.Value 		= "<%=Request("txtFrDlvyDt")%>"
				.frm1.hdnToDt2.Value 		= "<%=Request("txtToDlvyDt")%>"
				.frm1.hdnSoNo.value			= "<%=ConvSPChars(Request("txtSoNo"))%>"			
				.frm1.hdnTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"	
				.frm1.hdnSupplierCd.value	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnGroupCd.value		= "<%=ConvSPChars(Request("txtGroup"))%>"
				.frm1.hdnProcuType.value	= "<%=ConvSPChars(Request("txtProcure"))%>"
				.frm1.hdnPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
				if "<%=ConvSPChars(Request("txtProcure"))%>" = "P" then
					.frm1.hdnSubcontraflg.value	= "N"
				Else 
					.frm1.hdnSubcontraflg.value	= "Y"
				End if
				.ggoSpread.Source    = .frm1.vspdData1 
				.ggoSpread.SSShowData "<%=lgstrData_1%>"                            '��: Display data 
			
				.lgPageNo_1			 =  "<%=lgPageNo_1%>"							  '��: Next next data tag
				.DbQueryOk
			End If
		Else
			If "<%=lgDataExist%>" = "Yes" Then
				.frm1.hdnFrDt.Value 		= "<%=Request("txtFrPoDt")%>"
				.frm1.hdnToDt.Value 		= "<%=Request("txtToPoDt")%>"
				.frm1.hdnFrDt2.Value 		= "<%=Request("txtFrDlvyDt")%>"
				.frm1.hdnToDt2.Value 		= "<%=Request("txtToDlvyDt")%>"
				.frm1.hdnSoNo.value			= "<%=ConvSPChars(Request("txtSoNo"))%>"			
				.frm1.hdnTrackingNo.value	= "<%=ConvSPChars(Request("txtTrackingNo"))%>"	
				.frm1.hdnSupplierCd.value	= "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnGroupCd.value		= "<%=ConvSPChars(Request("txtGroup"))%>"
				.frm1.hdnProcuType.value	= "<%=ConvSPChars(Request("txtProcure"))%>"
				.frm1.hdnPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
				if "<%=ConvSPChars(Request("txtProcure"))%>" = "P" then
					.frm1.hdnSubcontraflg.value	= "N"
				Else 
					.frm1.hdnSubcontraflg.value	= "Y"
				End if
				.ggoSpread.Source    = .frm1.vspdData 
				.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
				.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
				.DbQuery2Ok
			End If
		End  if
	End with
</Script>