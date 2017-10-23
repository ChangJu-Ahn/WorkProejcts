
<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4212rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : �������(ADO)																	*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/22																			*
'*  9. Modifier (First)     : An Chang Hwan																*
'* 10. Modifier (Last)      : Kim Jin Ha 
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'* 13. History              :																			*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           '�� : DBAgent Parameter ���� 
   Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
   Dim iTotstrData
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim strPlantNm
   Dim strPurGrpNm
	
	Dim iPrevEndRow
	Dim iEndRow

    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
	 
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0
    
    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
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
    On Error Resume Next
    SetConditionData = false

   
    If Not(rs1.EOF Or rs1.BOF) Then
        strPurGrpNm = rs1("Pur_Grp_Nm")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	
    If Not(rs2.EOF Or rs2.BOF) Then
        strPlantNm = rs2("plant_nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtPlant"))) Then
			Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  

	SetConditionData = true
	
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	'call svrmsgbox(Request("gPurGrp"),0,1)

    Dim strVal
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "M4212RA101" 
     UNISqlId(1) = "S0000QA019"	'���ű׷� 
	 UNISqlId(2) = "M2111QA302"	'���� 

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

    strVal = " "	
	strVal = strVal & " AND A.QTY > (A.RECEIPT_QTY + ISNULL( K.CC_INSP_QTY,0)) "
 	'strVal = strVal & " AND H.SUBCONTRA_FLG =  " & FilterVar(Request("txtSubcontraFlg"), "''", "S") & " "
 	strVal = strVal & " AND H.RCPT_TYPE =  " & FilterVar(Request("txtRcptType"), "''", "S") & " "
 
	If Len(Request("txtPoNo")) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & " "
	End If

	If Len(Request("txtCcNo")) Then
		strVal = strVal & " AND A.CC_NO = " & FilterVar(UCase(Request("txtCcNo")), "''", "S") & " "
	End If

    If Len(Trim(Request("txtFrCcDt"))) Then
		strVal = strVal & " AND F.ID_DT >= " & FilterVar(UNIConvDate(Request("txtFrCcDt")), "''", "S") & ""
	Else
		strVal = strVal & " AND F.ID_DT >=" & "" & FilterVar("1900/01/01", "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToCcDt"))) Then
		strVal = strVal & " AND F.ID_DT <= " & FilterVar(UNIConvDate(Request("txtToCcDt")), "''", "S") & ""		
	Else
		strVal = strVal & " AND F.ID_DT <=" & "" & FilterVar("2900/12/30", "''", "S") & ""		
	End If
	
	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND F.BENEFICIARY = " & FilterVar(UCase(Request("txtSupplier")), "''", "S") & " "
	End If
	
	If Len(Request("txtGroup")) Then
		strVal = strVal & " AND F.PUR_GRP = " & FilterVar(Request("txtGroup"), "''", "S") & " "
	End If

	If Len(Request("txtPlant")) Then
		strVal = strVal & " AND A.PLANT_CD = " & FilterVar(Request("txtPlant"), "''", "S") & " "
	End If
	
	'---2003.07 TrackingNo �߰� 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & "  "		
	End If	

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND A.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND F.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND F.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     'If Request("gBizArea") <> "" Then
     '  strVal = strVal & " AND F.BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     'End If   

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtGroup"))), "" , "S") 						'���ű׷�   
    UNIValue(2,0) = " " & FilterVar(Trim(UCase(Request("txtPlant"))), "" , "S") & " "				'���ű׷�   

    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	If SetConditionData = False Then Exit Sub
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>
<Script Language=vbscript>
With parent
	.frm1.txtGroupNm.value = "<%=ConvSPChars(strPurGrpNm)%>" 
	.frm1.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>" 

    If "<%=lgDataExist%>" = "Yes" Then
       
       If "<%=lgPageNo%>" = "1" Then 
			.frm1.hdnPoNo.value			= "<%=ConvSPChars(Request("txtPoNo"))%>"
			.frm1.hdnCcNo.value			= "<%=ConvSPChars(Request("txtCcNo"))%>"
			.frm1.hdnFrCcDt.value			= "<%=Request("txtFrCcDt")%>"
			.frm1.hdnToCcDt.value			= "<%=Request("txtToCcDt")%>"
			'.frm1.hdnSubcontraFlg.value	= "<%=ConvSPChars(Request("txtSubcontraFlg"))%>"
			.frm1.hdnGroupCd.value		= "<%=ConvSPChars(Request("txtGroup"))%>" 
			.frm1.hdnGroupNm.value		= "<%=ConvSPChars(strPurGrpNm)%>" 
			.frm1.hdnPlantCd.value		= "<%=ConvSPChars(Request("txtPlant"))%>"  
       End If
       
	   .ggoSpread.Source  = .frm1.vspdData
       .frm1.vspdData.Redraw = False
       .ggoSpread.SSShowData "<%=iTotstrData%>"          '�� : Display data
       .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       .DbQueryOk
       .frm1.vspdData.Redraw = True
    End If 
End With  
</Script>	




