
<%'======================================================================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : B/L���� 
'*  3. Program ID           : m3111ra2
'*  4. Program Name         : �������� 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : M31118ListPoHdrForBlSvr
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kang Su-hwan	
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date ǥ������ 
'*							  2002/04/12 ADO ��ȯ 
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
                                                                     
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3, rs4, rs5			   '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iTotstrData

Dim strBpNm													'�ŷ�ó�� 
Dim strPurGrp												'���ű׷� 
Dim strPOType												'�������� 
Dim strPaymeth												'������� 
Dim strIncoterms											'�������� 

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
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
FUNCTION SetConditionData()
    
    SetConditionData = TRUE
    
    If Not(rs1.EOF Or rs1.BOF) Then			' �ŷ�ó�ڵ�/�� 
        strBpNm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtBeneficiary")) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = FALSE
			exit function
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then			' ���ű׷��ڵ�/�� 
        strPurGrp =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtPurGrp")) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = FALSE			
			exit function			
		End If			
    End If   	
    
    If Not(rs3.EOF Or rs3.BOF) Then			' ���������ڵ�/�� 
        strPOType =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtPOType")) Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = FALSE			
			exit function			
		End If				
    End If      

    If Not(rs4.EOF Or rs4.BOF) Then			' ��������ڵ�/�� 
        strPaymeth =  rs4(1)
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtPayMeth")) Then
			Call DisplayMsgBox("970000", vbInformation, "�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = FALSE			
			exit function			
		End If				
    End If      

    If Not(rs5.EOF Or rs5.BOF) Then			' ���������ڵ�/�� 
        strIncoterms =  rs5(1)
        Set rs5 = Nothing
    Else
		Set rs5 = Nothing
		If Len(Request("txtIncoterms")) Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			SetConditionData = FALSE			
			exit function			
		End If				
    End If      
End FUNCTION

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(5,2)

	Const strpaymethMajor 	= "B9004"									'������� 
	Const strIncotermsMajor = "B9006"									'�������� 

    UNISqlId(0) = "M3111QA001"  										' main query(spread sheet�� �ѷ����� query statement)
	UNISqlId(1) = "s0000qa002"  										' �ŷ�ó�ڵ�/�� 
	UNISqlId(2) = "s0000qa019"  										' ���ű׷��ڵ�/�� 
	UNISqlId(3) = "M3111QA103"  										' ���������ڵ�/�� 
	UNISqlId(4) = "s0000qa000"  										' ��������ڵ�/�� 
	UNISqlId(5) = "s0000qa000"  										' ���������ڵ�/�� 

	
	'--- 2003-08-19 by Byun jee Hyun for UNICODE
    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    

	strVal = " "
	
	IF Len(Trim(Request("txtBeneficiary"))) THEN
		strVal = "AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & "  " & chr(13)
	END IF
	arrVal(0) = FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S")
	
	IF Len(Trim(Request("txtPurGrp"))) THEN 
		strVal = strVal & "AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S") & "  " & chr(13)
	END IF
	arrVal(1) = FilterVar(Trim(UCase(Request("txtPurGrp"))), " " , "S")
	
	IF Len(Trim(Request("txtPOType"))) THEN 
		strVal = strVal & "AND A.PO_TYPE_CD = " & FilterVar(Trim(UCase(Request("txtPOType"))), " " , "S") & "  " & chr(13)
	END IF
	arrVal(2) = FilterVar(Trim(UCase(Request("txtPOType"))), " " , "S")

	IF Len(Trim(Request("txtPayMeth"))) THEN 
		strVal = strVal & "AND A.PAY_METH = " & FilterVar(Trim(UCase(Request("txtPayMeth"))), " " , "S") & "  " & chr(13)
	END IF
	arrVal(3) = FilterVar(Trim(UCase(Request("txtPayMeth"))), " " , "S")
	
	IF Len(Trim(Request("txtIncoterms"))) THEN 
		strVal = strVal & "AND A.INCOTERMS = " & FilterVar(Trim(UCase(Request("txtIncoterms"))), " " , "S") & "  " & chr(13)
	END IF
	arrVal(4) = FilterVar(Trim(UCase(Request("txtIncoterms"))), " " , "S")
	
	IF Len(Trim(Request("txtPOFrDt"))) THEN 
		strVal = strVal & "AND A.PO_DT >= " & FilterVar(UniconvDate(Trim(Request("txtPOFrDt"))), "''", "S") & " " & chr(13)
	END IF
	IF Len(Trim(Request("txtPOToDt"))) THEN 
		strVal = strVal & "AND A.PO_DT <= " & FilterVar(UniconvDate(Trim(Request("txtPOToDt"))), "''", "S") & " " & chr(13)
	END IF

     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND A.PUR_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND A.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND A.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
    
    UNIValue(0,1) = strVal    												'	UNISqlId(0)�� �ι�° ?�� �Էµ�	
	UNIValue(1,0) = arrVal(0)     					'��: �ŷ�ó�ڵ� 
	UNIValue(2,0) = arrVal(1)    					'��: ���ű׷��ڵ� 
	UNIValue(3,0) = arrVal(2)     		'��: ���������ڵ� 
	UNIValue(4,0) = FilterVar(strpaymethMajor, "''", "S")   				'��: ��������ڵ� 
	UNIValue(4,1) = arrVal(3)				     	'��: ��������ڵ� 
	UNIValue(5,0) = FilterVar(strIncotermsMajor, "''", "S") 				'��: ���������ڵ� 
	UNIValue(5,1) = arrVal(4)			     		'��: ���������ڵ� 
   
'	UNIValue(0,UBound(UNIValue,2)) = "ORDER BY A.PO_NO DESC "
	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If SetConditionData = FALSE THEN EXIT SUB

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("173100", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtBeneficiaryNm.value = "<%=ConvSPChars(strBpNm)%>" 
		.frm1.txtPurGrpNm.value = "<%=ConvSPChars(strPurGrp)%>" 
		.frm1.txtPOTypeNm.value = "<%=ConvSPChars(strPOType)%>" 
		.frm1.txtPayMethNm.value = "<%=ConvSPChars(strPaymeth)%>" 
		.frm1.txtIncotermsNm.value = "<%=ConvSPChars(strIncoterms)%>" 
		
		If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				parent.frm1.hdnBeneficiary.value	= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
				parent.frm1.hdnPurGrp.value			= "<%=ConvSPChars(Request("txtPurGrp"))%>"
				parent.frm1.hdnPOType.value			= "<%=ConvSPChars(Request("txtPOType"))%>"
				parent.frm1.hdnPayMeth.value		= "<%=ConvSPChars(Request("txtPayMeth"))%>"
				parent.frm1.hdnIncoterms.value		= "<%=ConvSPChars(Request("txtIncoterms"))%>"
				parent.frm1.hdnFrDt.value			= "<%=Request("txtPOFrDt")%>"
				parent.frm1.hdnToDt.value			= "<%=Request("txtPOToDt")%>"
			End If    
			       
			 .ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=iTotstrData%>"                            '��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
