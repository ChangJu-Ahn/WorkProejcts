<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2002/07/22
'*  7. Modified date(Last)  : 
'*                            
'*  8. Modifier (First)     : Oh Chang won
'*  9. Modifier (Last)      : 
'*                            
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort ���� 

Dim PotypeNm														'�� : �������¸� ���� 
Dim GroupNm										   				    '�� : ���ű׷�� ���� 
Dim SupplierNm														'�� : ����ó�� ���� 
Dim iFrPoint
   iFrPoint=0
   
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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

	Const C_SHEETMAXROWS_D  = 100
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
	Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint     = CLng(lgMaxCount) * CLng(lgPageNo)	
    End If
    
    iLoopCount = -1
    
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
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

    If iLoopCount < lgMaxCount Then                                 '��: Check if next data exists
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
    SetConditionData = True
    
	If Not(rs1.EOF Or rs1.BOF) Then
		PotypeNm = rs1("PO_TYPE_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtPotypeCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    SetConditionData = False
		End If
	End If   	
	
	If Not(rs2.EOF Or rs2.BOF) Then
		SupplierNm = rs2("BP_NM")
		Set rs2 = Nothing
	Else
		Set rs2 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    SetConditionData = False
		End If
	End If   	
	
	If Not(rs3.EOF Or rs3.BOF) Then
		GroupNm = rs3("PUR_GRP_NM")
		Set rs3 = Nothing
	Else
		Set rs3 = Nothing
		If Len(Request("txtGroupCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    SetConditionData = False
		End If
	End If   	

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	dim sTemp
	Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(3,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
    UNISqlId(0) = "M3111PA101"
    UNISqlId(1) = "s0000qa020"
    UNISqlId(2) = "s0000qa002"
    UNISqlId(3) = "s0000qa019"
    
    '--- 2004-08-19 by Byun Jee Hyun for UNICODE
    UNIValue(1,0) = FilterVar("zzzzz", "''", "S")
    UNIValue(2,0) = FilterVar("zzzzzzzzzz", "''", "S")
    UNIValue(3,0) = FilterVar("zzzz", "''", "S")
    
    sTemp = "1"
    
    '��������                    
    If Len(Trim(Request("txtPotypeCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PO_TYPE_CD =  " & FilterVar(Request("txtPotypeCd"), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PO_TYPE_CD =  " & FilterVar(Request("txtPotypeCd"), "''", "S") & " "	
		end if	
		UNIValue(1,0) = FilterVar(Trim(Request("txtPotypeCd")), "''", "S")
	End If

	'����ó 
    If Len(Trim(Request("txtSupplierCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.BP_CD =  " & FilterVar(Request("txtSupplierCd"), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.BP_CD =  " & FilterVar(Request("txtSupplierCd"), "''", "S") & " "	
		end if		    
		UNIValue(2,0) = FilterVar(Trim(Request("txtSupplierCd")), "''", "S")
	End If
    '������ 
    If Len(Trim(Request("txtFrPoDt"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
		end if		      
	End If
			
    If Len(Trim(Request("txtToPoDt"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
		end if		      
	End If
	'���ű��� 
	If Len(Trim(Request("txtGroupCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PUR_GRP =  " & FilterVar(Request("txtGroupCd"), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Request("txtGroupCd"), "''", "S") & " "	
		end if		  	
		UNIValue(3,0) = FilterVar(Trim(Request("txtGroupCd")), "''", "S")
	End If

	strVal = strVal & " AND A.SUBCONTRA_FLG = " & FilterVar("Y", "''", "S") & "  "

	strVal = strVal & " AND A.RET_FLG = " & FilterVar("N", "''", "S") & "  "		' ��޼ҿ䷮ �������� ��ǰ���� ��ȸ���� �ʵ��� ��.
		
'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
'	UNIValue(0,1) = strVal & " ORDER BY A.PO_NO DESC"
	UNIValue(0,1) = strVal & Trim(lgTailList)

    'UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 



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

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
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
		.frm1.txtPotypeNm.value 	= "<%=ConvSPChars(PotypeNm)%>"
		.frm1.txtSupplierNm.value 	= "<%=ConvSPChars(SupplierNm)%>"
		.frm1.txtGroupNm.value 		= "<%=ConvSPChars(GroupNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnPotype.value	= "<%=ConvSPChars(Request("txtPotypeCd"))%>"
				.frm1.hdnSupplier.value	= "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.frm1.hdnFrDt.value		= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
				.frm1.hdnToDt.value		= "<%=ConvSPChars(Request("txtToPoDt"))%>"
				.frm1.hdnGroup.value	= "<%=ConvSPChars(Request("txtGroupCd"))%>"
		
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			Parent.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>", "F"                  '��: Display data 
			
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",7),"A","I","X","X")
			
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
			Parent.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
