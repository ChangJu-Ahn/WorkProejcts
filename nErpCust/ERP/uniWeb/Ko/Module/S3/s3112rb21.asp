<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : s3112pb20
'*  4. Program Name         : ���ֹ�ȣ 
'*  5. Program Desc         : ���ֹ�ȣ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HJO
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
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
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
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist    = "No"
	    
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
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	   iFrPoint     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)	
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
    SetConditionData = False
    
	If Not(rs1.EOF Or rs1.BOF) Then
		SupplierNm = rs1("BP_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If   	
    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	dim sTemp
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
    
    UNISqlId(0) = "S3112RA2101"    
    UNISqlId(1) = "s0000qa002"
   '--For Unicode    
    UNIValue(1,0) = FilterVar("zzzzzzzzzz", " " , "S")  
    
    sTemp = "2"
    
	If Len(Trim(Request("txtFrPoDt"))) Then
		If UNIConvDate(Request("txtFrPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrPoDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If
	
	If Len(Trim(Request("txtToPoDt"))) Then
		If UNIConvDate(Request("txtToPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToPoDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If

   'default condition   
    strVal = strVal & " AND A.CLS_FLG = 'N'    AND F.RELEASE_FLG = 'Y' "
	'���� 
	if len(trim(request("txtPlant"))) then
		if sTemp="1" then
			strVal=strVal & " WHERE A.PLANT_CD = " &  FilterVar(Trim(UCase(Request("txtPlant"))), " " , "S") & "  "
			sTemp="2"
		else
			strVal=strVal & " AND A.PLANT_CD = " &  FilterVar(Trim(UCase(Request("txtPlant"))), " " , "S") & "  "
		end if	
	end if		
	
	'����ó 
    If Len(Trim(Request("txtSupplierCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE F.BP_CD =  " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND F.BP_CD =  " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & "  "	
		end if		    
		UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S")
	End If
    '������ 
    If Len(Trim(Request("txtFrPoDt"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE F.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND F.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
		end if		      
	End If
			
    If Len(Trim(Request("txtToPoDt"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE F.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND F.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
		end if		      
	End If
	
'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
	UNIValue(0,1) = strVal & UCase(Trim(lgTailList)) 

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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
	If SetConditionData = False Then Exit Sub

    If  rs0.EOF And rs0.BOF Then
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
		.frm1.txtSupplierNm.value = "<%=ConvSPChars(SupplierNm)%>"
		If "<%=lgDataExist%>" = "Yes" Then
		
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnSupplier.value	= "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.frm1.hdnFrDt.value 	= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
				.frm1.hdnToDt.value 	= "<%=ConvSPChars(Request("txtToPoDt"))%>"		
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
