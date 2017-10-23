<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212rb5.asp																*
'*  4. Program Name         : local l/c ��������(���Գ������)				   					    	*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2003/03/13																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : Lee Eun Hee																*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2003/03/13 : ȭ�� design												*
'********************************************************************************************************


On Error Resume Next

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           '�� : DBAgent Parameter ���� 
	Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
    Dim iTotstrData
    
	Dim strPurGrpNm
	Dim strPaymeth
	Dim strIncoterms
	Dim strPlantNm
	Dim strItemNm

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

	SELECT CASE REQUEST("txtMode")								 '�� : onChange ���� ȣ���Ұ��� ���������ΰ�� 
	
	CASE "changeItemPlant"
		Call FixUNISQLData2()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
		Call QueryData2()										 '�� : DB-Agent�� ���� ADO query
	
	CASE ELSE
		Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
		Call QueryData()											'�� : DB-Agent�� ���� ADO query
	END SELECT
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
        strPlantNm = rs1(1)
        Set rs1 = Nothing
	else
	    Set rs1 = Nothing
		If Len(Request("txtPlantCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit function
		End If
	End If  
	

	If Not(rs2.EOF Or rs2.BOF) Then
        strItemNm = rs2(1)
        Set rs2 = Nothing	
	else
	    Set rs2 = Nothing
		If Len(Request("txtItemcd")) Then
			Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit function
		End If

	End If  

	SetConditionData = True
	
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "m3212ra501" 
     UNISqlId(1) = "s0000qa009"											'���� 
	 UNISqlId(2) = "s0000qa016"											'ǰ�� 
     
     '--- 2004-08-20 by Byun Jee Hyun for UNICODE	
     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

	strVal = " "
	strVal = strVal & " AND E.LC_DOC_NO != " & FilterVar("", "''", "S") & "  AND E.OPEN_DT != " & FilterVar("", "''", "S")
	
	If Len(Request("txtLLCNo")) Then
		strVal = strVal & " AND A.LC_NO =  " & FilterVar(UCase(Request("txtLLCNo")), "''", "S") & " "
	End If

    If Len(Trim(Request("txtFrLCDt"))) Then
		strVal = strVal & " AND E.OPEN_DT >= " & FilterVar(UNIConvDate(Request("txtFrLCDt")), "''", "S") & ""
	Else
		strVal = strVal & " AND E.OPEN_DT >=" & "" & FilterVar("1900/01/01", "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToLCDt"))) Then
		strVal = strVal & " AND E.OPEN_DT <= " & FilterVar(UNIConvDate(Request("txtToLCDt")), "''", "S") & ""		
	Else
		strVal = strVal & " AND E.OPEN_DT <=" & "" & FilterVar("2900/12/30", "''", "S") & ""		
	End If
	
	If Len(Request("txtPlantCd")) Then
		strVal = strVal & " AND A.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S") & " "
	End If
	arrVal(0) = FilterVar(Trim(Request("txtPlantCd")), "", "S")
	
	If Len(Request("txtItemcd")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Request("txtItemcd"), "''", "S") & " "
	End If
	arrVal(1) = FilterVar(Trim(Request("txtItemCd")), "", "S")
	'--------------
	If Len(Trim(Request("txtPoNo"))) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Request("txtPoNo"), "''", "S") & " "		
	End If
	
	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND E.BENEFICIARY = " & FilterVar(UCase(Request("txtSupplier")), "''", "S") & " "
	End If
	
	If Len(Trim(Request("txtIvType"))) Then
		strVal = strVal & " AND F.IV_TYPE = " & FilterVar(Request("txtIvType"), "''", "S") & " "		
	End If
	
	If Len(Trim(Request("txtPoCur"))) Then
		strVal = strVal & " AND F.PO_CUR = " & FilterVar(Request("txtPoCur"), "''", "S") & " "		
	End If	
	'**����(2003.03.26)***
	If UCase(Trim(Request("txtLcKind"))) <> "N" Then
		'LC��ȣ�� �ִ°��(Local LC �� ����� �����ϴ� ���)
		strVal = strVal & " AND E.PAY_METHOD = " & FilterVar(Request("txtPayMeth"), "''", "S") & " "		
    End If
    
    '---2003.07 TrackingNo �߰� 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & "  "		
	End If
    
    UNIValue(0,1) = strVal 
    UNIValue(1,0) = arrVal(0)  	'���� 
    UNIValue(2,0) = arrVal(1)  	'ǰ��			'���ű׷�   
   
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
        
	If SetConditionData = false then Exit Sub
	
	If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
      
       Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� (ONCHAGE ���� ȣ��)
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData2()

    Dim strVal
	Dim arrVal(1)
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(1,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                        '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "s0000qa009"											'���� 
    UNISqlId(1) = "s0000qa016"											'ǰ�� 

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
	
    UNIValue(0,0) = Trim(lgSelectList)		                            '��: Select ������ Summary    �ʵ� 

	strVal = " "
	
	If Len(Request("txtPlantCd")) Then
		arrVal(0) = FilterVar(Trim(Request("txtPlantCd")), "", "S")
	End If
	
	If Len(Request("txtItemcd")) Then
		arrVal(1) = FilterVar(Trim(Request("txtItemCd")), "", "S")
	End If

    UNIValue(0,0) = arrVal(0)  	'���� 
    UNIValue(1,0) = arrVal(1)  	'ǰ�� 
   
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
                        '��: set ADO read mode
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData2()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    Call  SetConditionData()
End Sub


%>
<Script Language=vbscript>
	
parent.frm1.txtPlantNm.Value 		= "<%=ConvSPChars(strPlantNm)%>"
parent.frm1.txtItemNm.Value 		= "<%=ConvSPChars(strItemNm)%>"
	
    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			parent.frm1.hdnPoNo.value	= "<%=ConvSPChars(Request("txtPoNO"))%>"  
			parent.frm1.hdnFrLCDt.value	= "<%=Request("txtFrLCDt")%>"
			parent.frm1.hdnToLCDt.value	= "<%=Request("txtToLCDt")%>"
			parent.frm1.hdnSupplierCd.value	= "<%=ConvSPChars(request("txtSupplier"))%>"
			parent.frm1.hdnLLcNo.value = "<%=ConvSPChars(request("txtLLCNo"))%>"
			parent.frm1.hdnIvType.value = "<%=ConvSPChars(request("txtIvType"))%>"
			parent.frm1.hdnPoCur.value = "<%=ConvSPChars(request("txtPoCur"))%>"
			parent.frm1.hdnPlantCd.value = "<%=ConvSPChars(request("txtPlantCd"))%>"
			parent.frm1.hdnItemCd.value = "<%=ConvSPChars(request("txtItemCd"))%>"
				
	   End If
       Parent.frm1.vspdData.Redraw = False
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.ggoSpread.SSShowData "<%=iTotstrData%>", "F"          '�� : Display data
       
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",14),"C", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",15),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",22),"C", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",23),"A", "I" ,"X","X")
       
       parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
</Script>	
