<% Option explicit%>
<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : ���ֳ������� PopUp Transaction ó���� ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Sun-jung Lee
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				
Err.Clear

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4                '�� : DBAgent Parameter ���� 
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData
	
    Dim iPrevEndRow
    Dim iEndRow
   
	Dim strBeneficiaryNm
	Dim strPurGrpNm
	Dim strPaymeth
	Dim strIncoterms
	
	Dim strPlantNm
	Dim strItemNm
	Dim strVatTypeNm 	
	
	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
	
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0

	SELECT CASE REQUEST("txtMode")								 '�� : onChange ���� ȣ���Ұ��� ���������ΰ�� 
		CASE "changeItemPlant"
			Call FixUNISQLData2()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
			Call QueryData2()										 '�� : DB-Agent�� ���� ADO query
	
		CASE ELSE
			Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
			Call QueryData()										 '�� : DB-Agent�� ���� ADO query
	END SELECT

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(3)
	Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(3,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
    
    UNISqlId(0) = "M4111RA101"
    UNISqlId(1) = "s0000qa009"											'���� 
    UNISqlId(2) = "s0000qa016"											'ǰ�� 
    UNISqlId(3) = "s0000qa026"											'VAT��     

    UNIValue(0,0) = Trim(lgSelectList)		                            '��: Select ������ Summary    �ʵ� 

	strVal = ""
  
  	If Len(Request("txtMvmtNo")) Then
		strVal = strVal & " AND A.MVMT_RCPT_NO = " & FilterVar(Request("txtMvmtNo"), "''", "S") & " "
	End If
	
	'---2003.07 TrackingNo �߰� 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & "  "		
	End If

	If Len(Request("txtPlantCd")) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S") & " "
	End If
	arrVal(0) = FilterVar(Trim(Request("txtPlantCd")), "", "S")
	
	If Len(Request("txtItemcd")) Then
		strVal = strVal & " AND C.ITEM_CD = " & FilterVar(Request("txtItemcd"), "''", "S") & " "
	End If
	arrVal(1) = FilterVar(Trim(Request("txtItemCd")), "", "S")

	If Len(Request("txtVatType")) Then
		strVal = strVal & " AND i.minor_cd = " & FilterVar(Request("txtVatType"), "''", "S") & " "
	End If
	arrVal(2) = FilterVar(Trim(Request("txtVatType")), "", "S")

    If Len(Trim(Request("txtFrDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT >= " & FilterVar(UNIConvDate(Request("txtFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""		
	End If

	If Len(Trim(Request("txtPoNo"))) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Request("txtPoNo"), "''", "S") & " "		
	End If
	
	If Len(Trim(Request("txtSppl"))) Then
		strVal = strVal & " AND F.BP_CD = " & FilterVar(Request("txtSppl"), "''", "S") & " "		
	End If
	
	If Len(Trim(Request("txtIvType"))) Then
		strVal = strVal & " AND F.IV_TYPE = " & FilterVar(Request("txtIvType"), "''", "S") & " "		
	End If
	
	'2009-09-02 ȭ������� ������� �ҷ����� ����... ������ ���� ��û
	'If Len(Trim(Request("txtPoCur"))) Then
	'	strVal = strVal & " AND F.PO_CUR = " & FilterVar(Request("txtPoCur"), "''", "S") & " "		
	'End If
	
	If UCase(Trim(Request("txtLcKind"))) <> "N" Then
		'LC��ȣ�� �ִ°��(Local LC �� ����� �����ϴ� ���)
		strVal = strVal & " AND L.PAY_METHOD = " & FilterVar(Request("txtPayMeth"), "''", "S") & " "		
    End If
    
    '���Գ����� �԰������� �������� ������ ����� �Ǹ� ��ȸ�ǵ��� ����(2005-10-28)
    If Len(Trim(Request("txtIvDt"))) Then
		strVal = strVal & " AND A.MVMT_DT <= " & FilterVar(UNIConvDate(Request("txtIvDt")), "''", "S") & ""		
	End If		

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND a.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND a.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND a.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND a.MVMT_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   

	
    UNIValue(0,1) = strVal
    UNIValue(1,0) = arrVal(0)  	'���� 
    UNIValue(2,0) = arrVal(1)  	'ǰ�� 
    UNIValue(3,0) = arrVal(2)	'VAT
   
    'UNIValue(0,UBound(UNIValue,2)) = ""
    UNIValue(0,UBound(UNIValue,2)) = " ORDER BY A.MVMT_RCPT_NO DESC,A.PO_NO DESC,A.PO_SEQ_NO ASC "			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
                        '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� (ONCHAGE ���� ȣ��)
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData2()

    Dim strVal
	Dim arrVal(2)
	Redim UNISqlId(2)                                                   '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(2,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                        '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "s0000qa009"											'���� 
    UNISqlId(1) = "s0000qa016"											'ǰ�� 
    UNISqlId(2) = "s0000qa026"											'VAT��     

    UNIValue(0,0) = Trim(lgSelectList)		                            '��: Select ������ Summary    �ʵ� 

	strVal = " "
	
	'If Len(Request("txtPlantCd")) Then
		arrVal(0) = FilterVar(Trim(Request("txtPlantCd")), "", "S")
	'End If
	
	'If Len(Request("txtItemcd")) Then
		arrVal(1) = FilterVar(Trim(Request("txtItemCd")), "", "S")
	'End If

	'If Len(Request("txtVatType")) Then
		arrVal(2) = FilterVar(Trim(Request("txtVatType")), "", "S")
	'End If

    UNIValue(0,0) = arrVal(0)  	'���� 
    UNIValue(1,0) = arrVal(1)  	'ǰ�� 
    UNIValue(2,0) = arrVal(2)	'VAT
   
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
                        '��: set ADO read mode
	
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

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

'----------------------------------------------------------------------------------------------------------
' Query Data2
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
    
'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
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

	If Not(rs3.EOF Or rs3.BOF) Then
        strVatTypeNm = rs3(1)
        Set rs3 = Nothing
	else
	    Set rs3 = Nothing
		If Len(Request("txtVatType")) Then
			Call DisplayMsgBox("970000", vbInformation, "VAT����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    Exit function
		End If
    End If  

	
	SetConditionData = True
	
End Function

%>

<Script Language=vbscript>
parent.frm1.txtPlantNm.Value 		= "<%=ConvSPChars(strPlantNm)%>"
parent.frm1.txtItemNm.Value 		= "<%=ConvSPChars(strItemNm)%>"
parent.frm1.txtVatNm.Value 			= "<%=ConvSPChars(strVatTypeNm)%>"	

    With parent
		If "<%=lgDataExist%>" = "Yes" Then
		   If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.hdnPoNo.value = "<%=ConvSPChars(request("txtPoNo"))%>"
			.frm1.hdnFrMvmtDt.value = "<%=request("txtFrDt")%>"
			.frm1.hdnToMvmtDt.value = "<%=request("txtToDt")%>"
			.frm1.hdnSupplierCd.value = "<%=ConvSPChars(request("txtSppl"))%>"
			.frm1.hdnMvmtNo.value = "<%=ConvSPChars(request("txtMvmtNo"))%>"
			.frm1.hdnRefType.value = "<%=ConvSPChars(request("txtRefType"))%>"
			.frm1.hdnIvType.value = "<%=ConvSPChars(request("txtIvType"))%>"
			.frm1.hdnPoCur.value = "<%=ConvSPChars(request("txtPoCur"))%>"
			.frm1.hdnPlantCd.value = "<%=ConvSPChars(request("txtPlantCd"))%>"
			.frm1.hdnItemCd.value = "<%=ConvSPChars(request("txtItemCd"))%>"
			.frm1.hdnVatType.value = "<%=ConvSPChars(request("txtVatType"))%>"		   
		   End If
		   
		   .ggoSpread.Source  = .frm1.vspdData
		   Parent.frm1.vspdData.Redraw = false
		   .ggoSpread.SSShowData "<%=iTotstrData%>", "F"          '�� : Display data
		   
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",12),"D", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",19),"C", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",20),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",25),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",26),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",29),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",30),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",36),"C", "I" ,"X","X")
       
		   .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
 
		   .DbQueryOk
		   Parent.frm1.vspdData.Redraw = True
		End If  
	End with
</Script>	
 	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

