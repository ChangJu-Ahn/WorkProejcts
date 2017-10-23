<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : s3211ra1.asp	
'*  4. Program Name         : L/C����(L/C Amend��Ͽ���)
'*  5. Program Desc         : L/C����(L/C Amend��Ͽ���)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/25
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Seo Jinkyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3 , rs4  '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 

Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim strApplicnat , strSales_grp , strOpenBank
Dim BlankchkFlg  		
Dim iFrPoint
iFrPoint=0

Const C_SHEETMAXROWS_D  = 30                                       
Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)	
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
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
        
        rs0.MoveNext
	Loop

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
Sub SetConditionData()
    On Error Resume Next
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strApplicnat =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtApplicant")) Then
			Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
 %>
<Script Language=VBScript>
			Parent.frm1.txtApplicant.focus 
</Script>
<%		  					
			BlankchkFlg  =  True
			Response.End	
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSales_grp =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtSalesGroup")) Then
			Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtSalesGroup.focus 
</Script>
<%		  					
			BlankchkFlg  =  True
			Response.End
		End If			
    End If   
    
    If Not(rs4.EOF Or rs4.BOF) Then        
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtCurrency")) Then
			Call DisplayMsgBox("970000", vbInformation, "ȭ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtCurrency.focus 
</Script>
<%		  								
		    BlankchkFlg  =  True
		    Response.End
		End If				
    End If      
        	
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strOpenBank =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtOpenBank")) Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
 %>
<Script Language=VBScript>
			Parent.frm1.txtOpenBank.focus 
</Script>
<%		  								
		    BlankchkFlg  =  True
		    Response.End
		End If				
    End If      

    
        
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
																		  '�Ʒ��� ���� ȭ��ܿ��� �־� �ִ� query�� where�������� �� �� �ִ�.	
    Dim arrVal(4)														  '��: ȭ�鿡�� �˾��Ͽ� query
																		  '�Ʒ��� ���� UNISqlId(1),UNISqlId(2), UNISqlId(3)�� where�������� �� �� �ִ�.
    Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
																		  '��ȸȭ�鿡�� �ʿ��� query���ǹ����� ����(Statements table�� ����)
    Redim UNIValue(5,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 

    UNISqlId(0) = "S3211RA101"  ' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(1) = "s0000qa002"  ' ������  arrVal(0)
    UNISqlId(2) = "s0000qa005"  ' �����׷� arrVal(1)    
    UNISqlId(3) = "s0000qa008"  ' �������� arrVal(2)
    UNISqlId(4) = "s0000qa014"  ' ȭ��     arrVal(3)

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strVal = " "
	If Len(Request("txtFromDt")) > 0 Then
		strVal = "AND A.open_dt >= " & FilterVar(UNIConvDate(Trim(Request("txtFromDt"))), "''", "S") & ""			
	End If
	If Len(Request("txtToDt")) > 0 Then
		strVal = strVal &  "AND A.open_dt <= " & FilterVar(UNIConvDate(Trim(Request("txtToDt"))), "''", "S") & ""			
	End If
	
	If Len(Request("txtApplicant")) > 0 Then
		strVal =  strVal & "AND A.applicant = " & FilterVar(Request("txtApplicant"), "''", "S") & " "
		arrVal(0)= Trim(Request("txtApplicant")) 
	else
		arrVal(0)= " "
	End If
			
	If Len(Request("txtSalesGroup")) > 0 Then
		strVal =  strVal & "AND A.sales_grp = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "			
		arrVal(1)= Trim(Request("txtSalesGroup"))
	else
		arrVal(1)= " "		
	End If

	If Len(Request("txtLCDocNo")) > 0 Then
		strVal = strVal & " AND A.lc_doc_no = " & FilterVar(Request("txtLCDocNo"), "''", "S") & " "
	End If		
	
	If Len(Request("txtCurrency")) > 0 Then
		strVal = strVal & " AND A.cur = " & FilterVar(Request("txtCurrency"), "''", "S") & " "			
		arrVal(3)= Trim(Request("txtCurrency")) 
	End If		
	
	If Len(Request("txtOpenBank")) > 0 Then
		strVal = strVal & " AND A.issue_bank_cd = " & FilterVar(Request("txtOpenBank"), "''", "S") & " "			
		arrVal(2)= Trim(Request("txtOpenBank")) 
	else
		arrVal(2)= " "
	End If
	
    UNIValue(0,1) = strVal    '	UNISqlId(0)�� �ι�° ?�� �Էµ� 
    UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S") '	UNISqlId(1)�� ù��° ? 
    UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S") '	UNISqlId(2)�� ù��° ? 
    UNIValue(3,0) = FilterVar(Trim(Request("txtOpenBank")), " " , "S") '	UNISqlId(3)�� ù��° ? 
    UNIValue(4,0) = FilterVar(Trim(Request("txtCurrency")), " " , "S") '	UNISqlId(4)�� ù��° ? 
    
      
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
	Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

    Dim iStr
	Dim FalsechkFlg
    BlankchkFlg = False

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'��:ADO ��ü�� ���� 
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
                            

    FalsechkFlg = False

    iStr = Split(lgstrRetMsg,gColSep)
    
    
    call  SetConditionData
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    If BlankchkFlg = False Then
		If  rs0.EOF And rs0.BOF And BlankchkFlg = False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
 %>
<Script Language=VBScript>
			Parent.frm1.txtApplicant.focus 
</Script>
<%		  		
		    
		    rs0.Close
		    Set rs0 = Nothing
		Else   
		    Call  MakeSpreadSheetData()       
		End If
    End If
   
End Sub
%>
<Script Language=vbscript>
	
	With parent
		.frm1.txtApplicantNm.Value	= "<%=ConvSPChars(strApplicnat)%>"
	    .frm1.txtSalesGroupNm.Value	= "<%=ConvSPChars(strSales_grp)%>"
	    .frm1.txtOpenBankNm.Value	= "<%=ConvSPChars(strOpenBank)%>"			    
	
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				parent.frm1.txtHApplicant.value = "<%=ConvSPChars(Request("txtApplicant"))%>"
				parent.frm1.txtHSalesGroup.value = "<%=ConvSPChars(Request("txtSalesGroup"))%>"
				parent.frm1.txtHLCDocNo.value = "<%=ConvSPChars(Request("txtLCDocNo"))%>"
				parent.frm1.txtHCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"
				parent.frm1.txtHOpenBank.value = "<%=ConvSPChars(Request("txtOpenBank"))%>"
				parent.frm1.txtHFromDt.value = "<%=Request("txtFromDt")%>"
				parent.frm1.txtHToDt.value = "<%=Request("txtToDt")%>"				
			End If    
			
			'Show multi spreadsheet data from this line			       
			.ggoSpread.Source    = .frm1.vspdData 
			
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                            '��: Display data 
        
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",6),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
			        			
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True
		End If
		
	End with
</Script>	

