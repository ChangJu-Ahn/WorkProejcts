<%
'********************************************************************************************************
'*  1. Module Name          : ��������																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ma7.asp																*
'*  4. Program Name         : Local L/C��Ȳ��ȸ															*
'*  5. Program Desc         : Local L/C��Ȳ��ȸ															*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : ȭ�� design												*
'*							  2. 2000/03/22 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
Call LoadBNumericFormatB("Q","S","NOCOOKIE","QB")

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4         '�� : DBAgent Parameter ���� 
   Dim lgStrData														'�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
   
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo

'--------------- ������ coding part(��������,Start)----------------------------------------------------
	Dim strApplicantNm			'������û�� 
	Dim strSalesGrpNm			'�����׷� 
	Dim strOpenBankNm			'�������� 
	Dim BlankchkFlg
    Const C_SHEETMAXROWS_D  = 100                                          '��: Fetch max count at once
	MsgDisplayFlag = False
	Dim iFrPoint
    iFrPoint=0
'--------------- ������ coding part(��������,End)------------------------------------------------------
    Call HideStatusWnd 
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)   
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'============================================================================================================
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

'============================================================================================================
Sub SetConditionData()
    On Error Resume Next
	
    If Not(rs1.EOF Or rs1.BOF) Then
       strApplicantNm =  rs1("BP_NM")
    Else
   		If Len(Request("txtApplicantCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "������û��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtApplicantCd.focus    
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs1 = Nothing 
	
	If Not(rs2.EOF Or rs2.BOF) Then
       strSalesGrpNm =  rs2("SALES_GRP_NM")
    Else
   		If Len(Request("txtSalesGrpCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGrpCd.focus    
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs2 = Nothing 

	If Not(rs3.EOF Or rs3.BOF) Then
       strOpenBankNm =  rs3("BANK_NM")
    Else
   		If Len(Request("txtOpenBankCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtOpenBankCd.focus    
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs3 = Nothing 

	If Not(rs4.EOF Or rs4.BOF) Then
    Else
   		If Len(Request("txtCur")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "ȭ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtCur.focus    
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs4 = Nothing
	
End Sub

'============================================================================================================
Sub FixUNISQLData()

    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
																		  '�Ʒ��� ���� ȭ��ܿ��� �־� �ִ� query�� where�������� �� �� �ִ�.	
    Dim arrVal(3)														  '��: ȭ�鿡�� �˾��Ͽ� query
																		  '�Ʒ��� ���� UNISqlId(1),UNISqlId(2), UNISqlId(3)�� where�������� �� �� �ִ�.
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
																		  '��ȸȭ�鿡�� �ʿ��� query���ǹ����� ����(Statements table�� ����)
    Redim UNIValue(4,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 

    UNISqlId(0) = "S3211QA701"											  ' main query(spread sheet�� �ѷ����� query statement)
    UNISqlId(1) = "s0000qa002"
    UNISqlId(2) = "s0000qa005"
    UNISqlId(3) = "s0000qa008"
    UNISqlId(4) = "s0000qa014"
    									  
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
																		  '	UNISqlId(0)�� ù��° ?�� �Էµ�				
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
		
	strVal = ""
	'---������û�� 
    If Len(Request("txtApplicantCd")) Then
    	strVal	  = strVal & "AND b.bp_cd =  " & FilterVar(UCase(Request("txtApplicantCd")), "''", "S") & "  "
    	arrVal(0) = Trim(Request("txtApplicantCd")) 
    End If
    
	'---�����׷� 
	If Len(Request("txtSalesGrpCd")) Then
		strVal	  = strVal & "AND c.sales_grp =  " & FilterVar(UCase(Request("txtSalesGrpCd")), "''", "S") & "  "
		arrVal(1) = Trim(Request("txtSalesGrpCd"))
	End If
    
    '---�����ݾ� 
    If Len(Request("txtFromLocAmt")) Then
		Dim txtFromLocAmt
		txtFromLocAmt = Trim(Request("txtFromLocAmt"))
    	strVal 	= strVal & "AND a.lc_amt >= " & UNIConvNum(txtFromLocAmt, 0) & " "
    End If
    '2003-01-23 UNICDbl�Լ��� ������ ��� 
    If Len(Request("txtToLocAmt")) And UNICDbl(UNIConvNum(Request("txtToLocAmt"),0),0) <> 0 Then
		Dim txtToLocAmt
		txtToLocAmt = Trim(Request("txtToLocAmt"))
    	strVal	= strVal & "AND a.lc_amt <= " & UNIConvNum(txtToLocAmt, 0) & " "
    Else
		strVal	= strVal & "AND a.lc_amt <= 9999999999999.99 "
    End If    
    
    '---ȭ�� 
	If Len(Request("txtCur")) Then
    	strVal	= strVal & "AND a.cur =  " & FilterVar(UCase(Request("txtCur")), "''", "S") & "  "
    	arrVal(3) = Trim(Request("txtCur"))
    End If
    
     '---������ 
    If Len(Request("txtFromDate")) Then
    	strVal	= strVal & "AND a.open_dt >=  " & FilterVar(uniConvDate(Trim(Request("txtFromDate"))), "''", "S") & " "
    End If
    
    If Len(Request("txtToDate")) Then
    	strVal	= strVal & "AND a.open_dt <=  " & FilterVar(uniConvDate(Trim(Request("txtToDate"))), "''", "S") & " "
    End If    
    
    '---�������� 
	If Len(Request("txtOpenBankCd")) Then
    	strVal	= strVal & "AND issue.bank_cd =  " & FilterVar(UCase(Request("txtOpenBankCd")), "''", "S") & "  "
    	arrVal(2) = Trim(Request("txtOpenBankCd"))
    End If
    
       
'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0,1)  = strVal
	UNIValue(1,0)  = FilterVar(arrVal(0), " " , "S")
	UNIValue(2,0)  = FilterVar(arrVal(1), " " , "S")
	UNIValue(3,0)  = FilterVar(arrVal(2), " " , "S")
	UNIValue(4,0)  = FilterVar(arrVal(3), " " , "S")
		
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'============================================================================================================
Sub QueryData()
	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    
    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)

    
    Call  SetConditionData()
    
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg, vbInformation, I_MKSCRIPT)
    End If  
    
    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtApplicantCd.focus    
                </Script>
            <%
			' �� ��ġ�� �ִ� Response.End �� �����Ͽ��� ��. Client �ܿ��� Name�� ��� �ѷ��� �Ŀ� Response.End �� �����.
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	

        
   
End Sub

%>

<Script Language=vbscript>
    
    With parent
		.frm1.txtApplicantNm.value	 = "<%=ConvSPChars(strApplicantNm)%>"
		.frm1.txtSalesGrpNm.value	 = "<%=ConvSPChars(strSalesGrpNm)%>"
		.frm1.txtOpenBankNm.value	 = "<%=ConvSPChars(strOpenBankNm)%>"
		 
    If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.HApplicantCd.value	 = "<%=ConvSPChars(Request("txtApplicantCd"))%>"
			.frm1.HSalesGrpCd.value	  	 = "<%=ConvSPChars(Request("txtSalesGrpCd"))%>"
			.frm1.HFromLocAmt.value		 = "<%=ConvSPChars(Request("txtFromLocAmt"))%>"
			.frm1.HToLocAmt.value		 = "<%=ConvSPChars(Request("txtToLocAmt"))%>"
			.frm1.HCur.value			 = "<%=ConvSPChars(Request("txtCur"))%>"
			.frm1.HFromDate.value		 = "<%=ConvSPChars(Request("txtFromDate"))%>"
			.frm1.HToDate.value			 = "<%=ConvSPChars(Request("txtToDate"))%>"
			.frm1.HOpenBankCd.value		 = "<%=ConvSPChars(Request("txtOpenBankCd"))%>"
		End If
		
		.ggoSpread.Source  = .frm1.vspdData
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
		
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",6),"A", "Q" ,"X","X")		
		
		.lgPageNo	  	   =  "<%=lgPageNo%>"  				  '��: Next next data tag
        .DbQueryOk
        .frm1.vspdData.Redraw = True
	End If

	End with	
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>	
