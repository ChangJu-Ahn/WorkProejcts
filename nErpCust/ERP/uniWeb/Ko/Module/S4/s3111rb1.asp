<%
'********************************************************************************************************
'*  1. Module Name          : ����																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3111rb1.asp																*
'*  4. Program Name         : ��������(L/C���) 														*
'*  5. Program Desc         : ��������(L/C���)															*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				  '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5									  '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strApplicantNm											  ' �ֹ�ó�� 
Dim strSalesGroupNm											  ' �����׷�� 
Dim strSOTypeNm											      ' �������¸� 
Dim strPayTermsNm										      ' ������� 
Dim strIncoTermsNm
Dim BlankchkFlg											  ' �������� 
Dim iFrPoint
iFrPoint=0
Const C_SHEETMAXROWS_D  = 30                                         

    Call HideStatusWnd 

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
Sub SetConditionData()
    On Error Resume Next
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strApplicantNm =  rs1("BP_NM")
        rs1.Close
        Set rs1 = Nothing
    Else
		rs1.Close
		Set rs1 = Nothing
		If Len(Request("txtApplicant")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "�ֹ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			BlankchkFlg  =  True
		%>
            <Script language=vbs>
            Parent.txtApplicant.focus    
            </Script>
         <%		   						
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSalesGroupNm =  rs2("SALES_GRP_NM")
        rs2.Close
        Set rs2 = Nothing
    Else
		rs2.Close
		Set rs2 = Nothing
		If Len(Request("txtSalesGroup")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			BlankchkFlg  =  True
		%>
            <Script language=vbs>
            Parent.txtSalesGroup.focus    
            </Script>
         <%		   				
		End If			
    End If   	
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strSOTypeNm =  rs3("SO_TYPE_NM")
        rs3.Close
        Set rs3 = Nothing
    Else
		rs3.Close
		Set rs3 = Nothing
		If Len(Request("txtSOType")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    BlankchkFlg  =  True
		%>
            <Script language=vbs>
            Parent.txtSOType.focus    
            </Script>
         <%			
		End If				
    End If      
    
    If Not(rs4.EOF Or rs4.BOF) Then
        strPayTermsNm =  rs4("MINOR_NM")
        rs4.Close
        Set rs4 = Nothing
    Else
		rs4.Close
		Set rs4 = Nothing
		If Len(Request("txtPayTerms")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		    BlankchkFlg  =  True
		%>
            <Script language=vbs>
            Parent.txtPayTerms.focus    
            </Script>
         <%				
		End If				
    End If      
    
    If Not(rs5.EOF Or rs5.BOF) Then
        strIncoTermsNm =  rs5("MINOR_NM")
        rs5.Close
        Set rs5 = Nothing
    Else
		rs5.Close
		Set rs5 = Nothing
		If Len(Request("txtIncoTerms")) And BlankchkFlg =  False Then
			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			BlankchkFlg  =  True
		%>
            <Script language=vbs>
            Parent.txtIncoTerms.focus    
            </Script>
         <%						
		End If				
    End If      
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(5,2)

    UNISqlId(0) = "S3111RA101"
    UNISqlId(1) = "s0000qa002"					'�ֹ�ó�� 
    UNISqlId(2) = "s0000qa005"					'�����׷�� 
    UNISqlId(3) = "s0000qa007"					'�������¸� 
    UNISqlId(4) = "s0000qa000"					'���������  
    UNISqlId(5) = "s0000qa000"					'�������Ǹ�    
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = ""

	If Len(Request("txtApplicant")) Then
		strVal = strVal & "AND c.bp_cd = " & FilterVar(Request("txtApplicant"), "''", "S") & "  "	
		arrVal(0) = Trim(Request("txtApplicant")) 
	End If

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & "AND b.sales_grp = " & FilterVar(Request("txtSalesGroup"), "''", "S") & "  "		
		arrVal(1) = Trim(Request("txtSalesGroup")) 
	End If		
		
 	If Len(Request("txtSOType")) Then
		strVal = strVal & "AND d.so_type = " & FilterVar(Request("txtSOType"), "''", "S") & "  "		
		arrVal(2) = Trim(Request("txtSOType")) 
	End If	    
	
    If Len(Request("txtFromDt")) Then
		strVal = strVal & "AND a.so_dt >= " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & " "			
	End If		
	
	If Len(Request("txtToDt")) Then
		strVal = strVal & "AND a.so_dt <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & " "		
	End If
	
	If Len(Request("txtPayTerms")) Then
		strVal = strVal & "AND a.pay_meth = " & FilterVar(Request("txtPayTerms"), "''", "S") & "  "		
		arrVal(3) = Trim(Request("txtPayTerms")) 
	End If		
    
	If Len(Request("txtIncoTerms")) Then
		strVal = strVal & "AND a.incoterms = " & FilterVar(Request("txtIncoTerms"), "''", "S") & "  "	
		arrVal(4) = Trim(Request("txtIncoTerms"))
	End If		
	
	UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(Trim(Request("txtApplicant")), " " , "S")					'�ֹ�ó�ڵ� 
    UNIValue(2,0) = FilterVar(Trim(Request("txtSalesGroup")), " " , "S")					'�����׷��ڵ� 
    UNIValue(3,0) = FilterVar(Trim(Request("txtSOType")), " " , "S")					'���������ڵ� 
    UNIValue(4,0) = FilterVar("B9004", " " , "S")						'��������ڵ�    
    UNIValue(4,1) = FilterVar(Trim(Request("txtPayTerms")), " " , "S")					'��������ڵ�    
    UNIValue(5,0) = FilterVar("B9006", " " , "S")						'���������ڵ� 
    UNIValue(5,1) = FilterVar(Trim(Request("txtIncoTerms")), " " , "S")					'���������ڵ�                
    
'--------------- ������ coding part(�������,End)------------------------------------------------------

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
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)


	Call  SetConditionData()


	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    If BlankchkFlg = False Then         
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
		%>
            <Script language=vbs>
            Parent.txtApplicant.focus    
            </Script>
         <%		 		    
		Else    
		    Call  MakeSpreadSheetData()	    
		End If
	End If
		  
End Sub

%>

<Script Language=vbscript>
	With parent
		.txtApplicantNm.value	= "<%=ConvSPChars(strApplicantNm)%>" 
		.txtSalesGroupNm.value	= "<%=ConvSPChars(strSalesGroupNm)%>" 
        .txtSOTypeNm.value		= "<%=ConvSPChars(strSOTypeNm)%>"
        .txtPayTermsNm.value	= "<%=ConvSPChars(strPayTermsNm)%>"
        .txtIncoTermsNm.value	= "<%=ConvSPChars(strIncoTermsNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.txtHApplicant.value	= "<%=ConvSPChars(Request("txtApplicant"))%>"
				.txtHSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>" 
				.txtHSOType.value		= "<%=ConvSPChars(Request("txtSOType"))%>"
				.txtHFromDt.value		= "<%=ConvSPChars(Request("txtFromDt"))%>"
				.txtHToDt.value			= "<%=ConvSPChars(Request("txtToDt"))%>"
				.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"
				.txtHIncoTerms.value	= "<%=ConvSPChars(Request("txtIncoTerms"))%>"
			.DbQueryOk
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .vspdData 
			.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                            '��: Display data 
        
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.vspdData,"<%=iFrPoint+1%>",.vspddata.maxrows,Parent.GetKeyPos("A",6),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
			        
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
			.vspdData.Redraw = True
		End If
	End with
</Script>
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
