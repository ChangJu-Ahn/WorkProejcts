<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : s3111pb1.asp	
'*  4. Program Name         : ���ֹ�ȣ�˾� 
'*  5. Program Desc         : ���ֹ�ȣ�˾� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
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
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")
On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim BlankchkFlg
Dim lgPageNo
Const C_SHEETMAXROWS_D  = 30                                '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Dim arrRsVal(7)												'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Dim arrVal(3)															
    Dim MajorCd
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,2)

    UNISqlId(0) = "S3111pa101"									'* : ������ ��ȸ�� ���� SQL�� 
    UNISqlId(1) = "S0000QA002"									'* : ������ ��ȸ���Ǻθ��� Name �� �������� SQL ���� ���� 
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "s0000qa007"
    UNISqlId(4) = "S0000QA000"
 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "

	If Len(Request("txtBpCd")) Then
		strVal = " AND A.SOLD_TO_PARTY = " & FilterVar(Request("txtBpCd"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(Request("txtBpCd")),"","S")

	If Len(Request("txtSalesGroup")) Then		
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
	End If	
	arrVal(1) = FilterVar(Trim(Request("txtSalesGroup"))," ","S")
	
 	If Len(Request("txtSo_Type")) Then		
		strVal = strVal & " AND A.SO_TYPE = " & FilterVar(Request("txtSo_Type"), "''", "S") & " "		
	End If		
    arrVal(2) = FilterVar(Trim(Request("txtSo_Type"))," ","S")
    
 	If Len(Request("txtPay_terms")) Then		
		strVal = strVal & " AND A.PAY_METH = " & FilterVar(Request("txtPay_terms"), "''", "S") & " "		
		MajorCd = "B9004"
	End If
	arrVal(3) = FilterVar(Trim(Request("txtPay_terms"))," ","S")		


	If Trim(Request("txtRadio")) = "Y" Then
		strVal = strVal & " AND A.CFM_FLAG =" & FilterVar("Y", "''", "S") & " "
	ElseIf Trim(Request("txtRadio")) = "N" Then
		strVal = strVal & " AND A.CFM_FLAG =" & FilterVar("N", "''", "S") & " "
	End If			
		
    If Len(Trim(Request("txtSOFrDt"))) Then
		strVal = strVal & " AND A.SO_DT >= " & FilterVar(UNIConvDate(Request("txtSOFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtSoToDt"))) Then
		strVal = strVal & " AND A.SO_DT <= " & FilterVar(UNIConvDate(Request("txtSoToDt")), "''", "S") & ""		
	End If
	
	If Trim(Request("txtSTOFlag")) = "N" Then
		strVal = strVal & " AND A.STO_FLAG = " & FilterVar("N", "''", "S") & " "
	ElseIf Trim(Request("txtSTOFlag")) = "Y" Then
		strVal = strVal & " AND A.STO_FLAG = " & FilterVar("Y", "''", "S") & " "
	End If
	
	If Trim(Request("txtPopFlag")) = "invoice" Then
		strVal = strVal & " AND A.EXPORT_FLAG = " & FilterVar("Y", "''", "S") & " "
	End If	
	
	If Trim(Request("txtPopFlag")) = "alloc" Then
		strVal = strVal & " AND A.RET_ITEM_FLAG = " & FilterVar("N", "''", "S") & " "
	End If

    If Len(Trim(Request("gBizArea"))) Then
		strVal = strVal & " AND A.BIZ_AREA = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If		
	
	

    UNIValue(0,1) = strVal   '---������ 
    UNIValue(1,0) = arrVal(0)  
    UNIValue(2,0) = arrVal(1)  
    UNIValue(3,0) = arrVal(2) 
    UNIValue(4,0) =  FilterVar(MajorCd , "''", "S") 
    UNIValue(4,1) = arrVal(3)  
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4) '* : Record Set �� ���� ���� 
    iStr = Split(lgstrRetMsg,gColSep)

    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtBpCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�ֹ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True		       
	       	Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtBpCd.focus " & vbCr    
            Response.Write "</Script>      " & vbCr		    
            
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

	 	If Len(Request("txtSalesGroup")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True	
	        Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtSalesGroup.focus " & vbCr    
            Response.Write "</Script>      " & vbCr		    
            
		End If	
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
 
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

	 	If Len(Request("txtSo_Type")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True	
	        Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtSo_Type.focus " & vbCr    
            Response.Write "</Script>      " & vbCr		    	       
		End If	
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing

	 	If Len(Request("txtPay_terms")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True	
	        Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtPay_terms.focus " & vbCr    
            Response.Write "</Script>      " & vbCr		    	       
		End If	
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
		    		    
            Response.Write "<Script language=vbs>  " & vbCr   
            Response.Write "   Parent.frm1.txtBpCd.focus " & vbCr    
            Response.Write "</Script>      " & vbCr		    	
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub


%>
<Script Language=vbscript>
    With parent
        
        .frm1.txtBpNm.value				=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		.frm1.txtSalesGroupNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtSo_TypeNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>" 	
  		.frm1.txtPay_terms_nm.value		=  "<%=ConvSPChars(arrRsVal(7))%>" 
		
		'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.HBpCd.value		= "<%=ConvSPChars(Request("txtBpCd"))%>"
			.frm1.HSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
			.frm1.HSo_Type.value	= "<%=ConvSPChars(Request("txtSo_Type"))%>"
			.frm1.HPay_terms.value	= "<%=ConvSPChars(Request("txtPay_terms"))%>"
			.frm1.HRadio.value		= "<%=Request("txtRadio")%>"
			.frm1.HSOFrDt.value		= "<%=Request("txtSOFrDt")%>"
			.frm1.HSoToDt.value		= "<%=Request("txtSoToDt")%>"
		End If  	

		.ggoSpread.Source    = .frm1.vspdData 
		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip  "<%=lgstrData%>","F"                            '��: Display data 
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,-1,-1,.GetKeyPos("A",5),.GetKeyPos("A",6),"A","Q","X","X")
		.lgPageNo			 = "<%=lgPageNo%>"							  '��: Next next data tag
 		.DbQueryOk
  		.frm1.vspdData.Redraw = True
  		
	End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
