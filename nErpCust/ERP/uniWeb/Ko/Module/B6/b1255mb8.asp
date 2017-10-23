<%
'************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : B1255MB8
'*  4. Program Name         : ����������ȸ 
'*  5. Program Desc         : ����������ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2002/04/21
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													

On Error Resume Next

Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '�� : DBAgent Parameter ���� 
Dim rs1, rs2 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 

Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim BlankchkFlg

Dim strSalesOrg	                                                       
Dim strUpperSalesOrg	                                                           
Const C_SHEETMAXROWS_D = 30              
Dim arrRsVal(3)							
  
	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'��:
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag    
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgDataExist    = "No"

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
     Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(3)           '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "B1255MA801"			'* : ������ ��ȸ�� ���� SQL�� ���� 
	
	UNISqlId(1) = "B1255MA802"			'�������� 
	UNISqlId(2) = "B1255MA802"			'������������ 
		
	'--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list     
	UNIValue(1,0)  = UCase(strSalesOrg)
    UNIValue(2,0)  = UCase(strUpperSalesOrg)
        
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	strVal = ""            
	If Trim(Request("txtSales_Org")) <> "" Then
		strVal = strVal& " A.SALES_ORG >=  " & FilterVar(Trim(UCase(Request("txtSales_Org"))), " " , "S") & "  AND A.SALES_ORG <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	Else
		strVal = strVal& " A.SALES_ORG >= '' AND A.SALES_ORG <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	End If
	
	If Trim(Request("txtUpper_Sales_Org")) <> "" Then
		strVal = strVal& " AND A.UPPER_SALES_ORG >=  " & FilterVar(Trim(UCase(Request("txtUpper_Sales_Org"))), " " , "S") & "  AND A.UPPER_SALES_ORG <=  " & FilterVar(Trim(UCase(Request("txtUpper_Sales_Org"))), " " , "S") & " "
	Else		
		strVal = strVal
	End If
	
	If Trim(Request("txtlvl")) <> "" and Trim(Request("txtlvl")) <> "0"Then
		strVal = strVal& " AND A.LVL >= " & CInt(Trim(Request("txtlvl"))) & " AND A.LVL <= " & CInt(Trim(Request("txtlvl")))
	Else
		strVal = strVal& " AND A.LVL >= 0 "
	End If
  	
	If Trim(Request("txtRadio")) <> "" Then
		strVal = strVal& " AND A.USAGE_FLAG >= " & FilterVar(UCase(Request("txtRadio")), "''", "S") & " AND A.USAGE_FLAG <=  " & FilterVar(UCase(Request("txtRadio")), "''", "S") & ""
	Else
		strVal = strVal& " AND A.USAGE_FLAG >='' AND A.USAGE_FLAG <= " & FilterVar("zzzzzzzzz", "''", "S") & "  "
	End If
		
    UNIValue(0,1) = strVal    

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
    BlankchkFlg = False
    
	
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs0,rs1,rs2) '* : Record Set �� ���� ���� 
    
    Set lgADF   = Nothing

    
    iStr = Split(lgstrRetMsg,gColSep)
	
	'============================= �߰��� �κ� =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtSales_Org")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSales_Org.focus    
                </Script>
            <%	       	
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
        If Len(Request("txtUpper_Sales_Org")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "������������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtUpper_Sales_Org.focus    
                </Script>
            <%	       	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
			
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    			
			
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg = False Then
		    Call DisplayMsgBox("125500", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSales_Org.focus    
                </Script>
            <%		    
    
			' �� ��ġ�� �ִ� Response.End �� �����Ͽ��� ��. Client �ܿ��� Name�� ��� �ѷ��� �Ŀ� Response.End �� �����.
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
	
	'�������� 
    If Len(Trim(Request("txtSales_Org"))) Then
    	strSalesOrg = " " & FilterVar(Trim(Request("txtSales_Org")), " " , "S") & " "
    Else
    	strSalesOrg = "''"
    End If
    '������������ 
    If Len(Trim(Request("txtUpper_Sales_Org"))) Then
    	strUpperSalesOrg = " " & FilterVar(Trim(Request("txtUpper_Sales_Org")), " " , "S") & " "
    Else
    	strUpperSalesOrg = "''"
    End If
	
End Sub


%>
<Script Language=vbscript>
    parent.frm1.txtSales_Org_nm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  	parent.frm1.txtUpper_Sales_OrgNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HSales_Org.value			= "<%=ConvSPChars(Request("txtSales_Org"))%>"
				.frm1.HRadio.value				= "<%=Request("txtRadio")%>"
				.frm1.HUpper_Sales_Org.value	= "<%=ConvSPChars(Request("txtUpper_Sales_Org"))%>"
				.frm1.Hlvl.value				= "<%=Request("txtlvl")%>"
			End If
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>"          '��: Display data 
			.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag		
			.DbQueryOk
		
		End with
	End If   
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
