<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1411QB1
'*  4. Program Name         : ������Ȳ��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/11/01
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : KimTaeHyun
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
Dim arrRsVal(3)
Dim BlankchkFlg
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
Call LoadBasisGlobalInf()
Call HideStatusWnd 

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = CInt(100)                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
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
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
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
    Dim arrVal(1)
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,2)

    UNISqlId(0) = "S1411QA101"
	UNISqlId(1) = "S0000QA004"
	UNISqlId(2) = "S0000QA014"
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
		
	If Len(Request("txtCreditGrp")) Then
		strVal = " AND A.CREDIT_GRP =  " & FilterVar(Trim(UCase(Request("txtCreditGrp"))), "" , "S") & "  "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(Request("txtCreditGrp")), "", "S")
	
	If Len(Request("txtCurrency")) Then
		strVal = strVal & " AND A.CUR = " & FilterVar(Trim(UCase(Request("txtCurrency"))), "" , "S") & "  "				
	End If		
	arrVal(1) = FilterVar(Trim(Request("txtCurrency")), "", "S")

    UNIValue(0,1) = strVal   '---������ 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)

    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtCreditGrp")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���Ű����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtCreditGrp.focus    
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

		If Len(Request("txtCurrency")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "ȭ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtCurrency.focus    
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
		If  rs0.EOF And rs0.BOF And BlankchkFlg =False Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtCreditGrp.focus    
                </Script>
            <%	       
   		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub

%>

<Script Language=vbscript>
With parent

    .frm1.txtCreditGrpNm.value = "<%=ConvSPChars(arrRsVal(1))%>"
    
	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists

			.frm1.HtxtCreditGrp.value = "<%=ConvSPChars(Request("txtCreditGrp"))%>"
			.frm1.HtxtCurrency.value = "<%=ConvSPChars(Request("txtCurrency"))%>"
		
		End if 

        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '��: Display data 
        .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
        .DbQueryOk
		    
    End if

End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
