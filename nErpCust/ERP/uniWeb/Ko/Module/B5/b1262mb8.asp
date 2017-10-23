<%
'************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : �������� 
'*  3. Program ID           : B1262MB8
'*  4. Program Name         : �ŷ�ó���� ��ȸ 
'*  5. Program Desc         : �ŷ�ó���� ��ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : kim hyung suk
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : -2000/04/29 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'*                            -2002/04/11 : ADO��ȯ 
'**************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													

On Error Resume Next
'�߰� 
Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '�� : DBAgent Parameter ���� 
Dim rs1, rs2 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim BlankchkFlg
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strtxtBpcd	                                                       

'----------------------- �߰��� �κ� ----------------------------------------------------------------------
Dim arrRsVal(1)								'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
'----------------------------------------------------------------------------------------------------------
'--------------- ������ coding part(��������,End)----------------------------------------------------------
	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'��:
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(30)                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(2,2)

    UNISqlId(0) = "B1262MA801"									'* : ������ ��ȸ�� ���� SQL�� ���� 
	
	UNISqlId(1) = "B1261MA802"			'�ŷ�ó 
		

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
     
	UNIValue(1,0)  = UCase(Trim(strtxtBpcd))
        
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strVal = ""
    
	If Trim(Request("txtBp_cd")) <> "" Then
		strVal = strVal& " A.BP_CD = " & FilterVar(Trim(UCase(Request("txtBp_cd"))), " " , "S") & "  "
	End If
  		
    UNIValue(0,1) = strVal   
	
'================================================================================================================   
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs0,rs1) '* : Record Set �� ���� ���� 
    
    
    Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)
	
	'============================= �߰��� �κ� =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtBp_cd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�ŷ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True

            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBp_cd.focus    
                </Script>
            <%	       	       	
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
      	
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

		
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("126200", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtBp_cd.focus    
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

	'---�ŷ�ó 
    If Len(Trim(Request("txtBp_cd"))) Then
    	strtxtBpcd = " " & FilterVar(Trim(Request("txtBp_cd")), " " , "S") & " "
    	
    Else
    	strtxtBpcd = "''"
    End If
    		

End Sub


%>
<Script Language=vbscript>
    parent.frm1.txtBp_nm.value = "<%=ConvSPChars(arrRsVal(1))%>" 		
	
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HBp_cd.value = "<%=ConvSPChars(Request("txtBp_cd"))%>"
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
