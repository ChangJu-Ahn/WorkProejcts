<%
'************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : �������� 
'*  3. Program ID           : B1263MB8
'*  4. Program Name         : ������̷���ȸ 
'*  5. Program Desc         : ������̷���ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Sonbumyeol
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
Dim lgDataExist
Dim lgPageNo

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0			                   '�� : DBAgent Parameter ���� 
Dim rs1
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim BlankchkFlg
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strtxtConBpcd	                                                       

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

    UNISqlId(0) = "B1263MA801"									'* : ������ ��ȸ�� ���� SQL�� ���� 
	
	UNISqlId(1) = "B1261MA802"			'�ŷ�ó 
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
     
	UNIValue(1,0)  = UCase(Trim(strtxtConBpcd))
        
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strVal = ""
    
	If Trim(Request("txtConBp_cd")) <> "" Then
		strVal = strVal& " A.BP_CD = " & FilterVar(Trim(UCase(Request("txtConBp_cd"))), " " , "S") & "  "
	End If
	
	If Len(Trim(Request("txtConValidFromDt"))) Then
		strVal = strVal & " AND A.VALID_FROM_DT >= " & FilterVar(UNIConvDate(Request("txtConValidFromDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtConValidToDt"))) Then
		strVal = strVal & " AND A.VALID_FROM_DT <= " & FilterVar(UNIConvDate(Request("txtConValidToDt")), "''", "S") & ""		
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


    If Not(rs1.EOF Or rs1.BOF) Then
        arrRsVal(1) =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtConBp_cd")) Then
			Call DisplayMsgBox("970000", vbInformation, "�ŷ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			BlankchkFlg  =  True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtConBp_cd.focus    
                </Script>
            <%					
		End If
	End If   	

    


	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    	
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
		    Call DisplayMsgBox("126300", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtConBp_cd.focus    
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
    If Len(Trim(Request("txtConBp_cd"))) Then
    	strtxtConBpcd = " " & FilterVar(Trim(Request("txtConBp_cd")), " " , "S") & " "
    	
    Else
    	strtxtConBpcd = "''"
    End If
    
End Sub


%>
<Script Language=vbscript>
    parent.frm1.txtConBp_nm.value	=  "<%=ConvSPChars(arrRsVal(1))%>" 	
	
	If "<%=lgDataExist%>" = "Yes" Then
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHBpCode.value = "<%=ConvSPChars(Request("txtConBp_cd"))%>"
				.frm1.txtHValidFDt.value = "<%=Request("txtConValidFromDt")%>"
				.frm1.txtHValidTDt.value = "<%=Request("txtConValidToDt")%>"
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
