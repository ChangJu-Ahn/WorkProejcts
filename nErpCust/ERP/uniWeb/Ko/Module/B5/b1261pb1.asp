<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : �������� 
'*  3. Program ID           : S1261PB1
'*  4. Program Name         : �ŷ�ó�˾� 
'*  5. Program Desc         : �ŷ�ó������ �ŷ�ó�˾� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Cho inkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2000/12/09
'*                            2001/12/18  Date ǥ������ 
'*							  2002/04/12 ADO ��ȯ 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort ���� 

Dim strBiz_grp_nm											  ' �����׷�� 
Dim strPur_grp_nm
Dim BlankchkFlg 											  ' ���ű׷�� 

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(30)             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
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

    If iLoopCount < lgMaxCount Then                                 '��: Check if next data exists
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
        strBiz_grp_nm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtBiz_grp")) And BlankchkFlg = False  Then
			Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	
			BlankchkFlg = True
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strPur_grp_nm =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtPur_grp")) And BlankchkFlg = False  Then
			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	
			BlankchkFlg = True
		End If			
    End If     
     
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(3,2)

    UNISqlId(0) = "B1261PA101"
    UNISqlId(1) = "s0000qa005"					'�����׷�� 
    UNISqlId(2) = "s0000qa019"					'���ű׷��    
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtBp_cd")) Then
		strVal = "AND A.BP_CD LIKE " & FilterVar("%" & Trim(UCase(Request("txtBp_cd"))) & "%", "''", "S")	
	Else
		strVal = ""
	End If

	If Len(Request("txtBp_nm")) Then
		strVal = strVal & " AND A.BP_NM LIKE " & FilterVar("%" & Trim(UCase(Request("txtBp_nm"))) & "%", "''", "S")			
	End If		
		   
	If Len(Request("txtBiz_grp")) Then
		strVal = strVal & " AND A.BIZ_GRP = " & FilterVar(Trim(UCase(Request("txtBiz_grp"))), " " , "S") & " "		
		arrVal(0) = Trim(Request("txtBiz_grp"))
	End If		
    
 	If Len(Request("txtPur_grp")) Then
		strVal = strVal & " AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPur_grp"))), " " , "S") & " "		
		arrVal(1) = Trim(Request("txtPur_grp")) 
	End If
	
	If Trim(Request("txtRadio2")) = "C" Or Trim(Request("txtRadio2")) = "S" Then
		strVal = strVal & " AND A.BP_TYPE LIKE  " & FilterVar("%" & Trim(Request("txtRadio2")) & "%", "''", "S") & ""		
	End If
	
	If Trim(Request("txtRadio3")) = "Y" Or Trim(Request("txtRadio3")) = "N" Then
		strVal = strVal & " AND A.USAGE_FLAG = " & FilterVar(Request("txtRadio3"), "''", "S") & ""		
	End If   	
	
	If Len(Request("txtOwnRgstN")) Then
		strVal = strVal & " AND A.BP_RGST_NO = " & FilterVar(Request("txtOwnRgstN"), "''", "S") & " "		
	End If   	
	
	
	UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(arrVal(0), " " , "S")									'�����׷� 
    UNIValue(2,0) = FilterVar(arrVal(1), " " , "S")									'���ű׷�    
    
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))					  '��: ǥ�������� �Է� 
    UNILock = DISCONNREAD :	UNIFlag = "1"										  '��: set ADO read mode
 
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
    
    
    Call  SetConditionData()

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And BlankchkFlg  =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>
<Script Language=vbscript>
    With parent
		.frm1.txtSales_grp_nm.value	= "<%=ConvSPChars(strBiz_grp_nm)%>" 
		.frm1.txtPur_grp_nm.value	= "<%=ConvSPChars(strPur_grp_nm)%>"
'		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.HBp_cd.value	= "<%=ConvSPChars(Request("txtBp_cd"))%>"
				.frm1.HBp_nm.value	= "<%=ConvSPChars(Request("txtBp_nm"))%>"
				.frm1.HBiz_grp.value= "<%=ConvSPChars(Request("txtBiz_grp"))%>"
				.frm1.HPur_grp.value= "<%=ConvSPChars(Request("txtPur_grp"))%>"				
				.frm1.HRadio2.value	= "<%=Request("txtRadio2")%>"
				.frm1.HRadio3.value	= "<%=Request("txtRadio3")%>"					
				.frm1.HOwn_Rgst_N.value	= "<%=ConvSPChars(Request("txtOwnRgstN"))%>"					
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>"                  '��: Display data 																					
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
