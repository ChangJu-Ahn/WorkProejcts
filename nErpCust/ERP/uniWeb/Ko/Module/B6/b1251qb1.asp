<%@ LANGUAGE="VBSCRIPT" %>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : b1251qb1
'*  4. Program Name         : ���ű׷���ȸ 
'*  5. Program Desc         : ���ű׷���ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncServer.asp" -->

<%

On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort ���� 

Dim PurOrgNm														'�� : ���������� ���� 
Dim PurGrpNm									   				    '�� : ���ű׷�� ���� 
Dim BizAreaNm														'�� : ������ ���� 
 

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
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
	Const C_SHEETMAXROWS_D  = 100

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			PvArr(iLoopCount) = lgstrData
			lgstrData=""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	lgstrData = join(PvArr,"")
	
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
        PurOrgNm = rs1("PUR_ORG_NM")
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
'		If Len(Request("txtORGCd")) Then
'			Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
'		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        PurGrpNm = rs2("PUR_GRP_NM")
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
'		If Len(Request("txtGroupCd")) Then
'			Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
'		End If			
    End If     
    
	If Not(rs3.EOF Or rs3.BOF) Then
        BizAreaNm = rs3("BIZ_AREA_NM")
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtBACd")) Then
			Call DisplayMsgBox("970000", vbInformation, "�����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
		End If			
    End If         
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	dim sTemp
	Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(3,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
    UNISqlId(0) = "B1251QA101"
    UNISqlId(1) = "s0000qa021"
    UNISqlId(2) = "s0000qa022"
    UNISqlId(3) = "s0000qa013"
    
    UNIValue(1,0) = "'zzzz'"
    UNIValue(2,0) = "'zzzz'"
    UNIValue(3,0) = "'zzzzzzzzzz'"
    
    sTemp = "1"
    '��������                    
    If Len(Trim(Request("txtORGCd"))) Then
        if sTemp = "1" then
			strVal = strVal & " Where A.PUR_ORG >=  " & FilterVar(Request("txtORGCd"), "''", "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PUR_ORG >=  " & FilterVar(Request("txtORGCd"), "''", "S") & "  "	
		end if
		UNIValue(1,0) = FilterVar(Trim(Request("txtORGCd")), "''" ,  "S")     	'��: Select list
	End If
	'���ű׷� 
    If Len(Trim(Request("txtGroupCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " Where A.PUR_GRP >=  " & FilterVar(Request("txtGroupCd"), "''", "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PUR_GRP >=  " & FilterVar(Request("txtGroupCd"), "''", "S") & "  "	
		end if	
		UNIValue(2,0) = FilterVar(Trim(Request("txtGroupCd")), "''" ,  "S")     	'��: Select list
	End If
	'����� 
    If Len(Trim(Request("txtBACd"))) Then
		if sTemp = "1" then
			strVal = strVal & " Where C.BIZ_AREA_CD =  " & FilterVar(Request("txtBACd"), "''", "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND C.BIZ_AREA_CD =  " & FilterVar(Request("txtBACd"), "''", "S") & "  "	
		end if	
		UNIValue(3,0) = FilterVar(Trim(Request("txtBACd")), "''" ,  "S")     	'��: Select list
	End If
    '��뿩��			
    If Trim(Request("rdoUseflg")) = "Y" then
		if sTemp = "1" then
			strVal = strVal & " Where A.USAGE_FLG = " & FilterVar("Y", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.USAGE_FLG = " & FilterVar("Y", "''", "S") & "  "
		end if	
	ElseIf Trim(Request("rdoUseflg")) = "N" then
		if sTemp = "1" then
			strVal = strVal & " Where A.USAGE_FLG = " & FilterVar("N", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.USAGE_FLG = " & FilterVar("N", "''", "S") & "  "
		end if	
	End If
	
 
'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
	UNIValue(0,1) = strVal & " " & Trim(lgTailList)  '---������ 


    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtORGNm.value = "<%=ConvSPChars(PurOrgNm)%>"
		.frm1.txtGroupNm.value = "<%=ConvSPChars(PurGrpNm)%>"
		.frm1.txtBANm.value = "<%=ConvSPChars(BizAreaNm)%>"
'		
		If "<%=lgDataExist%>" = "Yes" Then
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                  '��: Display data 
			
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
		Else
			Parent.frm1.txtORGCd.focus
			Set Parent.gActiveElement = Parent.document.activeElement
		End If
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
