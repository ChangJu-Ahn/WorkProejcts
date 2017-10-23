<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Project Manage
'*  2. Function Name        : 
'*  3. Program ID           : Y7001PB1
'*  4. Program Name         : Project Popup
'*  5. Program Desc         : Project Popup
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2005/04/20
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Nhg
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	On Error Resume Next
                                                                         
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2      '�� : DBAgent Parameter ���� 
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim SortNo													  ' Sort ���� 

	Dim PlantNm														'�� : ����� ���� 
	Dim ItemNm										   				    '�� : ǰ��� ���� 
	Dim ProjectNm
	
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

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
	Const C_SHEETMAXROWS_D  = 1000            
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
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        PvArr(iLoopCount) = lgstrData
        lgstrData=""
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

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
Function SetConditionData()
    On Error Resume Next
	    
	SetConditionData = FALSE
	    
    If Not(rs1.EOF Or rs1.BOF) Then
        ProjectNm = rs1("Project_Nm")
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtProjectCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "������Ʈ", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code	
			Exit Function		'20030124 - leejt		
		End If
	End If   	
    
    SetConditionData = TRUE
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim strVal
	Dim sTemp
	Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
    
    UNISqlId(0) = "Y7001PA1"
    'UNISqlId(1) = "Y7111PA102"					'������Ʈ 
    
    UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPROJECTCD"))) ,"''","S")
    
    sTemp = "1"  
    
    If Len(Trim(Request("txtProjectCd"))) Then
        If sTemp = "1" Then
			strVal = strVal & "WHERE A.PROJECT_CODE =  " & FilterVar(Trim(UCase(Request("txtProjectCd"))), " " , "S") & "  "	
			sTemp = "2"
		'Else
		'	strVal = strVal & " AND A.PROJECT_CODE =  " & FilterVar(Trim(UCase(Request("txtProjectCd"))), " " , "S") & "  "	
		End If	
		UNIValue(1,0) = " " & FilterVar(Trim(UCase(Request("txtProjectCd"))), " " , "S") & " "    	'��: Select list
	End If
	
'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtProject"))) ,"''","S")
	'UNIValue(0,1) = strVal 
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If SetConditionData() = False Then Exit Sub
         
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call MakeSpreadSheetData()
    End If  
End Sub

%>
<Script Language=vbscript>
    With parent
		.frm1.txtProjectNm.value = "<%=ConvSPChars(ProjectNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnProjectCd.value	= "<%=Request("txtProjectCd")%>"
			End If    
			'Show multi spreadsheet data from this line
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                  '��: Display data 
																	'0: ����None 1 :��������  2: ��������					
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>