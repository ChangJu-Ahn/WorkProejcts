<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : s3111ab1
'*  4. Program Name         : �������� 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/12
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Cho inkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date ǥ������ 
'*							  2002/04/12 ADO ��ȯ 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%

On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2	  '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strItemNm											      ' ǰ��� 
Dim strItemSpec											      ' ����� 
Dim MsgDisplayFlag

	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 30							             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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

    If iLoopCount < lgMaxCount Then                                      '��: Check if next data exists
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
    
	SetConditionData = False

    If Not(rs1.EOF Or rs1.BOF) Then
        strItemNm =  rs1(1)
        Set rs1 = Nothing
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
       strItemSpec =  rs2(1)
       Set rs2 = Nothing
    End If   	
    
	SetConditionData = True

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(4,2)

    UNISqlId(0) = "S4112RA901"
    UNISqlId(1) = "s0000qa001"					
    UNISqlId(2) = "s0000qa027"					
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtDnNo")) Then
		strVal = "AND A.DOCUMENT_NO = " & FilterVar(Request("txtDnNo"), "''", "S") & ""	
	Else
		strVal = ""
	End If

	If Len(Request("txtDnSeq")) Then
		strVal = strVal & " AND A.DOCUMENT_SEQ_NO = " & FilterVar(Request("txtDnSeq"), "''", "S") & ""	
	Else
		strVal = ""
	End If

	If Len(Request("txtItem")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Request("txtItem"), "''", "S") & ""		
		arrVal = Trim(Request("txtItem"))
	End If		

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(arrVal, "''", "S")	
    UNIValue(2,0) = FilterVar(arrVal, "''", "S")
    
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
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If SetConditionData = False Then Exit Sub

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        MsgDisplayFlag = True
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtItemNm.value			= "<%=ConvSPChars(strItemNm)%>" 
		.frm1.txtSpec.value				= "<%=ConvSPChars(strItemSpec)%>" 
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.txtHDnNo.value		= "<%=ConvSPChars(Request("txtDnNo"))%>"
				.frm1.txtHDnSeq.value		= "<%=ConvSPChars(Request("txtDnSeq"))%>"
				.frm1.txtHItem.value		= "<%=ConvSPChars(Request("txtItem"))%>"
			End If    
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
