<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m4111rb2
'*  4. Program Name         : �����԰����� 
'*  5. Program Desc         : �����԰����� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/03/21	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Shin jin hyun
'* 10. Modifier (Last)      : Kang Su Hwan
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
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0	   		      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim strPtnBpNm												  ' ��ǰó�� 
Dim strDNTypeNm												  ' �������¸� 
Dim strSOTypeNm											      ' ����Ÿ�Ը� 

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
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        PvArr(iLoopCount) = lgstrData
        lgstrData=""
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

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
    Dim arrVal(2)
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,2)

    UNISqlId(0) = "M4111RA201"    
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	If Len(Request("txtSppl")) Then
		strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtSppl"))), " " , "S") & " "	
	End If

	If Len(Request("txtRefPONO")) Then
		strVal = strVal & " AND G.PO_NO = " & FilterVar(Trim(UCase(Request("txtRefPONO"))), " " , "S") & " "		
	End If		
		   
	If Len(Request("txtcur")) Then
		strVal = strVal & " AND A.MVMT_CUR = " & FilterVar(Trim(UCase(Request("txtcur"))), " " , "S") & " "		
	End If		
    
 	If Len(Request("txtMvmtNo")) Then
		strVal = strVal & " AND A.MVMT_RCPT_NO = " & FilterVar(Trim(UCase(Request("txtMvmtNo"))), " " , "S") & " "		
	End If	    
	
    If Len(Request("txtFrMvmtDt")) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT >= " & FilterVar(UNIConvDate(Request("txtFrMvmtDt")), "''", "S") & ""			
	End If		
	
	If Len(Request("txtToMvmtDt")) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <= " & FilterVar(UNIConvDate(Request("txtToMvmtDt")), "''", "S") & ""		
	End If

	' ���ְ������� �߰� 
	If Len(Request("txtSubcontraflg")) Then
		strVal = strVal & " AND G.SUBCONTRA_FLG = " & FilterVar(Trim(UCase(Request("txtSubcontraflg"))), " " , "S") & ""		
	End If
	
	' �԰��� ��ǰ�� ����Ҽ� �ִ� ������ 0���� Ŭ�븸 ��ȸ�ǵ��� �Ѵ�.
	strVal = strVal & " AND	A.MVMT_QTY - (A.IV_QTY + ISNULL(A.RET_ORD_QTY,0)) > 0 "

'    UNIValue(0,1) = strVal & " " &  "ORDER BY A.MVMT_NO DESC"
    UNIValue(0,1) = strVal & " " & UCase(Trim(lgTailList))
    
'--------------- ������ coding part(�������,End)------------------------------------------------------

'    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF Then
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
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.hdnFrMvmtDt.Value 	= "<%=Request("txtFrMvmtDt")%>"
				.frm1.hdnToMvmtDt.Value 	= "<%=Request("txtToMvmtDt")%>"
				.frm1.hdnMvmtNo.Value 		= "<%=ConvSPChars(Request("txtMvmtNo"))%>"
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
			.lgPageNo			 =  "<%=lgPageNo%>"							  '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
