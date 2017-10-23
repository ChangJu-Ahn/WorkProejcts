<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procuremant
'*  2. Function Name        : 
'*  3. Program ID           : mc602pa1
'*  4. Program Name         : L/C Reference ASP		
'*  5. Program Desc         : L/C Reference ASP		
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/28	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Ahn Jung Je	
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

	On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
	Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData

	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "PB")
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
 
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal															  '��:UNISqlId(0)�� ���� �Էº��� 
	
	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(0,2)													  '��: ������ SQL ID�� �Էµ� where ������ ������ �� 2���� �迭 
	
	
    UNISqlId(0) = "M4111PA301"											' main query(spread sheet�� �ѷ����� query statement)
    UNIValue(0,0) = lgSelectList                                          '��: Select list
	
	strVal = ""
	If Len(Request("txtFrRcptDt")) Then 
		strVal = strVal & " AND A.MVMT_RCPT_DT >=  " & FilterVar(UNIConvDate(Request("txtFrRcptDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtToRcptDt")) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <=  " & FilterVar(UNIConvDate(Request("txtToRcptDt")), "''", "S") & " "
	End If
	
	If Len(Request("cboMvmtType")) Then
		strVal = strVal & " AND B.IO_TYPE_CD =  " & FilterVar(Request("cboMvmtType"), "''", "S") & " "
	end if	    
    		
	If Len(Trim(Request("txtSupplier"))) Then
		strVal = strVal & " AND A.BP_CD =  " & FilterVar(Request("txtSupplier"), "''", "S") & " "
	End If
	
	IF LEN(Trim(Request("txtGroup"))) THEN
		strVal = strVal & " AND D.PUR_GRP =  " & FilterVar(Request("txtGroup"), "''", "S") & " "
	END IF	 

	IF Request("txtFlag") = "MC" THEN
		strVal = strVal & " AND A.DLVY_ORD_FLG = " & FilterVar("Y", "''", "S") & "  "
	END IF

	
	UNIValue(0,1) = strVal											'	UNISqlId(0)�� �ι�° ?�� �Էµ�	
    UNIValue(0,2) = UCase(Trim(lgTailList))							'	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
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
    
'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100 
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

	Response.Write "<Script Language=vbscript> " & vbCr   
	Response.Write " With Parent "               & vbCr
	
	Response.Write "	If """ & lgDataExist & """  = ""Yes"" Then " & vbCr  
	Response.Write "		If """ & lgPageNo & """ = ""1"" Then " & vbCr  
	Response.Write "			.frm1.hdnMvmtType.value  = """ & ConvSPChars(Request("cboMvmtType")) & """" & vbCr
	Response.Write "			.frm1.hdnSupplier.value  = """ & ConvSPChars(Request("txtSupplier")) & """" & vbCr
	Response.Write "			.frm1.hdnFrRcptDt.value  = """ & Request("txtFrRcptDt") & """" & vbCr
	Response.Write "			.frm1.hdnToRcptDt.value = """ & Request("txtToRcptDt") & """" & vbCr
	Response.Write "			.frm1.hdnGroup.value  = """ & ConvSPChars(Request("txtGroup")) & """" & vbCr
	Response.Write "		End If       " & vbCr                    
	Response.Write "	End If       " & vbCr                    	    
	
	Response.Write "	.ggoSpread.Source     = .frm1.vspdData "    & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iTotstrData  & """" & vbCr
	Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr
	Response.Write "	.DbQueryOk "                                                                                                   & vbCr  
	Response.Write "End With       " & vbCr                    
	Response.Write "</Script>      " & vbCr   
	    
	Response.End 

%>
