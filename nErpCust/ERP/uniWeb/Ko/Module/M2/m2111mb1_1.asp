<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MB1_1
'*  4. Program Name         : ���ſ�û��� 
'*  5. Program Desc         : ���ſ�û��ȣ�� ��ü�������� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/12/06
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim lgstrData
	Dim iStrPoNo
	Dim StrNextKey		' ���� �� 
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' ���� �� Return ���� ���� ������ ���� ����     
    Dim lgDataExist
    Dim lgPageNo
    
    dim lgPageNo1
    Dim lgStrPrevKeyM
    
    Dim lglngHiddenRows
    Dim lRow
    
	
	DIM MaxRow2
	
 
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call HideStatusWnd                                                               '��: Hide Processing message
	Call SubBizQueryMulti


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))

	Call FixUNISQLData()
	Call QueryData()	
End Sub    


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
	Dim arrVal(3)
	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,0)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "m2111ma201" 											' header

	strval = ""
	
	If Len(Request("txtPrno")) Then
		strVal = " " & FilterVar(Trim(Request("txtPrno")), " " , "S") & " "		
	End If

    UNIValue(0,UBound(UNIValue,2)) = strval			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
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
    
	If Not(rs0.EOF Or rs0.BOF) Then
		Call  MakeSpreadSheetData()
	Else
		lgPageNo = ""
        rs0.Close
        Set rs0 = Nothing
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    lgDataExist    = "Yes"
  
	If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
	Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1

        iRowStr = ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sppl_cd"))	    
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("bp_nm"))				 
		iRowStr = iRowStr & Chr(11) & UNIConvNum(rs0("quota_rate"),0)				         							
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("apportion_qty"),ggQty.DecPoint,0)			
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("pur_plan_dt"))		
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("pur_grp"))	                   
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("pur_grp_nm"))
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount    
		
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

	lgstrData  = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If

    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing
End Sub

%>

<Script Language=vbscript>
    With parent
		If "<%=lgDataExist%>" = "Yes" Then
		   .ggoSpread.Source  = .frm1.vspdData2
		   .ggoSpread.SSShowData "<%=lgstrData%>"          '�� : Display data
		   .lgPageNo2      =  "<%=lgPageNo%>"               '�� : Next next data tag
		   .DbQueryOk2
		End If  
	End with
</Script>	
 	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
