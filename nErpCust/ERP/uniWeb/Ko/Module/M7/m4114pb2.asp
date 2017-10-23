<!--
'********************************************************************************************************
'*  1. Module Name          : Prucurement																		*
'*  2. Function Name        : 																	*
'*  3. Program ID           : m4114pb2																	*
'*  4. Program Name         : �������԰�������Ȳ-�����ݾ�(IR) �˾�																*
'*  5. Program Desc         :  																			*
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/10/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next													   '���� ������ �߻��� �� ������ �߻��� ���� �ٷ� ������ ������ ��ӵ� �� �ִ� ������ ��Ʈ���� �ű� �� �ֵ��� �����մϴ�.				

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1       	'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgpageNo	                                            '�� : ���� �� 
Dim lgTailList
Dim lgDataExist

Dim istrData

Dim strBeneficiaryNm

Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)              '�� : Next key flag
lgDataExist     = "No"							                    '�� : Orderby value

Call FixUNISQLData()
Call QueryData()


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
   	Dim strVal
	ReDim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Redim UNIValue(0,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
	UNISqlId(0) = "M4114PA2"		

	If Trim(Request("txtSearchDt")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSearchDt"))), " " , "SNM") & "' "
	Else
	    UNIValue(0,0) = "|"
	End If

	If Trim(Request("txtBpCd")) <> "" Then
		UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "SNM") & "' "
	Else
	    UNIValue(0,1) = "|"
	End If


	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim iStr
   
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    Set lgADF   = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
   
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    
    Else
		Call  MakeSpreadSheetData()
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_S  = 100
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt

	Const C_IV_TYPE_CD	= 0	'�������� 
	Const C_IV_TYPE_NM	= 1	'���������� 
	Const C_IV_LOC_AMT	= 2	'�����ݾ�(IR)


	lgDataExist    = "Yes"

	If CLng(lgPageNo) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_S) * CLng(lgPageNo)                
		intTRows	= CLng(C_SHEETMAXROWS_S) * CLng(lgPageNo)
	End If

	'//Response.end

	'----- ���ڵ�� Į�� ���� ----------
	'-----------------------------------
	iLoopCount = 0

    	ReDim PvArr(C_SHEETMAXROWS_S - 1)

	Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""


		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_IV_TYPE_CD))	        
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_IV_TYPE_NM))	        
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(C_IV_LOC_AMT), ggQty.DecPoint, 0)

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount < C_SHEETMAXROWS_S Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)

        Else
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop

	
	intARows = iLoopCount
	If iLoopCount  < C_SHEETMAXROWS_S Then                                      '��: Check if next data exists
	  lgPageNo = ""
	End If

	rs0.Close                                                       '��: Close recordset object
	Set rs0 = Nothing	 

End Sub

											'��: �����Ͻ� ���� ó���� ������ 
%>
<Script Language=vbscript>
    With parent

		If "<%=lgDataExist%>" = "Yes" Then
			.ggoSpread.Source = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=istrData%>"                  '��: Display data
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
