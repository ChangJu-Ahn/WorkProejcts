<%
'======================================================================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ְ��� 
'*  3. Program ID           : S3135QB1
'*  4. Program Name         : �������ະ��ȸ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/08
'*  8. Modified date(Last)  : 2002/02/15
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn tae hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Dateǥ������ 
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

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgPageNo																'�� : ���� �� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Const C_SHEETMAXROWS_D  = 100                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim arrRsVal(1)
'--------------- ������ coding part(��������,End)----------------------------------------------------------
Call HideStatusWnd 

lgPageNo	   = UNICInt(Trim(Request("lgPageNo")),0)
lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Orderby value

Call TrimData()
Call FixUNISQLData()
Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
	Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(1,2)

    UNISqlId(0) = "S3135QA101"
    UNISqlId(1) = "s0000qa001"
   

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	strVal = "AND A.SO_QTY > 0 "
	
	If Len(Request("txtTrackingNo")) Then
		strVal = strVal & "AND A.TRACKING_NO = " & FilterVar(Request("txtTrackingNo"), "''", "S") & " "
	End If

	If Len(Request("txtItemCode")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Request("txtItemCode"), "''", "S") & " "		
		arrVal = Request("txtItemCode")
	End If		
		   
	If Len(Request("txtSoNo")) Then
		strVal = strVal & " AND A.SO_NO = " & FilterVar(Request("txtSoNo"), "''", "S") & " "
	End If		

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(arrVal, "" , "S")
        
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

    FalsechkFlg = False
    
    iStr = Split(lgstrRetMsg,gColSep)
	
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtItemCode")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtItemCode.focus    
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
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtTrackingNo.focus    
            </Script>
        <%        
    Else   
        Call  MakeSpreadSheetData()
    End If
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub
%>

<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"			'��: Display data 
        .lgPageNo        =  "<%=lgPageNo%>"		 '��: set next data tag         
      	.frm1.txtHTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtSoNo"))%>"
		.frm1.txtHItemCode.value = "<%=ConvSPChars(Request("txtItemCode"))%>"            
        .frm1.txtItemCodeNm.value = "<%=ConvSPChars(arrRsVal(1))%>"  
		.DbQueryOk      
    End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
