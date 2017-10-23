<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Common Popup
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/26
'*  7. Modified date(Last)  : 2000/09/26
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

%>
<!-- #Include file="../inc/IncServer.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

    Dim ADF														'ActiveX Data Factory ���� �������� 
    Dim strRetMsg												'Record Set Return Message �������� 
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 
	Dim StrData
	Dim iLoop,jLoop
	Dim isOverFlowKey
	Dim isOverFlowName
	
    Const C_SHEETMAXROWS = 30									'��ȭ�鿡 ���ϼ� �ִ� �ִ� Row �� 


    Call HideStatusWnd

If Request("arrField") <> "" Then
	Dim strSelect					'SELECT �� Field �������� ���� 
	Dim strTable					'SELECT �ϰ����ϴ� Table�� ���� ���� 
	Dim strWhere					'SELECT �ϰ����ϴ� SQL������ WHERE ������ ���� ���� 
	Dim intDataCount

	Redim UNISqlId(0)
	Redim UNIValue(0, 2)
	
	intDataCount = Request("gintDataCnt")
	strTable     = Request("txtTable")
	strWhere     = Request("txtWhere")

    strSelect = replace(Request("arrField"),gColSep,",")
    strSelect = Left(strSelect,Len(Trim(strSelect)) - 1)
    
	
	
	UNISqlId(0) = "compopup"
	UNIValue(0, 0) = strSelect
	UNIValue(0, 1) = strTable
	UNIValue(0, 2) = strWhere
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	If Not (rs0.EOF And rs0.BOF) Then

       isOverFlowKey  = ""
       isOverFlowName = ""
       strData        = ""
 
       For iLoop = 0 to rs0.RecordCount-1
         If iLoop < C_SHEETMAXROWS Then
		    For jLoop = 0 To intDataCount - 1
                strData = strData & Chr(11) & rs0(jLoop)
            Next    
			strData = strData & Chr(11) & Chr(12)
         Else
		    isOverFlowKey  = rs0(0)
			isOverFlowName = rs0(1)
			Exit For
		End If
        rs0.MoveNext
	   Next
	End If   

    rs0.Close
    Set rs0 = Nothing
    Set ADF = Nothing
End If    
%>		

<Script Language="vbscript">   
  On Error Resume Next
	With parent
        .ggoSpread.SSShowData  "<%=ConvSPChars(strData)%>"
        .lgStrCodeKey        = "<%=ConvSPChars(isOverFlowKey)%>"
        .lgStrNameKey        = "<%=ConvSPChars(isOverFlowName)%>"
        .vspdData.focus		
        .DbQueryOk()
	End With

</Script>
