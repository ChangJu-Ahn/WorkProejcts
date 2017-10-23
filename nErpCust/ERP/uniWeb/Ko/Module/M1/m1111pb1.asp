<%@ LANGUAGE="VBSCRIPT" %>
<% Option explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1111PB1
'*  4. Program Name         : ǰ�������˾� 
'*  5. Program Desc         : ǰ�������˾� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : KimTaeHyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	on Error Resume Next

    Dim ADF														'ActiveX Data Factory ���� �������� 
    Dim strRetMsg												'Record Set Return Message �������� 
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 
	Dim StrData

	Dim iLoop,jLoop
	Dim isOverFlowKey
	Dim isOverFlowName
	Dim arrStrDT
	Dim iStr
	Dim PvArr,iRsCnt
    Const C_SHEETMAXROWS = 30									'��ȭ�鿡 ���ϼ� �ִ� �ִ� Row �� 

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
    
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
	    
	    arrStrDT  = Split(Request("arrStrDT"),gColSep)    	

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
		   iRsCnt		  = rs0.RecordCount 	
		   Redim PvArr(iRsCnt)
		   For iLoop = 0 to iRsCnt-1
		     If iLoop < C_SHEETMAXROWS Then
			    For jLoop = 0 To intDataCount
			        Select Case arrStrDT(jLoop)
			           Case "DD"  :    strData = strData & Chr(11) & UNIDateClientFormat(rs0(jLoop))
			           Case "F2"  :    strData = strData & Chr(11) & UNINumClientFormat(rs0(jLoop), ggAmtOfMoney.DecPoint, 0)
			           Case "F3"  :    strData = strData & Chr(11) & UNINumClientFormat(rs0(jLoop), ggQty.DecPoint       , 0)
			           Case "F4"  :    strData = strData & Chr(11) & UNINumClientFormat(rs0(jLoop), ggUnitCost.DecPoint  , 0)
			           Case "F5"  :    strData = strData & Chr(11) & UNINumClientFormat(rs0(jLoop), ggExchRate.DecPoint  , 0)
			           Case Else  :    strData = strData & Chr(11) & rs0(jLoop)                    
			        End Select    
		        Next    
				strData = strData & Chr(11) & jLoop + 1
				strData = strData & Chr(11) & Chr(12)
		
		     Else
			    isOverFlowKey  = rs0(0)
				isOverFlowName = rs0(1)
				Exit For
			End If
			PvArr(iLoop) = strData
			strData=""
		    rs0.MoveNext
		   Next
		End If   
		rs0.Close
		strData = Join(PvArr, "")
		
		Set rs0 = Nothing
		Set ADF = Nothing
	End If  


	'ǰ������� FETCH -- Corrected by Min
	If  Request("txtJnlItem") <> "" or Request("txtJnlItem") <> Null then 
		Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
		 
		lgStrSQL = "Select MINOR_CD,MINOR_NM " 
		lgStrSQL = lgStrSQL & " From  B_MINOR "
		lgStrSQL = lgStrSQL & " WHERE MAJOR_CD=" & FilterVar("P1001", "''", "S") & " "
		lgStrSQL = lgStrSQL & " AND   MINOR_CD =  " & FilterVar(Request("txtJnlItem"), "''", "S") & " "

		call FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") 
	end if


%>		

<Script Language="vbscript">   
  'On Error Resume Next
	With parent
	    .ggoSpread.Source    = .vspdData 
	    .ggoSpread.SSShowData  "<%=ConvSPChars(strData)%>"
        .lgStrCodeKey        = "<%=ConvSPChars(isOverFlowKey)%>"
        .lgStrNameKey        = "<%=ConvSPChars(isOverFlowName)%>"

        '�߰� 
     	.txtJnlItemNm.value = "<%=ConvSPChars(lgObjRs(1))%>"
	    .vspdData.focus		
        .DbQueryOk()
	End With

</Script>

<%
'�߰� 
If  Request("txtJnlItem") <> "" or Request("txtJnlItem") <> Null then 
	Call SubCloseRs(lgObjRs)                                                    '�� : Release RecordSSet
end if
%>
