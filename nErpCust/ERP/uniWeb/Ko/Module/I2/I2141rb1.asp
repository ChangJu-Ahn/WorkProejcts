<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Goods mvmt List After Phy Inv Conting
'*  2. Function Name        : 
'*  3. Program ID           : I2141rb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 실사선별후 수불참조 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2006/08/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : LEE SEUNG Wook
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%             
on Error Resume Next

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd 

Err.Clear

Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

Dim i

	Const C_SHEETMAXROWS_D = 100

Dim strDocumentDt,strDocumentNo,strSeqNo,strSubSeqNo,strItemCd
Dim strkeyval

	Redim UNISqlId(0)
	Redim UNIValue(0, 5)
	
	UNISqlId(0) = "i2141rb1"
	
	If Request("lgStrPrevKey1") <> "" And Request("lgStrPrevKey2") <> "" And _
		 Request("lgStrPrevKey3") <> "" And Request("lgStrPrevKey4") <> "" Then
		
		strDocumentDt = FilterVar(Request("lgStrPrevKey1"),"''","S")
		strDocumentNo = FilterVar(Request("lgStrPrevKey2"),"''","S")
		strSeqNo = FilterVar(Request("lgStrPrevKey3"),"''","S")
		strSubSeqNo = FilterVar(Request("lgStrPrevKey4"),"''","S")
		
		strkeyval = " ( A.DOCUMENT_DT > "& strDocumentDt & _
					" Or ( A.DOCUMENT_DT = "& strDocumentDt &" AND B.ITEM_DOCUMENT_NO > "& strdocumentNo &")" & _
					" Or ( A.DOCUMENT_DT = "& strDocumentDt &" AND B.ITEM_DOCUMENT_NO = "& strdocumentNo &" and B.SEQ_NO > "& strSeqNo &" )" & _
					" Or ( A.DOCUMENT_DT = "& strDocumentDt &" AND B.ITEM_DOCUMENT_NO = "& strdocumentNo &" and B.SEQ_NO = "& strSeqNo &" and B.SUB_SEQ_NO >= "& strSubSeqNo &")) "
		
	Else
		strkeyval = "|"	
	End If
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar((Trim(Request("txtPhyInvNo"))),"''","S")
	UNIValue(0, 3) = FilterVar((Trim(Request("txtInspDt"))), "''", "S")
	UNIValue(0, 4) = FilterVar((Trim(Request("txtItemCd"))), "''", "S")
	UNIValue(0, 5) = strkeyval
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent
	LngMaxRow = .vspdData.MaxRows
	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
			ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
		Else
%>			
			ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
		End If
			
		For i=0 to rs0.RecordCount-1
			If i < C_SHEETMAXROWS_D Then 
%>
				strData = "" _
				& Chr(11) & "<%=ConvSPChars(rs0("ITEM_DOCUMENT_NO"))%>" _
				& Chr(11) & "<%=UNIDateClientFormat(rs0("DOCUMENT_DT"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("TRNS_NM"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("MOV_NM"))%>"
				
				Select Case "<%=ConvSPChars(rs0("DEBIT_CREDIT_FLAG"))%>"
					Case "D"
						strData = strData & Chr(11) & "증가"
					Case "C"
						strData = strData & Chr(11) & "감소"	
				End Select
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SL_CD"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("LOT_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("LOT_SUB_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SEQ_NO"))%>" _
				& Chr(11) & "<%=ConvSPChars(rs0("SUB_SEQ_NO"))%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("AMOUNT"),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit,0)%>" _
				& Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PRICE"),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)%>" _
				& Chr(11) & LngMaxRow + <%=i%> _
				& Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		.lgStrPrevKey1 = "<%=Trim(rs0("DOCUMENT_DT"))%>"
		.lgStrPrevKey2 = "<%=Trim(rs0("ITEM_DOCUMENT_NO"))%>"
		.lgStrPrevKey3 = "<%=Trim(rs0("SEQ_NO"))%>"
		.lgStrPrevKey4 = "<%=Trim(rs0("SUB_SEQ_NO"))%>"
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>
	On Error Resume Next
	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey1 <> "" And _
		.lgStrPrevKey2 <> "" And .lgStrPrevKey3 <> "" And .lgStrPrevKey4 <> "" Then
		.DbQuery
	Else
		.DbQueryOk	
	End If
End With

</Script>	
<%
Set ADF = Nothing
%>

 
