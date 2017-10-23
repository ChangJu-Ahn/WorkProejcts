<%@LANGUAGE = VBScript%>
<%'********************************************************************************************
'*  1. Module Name			: DT
'*  2. Function Name		: 
'*  3. Program ID			: d1211PB1.asp
'*  4. Program Name			: Digital Tax (Query)
'*  5. Program Desc			:
'*  6. Comproxy List		: DB Agent
'*  7. Modified date(First)	: 2009/12/20
'*  8. Modified date(Last)	: 2009/12/22
'*  9. Modifier (First)		: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'********************************************************************************************
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "M","NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs1										'DBAgent Parameter 선언 
Dim strMode									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim i
Dim strData

Dim iTotalStr
Dim TmpBuffer1


    Call HideStatusWnd

	Redim UNISqlId(0)
	Redim UNIValue(0, 5)
	
	UNISqlId(0) = "D1511QA12"
	
    UNIValue(0, 0) = FilterVar(Request("txtTaxBillNo"), "''", "S")

	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
	
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	End If
		
		If Not(rs1.EOF And rs1.BOF) Then
		
			Redim TmpBuffer1(rs1.RecordCount-1)
			
			For i=0 to rs1.RecordCount-1
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs1("INV_NO"))
				strData = strData & Chr(11) & ConvSPChars(rs1("INV_ITEM_SEQ_NO"))
				strData = strData & Chr(11) & ConvSPChars(rs1("ITEM"))
				strData = strData & Chr(11) & ConvSPChars(rs1("ITEM_STD"))
				strData = strData & Chr(11) & UniNumClientFormat(rs1("ITEM_PRC") ,ggAmtOfMoney.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs1("ITEM_QTY") ,ggAmtOfMoney.DecPoint,0)
				strData = strData & Chr(11) & UNIDateClientFormat(rs1("ITEM_DATE"))
				strData = strData & Chr(11) & UniNumClientFormat(rs1("ITEM_AMT") ,ggAmtOfMoney.DecPoint,0)
				strData = strData & Chr(11) & UniNumClientFormat(rs1("ITEM_TAX") ,ggAmtOfMoney.DecPoint,0) 
				strData = strData & Chr(11) & ConvSPChars(rs1("ITEM_MEMO"))
				
				strData = strData & Chr(11) & i
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer1(i) = strData
				rs1.MoveNext
				
			Next
			
		iTotalStr = Join(TmpBuffer1,"") 

		End If
		

		rs1.close

		Set rs1 = Nothing

        Set ADF = Nothing
        
%>	
		
    
<Script Language=vbscript>

    
    With parent												'☜: 화면 처리 ASP 를 지칭함 
		
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip  "<%=iTotalStr%>"
		
    End With
    
</Script>	
