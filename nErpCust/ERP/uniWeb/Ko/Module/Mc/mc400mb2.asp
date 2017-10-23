<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc400mb2
'*  4. Program Name         : ��������Ȯ����� 
'*  5. Program Desc         : ��������Ȯ����� 
'*  6. Component List       : PMCG350.cMMaintDeliveryOrder
'*  7. Modified date(First) : 2003/02/27
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf


Dim oPMCG350																	'�� : ����� ComProxy Dll ��� ���� 

Dim StrNextKey											'��: ���� �� 
Dim lgStrPrevKey										'��: ���� �� 
Dim strMode												'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim LngMaxRow											'��: ���� �׸����� �ִ�Row
Dim intFlgMode
Dim LngRow
Dim LngIdx

Dim I1_M_DLVY_ORD_HDR
Dim IG1_M_DLVY_ORD_KEY

Dim arrCols, arrRows									'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus											'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt												'��: Group Count
Dim strCode												'��: Lookup �� ���� ���� 

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim ii

'Export Group
Const C_IG1_Prodt_Order_No = 0
Const C_IG1_Opr_No = 1
Const C_IG1_Seq = 2
Const C_IG1_Sub_Seq = 3

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 
	intFlgMode = CInt(Request("txtFlgMode"))	
	LngMaxRow = CInt(Request("txtMaxRows"))									'��: �ִ� ������Ʈ�� ���� 
    lgStrPrevKey = Trim(Request("lgStrPrevKey"))
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   

	arrRows = Split(itxtSpread, gRowSep)							'��: Spread Sheet ������ ��� �ִ� Element�� 
    ReDim IG1_M_DLVY_ORD_KEY(UBound(arrRows,1),3)

	For LngRow = 0 To LngMaxRow - 1
		arrCols = Split(arrRows(LngRow), gColSep)

		IG1_M_DLVY_ORD_KEY(LngRow, C_IG1_Prodt_Order_No)= UCase(Trim(arrCols(2)))
		IG1_M_DLVY_ORD_KEY(LngRow, C_IG1_Opr_No)		= UCase(Trim(arrCols(3)))
		IG1_M_DLVY_ORD_KEY(LngRow, C_IG1_Seq)			= UniConvNum(arrCols(4), 0)
		IG1_M_DLVY_ORD_KEY(LngRow, C_IG1_Sub_Seq)		= UniConvNum(arrCols(5), 0)
	Next
	
	Set oPMCG350 = Server.CreateObject("PMCG350.cMMaintDeliveryOrder")

	If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

	Call oPMCG350.M_CANCEL_DELIVERY_ORDER(gStrGlobalCollection, _
										IG1_M_DLVY_ORD_KEY)

	If CheckSYSTEMError(Err,True) = True Then
       Set oPMCG350 = Nothing		                                        '��: Unload Comproxy DLL
       Response.End		
    End If

	Set oPMCG350 = Nothing                                                  '��: Unload Comproxy

	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr					
	
	'==============================================================================
	' Function : SheetFocus
	' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
	'==============================================================================
	Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
		Dim strHTML
		If iLoc = I_INSCRIPT Then
			strHTML = "parent.frm1.vspdData.focus" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
			Response.Write strHTML
		ElseIf iLoc = I_MKSCRIPT Then
			strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
			strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
			strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
			Response.Write strHTML
		End If
	End Function

%>
