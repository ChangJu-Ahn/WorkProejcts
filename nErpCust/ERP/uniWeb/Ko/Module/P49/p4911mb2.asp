<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4911mb2.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005-01-25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd

Dim pPP4G901

Dim lgIntFlgMode
Dim LngMaxRow

Dim arrRowVal								'��: Spread Sheet �� ���� ���� Array ���� 
Dim arrColVal								'��: Spread Sheet �� ���� ���� Array ���� 

Dim iErrorPosition
Dim txtSpread
Dim LngRow
Dim strPlantCd

strPlantCd = UCase(Trim(Request("txtPlantCd")))
txtSpread = Request("txtSpread")		' Create, Update

'-------------------------------------------------------------------------------
'	COM+ Action
'-------------------------------------------------------------------------------

Set pPP4G901 = Server.CreateObject("PP4G901.cPMngStdrdPrdtTime")

If CheckSYSTEMError(Err,True) = True Then
   Response.End
End If

Call pPP4G901.P_MANAGE_STANDARD_PRODUCT_TIME(gStrGlobalCollection, strPlantCd, txtSpread, iErrorPosition)

If CheckSYSTEMError2(Err,True, "","","","","") = True Then
	If iErrorPosition <> "" Then
		Call SheetFocus(iErrorPosition, 1 ,I_MKSCRIPT)
		Set pPP4G901 = Nothing
		Response.End
	End If
End If

Set pPP4G901 = Nothing

%>
<Script Language=vbscript>
With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
	.DbSaveOk
End With
</Script>
<%
' Server Side ������ ���⼭ ���� 
'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
%>
<Script Language=vbscript RUNAT=server>
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
</Script>