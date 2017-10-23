<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")%>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1213MB2
'*  4. Program Name         : ������ (������) �˻����� ��� 
'*  5. Program Desc         : 
'*  6. Component List       : PQBG150
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd

Dim PQBG150																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim lgIntFlgMode
Dim LngMaxRow
Dim arrRowVal								'��: Spread Sheet �� ���� ���� Array ���� 
Dim arrColVal								'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus								'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim lGrpCnt								'��: Group Count
Dim strUserId
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strInspMthdCd
Dim strInspItemCd
Dim txtSpread
Dim iErrorPosition
	
LngMaxRow = CInt(Request("txtMaxRows"))					'��: �ִ� ������Ʈ�� ���� 
lgIntFlgMode = CInt(Request("txtFlgMode"))					'��: ����� Create/Update �Ǻ� 
	
strInspItemCd = UCase(Request("txtInspItemCd"))
strPlantCd = UCase(Request("txtPlantCd"))
strItemCd = UCase(Request("txtItemCd"))
strInspClassCd = UCase(Request("cboInspClassCd"))
strInspMthdCd = UCase(Request("txtInspMthdCd"))
	
txtSpread = Request("txtSpread") 
	
Set PQBG150 = Server.CreateObject("PQBG150.cQMtInspStdDtl1Svr")
 
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If	 
	
Call PQBG150.Q_MAINT_INSP_STD_DTL1_SVR(gStrGlobalCollection, strInspItemCd, strPlantCd, strItemCd, strInspClassCd, strInspMthdCd, txtSpread, iErrorPosition)
	 
If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then
	If iErrorPosition <> "" Then	
		Call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
		Set PQBG150 = Nothing
		Response.End
	End If
End If	

Set PQBG150 = Nothing 
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