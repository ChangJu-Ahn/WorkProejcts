<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1111MB1
'*  4. Program Name         : ������ ������� 
'*  5. Program Desc         : ������ ������� 
'*  6. Component List       : PQBG010
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/09
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
												'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd 
		
Dim PQBG010																	'�� : ��ȸ�� ComProxy Dll ��� ���� 
Dim txtSpread
Dim iErrorPosition		

	txtSpread	= Request("txtSpread")
			
	Set PQBG010 = Server.CreateObject("PQBG010.cQMaintMeaEquSvr")
		
	Call PQBG010.Q_MAINT_MEA_EQU_SVR (gstrGlobalCollection, _
									  txtSpread, _
									  iErrorPosition)
			
	If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then
	   Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
	   Set PQBG010 = Nothing 
	   Response.End
	End if
		
	Set PQBG010 = Nothing 
%>
<Script Language=vbscript>
	With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
		.DbSaveOk
	End With
</Script>
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