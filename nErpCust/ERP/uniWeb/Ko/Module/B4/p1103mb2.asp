<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1103mb2.asp
'*  4. Program Name         : Mfg Calendar Type Save
'*  5. Program Desc         :
'*  6. Component List       : +PB4G102.P_MANAGE_MFG_CALENDAR_TYPE.P_MANAGE_MFG_CALENDAR_TYPE
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2000/04/17
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'��: 

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Call LoadBasisGlobalInf() 

Dim pPB4G102																	'�� : ����� ComProxy Dll ��� ���� 
Dim strSpread

 	Err.Clear																'��: Protect system from crashing
 	
 	strSpread = Request("txtSpread")

    Set pPB4G102 = Server.CreateObject("PB4G102.cPMngMfgCalenType")    

	'-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
															
	call pPB4G102.P_MANAGE_MFG_CALENDAR_TYPE (gStrGlobalCollection, strSpread)
				
	If CheckSYSTEMError(Err,True) = True Then
		Set pPB4G102 = Nothing												'��: ComProxy Unload
		Response.End
	End If
	
	Set pPB4G102 = Nothing												'��: ComProxy Unload

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