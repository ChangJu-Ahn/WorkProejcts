<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<%
	Call LoadBasisGlobalInf() 
	Call HideStatusWnd

	On Error Resume Next
    Err.Clear

	Call SubBizSaveMulti()

Sub SubBizSaveMulti()
    On Error Resume Next
    Err.Clear

    Dim iPD1G040                 '�� : ��ȸ�� ComProxy Dll ��� ���� 
    Dim lgIntFlgMode
    Dim iErrorPosition

	IF Request("txtMode") <> "" Then
		lgIntFlgMode = CInt(Request("txtMode"))         '��: ����� Create/Update �Ǻ� 
	END IF


	If Request("txtSpread") <> "" Then

		Set iPD1G040 = Server.CreateObject("PD1G040.cAMngAcctClssSvr")
		If CheckSYSTEMError(Err, True) = True Then
			Set iPD1G040 = Nothing
			Exit Sub
		End If

		Call iPD1G040.A_MANAGE_ACCT_CLASS_SVR(gStrGlobalCollection, Request("txtClassType"), Request("txtSpread"),iErrorPosition)

		If CheckSYSTEMError2(Err, True,iErrorPosition & "��","","","","") = True Then
		Set iPD1G040 = Nothing
		Exit Sub
		End If

		Set iPD1G040 = Nothing

	End If
End Sub

%>
<Script Language=vbscript>
 parent.DbSaveOk
</Script>
