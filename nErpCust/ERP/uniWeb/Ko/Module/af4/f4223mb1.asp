<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")  
	Call HideStatusWnd

    Select Case Request("txtMode")
        Case CStr(UID_M0001)                                                         '��: Query
          '  Call SubBizQuery()
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
          '  Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMulti()

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Dim PAFG415CUD
    Dim I2_f_ln_info
    Dim importArray
    
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbOKOnly, "", "", I_MKSCRIPT)		'txtMaxRows ���ǰ��� ����ֽ��ϴ�!
		Response.End
	End If

    Set PAFG415CUD = Server.CreateObject("PAFG415.cFMngLnPlnAnRsltSvr")

    If CheckSYSTEMError(Err, True) = True Then
       Exit Sub
    End If
    
    importArray = Split(Trim(Request("txtSpread")), gRowSep)

    Call PAFG415CUD.F_MANAGE_LN_PLAN_AND_RESULT_SVR(gStrGloBalCollection, importArray, I2_f_ln_info)		

    If CheckSYSTEMError(Err, True) = True Then
       Set PAFG415CUD = Nothing
       Exit Sub
    End If    

    Set PAFG415CUD = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
	
End Sub
%>
