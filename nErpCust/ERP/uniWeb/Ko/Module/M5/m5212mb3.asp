<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%
	call LoadBasisGlobalInf()
'********************************************************************************************************
'*  1. Module Name          : Procuremant                *
'*  2. Function Name        :                   *
'*  3. Program ID           : m5212mb3.asp                *
'*  4. Program Name         :                   *
'*  5. Program Desc         : ���� B/L���� ȸ��ó�� Transaction ó���� ASP       *
'*  7. Modified date(First) : 2000/04/21                *
'*  8. Modified date(Last)  : 2000/04/21                *
'*  9. Modifier (First)     :                   *
'* 10. Modifier (Last)      :                   *
'* 11. Comment              :                   *
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"         *
'*                            this mark(��) Means that "may  change"         *
'*                            this mark(��) Means that "must change"         *
'* 13. History              : 1. 2000/04/21 : Coding Start            *
'********************************************************************************************************

Response.Expires = -1               '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True               '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'   ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
                    '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call HideStatusWnd

Dim strMode                  '��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

strMode = Request("txtMode")             '�� : ���� ���¸� ���� 

Select Case strMode
 Case CStr(5211)                '��: ���� Save ��û�� ���� 

  Dim M52115                ' ���� B/L Header Save�� Object

  Err.Clear                '��: Protect system from crashing
   
  Set M52115 = Server.CreateObject("M52115.M52115PostBlInOpenApSvr")

  '-----------------------
  'Com action result check area(OS,internal)
  '-----------------------
  If Err.Number <> 0 Then
   Set M52115 = Nothing            '��: ComProxy UnLoad
   Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
   Response.End              '��: Process End
  End If

  '-----------------------
  'Data manipulate  area(import view match)
  '-----------------------
  M52115.ImportMBlHdrBlNo = UCase(Trim(Request("txtBLNo")))
  M52115.CommandSent = "POST"
  M52115.ServerLocation = ggServerIP

  '-----------------------
  'Com action area
  '-----------------------
  M52115.ComCfg = gConnectionString
  M52115.Execute

  '-----------------------
  'Com action result check area(OS,internal)
  '-----------------------
  If Err.Number <> 0 Then
   Set M52115 = Nothing            '��: ComProxy UnLoad
   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)
   Response.End              '��: Process End
  End If

  '-----------------------
  'Com action result check area(DB,internal)
  '-----------------------
  If Not (M52115.OperationStatusMessage = MSG_OK_STR) Then
   Select Case M52115.OperationStatusMessage
    Case MSG_DEADLOCK_STR
     Call DisplayMsgBox2("999999", "25", "deadlock or timeout" , I_MKSCRIPT)
    Case MSG_DBERROR_STR
     Call DisplayMsgBox2(M52115.ExportErrEabSqlCodeSqlcode, _
           M52115.ExportErrEabSqlCodeSeverity, _
           M52115.ExportErrEabSqlCodeErrorMsg, I_MKSCRIPT)
    Case Else
     Call DisplayMsgBox(M52115.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
   End Select
  
   Set M52115 = Nothing
   Response.End 
  End If
   
%>
<Script Language=VBScript>
 With parent.frm1
  Dim strDt
  
  'If "<%=Request("txtPost")%>" = "Y" Then
  ' parent.dbSaveOK
  'Else
   parent.dbSaveOK
  'End If
 End With
</Script>
<%
  '-----------------------
  'Result data display area
  '-----------------------
  Set M52115 = Nothing              '��: Unload Comproxy
  Response.End                '��: Process End
  
 Case Else
  Response.End
End Select
%>
