<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_Confirm_TempGl,A_Unconfirm_TempGl
'*  3. Program ID        : a5103mb2
'*  4. Program �̸�      : ������ ������ǥ �ϰ�����,������� 
'*  5. Program ����      : ������ ������ǥ �ϰ�����,������Ҹ� ���� Logic
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2001/02/07
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : hersheys
'*  9. ���� �ۼ���       : 
'* 10. ��ü comment      :
'* 11. ���� Coding Guide : �ּ��� mark(��)�� �Ǿ��ִ� �κ��� ���� ����ڰ� ����(X)
'*                         �ּ��� mark(��)�� �Ǿ��ִ� �κ��� ���� ����ڰ� ����(O)
'* 12. History           : 
'**********************************************************************************************
Response.Expires = -1		'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True		'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next			' ��: 

Dim a53013					' ������ǥ���� ComProxy Dll ��� ���� 
Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strWkfg

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
		
strMode = Request("txtMode")	'�� : ���� ���¸� ���� 
strWkfg = Request("htxtWorkFg")

Select Case strMode
	'-------------------------------------------------------------------------------
	'								    �ϰ�ó�� ���� 
	'-------------------------------------------------------------------------------   
	Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
	    					
	    Err.Clear 												'��: Protect system from crashing

		Set a53013 = Server.CreateObject("A53013.A53012HqConfirmTempGlSvr")

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.Number <> 0 Then
			Set a53013 = Nothing												'��: ComProxy Unload
			Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
			Response.End														'��: �����Ͻ� ���� ó���� ������ 
		End If	

		'-----------------------
		'Data manipulate area
		'-----------------------
		a53013.ImportFromDtATempGlTempGlDt      = UNIConvDate(Request("txtFromTempGlDt"))
		a53013.ImportToDtATempGlTempGlDt        = UNIConvDate(Request("txtToTempGlDt"))
		A53013.ImportBBizAreaBizAreaCd			= Trim(Request("txtBizAreaCd"))
		a53013.ImportATempGlUpdtUserId          = gUsrId
		a53013.ImportBAcctDeptOrgChangeId       = gChangeOrgId
		a53013.ServerLocation                   = ggServerIP

		If UCase(strWkfg) = "CONF"  then		
			a53013.ImportIefSuppliedCommand = "CONF"			
		Elseif UCase(strWkfg) = "UNCONF"  then		
			a53013.ImportIefSuppliedCommand = "UNCONF"
		End if					

		a53013.ComCfg = gConnectionString
'		a53013.ComCfg = "TCP letitbe 2055"
        a53013.Execute

        '-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If Err.number <> 0 Then
			Set a53013 = Nothing
			Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'��:
			Response.End 
		End If

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------				
		If Not (a53013.OperationStatusMessage = MSG_OK_STR) Then
			Call DisplayMsgBox(a53013.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
			Set a53013 = Nothing
			Response.End 
		End If
                   
		a53013.Clear
		lGrpCnt = 0
		
		'ggoSpread.SSDeleteFlag lStartRow, lEndRow	'���ϴ°��� �𸣰���       
	
	Set a53013 = Nothing                                                   '��: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
        .InitSpreadSheet
        .InitComboBox
        .InitVariables 															'��: Initializes local global variables	
		.dbQuery		
	End With
 
</Script>

<%					

End Select

%>

