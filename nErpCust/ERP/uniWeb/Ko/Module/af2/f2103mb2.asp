<%
'**********************************************************************************************
'*  1. Module Name          : �ڱ� 
'*  2. Function Name        : �߰����� 
'*  3. Program ID           : a2103mb2(�����Ⱓ���� lookup)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +FLookupBdgAcctSvr
'*  7. Modified date(First) : 2000/9/07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : You. So. Eun.
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim pFU0019
Dim strAdd
Dim pB1a028
Call HideStatusWnd

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case strMode
Case "UnitFg"
'********************************************************  
'              LOOKUP For Business Partner's name
'********************************************************  

	Err.Clear                                                  
   
	set pFU0019 = Server.CreateObject("FU0019.FLookupBdgAcctSvr")  	
     
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set pFU0019 = Nothing																'��: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'��:
		Response.End																		'��: Process End
	End If
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------

	pFU0019.ImportFBdgAcctBdgCd   = Trim(Request("Unit"))
        
    pFU0019.ServerLocation = ggServerIP
    pFU0019.CommandSent    = "LOOKUP"
    
    '-----------------------
    'Com action area
    '-----------------------       
    pFU0019.ComCfg = gConnectionString
    pFU0019.Execute 
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If Err.Number <> 0 Then
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                          '��:
	   Set pFU0019 = Nothing																	    '��: ComProxy UnLoad
	   Response.End																				'��: Process End
	End If
    
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pFU0019.OperationStatusMessage = MSG_OK_STR) Then
	   Call DisplayMsgBox(pFU0019.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	   Set pFU0019 = Nothing																	'��: ComProxy UnLoad
	   Response.End																				'��: Process End
	End If

    
%>
<Script Language=vbscript>
	
        parent.frm1.txtCtrl_Unit.value  = "<%=ConvSPChars(pFU0019.ExportCtrlUnitBMinorMinorNm)%>"
        strAdd = "<%=ConvSPChars(pFU0019.ExportFBdgAcctAddFg)%>"
        
        if strAdd = "1" Then
			parent.frm1.txtadd.value = "�߰�����"
        End If
</Script>
<%
    Set pFU0019 = Nothing															    '��: Unload Comproxy

	Response.End																		'��: Process End   
End Select
%>

	
    
    

