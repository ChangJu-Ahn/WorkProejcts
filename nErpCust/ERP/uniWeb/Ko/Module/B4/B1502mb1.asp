<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Common Holiday)
'*  3. Program ID           : B1502mb1
'*  4. Program Name         : �������ϵ�� 
'*  5. Program Desc         :
'*  6. Comproxy List        : +B15021ControlCommonHoliday
'                             +B15028ListCommonHoliday
'*  7. Modified date(First) : 2000/09/14
'*  8. Modified date(Last)  : 2002/12/13
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													                       '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
                           
Call HideStatusWnd													      	'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Dim PB4G011												                  '  ��  ComProxy Dll ��� ���� 
Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strSpread
Dim lgstrdata
Dim iErrPosition

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strSpread = Trim(Request("txtSpread"))

Select Case strMode
Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

    on error resume next	
    
    Set PB4G011 = Server.CreateObject("PB4G011.cBListCommonHoliday")    
    If CheckSYSTEMError(Err,True) = True Then
        set PB4G011 = nothing
        Response.End  
    End If	
	on error goto 0
	
    on error resume next
    lgstrdata = PB4G011.B_LIST_COMMON_HOLIDAY(gStrGlobalCollection)
    If CheckSYSTEMError(Err,True) = True Then
        set PB4G011 = nothing
        Response.End  
    End If	
	on error goto 0
%>
<Script Language=vbscript>    
	
	With parent			
	
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"	
	.DbQueryOk
	
	End With
</Script>	
<%    
    Set PB4G011 = Nothing    
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 

    on error resume next
    Set PB4G011 = Server.CreateObject("PB4G011.cBCrlCommonHoliday")    
    If CheckSYSTEMError(Err,True) = True Then
        set PB4G011 = nothing
        Response.End  
    End If	
	on error goto 0
    
    on error resume next
    Call PB4G011.B_CONTROL_COMMON_HOLIDAY(gStrGlobalCollection,strSpread)    
 	If CheckSYSTEMError(Err,True) = True Then
        set PB4G011 = nothing
        Response.End  
    End If	
	on error goto 0

    Set PB4G011 = Nothing                                                   '��: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
		.DbSaveOk
	End With
</Script>
<%					
End Select
%>