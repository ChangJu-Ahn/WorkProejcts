<% 
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(ī���� ����)
'*  3. Program ID           : B1503mb1.asp
'*  4. Program Name         : B1503mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B15011ControlCalendar
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2002/12/16
'*  9. Modifier (First)     : Hwnag, Jeong-won
'* 10. Modifier (Last)      : Sim Hae Yong
'* 11. Comment              :
'**********************************************************************************************

%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd

Dim PB4G021											'�� : �Է�/������ ComProxy Dll ��� ���� 

Dim I1_year_month  
Dim Import_Array  ''Import Array
''Import		           
Const B345_I2_SUN = 0
Const B345_I2_MON = 1
Const B345_I2_TUE = 2
Const B345_I2_WED = 3
Const B345_I2_THU = 4
Const B345_I2_FRI = 5
Const B345_I2_SAT = 6

Call LoadBasisGlobalInf()
        
REDIM Import_Array(B345_I2_SAT)        
'-----------------------        
I1_year_month = Request("txtYear")
Import_Array(B345_I2_SUN) = Request("chkSun")
Import_Array(B345_I2_MON) = Request("chkMon")
Import_Array(B345_I2_TUE) = Request("chkTue")
Import_Array(B345_I2_WED) = Request("chkWed")
Import_Array(B345_I2_THU) = Request("chkThu")
Import_Array(B345_I2_FRI) = Request("chkFri")
Import_Array(B345_I2_SAT) = Request("chkSat")
	
'''''''''''''''''''''''''''
Set PB4G021 = Server.CreateObject("PB4G021.cBControlCalendar")	
On Error Resume Next    
	    
Err.Clear 
CALL PB4G021.B_CREATE_CALENDAR(gStrGlobalCollection,I1_year_month,Import_Array)
Set PB4G021 = Nothing
		
If CheckSYSTEMError(Err,True) = True Then                               
	Response.End														'��: �����Ͻ� ���� ó���� ������ 
End If
%>
<Script Language=vbscript>
With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
	.DbSaveOk
End With
</Script>