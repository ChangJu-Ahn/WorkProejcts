<% 
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(카렌더 생성)
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
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd

Dim PB4G021											'☆ : 입력/수정용 ComProxy Dll 사용 변수 

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
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
%>
<Script Language=vbscript>
With parent																		'☜: 화면 처리 ASP 를 지칭함 
	.DbSaveOk
End With
</Script>