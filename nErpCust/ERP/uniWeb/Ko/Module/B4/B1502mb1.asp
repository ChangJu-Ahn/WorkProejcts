<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Common Holiday)
'*  3. Program ID           : B1502mb1
'*  4. Program Name         : 공통휴일등록 
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
<%													                       '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
                           
Call HideStatusWnd													      	'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Dim PB4G011												                  '  ☆  ComProxy Dll 사용 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim lgstrdata
Dim iErrPosition

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

Select Case strMode
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

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
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

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

    Set PB4G011 = Nothing                                                   '☜: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%					
End Select
%>