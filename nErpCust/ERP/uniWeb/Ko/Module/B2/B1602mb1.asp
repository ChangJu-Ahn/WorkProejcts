<%
'**********************************************************************************************
'*  1. Module Name          : Sale,Production
'*  2. Function Name        : Sales Order,....
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B16021ControlProvince
'                             +B16028ListProvince
'                             +B16019LookupCountry
'                             +B16029LookupProvince
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2002/12/13
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Dim PB2G141												'  ☆  ComProxy Dll 사용 변수 
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
	
	importtxtCountry = request("txtCountry")
	importtxtProvince = request("txtProvince")
    
%>
<Script Language=vbscript>    
	With parent			
        .DbLookUp
	End With
</Script>	
<%      

    Set PB2G141 = Server.CreateObject("PB2G141.cBListProvince")    
    If CheckSYSTEMError(Err,True) = True Then
        set PB2G141 = nothing
        Response.End  
    End If	
	on error goto 0
	
    on error resume next
    lgstrdata = PB2G141.B_READ_PROVINCE(gStrGlobalCollection,importtxtCountry,importtxtProvince)
    If CheckSYSTEMError(Err,True) = True Then
        set PB2G141 = nothing
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
    Set PB2G141 = Nothing
    
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 

    on error resume next
    Set PB2G141 = Server.CreateObject("PB2G141.cBControlProvince")    
    If CheckSYSTEMError(Err,True) = True Then
        set PB2G141 = nothing
        Response.End  
    End If	
	on error goto 0
    
    on error resume next
    Call PB2G141.B_CONTROL_PROVINCE(gStrGlobalCollection,strSpread)    
 	If CheckSYSTEMError(Err,True) = True Then
        set PB2G141 = nothing
        Response.End  
    End If	
	on error goto 0

    Set PB2G141 = Nothing                                                   '☜: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		'window.status = "저장 성공"
		.DbSaveOk
	End With
</Script>
<%					
End Select
%>

