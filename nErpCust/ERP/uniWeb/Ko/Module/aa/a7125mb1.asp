<%
'======================================================================================================
'*  1. Module Name          : accounting
'*  2. Function Name        : 
'*  3. Program ID           : a7125ma1
'*  4. Program Name         : 고정자산취득상세내역등록 
'*  5. Program Desc         : 고정자산취득상세내역등록 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2000/08/23
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 

Dim pAS0101											'입력/수정용 ComProxy Dll 사용 변수 

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          
Dim iPAAG011


Call HideStatusWnd

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")


strMode = Request("txtMode")						'☜ : 현재 상태를 받음 

GetGlobalVar

On Error Resume Next

Select Case strMode

    
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
									
    Err.Clear																		'☜: Protect system from crashing

    LngMaxRow = CInt(Request("txtMaxRows_2"))											'☜: 최대 업데이트된 갯수 

    Set iPAAG011 = Server.CreateObject("PAAG011.cAMngAsItmSvr")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		set iPAAG011 = Nothing
		Response.End
	End IF	
	
	arrTemp = Split(Request("txtSpread_m"), gRowSep)									'Spread Sheet 내용을 담고 있는 Element명 

    call iPAAG011.A_MANAGE_ASSET_ITEM_SVR( gStrGloBalCollection , Request("txtSpread_m"))
            
        '-----------------------
        'Com action result check area(OS,internal)
        '-----------------------                   
    If CheckSYSTEMError(Err,True) = True Then
		set pAS0101 = Nothing
		Response.End			
	End IF
                                              

    Set iPAAG011 = Nothing														    '☜: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%					

End Select

%>
