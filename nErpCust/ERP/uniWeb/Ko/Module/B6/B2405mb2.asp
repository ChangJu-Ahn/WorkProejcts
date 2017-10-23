<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(내부부서코드생성)
'*  3. Program ID           : B2405mb2.asp
'*  4. Program Name         : B2405mb2.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B24052MakeInternalCd
'*  7. Modified date(First) : 2000/10/30
'*  8. Modified date(Last)  : 2002/12/03
'*  9. Modifier (First)     : Hwnag Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													                       '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd
On Error Resume Next														'☜: 
Err.Clear   

Dim PB6G062																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strCase

Call LoadBasisGlobalInf()

strCase = Request("txtOrgId")

If strCase <> "" Then

    Set PB6G062 = Server.CreateObject("PB6G062.bBMakeInternalCd")
    call PB6G062.B_MAKE_INTERNAL_CD(gStrGlobalCollection,strCase)
    Set PB6G062 = Nothing

	If CheckSYSTEMError(Err,True) = True Then
	    ''Response.End 
	End If	
    On error goto 0
%>
<Script Language=vbscript>	
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		'window.status = "생성 성공"
		.Batch_OK
	End With
</Script>
<%
End If 
    
%>
