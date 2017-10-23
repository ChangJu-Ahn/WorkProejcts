<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow

On Error Resume Next

Call LoadBasisGlobalInf()

Call HideStatusWnd

strMode      = Request("txtMode")												'☜ : 현재 상태를 받음 

Dim PB0C008
    
Dim strLogCntList
        
Err.Clear 
On Error Resume Next		
			
Set PB0C008 = Server.CreateObject("PB0C008.CB0C008")
		        
strLogCntList = PB0C008.Z_GET_CHECK_LONGIN_USER_LIST(gStrGlobalCollection)

If Err.Number <> 0 Then
	Set PB0C008 = Nothing												'☜: ComProxy Unload
	Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'⊙:
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
		    
Set PB0C008 = Nothing

%>		    
<Script Language="vbscript">  
With parent		
    .ggoSpread.Source = parent.vspdData
	.ggoSpread.SSShowData "<%=Trim(strLogCntList)%>"
	.vspdData.focus

	If .vspdData.MaxRows = 0 Then
		parent.UNIMsgBox "There's no data.", 48, parent.top.document.title
	End If

End With

</Script>
