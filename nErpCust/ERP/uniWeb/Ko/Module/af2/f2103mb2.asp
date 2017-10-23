<%
'**********************************************************************************************
'*  1. Module Name          : 자금 
'*  2. Function Name        : 추가예산 
'*  3. Program ID           : a2103mb2(통제기간단위 lookup)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +FLookupBdgAcctSvr
'*  7. Modified date(First) : 2000/9/07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : You. So. Eun.
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncServer.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim pFU0019
Dim strAdd
Dim pB1a028
Call HideStatusWnd

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

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
		Set pFU0019 = Nothing																'☜: ComProxy UnLoad
		Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:
		Response.End																		'☜: Process End
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
	   Call ServerMesgBox(Err.description , vbInformation, I_MKSCRIPT)                          '⊙:
	   Set pFU0019 = Nothing																	    '☜: ComProxy UnLoad
	   Response.End																				'☜: Process End
	End If
    
	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If Not (pFU0019.OperationStatusMessage = MSG_OK_STR) Then
	   Call DisplayMsgBox(pFU0019.OperationStatusMessage, vbInformation, "", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	   Set pFU0019 = Nothing																	'☜: ComProxy UnLoad
	   Response.End																				'☜: Process End
	End If

    
%>
<Script Language=vbscript>
	
        parent.frm1.txtCtrl_Unit.value  = "<%=ConvSPChars(pFU0019.ExportCtrlUnitBMinorMinorNm)%>"
        strAdd = "<%=ConvSPChars(pFU0019.ExportFBdgAcctAddFg)%>"
        
        if strAdd = "1" Then
			parent.frm1.txtadd.value = "추가가능"
        End If
</Script>
<%
    Set pFU0019 = Nothing															    '☜: Unload Comproxy

	Response.End																		'☜: Process End   
End Select
%>

	
    
    

