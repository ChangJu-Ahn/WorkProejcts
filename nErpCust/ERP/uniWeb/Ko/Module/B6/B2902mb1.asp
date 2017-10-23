<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Organization(내부부서코드반영 Table등록)
'*  3. Program ID           : B2902mb1.asp
'*  4. Program Name         : B2902mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +cBListTableReflection
'                             +PB6G241ControlTableReflection
'*  7. Modified date(First) : 2000/09/26
'*  8. Modified date(Last)  : 2002/12/03
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************
option explicit
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPB6G241												'  ☆  ComProxy Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          
Dim StrModuleCD,VarExportGroup
Dim iErrPosition

' 조회용 상수 
Const B488_EG1_minor_nm = 0
Const B488_EG1_module_cd = 1
Const B488_EG1_table_id = 2
Const B488_EG1_use_flag = 3
Const B488_EG1_calendar_dt = 4
Const B488_EG1_change_id = 5
Const B488_EG1_success_flag = 6

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 
strSpread = Request("txtSpread")

Select Case strMode
Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 

   on error resume next
    Set pPB6G241 = Server.CreateObject("PB6G241.cBListTblReflect")

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
    If CheckSYSTEMError(Err,True) = True Then
        set pPB6G241 = nothing
        Response.End  
    End If	
	on error goto 0

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    If Request("cboModuleCd") = "*" Then
        StrModuleCD = ""
    Else    
		StrModuleCD = Request("cboModuleCd")
    End If
    
    '-----------------------
    'Com Action Area
    '-----------------------
    on error resume next
    call pPB6G241.B_LIST_TABLE_REFLECTION(gStrGlobalCollection,StrModuleCD,Request("txtTable"),VarExportGroup)

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
    If CheckSYSTEMError(Err,True) = True Then
        set pPB6G241 = nothing
        Response.End  
    End If	
	on error goto 0

%>
<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		
		LngMaxRow = 0										'Save previous Maxrow                                                
<%      
    GroupCount = Ubound(VarExportGroup,1)
	For LngRow = 0 To GroupCount
%>
        strData = strData & Chr(11) & "<%=ConvSPChars(VarExportGroup(LngRow,B488_EG1_minor_nm))%>"	            '  Minor Name
        strData = strData & Chr(11) & "<%=ConvSPChars(VarExportGroup(LngRow,B488_EG1_module_cd))%>"	'  Module Code
        strData = strData & Chr(11) & "<%=ConvSPChars(VarExportGroup(LngRow,B488_EG1_table_id))%>"	'  Table ID
        strData = strData & Chr(11) & " " '4 PopupButton
        
	  If "<%=VarExportGroup(LngRow,B488_EG1_use_flag)%>" = "Y" Then                                   ' Use Flag
		strData = strData & Chr(11) & "1" '5
      Else
		strData = strData & Chr(11) & "0" '5
	  End If
		
        strData = strData & Chr(11) & "<%=UNIDateClientFormat(VarExportGroup(LngRow,B488_EG1_calendar_dt))%>"    '  Calendar Date
        strData = strData & Chr(11) & "<%=ConvSPChars(VarExportGroup(LngRow,B488_EG1_change_id))%>"	'  Change ID
        
      If "<%=VarExportGroup(LngRow,B488_EG1_success_flag)%>" = "Y" Then                                   ' Success Flag
		strData = strData & Chr(11) & "1" '8
      Elseif "<%=VarExportGroup(LngRow,B488_EG1_success_flag)%>" = "N" Then
		strData = strData & Chr(11) & "0" '8 
	  Else
		strData = strData & Chr(11) & ""  '8
	  End If
	          
        strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
        strData = strData & Chr(11) & Chr(12)
<%      
    Next
%>    
	.ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData strData
				
	.frm1.hTable.value = "<%=ConvSPChars(Request("txtTable"))%>"
	.frm1.hModuleCd.value = "<%=Request("cboModuleCd")%>"
	.DbQueryOk
	End With
</Script>	
<%    
    Set pPB6G241 = Nothing
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
									
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If

     on error resume next
    Set pPB6G241 = Server.CreateObject("PB6G241.cBControlTblReflect")    
    If CheckSYSTEMError(Err,True) = True Then
        set pPB6G241 = nothing
        Response.End  
    End If	
	on error goto 0
    
    on error resume next
    Call pPB6G241.B_CONTROL_TABLE_REFLECTION(gStrGlobalCollection,strSpread,iErrPosition)
    If CheckSYSTEMError2(Err,True,iErrPosition & "행","","","","") = True Then
        set pPB6G241 = nothing
        Response.End  
    End If
 	on error goto 0

    Set pPB6G241 = Nothing                                                   '☜: Unload Comproxy
    
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
