
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call HideStatusWnd

On Error Resume Next														'☜: 

Dim vAg0042                   				                                '☆ : 입력/수정용 ComProxy Dll 사용 변수(as0031
           																	'☆ : 조회용 ComProxy Dll 사용 변수 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strFromYYYYMM
Dim strTOYYYYMM
Dim Conf_fg
Dim IntRetCD
Dim lgObjComm
Dim lgErrorStatus 

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

strFromYYYYMM = Trim(Request("txtFromdt"))
strTOYYYYMM = Trim(Request("txtTodt"))
Conf_fg = Request("txtRadio")

Select Case strMode

Case CStr(UID_M0002)														'☜: 현재 조회/Prev/Next 요청을 받음 
  '********************************************************  
  '                        Execution
  '********************************************************  

    Err.Clear
    Call SubCreateCommandObject(lgObjComm)
    
    With lgObjComm

        .CommandText = "usp_a_close_gl"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@conf_fg"     ,adVarWChar,adParamInput,LEN(Conf_fg), Conf_fg)
		.Parameters.Append .CreateParameter("@from_date"     ,adVarWChar,adParamInput,LEN(strFromYYYYMM), strFromYYYYMM)
		.Parameters.Append .CreateParameter("@to_date"     ,adVarWChar,adParamInput,LEN(strToYYYYMM), strToYYYYMM)
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarWChar,adParamInput,13, gUsrID)
        .Parameters.Append .CreateParameter("@msg_cd"     ,adVarWChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords

    End With


  '  If  Err.number = 0 Then '2006.10 lee wol san 
  
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
        end if
   ' Else
      '  lgErrorStatus     = "YES"                                                         '☜: Set error status
        'Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    'End if
    
    
    Call SubCloseCommandObject(lgObjComm)



End Select
%>	

<Script Language="VBScript">
    Parent.fnButtonExecOk
</Script>