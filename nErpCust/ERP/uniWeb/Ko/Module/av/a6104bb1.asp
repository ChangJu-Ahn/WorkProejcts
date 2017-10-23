<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
Call LoadBasisGlobalInf()
							
On Error Resume Next
Err.Clear

Call HideStatusWnd

lgOpModeCRUD = Request("txtMode")											'☜: Read Operation Mode (CRUD)

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)														
        'Call SubBizQuery()														
        'Call SubBizQueryMulti()												
    Case CStr(UID_M0002)														
        'Call SubBizSave()														
         Call SubExecuteSP()														
    Case CStr(UID_M0003)														
        'Call SubBizDelete()													
End Select

Sub SubExecuteSP()
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    Dim IntRetCD
    Dim strFromDt, strToDt, strMsgCd

	strFromDt = UNIConvDateAToB(UNIConvDate(Request("txtFromDt")), gAPDateFormat, gServerDateFormat)	
	strToDt = UNIConvDateAToB(UNIConvDate(Request("txtToDt")), gAPDateFormat, gServerDateFormat)	

	 Call SubOpenDB(lgObjConn)                    
	 Call SubCreateCommandObject(lgObjComm)

		With lgObjComm
			.CommandText = "usp_a_update_a_vat_issue_dt_fg"
			.CommandType = adCmdStoredProc

			lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_dt"    , adDate,adParamInput,Len(Trim(strFromDt)), strFromDt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_dt"      , adDate,adParamInput,Len(Trim(strToDt)), strToDt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id"    , advarxchar,adParamInput,Len(Trim(gUsrID)), gUsrID)
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     , advarxchar ,adParamOutput,6)

			lgObjComm.Execute ,, adExecuteNoRecords
		End With

		IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

		If  IntRetCD <> 0 Then
            strMsgCd = lgObjComm.Parameters("@msg_cd").Value        
		    
            Call DisplayMsgBox(strMsgCd, vbInformation, "", "", I_MKSCRIPT)

			Call SubCloseCommandObject(lgObjComm)
			Call SubCloseRs(pvObjRs)
			Call SubCloseDB(lgObjConn)
			Response.end
		End If

		Call SubCloseCommandObject(lgObjComm)
        Call SubCloseDB(lgObjConn)       
    
        Call DisplayMsgBox("800154", vbInformation, "", "", I_MKSCRIPT)  ' msgno 추가되면사용~ 

End Sub
%>
