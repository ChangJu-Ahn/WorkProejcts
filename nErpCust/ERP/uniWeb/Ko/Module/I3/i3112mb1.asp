<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3112MB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/09/01
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : LEE SEUNG WOOK
'* 11. Comment
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change" 
'*                            this mark(¢Á) Means that "may  change"
'*                            this mark(¡Ù) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<%
Call LoadBasisGlobalInf()			
Call HideStatusWnd
	
Dim iStrReAnalFlag
Dim iStrPlantCd
Dim iStrAnalyyyymm
Dim iStrLongterm
Dim iStrPernicious
Dim IntRetCD
Dim strMsg_cd

	On Error Resume Next														
 	Err.Clear
 	
 	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)

	iStrReAnalFlag	= Request("txtReAnalFlag")
	iStrPlantCd		= Request("txtPlantCd")
	iStrAnalyyyymm	= Request("txtAnalYYYYMM")
	iStrLongterm	= Request("txtLongterm")
	iStrPernicious	= Request("txtPernicious")

	With lgObjComm

		.CommandTimeOut = 1800
	    .CommandText = "USP_I_LONGTERM_ANALYSIS"
	    .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    .Parameters.Append .CreateParameter("@yyyymm",				adVarXChar,adParamInput,LEN(iStrAnalyyyymm), iStrAnalyyyymm)
	    .Parameters.Append .CreateParameter("@plant_cd",			adVarXChar,adParamInput,LEN(iStrPlantCd), iStrPlantCd)
		.Parameters.Append .CreateParameter("@Longterm_period",		adVarXChar,adParamInput,13, iStrLongterm)
	    .Parameters.Append .CreateParameter("@Pernicious_period",	adVarXChar,adParamInput,13, iStrPernicious)
	    .Parameters.Append .CreateParameter("@insrt_user_id",		adVarXChar,adParamInput,13, gUsrID)
	    .Parameters.Append .CreateParameter("@msg_cd"		,		adVarXChar,adParamOutput,8)

	    .Execute ,, adExecuteNoRecords

	End With

	If  Err.number = 0 Then
	    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

	    if  IntRetCD <> 1 then
	        strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
	        Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '¢Ð: Protect system from crashing   
			Response.end
	    end if
	Else
	    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	    Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if

	Call SubCloseCommandObject(lgObjComm)
	Call SubCloseDB(lgObjConn)
%>

<Script Language=vbscript>
'Dim strData
	Call Parent.DbAnalysisOk
</Script>	
