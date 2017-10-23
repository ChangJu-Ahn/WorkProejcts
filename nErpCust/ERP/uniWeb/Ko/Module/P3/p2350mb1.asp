<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/adovbs.inc" -->
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2350mb1
'*  4. Program Name         : MRP 예시전개 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002-04-16
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'**********************************************************************************************-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd

    On Error Resume Next
    Err.Clear

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = "" 
    lgOpModeCRUD      = Request("txtMode")
    
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim IntRetCd
	Dim CurDate

	CurDate = GetSvrDate
	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
    Select Case UniCInt(lgOpModeCRUD, 0)
        Case UID_M0001				'☜: 전체Query
			Call SubBizQuery("P")
		Case UID_M0002
			Call SubBizQuery("CK")
			Call SubCreateCommandObject(lgObjComm)
			Call SubBizBatch()
			Call SubCloseCommandObject(lgObjComm)
    End Select
         
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pQryMode)

	On Error Resume Next
    Err.Clear
    	
	Dim strPlantCd
	Dim iKey1
	Dim strStatus_mrp
		
	iKey1 = FilterVar(lgKeyStream(0),"''","S")
	
	Select Case pQryMode
		Case "P"
				
			'--------------
			'공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("P",iKey1,"","","")

			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT) 
				Call SetErrorStatus()
				IntRetCd = -1
%>
				<Script Language=vbscript>
					With Parent	
						.Frm1.txtPlantNm.Value		= ""
						.frm1.txtFixExecToDt.text	= ""
						.frm1.txtPlanExecToDt.text	= ""
						.frm1.txtPlantCd.focus
				    End With 
				</Script>       
<%		
			Else
				IntRetCd = 1
%>
				<Script Language=vbscript>
					With Parent	
						.Frm1.txtPlantNm.Value		= "<%=ConvSPChars(lgObjRs(1))%>"                   'Set condition area
						.frm1.txtFixExecToDt.text	= "<%= UNIDateClientFormat(DateAdd("d",UniCInt(lgObjRs(11), 0), CurDate)) %>"
						.frm1.txtPlanExecToDt.text	= "<%= UNIDateClientFormat(DateAdd("d",UniCInt(lgObjRs(8), 0), CurDate)) %>"
						.lgInvCloseDt				= "<%= UNIDateClientFormat(lgObjRs("inv_cls_dt")) %>"
						
				    End With  
				</Script>       
<%			
			End If
			
		Case "CK"
				
			'--------------
			'공장 체크		
			'--------------	
			lgStrSQL = ""
			iKey1 = FilterVar(Trim(Request("txtPlantCd")),"''","S")
			Call SubMakeSQLStatements("CK",iKey1,"","","")

			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
				IntRetCd = 1
			Else
				IntRetCd = -1
				
				strStatus_mrp = lgObjRs(17)
				
				If strStatus_mrp = "1" Then
					Call DisplayMsgBox("187731", vbInformation, "", "", I_MKSCRIPT)
				    Response.End
				ElseIF  strStatus_mrp = "2" Or strStatus_mrp = "3" Then
					Call DisplayMsgBox("187732", vbInformation, "", "", I_MKSCRIPT)
				    Response.End
				End If
			End If
	End Select
					
	Call SubCloseRs(lgObjRs) 
   
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType
		Case "P"
			lgStrSQL = "SELECT * " 
            lgStrSQL = lgStrSQL & " FROM  b_plant "
            lgStrSQL = lgStrSQL & " WHERE plant_cd = " & pCode
        Case "CK"
			lgStrSQL = "SELECT * " 
            lgStrSQL = lgStrSQL & " FROM  p_mrp_history "
            lgStrSQL = lgStrSQL & " WHERE plant_cd = " & pCode
            lgStrSQL = lgStrSQL & " AND run_no = (select max(run_no) from p_mrp_history where plant_cd = " & pCode & ")"

    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()
	
	Dim strMsg_cd
    Dim strMsg_text
	Dim strPlantCd, strCurDate, strOpenDate, strPlanDate, strFlag, strSafeFlg, strInvFlg
	Dim strIdepFlg, strOptionFlg, strItemCd, strWrnFlg, strOrdNo, strCodrFlg
	Dim strNetFlg, strPackFlg, strScrap, strForward, strMpsScope
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	'------ Parameter Setting (Start ) ------------------------------------------------------------------ 
	strPlantCd = UCase(Trim(Request("txtPlantCd")))										'☆: Plant Code    
    strCurDate = UniConvDateToYYYYMMDD(Request("txtFixExecFromDt"),gDateFormat,"")
	strOpenDate = UniConvDateToYYYYMMDD(Request("txtFixExecToDt"),gDateFormat,"")
	strPlanDate = UniConvDateToYYYYMMDD(Request("txtPlanExecToDt"),gDateFormat,"")
	
    strFlag = " "

    If Request("rdoSafeInvFlg") = "Y" Then
         strSafeFlg  = "Y"
    Else
    	 strSafeFlg  = "N"
    End If

    If Request("rdoAvailInvFlg") = "Y" Then
         strInvFlg  = "Y"
    Else
    	 strInvFlg  = "N"
    End If

    strIdepFlg = "Y"
    strOptionFlg = "%"
    strItemCd = "%"
    strWrnFlg = "N"
    strOrdNo = " "
    strCodrFlg = "Y"
    If Request("rdoAvailInvFlg") = "Y" Then
         strNetFlg  = "Y"
    Else
    	 strNetFlg  = "N"
    End If

    strPackFlg = "N"
    strScrap = " "
    strForward = " "
    strMpsScope = " "
	'------ Parameter Setting (End ) ------------------------------------------------------------------ 
	
    With lgObjComm
        .CommandText = "usp_mmp000p"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@orgid",			advarXchar,adParamInput,4, strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@string_curdate",advarXchar,adParamInput,Len(strCurDate), strCurDate)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@string_p_date",	advarXchar,adParamInput,Len(strPlanDate), strPlanDate)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@string_o_date",	advarXchar,adParamInput,Len(strOpenDate),strOpenDate)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@flag",			advarXchar,adParamInput,1,strFlag)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@safe_flg",		advarXchar,adParamInput,1, strSafeFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@inv_flg",		advarXchar,adParamInput,1, strInvFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@idep_flg",		advarXchar,adParamInput,1,strIdepFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@option_flg",	advarXchar,adParamInput,1,strOptionFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@partno",		advarXchar,adParamInput,18, strItemCd)
	    
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@entusr",		advarXchar,adParamInput,13, gUsrID)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@wrn_flg",		advarXchar,adParamInput,1,strWrnFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@odrno",			advarXchar,adParamInput,18,strOrdNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@codr_flg",		advarXchar,adParamInput,18,strCodrFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@net_flg",		advarXchar,adParamInput,18, strNetFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pack_flg",		advarXchar,adParamInput,18, strPackFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@scrap",			advarXchar,adParamInput,18,strScrap)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@forward",		advarXchar,adParamInput,18,strForward)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@mpsscope",		advarXchar,adParamInput,18,strMpsScope)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)

        lgObjComm.Execute ,, adExecuteNoRecords
     
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        If  IntRetCD = 0 or IntRetCD = -1 then
			Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT) 
			IntRetCD = 1
        Else
        Call SvrMsgBox(err.number , vbinformation, i_mkscript)
			Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
			Call SubHandleError("MB",lgObjComm.ActiveConnection,lgObjRs,Err)
            IntRetCD = -1			
        End if
    Else           
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError("MB",lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
End Sub	


'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
        Case "MU"
        Case "MB"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If      
    End Select
End Sub
		
%>
<Script Language="VBScript">
	Select Case <%=lgOpModeCRUD%>
		Case <%=UID_M0001%>
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				Call parent.LookUpPlantOk()
	        End If   
		Case <%=UID_M0002%>
			If Trim("<%=lgErrorStatus%>") = "NO" And <%=IntRetCd%> <> -1 Then
				Call parent.ExecuteOk() 
	        End If
	End Select    
       
</Script>	
  