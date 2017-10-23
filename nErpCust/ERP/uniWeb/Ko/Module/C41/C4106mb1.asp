<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3980mb2
'*  4. Program Name         : 마감정보관리 
'*  5. Program Desc         : 원가계산을 위한 마감정보관리 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2004/12/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True												'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->

<%													

On Error Resume Next
Err.Clear 

Call LoadBasisGlobalInf() 

Dim ADF																'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg														'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0						'DBAgent Parameter 선언 

Const C_CLOSE_GB            = 0
Const C_TAGET_WORKING_MNTH  = 1
Const C_CLOSE               = 2
Const C_CANCEL              = 3

'---------------------------------------------------------------------------------------------------------

Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""  
    lgOpModeCRUD      = Request("txtMode")												'☜: Read Operation Mode (CRUD)   

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             'Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizSave
' Desc : 
'============================================================================================================
Sub SubBizSave()
	On Error Resume Next
	Err.Clear

	Dim ii,arrTemp,arrVal
	Dim iNextMnth

	arrTemp = Split(Request("txtSpread"), gRowSep)		'ITEM SPREAD

	For ii = 0 To UBound(arrTemp,1) - 1
		arrVal = Split(arrTemp(ii), gColSep)

		If arrVal(C_CLOSE_GB) = "MC" Then
			If arrVal(C_CLOSE) = "1" And arrVal(C_CANCEL) = "0" Then
				Call SubMovementClose(arrVal(C_TAGET_WORKING_MNTH),"CC")
			ElseIf 	arrVal(C_CLOSE) = "0" And arrVal(C_CANCEL) = "1" Then
				Call SubMovementClose(arrVal(C_TAGET_WORKING_MNTH),"CD")
			End If
		ElseIf 	arrVal(C_CLOSE_GB) = "CC" Then
			If arrVal(C_CLOSE) = "1" And  arrVal(C_CANCEL) = "0" Then
				Call SubCostClose(arrVal(C_TAGET_WORKING_MNTH),"C")
			ElseIf 	arrVal(C_CLOSE) = "0" And arrVal(C_CANCEL) = "1" Then
				Call SubCostClose(arrVal(C_TAGET_WORKING_MNTH),"D")
			End If
		ElseIf 	arrVal(C_CLOSE_GB) = "AC" Then
			If arrVal(C_CLOSE) = "1" And  arrVal(C_CANCEL) = "0" Then
				iNextMnth = uniConvYYYYMMDDToDate(gDateFormat,Left(arrVal(C_TAGET_WORKING_MNTH),4),Right(arrVal(C_TAGET_WORKING_MNTH),2),"01")
				iNextMnth = uniDateadd("M",1,iNextMnth,gServerDateFormat)
				iNextMnth = uniConvDateToYYYYMM(iNextMnth,gDateFormat,"")

				Call SubAccountClose(iNextMnth,"1")
			ElseIf 	arrVal(C_CLOSE) = "0" And arrVal(C_CANCEL) = "1" Then
				Call SubAccountClose(arrVal(C_TAGET_WORKING_MNTH),"2")
			End If			
		End If	
	Next

	Response.Write " <Script Language=vbscript>                      " & vbCr
	Response.Write " With parent                                     " & vbCr
	Response.Write " 	If """ & lgErrorStatus & """ <> ""YES"" Then " & vbCr															    '☜: 화면 처리 ASP 를 지칭함 
	Response.Write "		.DbSaveOk                                " & vbCr
	Response.Write "	End If                                       " & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write " </Script>	                                     " & vbCr
End Sub

'============================================================================================================
' Name : SubCostClose
' Desc : 
'============================================================================================================
Sub SubCostClose(ByVal working_mnth,ByVal woking_type)
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    Dim iPC4G106
		
    Set iPC4G106 = Server.CreateObject("PC4G106.cCMngCostClosingNewSvr")

    If CheckSYSTEMError(Err, True) = True Then
		SetErrorStatus()						
		Exit Sub
    End If    

    Call iPC4G106.C_MANAGE_COST_CLOSING_NEW_SVR (gStrGloBalCollection,woking_type,working_mnth)		
		
    If CheckSYSTEMError(Err, True) = True Then					
		Set iPC4G106 = Nothing
		SetErrorStatus()
		Exit Sub
    End If    
    
    Set iPC4G106 = Nothing
End Sub    

'============================================================================================================
' Name : SubMovementClose
' Desc : 
'============================================================================================================
Sub SubMovementClose(ByVal Working_mnth,ByVal woking_type)
	Dim IntRetCD
	Dim strMsg_cd

    Call SubCreateCommandObject(lgObjComm)	 

    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "usp_c_close_movement"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    .Parameters.Append .CreateParameter("@work_type"  ,adVarChar,adParamInput,LEN(woking_type), Trim(woking_type))
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarChar,adParamInput,13, gUsrID)
		.Parameters.Append .CreateParameter("@yyyymm"     ,adVarChar,adParamInput,LEN(Working_mnth), Trim(Working_mnth))
        .Parameters.Append .CreateParameter("@out_date"   ,adVarChar,adParamOutput,8)	    
        .Parameters.Append .CreateParameter("@msg_cd"     ,adVarChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        If  IntRetCD <> 1 Then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
        End If
    Else
        lgErrorStatus = "YES"                                                         '☜: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End If

	Call SubCloseCommandObject(lgObjComm)	
End Sub

'============================================================================================================
' Name : SubMovementClose
' Desc : 
'============================================================================================================
Sub SubAccountClose(ByVal Working_mnth,ByVal woking_type)
	Dim IntRetCD
	Dim strMsg_cd

    Call SubCreateCommandObject(lgObjComm)	 

    With lgObjComm
        .CommandText = "usp_a_close_gl"
        .CommandType = adCmdStoredProc

        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@conf_fg"    ,adVarChar,adParamInput,LEN(woking_type), woking_type)
		.Parameters.Append .CreateParameter("@from_date"  ,adVarChar,adParamInput,LEN(Working_mnth), Working_mnth)
		.Parameters.Append .CreateParameter("@to_date"    ,adVarChar,adParamInput,LEN(Working_mnth), Working_mnth)
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarChar,adParamInput,13, gUsrID)
        .Parameters.Append .CreateParameter("@msg_cd"     ,adVarChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        If  IntRetCD <> 1 Then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
        End If
    Else
        lgErrorStatus = "YES"                                                         '☜: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End If

	Call SubCloseCommandObject(lgObjComm)	
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    
    If CheckSYSTEMError(pErr,True) = True Then
		ObjectContext.SetAbort
		Call SetErrorStatus
    Else
		If CheckSQLError(pConn,True) = True Then
			ObjectContext.SetAbort
			Call SetErrorStatus
		End If
	End If
End Sub   

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : 
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next
	Err.Clear 

	Dim strData
	Dim LngRow

	Redim UNISqlId(0)
	Redim UNIValue(0,0)

	UNISqlId(0) = "C5980MA101"
	UNILock = DISCONNREAD :	UNIFlag = "1"
		
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")

	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call ServerMesgBox("" , vbInformation, I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End																	'☜: 비지니스 로직 처리를 종료함 
	End If

    For LngRow =0 To rs0.RecordCount-1
    
'		Response.Write UNIMonthClientFormat(rs0("WORKING_MNTH")) & "<BR>"
		strData = strData & Chr(11) & ConvSPChars(rs0("MINOR_CD"))
		strData = strData & Chr(11) & ConvSPChars(rs0("MINOR_NM"))
		strData = strData & Chr(11) & ConvSPChars(rs0("PLANT_CD"))
		strData = strData & Chr(11) & ConvSPChars(rs0("CLOSE_DT"))
		strData = strData & Chr(11) & ConvSPChars(rs0("CLOSE_CREATE_DT"))
		strData = strData & Chr(11) & ConvSPChars(rs0("LAST_CLOSE_ID"))
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & ConvSPChars(rs0("CLOSE_FG"))
		strData = strData & Chr(11) & ConvSPChars(rs0("CANCEL_FG"))
		strData = strData & Chr(11) & LngRow
		strData = strData & Chr(11) & Chr(12)

		rs0.MoveNext
	Next

	rs0.Close
	Set rs0 = Nothing

	Response.Write " <Script Language=vbscript>               " & vbCr
	Response.Write " With Parent                              " & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData       " & vbCr
	Response.Write " .ggoSpread.SSShowData """ & strData & """" & vbCr
	Response.Write " .DbQueryOk                               " & vbCr
	Response.Write " End With                                 " & vbCr
	Response.Write " </Script>	                              " & vbCr

	Set ADF = Nothing																	'☜: ActiveX Data Factory Object Nothing
End Sub

%>

