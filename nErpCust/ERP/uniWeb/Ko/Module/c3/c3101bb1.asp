<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c3101bb1
'*  4. Program Name         : �������� ��� 
'*  5. Program Desc         : �������� ��� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Bong Hoon, Song
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
Server.ScriptTimeOut = 10000
%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'@Var_Declare
'--- Karrman_ADO
'Dim lgADF														'ActiveX Data Factory ���� �������� 
'Dim iStr
'Dim lgstrRetMsg												'Record Set Return Message �������� 
'Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 
Dim strQryMode												'������ Query ���¸� ���� �������� 

'Const DISCONNUPD  = "1"										'Disconnect + Update Mode
'Const DISCONNREAD = "2"										'Disconnect + ReadOnly Mode

'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
On Error Resume Next
Call LoadBasisGlobalInf() 														'��: 

Call HideStatusWnd 

Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRowItem
Dim arrVal, arrTemp							'��: Spread Sheet �� ���� ���� Array ����   
Dim lGrpCnt
Dim IntRetCD

lgStrPrevKey 	= Request("lgStrPrevKey")
LngMaxRow		= Request("txtMaxRows")

'--- Karrman_ADO

strQryMode = Request("lgIntFlgMode")						'�� : ���� Query ���¸� ���� 

Redim UNISqlId(0)
Redim UNIValue(0,2)


   arrTemp = Split(Request("txtSpread"), gRowSep)		'ITEM SPREAD
   Err.Clear
   
   
   
   For LngRowItem = 1 To LngMaxRow

	lGrpCnt = lGrpCnt +1
	arrVal = Split(arrTemp(LngRowItem -1), gColSep)
    
    Call SubCreateCommandObject(lgObjComm)	 
    With lgObjComm
		.CommandTimeout = 0
        .CommandText = "usp_c_actl_exe"
        .CommandType = adCmdStoredProc
        
        .Parameters.Append .CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)

	    .Parameters.Append .CreateParameter("@work_step"  ,adVarXChar,adParamInput,LEN(Trim(arrVal(0))), Trim(arrVal(0)))
		.Parameters.Append .CreateParameter("@yyyymm"     ,adVarXChar,adParamInput,LEN(Trim(arrVal(2))), Trim(arrVal(2)))
	    .Parameters.Append .CreateParameter("@usr_id"     ,adVarXChar,adParamInput,13, gUsrID)
        .Parameters.Append .CreateParameter("@msg_cd"     ,adVarXChar,adParamOutput,6)

        .Execute ,, adExecuteNoRecords

    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '��: Protect system from crashing   
			Response.end
        end if
    Else
        lgErrorStatus     = "YES"                                                         '��: Set error status
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
    End if
    


	
		Call SubCloseCommandObject(lgObjComm)		
   Next
   
 
   
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pConn,pRs,pErr)
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status
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

'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------
%>

<Script Language=vbscript>
	With parent
		IF "<%=lgErrorStatus%>"	<> "YES" Then																    '��: ȭ�� ó�� ASP �� ��Ī�� 
			.DbSaveOk
		ENd If
	End With
</Script>
	
<%
'Set lgADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
