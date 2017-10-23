<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425mb1.asp
'*  4. Program Name         : 오더별실적조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-02-19
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0,rs1,rs2

Dim	txtMode
Dim	txtFromYYYYMM
Dim	txtToYYYYMM
Dim	txtItemAcctCd
Dim txtItemAcctNm
Dim	txtMovTypeCd
Dim	txtMovTypeNm
Dim	txtRadio

Dim	txtMaxRows
Dim lgDataExist



									'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================

Call HideStatusWnd

On Error Resume Next
Err.Clear


	txtMode			= Trim(Request("txtMode"))						'☜ : 현재 상태를 받음 
	txtFromYyyymm	= Trim(Request("txtFromYyyymm"))
	txtToYyyymm		= Trim(Request("txtToYyyymm"))
	txtItemAcctCd	= Trim(Request("txtItemAcctCd"))
	txtMovTypeCd	= Trim(Request("txtMovTypeCd"))
	txtRadio		= Trim(Request("txtRadio"))
	txtMaxRows		= Trim(Request("txtMaxRows"))
	
		
	
    select case txtMode
		case CStr(UID_M0001)
			Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
			Call SubBizBatch("C")    
			Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection   
			Call FixUNISQLData()
			Call QueryData()
                                                 '☜: Close DB Connection   
	end select



Sub SubBizBatch(ByVal flag)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatchMulti(flag)                            '☜: Run Batch
    Call SubCloseCommandObject(lgObjComm)

End Sub


'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti(ByVal Flag)
	On Error Resume NEXT
	Err.Clear
	
	     
	Dim IntRetCD
	Dim strMsg_cd, strMsg_text
	Dim strSp
	Dim spid    
	dim temp
	strSp = "usp_c_mcs_temp1"

	With lgObjComm
	   .CommandText = strSp
	   .CommandType = adCmdStoredProc
		.CommandTimeOut = 0
	
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,	adParamReturnValue)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@from_yyyymm",	adVarXChar,	adParamInput,		6,	txtFromYyyymm)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@to_yyyymm",		adVarXChar,	adParamInput,		6,	txtToYyyymm)
	   
	   IF txtRadio = "S" or txtRadio = "D" Then
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@flag1",			adVarXChar,	adParamInput,		1,	"A")	   
	   Else
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@flag1",			adVarXChar,	adParamInput,		1,	"S")	   
	   End If
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamOutput,6)
	   
	   lgObjComm.Execute ,, adExecuteNoRecords	
	End With
		
	
	If Err.number = 0 Then
	   IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value		
        if  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )                                                              '☜: Protect system from crashing   
			Response.end
        end if
	Else    
	  lgErrorStatus     = "YES"
	  Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
	  Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if    
	
     
End Sub



Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
   
    
    lgDataExist    = "Yes"
    lgStrData      = ""


    iLoopCount = 0
  
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1

		IF txtRadio = "S" or txtRadio = "S1"  Then
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))
		    lgstrData = lgstrData & Chr(11) & ""	
		    lgstrData = lgstrData & Chr(11) & ""
		    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(3), ggAmtOfMoney.DecPoint, 0)		 
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(5))		
		ELSE
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))				
		    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(5), ggAmtOfMoney.DecPoint, 0)		 
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(6))		
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(7))		
		END IF 		

        lgstrData = lgstrData & Chr(11) & iLoopCount 
        lgstrData = lgstrData & Chr(11) & Chr(12)
 
        rs0.MoveNext
	Loop

	rs0.Close
    Set rs0 = Nothing 
    
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(2,0)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "COMMONQRY"
    UNISqlId(1) = "COMMONQRY"
    UNISqlId(2) = "COMMONQRY"
    
    IF txtRadio = "S" or txtRadio = "S1" Then
		UNIValue(0,0)  = "SELECT a.item_acct,isnull(b.minor_nm,''),a.remark,sum(a.amt),acct_seq,seq,type FROM C_MCS_TEMP a left outer join B_MINOR b  on a.item_acct = b.minor_cd and b.major_cd = " & FilterVar("P1001", "''", "S") & "  Where 1=1 "
		
		IF txtItemAcctCd <> "" Then
				UNIValue(0,0)  = UNIValue(0,0) & " and a.item_acct = " & FilterVar(txtItemAcctCd, "''", "S")
		END IF

		IF txtMovTypeCd <> "" Then
				UNIValue(0,0)  = UNIValue(0,0) & " and a.mov_type = " & FilterVar(txtMovTypeCd, "''", "S")
		END IF

		UNIValue(0,0)  = UNIValue(0,0) & " group by a.item_acct,isnull(b.minor_nm,''),a.remark,a.acct_seq,a.seq,a.type order by a.type,a.acct_seq,a.item_acct,a.seq "
	ELSE
		UNIValue(0,0)  = "SELECT a.item_acct,isnull(b.minor_nm,''),a.remark,a.mov_type,isnull(c.minor_nm,''),sum(a.amt),acct_seq,seq,type FROM C_MCS_TEMP a left outer join B_MINOR b  on a.item_acct = b.minor_cd and b.major_cd = " & FilterVar("P1001", "''", "S") & "  left outer join B_MINOR c on a.mov_type = c.minor_cd and c.major_cd = " & FilterVar("I0001", "''", "S") & "  Where 1=1 "
		
		IF txtItemAcctCd <> "" Then
				UNIValue(0,0)  = UNIValue(0,0) & " and a.item_acct = " & FilterVar(txtItemAcctCd, "''", "S")
		END IF

		IF txtMovTypeCd <> "" Then
				UNIValue(0,0)  = UNIValue(0,0) & " and a.mov_type = " & FilterVar(txtMovTypeCd, "''", "S")
		END IF

		UNIValue(0,0)  = UNIValue(0,0) & " group by a.item_acct,isnull(b.minor_nm,''),a.remark,a.mov_type,isnull(c.minor_nm,''),a.acct_seq,a.seq,a.type order by a.type,a.acct_seq,a.item_acct,a.seq "
	
	
	END IF

	UNIValue(1,0)  = "select minor_nm from b_minor where major_Cd = " & FilterVar("P1001", "''", "S") & "  and minor_Cd = " & FilterVar(txtItemAcctCd, "''", "S")
	UNIValue(2,0)  = "select minor_nm from b_minor where major_Cd = " & FilterVar("I0001", "''", "S") & "  and minor_Cd = " & FilterVar(txtMovTypeCd, "''", "S")
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs1.EOF And rs1.BOF Then    
		txtItemAcctNm = ""
	ELSE
		txtItemAcctNm = rs1(0)
	END IF	
	
	If  rs2.EOF And rs2.BOF Then  
		txtMovTypeNm = ""
	ELSE	
		txtMovTypeNm = rs2(0)
	END IF
	
    rs1.Close
    Set rs1 = Nothing
    rs2.Close
    Set rs2 = Nothing

	    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If


		
End Sub




'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
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


%>

<Script Language=vbscript>
 
	With Parent
		.frm1.txtItemAcctNm.Value = "<%=ConvSpChars(txtItemAcctNm)%>"
		.frm1.txtMovTypeNm.Value = "<%=ConvSpChars(txtMovTypeNm)%>"
		select case "<%=txtMode%>"	
			case "<%=UID_M0001%>" 
				If "<%=lgDataExist%>" = "Yes" AND "<%=lgErrorStatus%>" <> "YES" Then
				   .ggoSpread.Source  = Parent.frm1.vspdData1
				   .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
					
				   .DbQueryOk
				End If
		end select			    
    End With

</Script>	
	

<%
Set lgADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
