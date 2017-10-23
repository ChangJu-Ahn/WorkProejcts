<%'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 
'*  3. Program ID           : C3612MB1.ASP
'*  4. Program Name         : ���������γ�����ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
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

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4							'DBAgent Parameter ���� 

Dim	txtMode
Dim	txtYyyymm
Dim	txtPlantCd
Dim	txtCostCd
Dim	txtItemAcctCd
Dim	txtMaxRows
Dim	lgSpid
Dim lgSum
Dim lgPlantNm
Dim lgCostNm
Dim lgItemAcctNm

Dim lgDataExist
Dim lgCodeCond


									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

Call HideStatusWnd

On Error Resume Next
Err.Clear


	txtMode			= Trim(Request("txtMode"))						'�� : ���� ���¸� ���� 
	txtYyyymm		= Trim(Request("txtYyyymm"))
	txtItemAcctCd	= Trim(Request("txtItemAcctCd"))
	txtPlantCd		= Trim(Request("txtPlantCd"))
	txtCostCd		= Trim(Request("txtCostCd"))
	
		
	IF Trim(Request("txtSpid")) <> "" Then
		lgSpid			= Trim(Request("txtSpid"))
	Else
		lgSpid			= ""
	END If
	
	lgCodeCond = ""
	IF Trim(txtItemAcctCd) <> "" Then
		lgCodeCond = lgCodeCond & " and a.child_item_acct = " & FilterVar(txtItemAcctCd, "''", "S")
	END IF
	
	IF Trim(txtPlantCd) <> "" Then
		lgCodeCond = lgCodeCond & " and a.child_plant_cd = " & FilterVar(txtPlantCd, "''", "S")
	END IF

	IF Trim(txtCostCd) <> "" Then
		lgCodeCond = lgCodeCond & " and a.cost_cd = " & FilterVar(txtCostCd, "''", "S")
	END IF
	
    select case txtMode
		case CStr(UID_M0001)
			Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
			Call SubBizBatch("C")    
			Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection   
			Call FixUNISQLData()
			Call QueryData()
		case CStr(UID_M0003)
			Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
			Call SubBizBatch("D")    
			Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection   
	end select




	

	
	

Sub SubBizBatch(ByVal flag)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatchMulti(flag)                            '��: Run Batch
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
	strSp = "usp_c_common_material_dstb_detail"

	With lgObjComm
	   .CommandText = strSp
	   .CommandType = adCmdStoredProc
	
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,	adParamReturnValue)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@flag",			adVarXChar,	adParamInput,		1,	Flag)	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@yyyymm",			adVarXChar,	adParamInput,		6,	txtYyyymm)
	   	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id",			adVarXChar,	adParamInput,		10,	lgSpid)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@out_spid",		adVarXChar,	adParamOutput,10)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",		    adVarXChar,	adParamOutput,6)	   		  
	   
	   
	   lgObjComm.Execute ,, adExecuteNoRecords	
	End With

	If Err.number = 0 Then
	   IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value		
	   If IntRetCD <> 1 then
	      strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
	      If strMsg_Cd <> "" Then
		       Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
		  End If
	   End If
	   lgSpid = lgObjComm.Parameters("@out_spid").Value  
	Else    
	  lgErrorStatus     = "YES"
	  Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
	  Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if    
	
     
End Sub



Sub MakeSpreadSheetData()

    Dim  iLoopCount
    Dim  iRowStr
   
    
    lgDataExist    = "Yes"
    lgStrData      = ""
	

    iLoopCount = 0
  
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        		 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		'���� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'ǰ�� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))		'ǰ��� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))		'C/C
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))		'C/C�� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(5))		'ǰ����� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(6))		'ǰ������� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(7))		'���ޱ��� 
 		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(8), ggQty.DecPoint, 0) '���Լ��� 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(9), ggAmtOfMoney.DecPoint, 0) '���Աݾ� 
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

    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "C3612MA01"
    UNISqlId(1) = "C3612MA02"
    UNISqlId(2) = "COMMONQRY"					
    UNISqlId(3) = "COMMONQRY"					
    UNISqlId(4) = "COMMONQRY"					
    
    UNIValue(0,0) = FilterVar(txtYyyymm, "''", "S") 
    UNIValue(0,1) = lgCodeCond
	
    UNIValue(1,0) = FilterVar(txtYyyymm, "''", "S") 
    UNIValue(1,1) = lgCodeCond
    
    UNIValue(2,0)  = "SELECT MINOR_NM FROM b_minor where major_Cd = " & FilterVar("P1001", "''", "S") & "  and minor_Cd = " & FilterVar(txtItemAcctCd, "''", "S") '
	UNIValue(3,0)  = "select plant_nm from b_plant where plant_cd = " & FilterVar(txtPlantCd, "''", "S") 
	UNIValue(4,0)  = "select cost_nm from b_cost_center where cost_cd = " & FilterVar(txtCostCd, "''", "S") 
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    IF NOT (rs1.EOF or rs1.BOF) then
		lgSum = rs1(0)				' SUM(A.AMT)
	ELSE
		lgSum = 0
	End if
    rs1.Close
    Set rs1 = Nothing 
    

    IF NOT (rs2.EOF or rs2.BOF) then
 		lgItemAcctNm = rs2(0)				' SUM(A.AMT)
	ELSE
		lgItemAcctNm = ""
	End if
    rs2.Close
    Set rs2 = Nothing 

    IF NOT (rs3.EOF or rs3.BOF) then
		lgPlantNm = rs3(0)				' SUM(A.AMT)
	ELSE
		lgPlantNm = ""
	End if
    rs3.Close
    Set rs3 = Nothing 

    IF NOT (rs4.EOF or rs4.BOF) then
		lgCostNm = rs4(0)				' SUM(A.AMT)
	ELSE
		lgCostNm = ""
	End if
    rs4.Close
    Set rs4 = Nothing 

	    
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
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


%>

<Script Language=vbscript>
 
	With Parent
		select case "<%=txtMode%>"	
			case "<%=UID_M0001%>" 
				   .frm1.hSpid.value	=  "<%=ConvSPChars(lgSpid)%>"
				   .frm1.hYYYYMM.value	=  "<%=ConvSPChars(txtYYYYMM)%>"
				   .frm1.hItemAcctCd.value	=  "<%=ConvSPChars(txtItemAcctCd)%>"
				   .frm1.hPlantCd.value	=  "<%=ConvSPChars(txtPlantCd)%>"
				   .frm1.hCostCd.value	=  "<%=ConvSPChars(txtCostCd)%>"
				   				   				   				   
				   .frm1.txtSum1.text	= "<%=UniNumClientFormat(lgSum,ggAmtOfMoney.Decpoint,0)%>" 
				   .frm1.txtItemAcctNm.value = "<%=ConvSPChars(lgItemAcctNm)%>"
				   .frm1.txtPlantNm.value = "<%=ConvSPChars(lgPlantNm)%>"
				   .frm1.txtCostNm.value = "<%=ConvSPChars(lgCostNm)%>"
				   
				If "<%=lgDataExist%>" = "Yes" AND "<%=lgErrorStatus%>" <> "YES" Then
				   .ggoSpread.Source  = Parent.frm1.vspdData1
				   .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
					

				   .DbQueryOk
				End If
		end select			    
    End With

</Script>	
	

<%
Set lgADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
