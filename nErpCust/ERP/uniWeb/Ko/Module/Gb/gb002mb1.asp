<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425mb1.asp
'*  4. Program Name         : ������������ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-02-19
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->


<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "C", "NOCOOKIE","MB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2,rs3,rs4


Dim	txtYyyymm
Dim	txtBizUnitCd
Dim	txtCostCd
Dim	txtAcctGp
Dim	txtCtrlCd


Dim lgExpenseSum
Dim lgProfitSum
Dim lgCostNm
Dim lgAcctGpNm
Dim lgCtrlNm

Dim lgDataExist
Dim lgCodeCond
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgPageNo
Dim lgErrorStatus
Dim lgStrData
									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
									
Const C_SHEETMAXROWS_D  = 100                      									
lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

Call HideStatusWnd

	On Error Resume Next
	Err.Clear
	
	txtYyyymm			= Trim(Request("txtYyyymm"))
	txtCostCd			= Trim(Request("txtCostCd"))
	txtAcctGp			= Trim(Request("txtAcctGp"))
	txtCtrlCd			= Trim(Request("txtCtrlCd"))


	
    lgPageNo       = Cint(Trim(Request("lgPageNo")))    
'   lgMaxCount     = Cint(Trim(Request("lgMaxCount")))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgDataExist    = "No"
	lgErrorStatus  = "No" 
	lgCodeCond = ""


	IF txtCostCd = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and a.cost_cd = " & FilterVar(txtCostCd, "''", "S")
	END IF

	IF txtAcctGp = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and c.gp_cd = " & FilterVar(txtAcctGp, "''", "S")
	END IF


	IF txtCtrlCd = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and a.ctrl_cd = " & FilterVar(txtCtrlCd, "''", "S")
	END IF

	

	Call FixUNISQLData()
	Call QueryData()

	


Sub MakeSpreadSheetData()
   On Error Resume Next
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
   
    
    lgDataExist    = "Yes"
    lgStrData      = ""


    iLoopCount = 0
  

    If lgPageNo > 0 Then
       rs0.Move     = lgMaxCount * lgPageNo                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If  
  
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
			
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		'C/C
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'C/C�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))		'Cost Type
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))		'���� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))		'������ 
		IF ConvSPChars(rs0(5)) = "F" Then							'����Ư�� 
			lgstrData = lgstrData & Chr(11) & "����"
		ELSEIF ConvSPChars(rs0(5)) = "T" Then
			lgstrData = lgstrData & Chr(11) & "����ȯ��"
		ELSE
			lgstrData = lgstrData & Chr(11) & "����"
		END IF
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(6))		'�����׸� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(7))		'�����׸�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(8))		'�����׸�Value
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(9))		'�����׸�Value�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(10))		'�����׸� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(11))		'�����׸�� 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(12), ggAmtOfMoney.DecPoint, 0) '�ݾ� 
					
		lgstrData = lgstrData & Chr(11) & iLoopCount 
		lgstrData = lgstrData & Chr(11) & Chr(12)		
			
        
        If  iLoopCount >= lgMaxCount Then
            lgPageNo = lgPageNo + 1
            Exit Do
        End If        
 
        rs0.MoveNext
	Loop

	
    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = 0													'��: ���� ����Ÿ ����.
    End If

  	
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

    UNISqlId(0) = "GB002MA01"
    UNISqlId(1) = "GB002MA02"					
    UNISqlId(2) = "commonqry"
    UNISqlId(3) = "commonqry"	
    UNISqlId(4) = "commonqry"	
    					
 
    UNIValue(0,0) = FilterVar(txtYYYYMM, "''", "S")
    UNIValue(0,1) = lgCodeCond
    
    	
    UNIValue(1,0) = FilterVar(txtYYYYMM, "''", "S")
	UNIValue(1,1) = lgCodeCond

   
	UNIValue(2,0)  = "SELECT cost_nm from b_cost_center where cost_cd = " & FilterVar(txtCostCd, "''", "S") 
	UNIValue(3,0)  = "SELECT gp_nm from a_acct_gp where gp_cd = " & FilterVar(txtAcctGp, "''", "S") 
	UNIValue(4,0)  = "SELECT ctrl_nm from a_ctrl_item where ctrl_cd = " & FilterVar(txtCtrlCd, "''", "S") 
		
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2,rs3,rs4)
    


    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
  
        
    IF NOT (rs1.EOF or rs1.BOF) then
		lgExpenseSum	= rs1(0)
		lgProfitSum	= rs1(1)
	ELSE
		lgExpenseSum	= 0
		lgProfitSum	= 0
	End if
	
	
    rs1.Close
    Set rs1 = Nothing 
    

    IF NOT (rs2.EOF or rs2.BOF) then
		lgCostNm = rs2(0)				
	ELSE
		lgCostNm = ""
	End if
    rs2.Close
    Set rs2 = Nothing 

    IF NOT (rs3.EOF or rs3.BOF) then
		lgAcctGpNm = rs3(0)				
	ELSE
		lgAcctGpNm = ""
	End if
    rs3.Close
    Set rs3 = Nothing 

    IF NOT (rs4.EOF or rs4.BOF) then
		lgCtrlNm = rs4(0)			
	ELSE
		lgCtrlNm = ""
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


    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
		
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
	   .frm1.hYyyymm.value		=  "<%=ConvSPChars(txtYyyymm)%>"
	   .frm1.hCostCd.value		=  "<%=ConvSPChars(txtCostCd)%>"	 	
	   .frm1.hAcctGp.value	=  "<%=ConvSPChars(txtAcctGp)%>"	 	
	   .frm1.hCtrlCd.value	=  "<%=ConvSPChars(txtCtrlCd)%>"	 		 	
	   
	   .frm1.txtCostNm.value	= "<%=ConvSPChars(lgCostNm)%>"
	   .frm1.txtAcctGpNm.value	= "<%=ConvSPChars(lgAcctGpNm)%>"
	   .frm1.txtCtrlNm.value = "<%=ConvSPChars(lgCtrlNm)%>"
	   
	   
	   .frm1.txtExpenseSum.text		= "<%=UniNumClientFormat(lgExpenseSum,ggAmtOfMoney.Decpoint,0)%>" 
	   .frm1.txtProfitSum.text		= "<%=UniNumClientFormat(lgProfitSum,ggAmtOfMoney.Decpoint,0)%>" 
	   
		If "<%=lgDataExist%>" = "Yes" AND "<%=lgErrorStatus%>" <> "YES" Then
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
			.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		   .DbQueryOk
		End If

    End With

</Script>	
	

<%
Set lgADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
