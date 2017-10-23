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
<%Option Explicit%>
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
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7							'DBAgent Parameter ���� 


Dim	txtYyyymm
Dim	txtPlantCd
Dim	txtCostCd
Dim	txtItemAcctCd
Dim	txtItemCd
Dim	txtTrnsTypeCd
Dim txtMovTypeCd
Dim txtFlag

Dim lgQtySum
Dim lgAmtsum
Dim lgDiffSum
Dim lgPlantNm
Dim lgCostNm
Dim lgItemAcctNm
Dim lgItemNm
Dim lgTrnsTypeNm
Dim lgMovTypeNm

Dim lgDataExist
Dim lgCodeCond
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgPageNo
Dim lgErrorStatus
Dim lgStrData
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

	
	txtYyyymm		= Trim(Request("txtYyyymm"))
	txtPlantCd		= Trim(Request("txtPlantCd"))
	txtCostCd		= Trim(Request("txtCostCd"))
	txtItemAcctCd	= Trim(Request("txtItemAcctCd"))
	txtItemCd		= Trim(Request("txtItemCd"))
	txtTrnsTypeCd	= Trim(Request("txtTrnsTypeCd"))
	txtMovTypeCd	= Trim(Request("txtMovTypeCd"))
	txtFlag			= UCase(Trim(Request("txtFlag")))
	
    lgPageNo       = Cint(Trim(Request("lgPageNo")))    
'   lgMaxCount     = Cint(Trim(Request("lgMaxCount")))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgDataExist    = "No"
	lgErrorStatus  = "No" 


	lgCodeCond = " and a.yyyymm = " & FilterVar(txtYyyymm, "''", "S")

	IF txtPlantCd = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = " and a.plant_cd = " & FilterVar(txtPlantCd, "''", "S")
	END IF
	
	IF Trim(txtCostCd) = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = " and a.cost_cd = " & FilterVar(txtCostCd, "''", "S")
	END IF

	IF Trim(txtTrnsTypeCd) = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = " and a.trns_type = " & FilterVar(txtTrnsTypeCd, "''", "S")
	END IF


	IF Trim(txtMoveTypeCd) = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = " and a.mov_type = " & FilterVar(txtMovTypeCd, "''", "S")
	END IF

	
	IF Trim(txtItemAcctCd) = "" Then
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and a.item_acct = " & FilterVar(txtItemAcctCd, "''", "S")
	END IF
	
	
	IF Trim(txtItemCd) = "" Then
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and a.item_cd = " & FilterVar(txtItemCd, "''", "S")
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

	Const C_SHEETMAXROWS_D  = 100   
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time

    iLoopCount = 0
  

    If lgPageNo > 0 Then
       rs0.Move     = lgMaxCount * lgPageNo                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If  
  
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        
    

		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		'���� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'�켱���� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))		'ǰ������� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))		'ǰ�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))		'ǰ��� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(5))		'�������� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(6))		'���������� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(7))		'�̵����� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(8))		'�̵������� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(9))		'Cost Center
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(10))		'Cost Center�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(11))		'�̵����� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(12))		'�̵�����켱���� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(13))		'�̵�â�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(14))		'�̵�ǰ����� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(15))		'�̵�ǰ�� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(16))		'�̵�ǰ��� 
			
 		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(17), ggQty.DecPoint, 0) '���� 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(18), ggAmtOfMoney.DecPoint, 0) '�ݾ� 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(19), ggAmtOfMoney.DecPoint, 0) '���̱ݾ� 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(20), ggQty.DecPoint, 0) '�ܰ�		
		
		
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

    Redim UNISqlId(7)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------


    Redim UNIValue(7,0)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

	IF txtFlag = "ITEMACCT" Then
	    UNISqlId(0) = "C3950MA01"
	ELSE
		UNISqlId(0) = "C3950MA02"
	END IF
	
	
    UNISqlId(1) = "C3950MA03"
    UNISqlId(2) = "COMMONQRY"					'PLANT_NM
    UNISqlId(3) = "COMMONQRY"					'COST CENTER NM
    UNISqlId(4) = "COMMONQRY"					'TRNS_TPYE_NM
    UNISqlId(5) = "COMMONQRY"					'MOV_TYPE_NM
    UNISqlId(6) = "COMMONQRY"					'ITEM_ACCT_NM
    UNISqlId(7) = "COMMONQRY"					'ITEM_NM
    
    UNIValue(0,0) = lgCodeCond
    UNIValue(1,0) = lgCodeCond
	
   
    UNIValue(2,0)  = "SELECT PLANT_NM FROM B_PLANT WHERE PLANT_CD=" & FilterVar(txtPlantCd, "''", "S") 
	UNIValue(3,0)  = "SELECT COST_NM FROM B_COST_CENTER WHERE COST_CD=" & FilterVar(txtCostCd, "''", "S") 
	UNIValue(4,0)  = "SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = " & FilterVar("I0002", "''", "S") & "  AND MINOR_CD= " & FilterVar(txtTrnsTypeCd, "''", "S") 
	UNIValue(5,0)  = "SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = " & FilterVar("I0001", "''", "S") & "  AND MINOR_CD= " & FilterVar(txtMovTypeCd, "''", "S") 
	UNIValue(6,0)  = "SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  AND MINOR_CD= " & FilterVar(txtItemAcctCd, "''", "S") 
	UNIValue(7,0)  = "SELECT ITEM_NM FROM B_ITEM WHERE ITEM_CD= " & FilterVar(txtItemCd, "''", "S") 
	
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2, rs3, rs4, rs5, rs6, rs7)
    


    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
        
        
    IF NOT (rs1.EOF or rs1.BOF) then
		lgQtySum = rs1(0)
		lgAmtSum = rs1(1)
		lgDiffSum = rs1(2)						
	ELSE
		lgQtySum = 0
		lgAmtSum = 0
		lgDiffSum = 0		
	End if
    rs1.Close
    Set rs1 = Nothing 
    

    IF NOT (rs2.EOF or rs2.BOF) then
		lgPlantnm = rs2(0)				
	ELSE
		lgPlantnm = ""
	End if
    rs2.Close
    Set rs2 = Nothing 

    IF NOT (rs3.EOF or rs3.BOF) then
		lgCostNm = rs3(0)				
	ELSE
		lgCostNm = ""
	End if
    rs3.Close
    Set rs3 = Nothing 

    IF NOT (rs4.EOF or rs4.BOF) then
		lgTrnsTypeNm = rs4(0)			
	ELSE
		lgTrnsTypeNm = ""
	End if
    rs4.Close
    Set rs4 = Nothing 
    
    IF NOT (rs5.EOF or rs5.BOF) then
		lgMovTypeNm = rs5(0)			
	ELSE
		lgMovTypeNm = ""
	End if
    rs5.Close
    Set rs5 = Nothing
        
    IF NOT (rs6.EOF or rs6.BOF) then
		lgItemAcctNm = rs6(0)			
	ELSE
		lgItemAcctNm = ""
	End if
    rs6.Close
    Set rs6 = Nothing  
    
        
    IF NOT (rs7.EOF or rs7.BOF) then
		lgItemNm = rs7(0)				
	ELSE
		lgItemNm = ""
	End if
    rs7.Close
    Set rs7 = Nothing      

		    
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
	   .frm1.hPlantCd.value		=  "<%=ConvSPChars(txtPlantCd)%>"
	   .frm1.hCostCd.value		=  "<%=ConvSPChars(txtCostCd)%>"	 	
	   .frm1.hTrnsTypeCd.value	=  "<%=ConvSPChars(txtTrnsTypeCd)%>"	 	
	   .frm1.hMovTypeCd.value	=  "<%=ConvSPChars(txtMovTypeCd)%>"	 		 	
	   .frm1.hItemAcctCd.value	=  "<%=ConvSPChars(txtItemAcctCd)%>"	 		 	
	   .frm1.hItemCd.value		=  "<%=ConvSPChars(txtItemCd)%>"	 		 	
	   
	   .frm1.txtPlantNm.value	= "<%=ConvSPChars(lgPlantNm)%>"
	   .frm1.txtCostNm.value	= "<%=ConvSPChars(lgCostnm)%>"
	   .frm1.txtTrnsTypeNm.value = "<%=ConvSPChars(lgTrnsTypeNm)%>"
	   .frm1.txtMovTypeNm.value	= "<%=ConvSPChars(lgMovTypeNm)%>"
	   .frm1.txtItemAcctNm.value = "<%=ConvSPChars(lgItemAcctNm)%>"
	   .frm1.txtItemNm.value	= "<%=ConvSPChars(lgItemNm)%>"

	   
	   .frm1.txtQtySum.text		= "<%=UniNumClientFormat(lgQtySum,ggQty.Decpoint,0)%>" 
	   .frm1.txtAmtSum.text		= "<%=UniNumClientFormat(lgAmtSum,ggAmtOfMoney.Decpoint,0)%>" 
	   .frm1.txtDiffSum.text	= "<%=UniNumClientFormat(lgDiffSum,ggAmtOfMoney.Decpoint,0)%>" 
	   
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
