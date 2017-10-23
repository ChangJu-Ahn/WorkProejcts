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
Dim rs0, rs1, rs2,rs3,rs4,rs5,rs6,rs7,rs8,rs9							'DBAgent Parameter ���� 


Dim	txtYyyymm
Dim	txtBizUnitCd
Dim	txtCostCd
Dim	txtSalesOrg
Dim	txtSalesGrp
Dim	txtBpCd
Dim txtSoType
Dim txtItemAcct
Dim txtItemGroupCd


Dim	txtPrevCostCd
Dim	txtPrevSalesGrp
Dim	txtPrevBpCd
Dim txtPrevSoType
Dim txtPrevItemCd


Dim lgSalesAmtSum
Dim lgCostAmtSum
Dim lgProfitSum
Dim lgBizUnitNm
Dim lgCostNm
Dim lgSalesOrgNm
Dim lgSalesGrpNm
Dim lgBpNm
Dim lgSoTypeNm
Dim lgItemAcctNm
Dim lgItemGroupNm

Dim lgDataExist
Dim lgCodeCond
Dim lgCodeCond1
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
	txtBizUnitCd		= Trim(Request("txtBizUnitCd"))
	txtCostCd			= Trim(Request("txtCostCd"))
	txtSalesOrg			= Trim(Request("txtSalesOrg"))
	txtSalesGrp			= Trim(Request("txtSalesGrp"))
	txtBpCd				= Trim(Request("txtBpCd"))
	txtSoType			= Trim(Request("txtSoType"))
	txtItemAcct			= Trim(Request("txtItemAcct"))
	txtItemGroupCd		= Trim(Request("txtItemGroupCd"))

	txtPrevCostCd		= Trim(Request("txtPrevCostCd"))
	txtPrevSalesGrp		= Trim(Request("txtPrevSalesGrp"))
	txtPrevBpCd			= Trim(Request("txtPrevBpCd"))
	txtPrevSoType		= Trim(Request("txtPrevSoType"))
	txtPrevItemCd		= Trim(Request("txtPrevItemCd"))


	
    lgPageNo       = Cint(Trim(Request("lgPageNo")))    
'   lgMaxCount     = Cint(Trim(Request("lgMaxCount")))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgDataExist    = "No"
	lgErrorStatus  = "No" 
	lgCodeCond = ""

	IF txtBizUnitCd = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and c.biz_unit_cd = " & FilterVar(txtBizUnitCd, "''", "S")
	END IF
	
	IF txtCostCd = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and c.cost_cd = " & FilterVar(txtCostCd, "''", "S")
	END IF

	IF txtSalesOrg = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and c.sales_org = " & FilterVar(txtSalesOrg, "''", "S")
	END IF


	IF txtSalesGrp = "" Then
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and c.sales_grp = " & FilterVar(txtSalesGrp, "''", "S")
	END IF

	
	IF txtBpCd = "" Then
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and c.bp_cd = " & FilterVar(txtBpCd, "''", "S")
	END IF
	
	
	IF txtSoType = "" Then
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and c.deal_type = " & FilterVar(txtSoType, "''", "S")
	END IF


	IF txtItemAcct = "" Then
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and h.item_acct = " & FilterVar(txtItemAcct, "''", "S")
	END IF
	
	IF txtItemGroupCd = "" Then
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and c.item_group_Cd = " & FilterVar(txtItemGroupCd, "''", "S")
	END IF

	lgCodeCond1 = lgCodeCond
	lgCodeCond = lgCodeCond & " order by c.cost_cd,c.sales_grp,c.bp_cd,c.deal_type,c.item_cd,c.gain_cd"
	
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
        
        IF txtPrevCostCd <> Trim(ConvSPChars(rs0(1))) or txtPrevSalesGrp <> Trim(ConvSPChars(rs0(4))) or txtPrevBpCd <> Trim(ConvSPChars(rs0(6))) _ 
			or txtPrevSoType <> Trim(ConvSPChars(rs0(8))) or txtPrevItemCd <> Trim(ConvSPChars(rs0(11))) Then
					
			txtPrevCostCd	= Trim(ConvSPChars(rs0(1)))
			txtPrevSalesGrp	= Trim(ConvSPChars(rs0(4)))
			txtPrevBpCd		= Trim(ConvSPChars(rs0(6)))
			txtPrevSoType	= Trim(ConvSPChars(rs0(8)))
			txtPrevItemCd	= Trim(ConvSPChars(rs0(11)))
			
			
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		'����� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'C/C
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))		'C/C�� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))		'�������� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))		'�����׷� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(5))		'�����׷�� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(6))		'�ŷ�ó�ڵ� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(7))		'�ŷ�ó�� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(8))		'�ŷ����� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(9))		'�ŷ������� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(10))		'�𵨸� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(11))		'ǰ�� 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(12))		'ǰ��� 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(13), ggQty.DecPoint, 0) '���� 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(14), ggAmtOfMoney.DecPoint, 0) '����ݾ� 
		ELSE
			lgstrData = lgstrData & Chr(11) & ""		'����� 
			lgstrData = lgstrData & Chr(11) & ""		'C/C
			lgstrData = lgstrData & Chr(11) & ""		'C/C�� 
			lgstrData = lgstrData & Chr(11) & ""		'�������� 
			lgstrData = lgstrData & Chr(11) & ""		'�����׷� 
			lgstrData = lgstrData & Chr(11) & ""		'�����׷�� 
			lgstrData = lgstrData & Chr(11) & ""		'�ŷ�ó�ڵ� 
			lgstrData = lgstrData & Chr(11) & ""		'�ŷ�ó�� 
			lgstrData = lgstrData & Chr(11) & ""		'�ŷ����� 
			lgstrData = lgstrData & Chr(11) & ""		'�ŷ������� 
			lgstrData = lgstrData & Chr(11) & ""		'�𵨸� 
			lgstrData = lgstrData & Chr(11) & ""		'ǰ�� 
			lgstrData = lgstrData & Chr(11) & ""		'ǰ��� 
			lgstrData = lgstrData & Chr(11) & "" '���� 
			lgstrData = lgstrData & Chr(11) & "" '����ݾ� 
		END IF
		
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(15))		'��������׸� 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(16))		'��������׸�� 
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(17), ggAmtOfMoney.DecPoint, 0) '������� 
					
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

    Redim UNISqlId(9)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------


    Redim UNIValue(9,5)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "GB003MA01"
    UNISqlId(1) = "GB003MA02"					
    UNISqlId(2) = "commonqry"
    UNISqlId(3) = "commonqry"	
    UNISqlId(4) = "commonqry"	
    UNISqlId(5) = "commonqry"	
    UNISqlId(6) = "commonqry"	
    UNISqlId(7) = "commonqry"	
    UNISqlId(8) = "commonqry"	
    UNISqlId(9) = "commonqry"	
    					
 
    UNIValue(0,0) = FilterVar(txtYYYYMM, "''", "S")
    UNIValue(0,1) = FilterVar(txtYYYYMM, "''", "S")
    UNIValue(0,2) = FilterVar(txtYYYYMM, "''", "S")
    UNIValue(0,3) = FilterVar(txtYYYYMM, "''", "S")
    UNIValue(0,4) = FilterVar(txtYYYYMM, "''", "S")
    UNIValue(0,5) = lgCodeCond
    
    	
    UNIValue(1,0) = FilterVar(txtYYYYMM, "''", "S")
	UNIValue(1,1) = lgCodeCond1
	UNIValue(1,2) = FilterVar(txtYYYYMM, "''", "S")
	UNIValue(1,3) = lgCodeCond1	

   
	UNIValue(2,0)  = "select biz_unit_nm from b_biz_unit where biz_unit_cd = " & FilterVar(txtBizUnitCd, "''", "S") 
	UNIValue(3,0)  = "SELECT cost_nm from b_cost_center where cost_cd = " & FilterVar(txtCostCd, "''", "S") 
	UNIValue(4,0)  = "SELECT sales_org_nm from b_sales_org where sales_org = " & FilterVar(txtSalesOrg, "''", "S") 
	UNIValue(5,0)  = "SELECT sales_grp_nm from b_sales_grp where sales_grp = " & FilterVar(txtSalesGrp, "''", "S") 
	UNIValue(6,0)  = "SELECT bp_nm from b_biz_partner where bp_cd = " & FilterVar(txtBpCd, "''", "S") 
	UNIValue(7,0)  = "SELECT so_type_nm from (select so_type,so_type_nm from s_so_type_config union all select minor_cd as so_type,minor_nm as so_type_nm from b_minor where major_cd = " & FilterVar("G1025", "''", "S") & " ) a where so_type = " & FilterVar(txtSoType, "''", "S") 
	UNIValue(8,0)  = "SELECT minor_nm from b_minor where major_Cd = " & FilterVar("P1001", "''", "S") & " and minor_Cd = " & FilterVar(txtItemAcct, "''", "S") 
	UNIValue(9,0)  = "SELECT item_group_nm from b_item_group where item_group_cd = " & FilterVar(txtItemGroupCd, "''", "S") 
		
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2,rs3,rs4,rs5,rs6,rs7,rs8,rs9)
    


    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
  
        
    IF NOT (rs1.EOF or rs1.BOF) then
		lgSalesAmtSum	= rs1(0)
		lgCostAmtSum	= rs1(1)
		lgProfitSum		= rs1(2)
		
	ELSE
		lgSalesAmtSum	= 0
		lgCostAmtSum	= 0
		lgProfitSum		= 0		
	End if
	
	
    rs1.Close
    Set rs1 = Nothing 
    

    IF NOT (rs2.EOF or rs2.BOF) then
		lgBizUnitNm = rs2(0)				
	ELSE
		lgBizUnitNm = ""
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
		lgSalesOrgNm = rs4(0)			
	ELSE
		lgSalesOrgNm = ""
	End if
    rs4.Close
    Set rs4 = Nothing 
    
    IF NOT (rs5.EOF or rs5.BOF) then
		lgSalesGrpNm = rs5(0)			
	ELSE
		lgSalesGrpNm = ""
	End if
    rs5.Close
    Set rs5 = Nothing
        
    IF NOT (rs6.EOF or rs6.BOF) then
		lgBpNm = rs6(0)			
	ELSE
		lgBpNm = ""
	End if
    rs6.Close
    Set rs6 = Nothing  
    
        
    IF NOT (rs7.EOF or rs7.BOF) then
		lgSoTypeNm = rs7(0)				
	ELSE
		lgSoTypeNm = ""
	End if
    rs7.Close
    Set rs7 = Nothing      


    IF NOT (rs8.EOF or rs8.BOF) then
		lgItemAcctNm = rs8(0)				
	ELSE
		lgItemAcctNm = ""
	End if
    rs8.Close
    Set rs8 = Nothing   
    
    
    IF NOT (rs9.EOF or rs9.BOF) then
		lgItemGroupNm = rs9(0)				
	ELSE
		lgItemGroupNm = ""
	End if
    rs9.Close
    Set rs9 = Nothing   
		    
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
	   .frm1.hBizUnitCd.value		=  "<%=ConvSPChars(txtBizUnitCd)%>"
	   .frm1.hCostCd.value		=  "<%=ConvSPChars(txtCostCd)%>"	 	
	   .frm1.hSalesOrg.value	=  "<%=ConvSPChars(txtSalesOrg)%>"	 	
	   .frm1.hSalesGrp.value	=  "<%=ConvSPChars(txtSalesGrp)%>"	 		 	
	   .frm1.hBpCd.value	=  "<%=ConvSPChars(txtBpCd)%>"	 		 	
	   .frm1.hSoTypeCd.value		=  "<%=ConvSPChars(txtSoType)%>"
	   .frm1.hItemAcct.value		=  "<%=ConvSPChars(txtItemAcct)%>"
	   .frm1.hItemGroupCd.value		=  "<%=ConvSPChars(txtItemGroupCd)%>"	   	 		 	
	   
	   .frm1.txtBizUnitNm.value	= "<%=ConvSPChars(lgBizUnitNm)%>"
	   .frm1.txtCostNm.value	= "<%=ConvSPChars(lgCostNm)%>"
	   .frm1.txtSalesOrgNm.value = "<%=ConvSPChars(lgSalesOrgNm)%>"
	   .frm1.txtSalesGrpNm.value	= "<%=ConvSPChars(lgSalesGrpNm)%>"
	   .frm1.txtBpNm.value = "<%=ConvSPChars(lgBpNm)%>"
	   .frm1.txtSoTypeNm.value	= "<%=ConvSPChars(lgSoTypeNm)%>"
	   .frm1.txtItemAcctNm.value = "<%=ConvSPChars(lgItemAcctNm)%>"
	   .frm1.txtItemGroupNm.value	= "<%=ConvSPChars(lgItemGroupNm)%>"
		
		.frm1.txtPrevItemCd.value	= "<%=ConvSPChars(txtPrevItemCd)%>"
	   .frm1.txtPrevCostCd.value	= "<%=ConvSPChars(txtPrevCostCd)%>"
	   .frm1.txtPrevSalesGrp.value = "<%=ConvSPChars(txtPrevSalesGrp)%>"
	   .frm1.txtPrevBpCd.value	= "<%=ConvSPChars(txtPrevBpCd)%>"
	   .frm1.txtPrevSoType.value = "<%=ConvSPChars(txtPrevSoType)%>"
	   
	   
	   .frm1.txtSalesAmtSum.text		= "<%=UniNumClientFormat(lgSalesAmtSum,ggAmtOfMoney.Decpoint,0)%>" 
	   .frm1.txtCostAmtSum.text		= "<%=UniNumClientFormat(lgCostAmtSum,ggAmtOfMoney.Decpoint,0)%>" 
	   .frm1.txtProfitSum.text	= "<%=UniNumClientFormat(lgProfitSum,ggAmtOfMoney.Decpoint,0)%>" 
	   
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
