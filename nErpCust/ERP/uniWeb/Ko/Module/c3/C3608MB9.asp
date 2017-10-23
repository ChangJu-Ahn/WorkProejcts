<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")      

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPlantNm
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

Dim lgTotProdtQty		'�ѻ��귮 
Dim lgStdMCost			'����(ǥ��)
Dim lgStdLCost			'�빫��(ǥ��)
Dim lgStdECost			'�� ��(ǥ��)
Dim lgActlMCost			'����(����)
Dim lgActlLCost			'�빫��(����)
Dim lgActlECost			'�� ��(����)
Dim lgStdCost			'ǥ�ؿ��� 
Dim lgActlCost			'�������� 
Dim lgCostDiff			'���̱ݾ� 
Dim lgMCostDiff			'�������̱ݾ� 
Dim lgLCostDiff			'�빫�����̱ݾ� 
Dim lgECostDiff			'�� �����̱ݾ� 
Dim lgUnitStdMCost		'������ ����(ǥ��)
Dim lgUnitStdLCost		'������ �빫��(ǥ��)
Dim lgUnitStdECost		'������ �� ��(ǥ��)
Dim lgUnitActlMCost		'������ ����(����)
Dim lgUnitActlLCost		'������ �빫��(����)
Dim lgUnitActlECost		'������ �� ��(����)
Dim lgUnitStdCost		'������ ǥ�ؿ��� 
Dim lgUnitActlCost		'������ �������� 
Dim lgUnitCostDiff		'������ ���̱ݾ� 
Dim lgUnitMCostDiff		'������ �������̱ݾ� 
Dim lgUnitLCostDiff		'������ �빫�����̱ݾ� 
Dim lgUnitECostDiff		'������ �� �����̱ݾ� 



'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

'   lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    
    Call FixUNISQLData()
    Call QueryData()


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------



    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "c3608ma102"		'tot production qty
	UNISqlId(1) = "c3608ma103"		'Actl Cost
	UNISqlId(2) = "c3608ma104"		'Std Cost
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,0)  = FilterVar(Request("txtYyyymm")   , "''", "S")               
	UNIValue(0,1)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               
	UNIValue(0,2)  = FilterVar(Request("txtItemCd")   , "''", "S")               
    
    UNIValue(1,0)  = FilterVar(Request("txtYyyymm")   , "''", "S")               
	UNIValue(1,1)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               
	UNIValue(1,2)  = FilterVar(Request("txtItemCd")   , "''", "S")               

	UNIValue(2,0)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               
	UNIValue(2,1)  = FilterVar(Request("txtItemCd")   , "''", "S")               
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
    
    

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
        
    If  rs0.EOF And rs0.BOF Then
        lgTotProdtQty = 0
    Else    
        lgTotProdtQty = Cdbl(rs0("qty"))
    End If
    
    If  rs1.EOF And rs1.BOF Then
		lgActlMCost= 0
		lgActlLCost= 0 
		lgActlECost= 0
    Else    
		lgActlMCost= Cdbl(rs1("m_cost"))
		lgActlLCost= Cdbl(rs1("l_cost"))
		lgActlECost= Cdbl(rs1("e_cost"))
    End If

    If  rs1.EOF And rs1.BOF Then
		lgUnitStdMCost = 0
		lgUnitStdLCost = 0
		lgUnitStdECost = 0
    Else    
		lgUnitStdMCost = Cdbl(rs2("m_cost"))
		lgUnitStdLCost = Cdbl(rs2("l_cost"))
		lgUnitStdECost = Cdbl(rs2("e_cost"))
    End If

	rs0.Close
    rs1.Close
    rs2.Close
    Set rs0 = Nothing
	Set rs1 = Nothing
	Set rs2 = Nothing
	
	lgStdMCost = lgTotProdtQty*lgUnitStdMCost
	lgStdLCost = lgTotProdtQty*lgUnitStdLCost
	lgStdECost = lgTotProdtQty*lgUnitStdECost

	if lgTotProdtQty = 0 then
		lgUnitActlMCost= 0
		lgUnitActlLCost= 0
		lgUnitActlECost= 0
	else
		lgUnitActlMCost= lgActlMCost / lgTotProdtQty
		lgUnitActlLCost= lgActlLCost / lgTotProdtQty
		lgUnitActlECost= lgActlECost / lgTotProdtQty
	end if

	lgStdCost  = lgStdMCost+lgStdLCost+lgStdECost
	lgActlCost = lgActlMCost+lgActlLCost+lgActlECost
	lgCostDiff = lgActlCost-lgStdCost

	lgMCostDiff= lgActlMCost-lgStdMCost
	lgLCostDiff= lgActlLCost-lgStdLCost
	lgECostDiff= lgActlECost-lgStdECost

	lgUnitStdCost=lgUnitStdMCost+lgUnitStdLCost+lgUnitStdECost
	lgUnitActlCost=lgUnitActlMCost+lgUnitActlLCost+lgUnitActlECost
	lgUnitCostDiff=lgUnitActlCost-lgUnitStdCost

	lgUnitMCostDiff=lgUnitActlMCost-lgUnitStdMCost
	lgUnitLCostDiff=lgUnitActlLCost-lgUnitStdLCost
	lgUnitECostDiff=lgUnitActlECost-lgUnitStdECost

End Sub

%>

<Script Language=vbscript>

	With parent.frm1
		.txtTotQty.text = "<%=UNINumClientFormat(lgTotProdtQty, ggQty.DecPoint, 0)%>"
		
		'�ѻ��귮 ���� 
		.txtStd_Mcost1.text = "<%=UNINumClientFormat(lgStdMCost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtReal_Mcost1.text= "<%=UNINumClientFormat(lgActlMCost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtDiff_Mcost1.text= "<%=UNINumClientFormat(lgMCostDiff, ggAmtOfMoney.DecPoint, 0)%>"
		
		.txtStd_Lcost1.text = "<%=UNINumClientFormat(lgStdLCost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtReal_Lcost1.text= "<%=UNINumClientFormat(lgActlLCost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtDiff_Lcost1.text= "<%=UNINumClientFormat(lgLCostDiff, ggAmtOfMoney.DecPoint, 0)%>"
		
		.txtStd_Ecost1.text = "<%=UNINumClientFormat(lgStdECost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtReal_Ecost1.text= "<%=UNINumClientFormat(lgActlECost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtDiff_Ecost1.text= "<%=UNINumClientFormat(lgECostDiff, ggAmtOfMoney.DecPoint, 0)%>"
		
		.txtStd_Sum1.text   = "<%=UNINumClientFormat(lgStdCost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtReal_Sum1.text  = "<%=UNINumClientFormat(lgActlCost, ggAmtOfMoney.DecPoint, 0)%>"
		.txtDiff_Sum1.text  = "<%=UNINumClientFormat(lgCostDiff, ggAmtOfMoney.DecPoint, 0)%>"
		
		'���� ���� 
		.txtStd_Mcost2.text = "<%=UNINumClientFormat(lgUnitStdMCost, ggUnitCost.DecPoint, 0)%>"
		.txtReal_Mcost2.text= "<%=UNINumClientFormat(lgUnitActlMCost, ggUnitCost.DecPoint, 0)%>"
		.txtDiff_Mcost2.text= "<%=UNINumClientFormat(lgUnitMCostDiff, ggUnitCost.DecPoint, 0)%>"
		
		.txtStd_Lcost2.text = "<%=UNINumClientFormat(lgUnitStdLCost, ggUnitCost.DecPoint, 0)%>"
		.txtReal_Lcost2.text= "<%=UNINumClientFormat(lgUnitActlLCost, ggUnitCost.DecPoint, 0)%>"
		.txtDiff_Lcost2.text= "<%=UNINumClientFormat(lgUnitLCostDiff, ggUnitCost.DecPoint, 0)%>"
		
		.txtStd_Ecost2.text = "<%=UNINumClientFormat(lgUnitStdECost, ggUnitCost.DecPoint, 0)%>"
		.txtReal_Ecost2.text= "<%=UNINumClientFormat(lgUnitActlECost, ggUnitCost.DecPoint, 0)%>"
		.txtDiff_Ecost2.text= "<%=UNINumClientFormat(lgUnitECostDiff, ggUnitCost.DecPoint, 0)%>"
		
		.txtStd_Sum2.text   = "<%=UNINumClientFormat(lgUnitStdCost, ggUnitCost.DecPoint, 0)%>"
		.txtReal_Sum2.text  = "<%=UNINumClientFormat(lgUnitActlCost, ggUnitCost.DecPoint, 0)%>"
		.txtDiff_Sum2.text  = "<%=UNINumClientFormat(lgUnitCostDiff, ggUnitCost.DecPoint, 0)%>"		
		
    End With   
       'Show multi spreadsheet data from this line
       
    Parent.DbQuery2Ok
	
</Script>	

