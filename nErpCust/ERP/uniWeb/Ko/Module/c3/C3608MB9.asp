<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")      

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPlantNm
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

Dim lgTotProdtQty		'총생산량 
Dim lgStdMCost			'재료비(표준)
Dim lgStdLCost			'노무비(표준)
Dim lgStdECost			'경 비(표준)
Dim lgActlMCost			'재료비(실제)
Dim lgActlLCost			'노무비(실제)
Dim lgActlECost			'경 비(실제)
Dim lgStdCost			'표준원가 
Dim lgActlCost			'실제원가 
Dim lgCostDiff			'차이금액 
Dim lgMCostDiff			'재료비차이금액 
Dim lgLCostDiff			'노무비차이금액 
Dim lgECostDiff			'경 비차이금액 
Dim lgUnitStdMCost		'단위당 재료비(표준)
Dim lgUnitStdLCost		'단위당 노무비(표준)
Dim lgUnitStdECost		'단위당 경 비(표준)
Dim lgUnitActlMCost		'단위당 재료비(실제)
Dim lgUnitActlLCost		'단위당 노무비(실제)
Dim lgUnitActlECost		'단위당 경 비(실제)
Dim lgUnitStdCost		'단위당 표준원가 
Dim lgUnitActlCost		'단위당 실제원가 
Dim lgUnitCostDiff		'단위당 차이금액 
Dim lgUnitMCostDiff		'단위당 재료비차이금액 
Dim lgUnitLCostDiff		'단위당 노무비차이금액 
Dim lgUnitECostDiff		'단위당 경 비차이금액 



'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

'   lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    
    Call FixUNISQLData()
    Call QueryData()


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------



    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "c3608ma102"		'tot production qty
	UNISqlId(1) = "c3608ma103"		'Actl Cost
	UNISqlId(2) = "c3608ma104"		'Std Cost
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	UNIValue(0,0)  = FilterVar(Request("txtYyyymm")   , "''", "S")               
	UNIValue(0,1)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               
	UNIValue(0,2)  = FilterVar(Request("txtItemCd")   , "''", "S")               
    
    UNIValue(1,0)  = FilterVar(Request("txtYyyymm")   , "''", "S")               
	UNIValue(1,1)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               
	UNIValue(1,2)  = FilterVar(Request("txtItemCd")   , "''", "S")               

	UNIValue(2,0)  = FilterVar(Request("txtPlantCd"),"''"       ,"S")               
	UNIValue(2,1)  = FilterVar(Request("txtItemCd")   , "''", "S")               
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    
    

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
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
		
		'총생산량 기준 
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
		
		'단위 기준 
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

