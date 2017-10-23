<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2214MB1
'*  4. Program Name         : ����ǰ���ǸŰ�ȹ��� 
'*  5. Program Desc         : ����ǰ���ǸŰ�ȹ��� 
'*  6. Comproxy List        : PS2G241.dll
'*  7. Modified date(First) : 2003/01/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang seong bae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
Call LoadBNumericFormatB("I","*","NOCOOKIE","MB") %>

<%													

On Error Resume Next														

Call HideStatusWnd

Const	C_SpType		= 0
Const	C_LocExpFlag	= 1
Const	C_FrSpPeriod	= 3
Const	C_ToSpPeriod	= 4
Const	C_SalesGrp		= 6
Const	C_SoldToParty	= 7
Const	C_ItemCd		= 8

Dim iStrMode
Dim iStrSvrData, iStrSvrData2, iStrNextKey
Dim iObjPS2G241
Dim iArrListOut			' Result of recordset.getrow(), it means iArrListOut is two dimension array (column, row)
Dim iArrListGroupOut	' Result of recordset.getrow(), it means iArrListGroupOut is two dimension array (column, row)
Dim iArrWhereIn, iArrWhereOut
Dim iLngRow
Dim iLngLastRow			' The last row number in the spread
Dim iLngSheetMaxRows	' Row numbers to be displayed in the spread.
Dim iLngErrorPosition

iStrMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case iStrMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    Err.Clear                                                                '��: Protect system from crashing

	iLngSheetMaxRows = CLng(100)
	
    Set iObjPS2G241 = Server.CreateObject("PS2G241.cListSSpItemByBp")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

    Call iObjPS2G241.ListRows (gStrGlobalCollection, iLngSheetMaxRows, Request("txtWhere"), Request("lgStrPrevKey"), _
						  iArrListOut, iArrListGroupOut, iArrWhereOut)
	
	If CheckSYSTEMError(Err,True) = True Then
	   Set iObjPS2G241 = Nothing		                                                 '��: Unload Comproxy DLL
	   Response.End 
	End If

    Set iObjPS2G241 = Nothing
    
    ' Check Query Condition
    If Request("lgStrPrevKey") = "" Then
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
        
		iArrWhereIn = Split(Request("txtWhere"), gColSep)
		' �����׷� 
		If iArrWhereIn(C_SalesGrp) = iArrWhereOut(0, C_SalesGrp) Then
			Response.Write "Parent.frm1.txtConSalesGrpNm.value = """ & ConvSPChars(iArrWhereOut(1, C_SalesGrp)) & """" & vbCr
			' �Է����� ����(�ǸŰ�ȹ����, �����׷�, �ŷ�����)
			Response.Write "Parent.frm1.cboSpType.value = """ & iArrWhereOut(0, C_SpType) & """" & vbCr
			Response.Write "Parent.frm1.txtSalesGrp.value = """ & iArrWhereOut(0, C_SalesGrp) & """" & vbCr
			Response.Write "Parent.frm1.txtSalesGrpNm.value = """ & ConvSPChars(iArrWhereOut(1, C_SalesGrp)) & """" & vbCr
			Response.Write "Parent.frm1.cboLocExpFlag.value = """ & iArrWhereOut(0, C_LocExpFlag) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", ""�����׷�"", ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConSalesGrpNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConSalesGrp.focus " & vbCr   
			Response.Write "Call parent.SetToolbar(""11000000000011"") " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End		
		End If
		' ��ȹ�Ⱓ(����)			
		If iArrWhereIn(C_FrSpPeriod) = iArrWhereOut(0, C_FrSpPeriod) Then
			Response.Write "Parent.frm1.txtConFrSpPeriodDesc.value = """ & ConvSPChars(iArrWhereOut(1,C_FrSpPeriod)) & """" & vbCr
		End If
		' ��ȹ�Ⱓ(��)			
		If iArrWhereIn(C_ToSpPeriod) = iArrWhereOut(0, C_ToSpPeriod) Then
			Response.Write "Parent.frm1.txtConToSpPeriodDesc.value = """ & ConvSPChars(iArrWhereOut(1,C_ToSpPeriod)) & """" & vbCr
		End If
		' �ŷ�ó			
		If iArrWhereIn(C_SoldToParty) = iArrWhereOut(0, C_SoldToParty) Then
			Response.Write "Parent.frm1.txtConSoldToPartyNm.value = """ & ConvSPChars(iArrWhereOut(1,C_SoldToParty)) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", ""�ŷ�ó"", ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConSoldToPartyNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConSoldToParty.focus " & vbCr   
			Response.Write "Call parent.SetToolbar(""11000000000011"") " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End		
		End If
		' ǰ�� 
		If iArrWhereIn(C_ItemCd) = iArrWhereOut(0, C_ItemCd) Then
			Response.Write "Parent.frm1.txtConItemNm.value = """ & ConvSPChars(iArrWhereOut(1,C_ItemCd)) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", ""ǰ��"", ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConItemNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConItemCd.focus " & vbCr   
			Response.Write "Call parent.SetToolbar(""11000000000011"") " & vbCr   
			Response.Write "</SCRIPT> " & VbCr
			Response.End		
		End If
		
		' ��ϵ� �ڷᰡ �������� �ʽ��ϴ�.
		If UBound(iArrListOut) < 0 Then
			Response.Write "Call Parent.DisplayMsgBox(""202258"", ""X"", ""X"", ""X"")" & vbCr
			Response.Write "Call parent.SetToolbar(""11001111001111"") " & vbCr
			Response.Write "Call parent.GetLastCfmSpPeriod() " & vbCr
			Response.Write "Call parent.GetSpConfig() " & vbCr
			Response.Write "parent.frm1.cboConSpType.focus " & vbCr   
			Response.Write "</SCRIPT> " & VbCr
			Response.End		
		Else
			Response.Write "</SCRIPT> " & VbCr
		End If
	End If
    
	'------------------------
	'Result data display area
	'------------------------ 
	iLngLastRow = CLng(Request("txtLastRow"))

	' Set Next key
	If Ubound(iArrListOut,2) = iLngSheetMaxRows Then
		'��ȹ�Ⱓ + �ŷ�ó + ǰ�� 
		iStrNextKey = iArrListOut(3, iLngSheetMaxRows) & gColSep & iArrListOut(7, iLngSheetMaxRows) & gColSep & iArrListOut(11, iLngSheetMaxRows)
		iLngSheetMaxRows  = iLngSheetMaxRows - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(iArrListOut,2)
	End If

	' Spread1
   	For iLngRow = 0 To iLngSheetMaxRows
   	''''' �׸��ʵ� list'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' " SPI.SP_TYPE , SPI.LOC_EXP_FLAG, SPI.SO_BILL_FLAG, SPI.SP_PERIOD, "
	' " SPP.SP_PERIOD_DESC, SPI.SP_SEQ, SPI.SALES_GRP, SPI.SOLD_TO_PARTY,"
	' " BP.BP_NM, SPI.CUR, SPI.XCHG_RATE, SPI.ITEM_CD,"
	' " IT.ITEM_NM, SPI.QTY, SPI.UNIT, SPI.PRICE,"
	' " SPI.AMT, SPI.AMT_LOC, SPI.CFM_FLAG, SPI.DISTR_FLAG "
   '
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(0,iLngRow))			' ��ȹ����	
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(1,iLngRow))			' �ŷ�����(����/����)
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(2,iLngRow))			' ���ָ��ⱸ�� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(3,iLngRow))			' ��ȹ�Ⱓ 
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(4,iLngRow))			' ��ȹ�Ⱓ ���� 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(5,iLngRow), 0, 0)	' ��ȹ���� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(6,iLngRow))			' �����׷� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(7,iLngRow))			' �ŷ�ó 
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(8,iLngRow))			' �ŷ�ó�� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(9,iLngRow))			' ȭ�� 
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(10,iLngRow), ggExchRate.DecPoint, 0)			' ȯ�� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(11,iLngRow))			' ǰ�� 
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(12,iLngRow))			' ǰ��� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(13,iLngRow))			' �԰� 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(14,iLngRow), ggQty.DecPoint, 0)	' ���� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(15,iLngRow))			' ���� 
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & UNIConvNumDBToCompanyByCurrency(iArrListOut(16,iLngRow),iArrListOut(9,iLngRow),ggUnitCostNo, "X" , "X")		' �ܰ� 
   		iStrSvrData = iStrSvrData & gColSep & UNIConvNumDBToCompanyByCurrency(iArrListOut(17,iLngRow),iArrListOut(9,iLngRow),ggAmtOfMoneyNo, "X" , "X")	' �ݾ� 
   		iStrSvrData = iStrSvrData & gColSep & UNIConvNumDBToCompanyByCurrency(iArrListOut(18,iLngRow),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo  , "X")
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(19,iLngRow))			' Ȯ������ 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(20,iLngRow))			' ��п��� 
   		iStrSvrData = iStrSvrData & gColSep	& UNIDateClientFormat(iArrListOut(21,iLngRow))	' ������ 
   		iStrSvrData = iStrSvrData & gColSep	& UNIDateClientFormat(iArrListOut(22,iLngRow))	' ������ 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(23,iLngRow))			' �� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(24,iLngRow))			' �� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(25,iLngRow))			' ȯ��������	
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(14,iLngRow), ggQty.DecPoint, 0)	' ���� 
   		iStrSvrData = iStrSvrData & gColSep & UNIConvNumDBToCompanyByCurrency(iArrListOut(17,iLngRow),iArrListOut(9,iLngRow),ggAmtOfMoneyNo, "X" , "X")	' �ݾ� 
   		iStrSvrData = iStrSvrData & gColSep & iLngLastRow + iLngRow 
   		iStrSvrData = iStrSvrData & gColSep & gRowSep
   	Next
    
    ' Spread2
    IF Request("lgStrPrevKey") = "" Then
	   	For iLngRow = 0 To Ubound(iArrListGroupOut,2)
   	''''' �׸��ʵ� list'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' " SPI.SP_PERIOD, SPP.SP_PERIOD_DESC, SPI.SOLD_TO_PARTY, BP.BP_NM, SPI.QTY, SPI.CUR, SPI.AMT, SPP.FROM_DT, SPP.TO_DT, SPP.SP_WEEK "

	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(0,iLngRow))			' ��ȹ�Ⱓ	
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(1,iLngRow))			' ��ȹ�Ⱓ���� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(2,iLngRow))			' �ŷ�ó 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(3,iLngRow))			' �ŷ�ó�� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & UNINumClientFormat(iArrListGroupOut(4,iLngRow), ggQty.DecPoint, 0)	' ���� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(5,iLngRow))			' ȭ�� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & UNIConvNumDBToCompanyByCurrency(iArrListGroupOut(6,iLngRow),iArrListGroupOut(5,iLngRow),ggAmtOfMoneyNo, "X" , "X")	' �ݾ� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & iLngRow 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & gRowSep
	   	Next
	End If

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
        
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData " & vbCr
    Response.Write  "Parent.frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write  "Parent.ggoSpread.SSShowDataByClip   """ & iStrSvrData & """ ,""F""" & vbCr
    Response.Write  "Call Parent.FormatSpreadCellByCurrency(parent.lgLngStartRow, parent.frm1.vspdData.MaxRows, ""Q"")" & vbCr
    
    If Request("lgStrPrevKey") = "" Then
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData2 " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip """ & iStrSvrData2 & """ ,""F""" & vbCr
    Response.Write  "Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2," & -1 & "," & -1  & ",Parent.C_Cur2,Parent.C_TotAmt,""A"" ,""I"",""X"",""X"")" & vbCr
	End If
	
    Response.Write " Parent.lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr  
    Response.Write " Parent.DbQueryOk" & vbCr   
	Response.Write  "Parent.frm1.vspdData.Redraw = True  "       & vbCr      
	Response.Write "</SCRIPT> "		

Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
    Err.Clear																		'��: Protect system from crashing

    Set iObjPS2G241 = Server.CreateObject("PS2G241.cMaintSSpItemByBp")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	Call iObjPS2G241.Maintain (gStrGlobalCollection, Trim(Request("txtSpreadIns")), Trim(Request("txtSpreadUpd")), _
								Trim(Request("txtSpreadDel")), iLngErrorPosition)
	
	If CheckSYSTEMError2(Err, True, iLngErrorPosition & "��","","","","") = True Then
       Set iObjPS2G241 = Nothing
       
		Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
		Response.Write " Call Parent.SubSetErrPos(" & iLngErrorPosition & ")" & vbCr
		Response.Write "</SCRIPT> "		
       
	   Response.End 
	End If

    Set iObjPS2G241 = Nothing	
    
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> " 													'��: Row �� ���� 
    
End Select
%>
