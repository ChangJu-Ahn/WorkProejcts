<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2216MB1
'*  4. Program Name         : ���庰�Ϻ�ǰ���ǸŰ�ȹ���� 
'*  5. Program Desc         : ���庰�Ϻ�ǰ���ǸŰ�ȹ���� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seong Bae
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

Const	C_PlantCd		= 1
Const	C_SalesGrp		= 2
Const	C_ItemCd		= 3
Const	C_SoldToParty	= 4
Const	C_LocExpFlag	= 5

Dim iStrMode
Dim iStrSvrData, iStrSvrData2, iStrNextKey
Dim iObjPS2G261
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

	iLngSheetMaxRows = CLng(50)
	
    Set iObjPS2G261 = Server.CreateObject("PS2G261.cListSSpItemByPlantDaily")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

    Call iObjPS2G261.ListRows (gStrGlobalCollection, iLngSheetMaxRows, Request("txtWhere"), Request("lgStrPrevKey"), _
						  iArrListOut, iArrListGroupOut, iArrWhereOut)
	
	If CheckSYSTEMError(Err,True) = True Then
       Set iObjPS2G261 = Nothing		                                                 '��: Unload Comproxy DLL
       Response.End 
    End If   

    Set iObjPS2G261 = Nothing
    
    ' Check Query Condition
    If Request("lgStrPrevKey") = "" Then
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
		iArrWhereIn = Split(Request("txtWhere"), gColSep)
		' ���� 
		If iArrWhereIn(C_PlantCd) = iArrWhereOut(0, C_PlantCd) Then
			Response.Write "Parent.frm1.txtConPlantNm.value = """ & ConvSPChars(iArrWhereOut(1, C_PlantCd)) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", ""����"", ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConPlantNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConPlantCd.focus " & vbCr   
			Response.Write "Call parent.SetToolbar(""11000000000011"") " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If

		' �����׷� 
		If iArrWhereIn(C_SalesGrp) = iArrWhereOut(0, C_SalesGrp) Then
			Response.Write "Parent.frm1.txtConSalesGrpNm.value = """ & ConvSPChars(iArrWhereOut(1, C_SalesGrp)) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", ""�����׷�"", ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConSalesGrpNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConSalesGrp.focus " & vbCr   
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

		' ��ϵ� �ڷᰡ �������� �ʽ��ϴ�.
		If UBound(iArrListOut) < 0 Then
			Response.Write "Call Parent.DisplayMsgBox(""211210"", ""X"", ""X"", ""X"")" & vbCr
			Response.Write "Call parent.SetToolbar(""11000000000011"") " & vbCr   
			Response.Write "Call parent.SetFocusToDocument(""M"") " & vbCr   
			Response.Write "parent.frm1.txtConFromDt.focus " & vbCr   
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
		'��ȹ�Ⱓ, ����, �����׷�, ǰ��, �ŷ�ó, �ŷ����� 
		iStrNextKey = iArrListOut(0, iLngSheetMaxRows) & gColSep & iArrListOut(1, iLngSheetMaxRows) & gColSep & iArrListOut(3, iLngSheetMaxRows) & gColSep & _
					  iArrListOut(9, iLngSheetMaxRows) & gColSep & iArrListOut(7, iLngSheetMaxRows) & gColSep & iArrListOut(5, iLngSheetMaxRows)
		iLngSheetMaxRows  = iLngSheetMaxRows - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(iArrListOut,2)
	End If

    ' SIP.SP_DT(0), SIP.PLANT_CD(1), PT.PLANT_NM(2), SIP.SALES_GRP(3), SG.SALES_GRP_NM(4),
    ' SIP.LOC_EXP_FLAG(5), MN.MINOR_NM(6), SIP.SOLD_TO_PARTY(7), BP.BP_NM(8),
    ' SIP.ITEM_CD(9), IT.ITEM_NM(10), SIP.QTY(11), SIP.UNIT(12), SIP.QTY_ORDER_UNIT_MFG(13), SIP.ORDER_UNIT_MFG(14),
    ' SIP.CFM_FLAG(15), SP.SP_PERIOD(16), SP.SP_PERIOD_DESC(17), SP.SP_MONTH(18), SP.SP_WEEK(19)
	' Spread1
   	For iLngRow = 0 To iLngSheetMaxRows
   		iStrSvrData = iStrSvrData & gColSep	& UNIDateClientFormat(iArrListOut(0,iLngRow))	' ��ȹ�� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(1,iLngRow))			' ���� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(2,iLngRow))			' ����� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(3,iLngRow))			' �����׷� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(4,iLngRow))			' �����׷�� 
   		iStrSvrData = iStrSvrData & gColSep & iArrListOut(5,iLngRow)						' �ŷ����� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(6,iLngRow))			' �ŷ����и� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(7,iLngRow))			' �ŷ�ó 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(8,iLngRow))			' �ŷ�ó�� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(9,iLngRow))			' ǰ�� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(10,iLngRow))			' ǰ��� 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(11,iLngRow), ggQty.DecPoint, 0)	' ���� 
   		iStrSvrData = iStrSvrData & gColSep & iArrListOut(12,iLngRow)						' ���� 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(13,iLngRow), ggQty.DecPoint, 0)	' ����������� 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(14,iLngRow))			' ������� 
   		iStrSvrData = iStrSvrData & gColSep & iArrListOut(15,iLngRow)						' Ȯ������ 
   		iStrSvrData = iStrSvrData & gColSep & iArrListOut(16,iLngRow)						' ��ȹ�Ⱓ 
   		iStrSvrData = iStrSvrData & gColSep & ConvSPChars(iArrListOut(17,iLngRow))			' ��ȹ������ 
   		iStrSvrData = iStrSvrData & gColSep & iArrListOut(18,iLngRow)						' �� 
   		iStrSvrData = iStrSvrData & gColSep & iArrListOut(19,iLngRow)						' �� 
   		iStrSvrData = iStrSvrData & gColSep	
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(11,iLngRow), ggQty.DecPoint, 0)	' ���� 
   		iStrSvrData = iStrSvrData & gColSep & UNINumClientFormat(iArrListOut(13,iLngRow), ggQty.DecPoint, 0)	' ���� 
   		iStrSvrData = iStrSvrData & gColSep & iLngLastRow + iLngRow 
   		iStrSvrData = iStrSvrData & gColSep & gRowSep
   	Next
    
    ' Spread2
    IF Request("lgStrPrevKey") = "" Then
	   	For iLngRow = 0 To Ubound(iArrListGroupOut,2)
	'T.SP_PERIOD(0), T.SP_PERIOD_DESC(1), T.PLANT_CD(2), PT.PLANT_NM(3), T.TOT_QTY(4), T.UNIT(5), T.TOT_QTY_MFG(6), ORDER_UNIT_MFG(7) 

	   		iStrSvrData2 = iStrSvrData2 & gColSep & iArrListGroupOut(0,iLngRow)							' ��ȹ�Ⱓ	
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(1,iLngRow))			' ��ȹ�Ⱓ���� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(2,iLngRow))			' ���� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & ConvSPChars(iArrListGroupOut(3,iLngRow))			' ����� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & UNINumClientFormat(iArrListGroupOut(4,iLngRow), ggQty.DecPoint, 0)	' ���� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & iArrListGroupOut(5,iLngRow)							' ���� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & UNINumClientFormat(iArrListGroupOut(6,iLngRow), ggQty.DecPoint, 0)	' ����������� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & iArrListGroupOut(7,iLngRow)							' ������� 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & iLngRow 
	   		iStrSvrData2 = iStrSvrData2 & gColSep & gRowSep
	   	Next
	End If

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
        
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip """ & iStrSvrData & """" & vbCr
    
    If Request("lgStrPrevKey") = "" Then
    Response.Write " Parent.ggoSpread.Source = Parent.frm1.vspdData2 " & vbCr
    Response.Write " Parent.ggoSpread.SSShowDataByClip """ & iStrSvrData2 & """" & vbCr
	End If
	
    Response.Write " Parent.lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr  
    Response.Write " Parent.DbQueryOk" & vbCr   
	Response.Write "</SCRIPT> "		

Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
    Err.Clear																		'��: Protect system from crashing

    Set iObjPS2G261 = Server.CreateObject("PS2G261.cMaintSSpItemByPlantDaily")  
    
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If
	
	Call iObjPS2G261.Maintain (gStrGlobalCollection, Trim(Request("txtSpreadIns")), Trim(Request("txtSpreadUpd")), _
								Trim(Request("txtSpreadDel")), iLngErrorPosition)
	
	If CheckSYSTEMError2(Err, True, iLngErrorPosition & "��","","","","") = True Then
       Set iObjPS2G261 = Nothing
       
		Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
		Response.Write " Call Parent.SubSetErrPos(" & iLngErrorPosition & ")" & vbCr
		Response.Write "</SCRIPT> "		
       
	   Response.End 
	End If

    Set iObjPS2G261 = Nothing	
    
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> " 													'��: Row �� ���� 
    
End Select
%>
