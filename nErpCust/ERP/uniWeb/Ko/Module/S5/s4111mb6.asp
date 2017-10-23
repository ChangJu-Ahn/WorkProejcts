<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S4111MB6
'*  4. Program Name         : �ϰ����ó�� 
'*  5. Program Desc         :
'*  6. DLL List				: PS5G116
'*  7. Modified date(First) : 2003/07/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd ȭ�� layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� layout
'*                            -2001/12/19 : Date ǥ������ 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

On Error Resume Next									

Call HideStatusWnd

Dim iStrMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim iArrCols, iArrRows 

iStrMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Select Case iStrMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	
	Dim iStrNextKey							' ���� �� 
	Dim iLngLastRow							' ���� �׸����� �ִ�Row
	Dim iStrPostFlag
	Dim iBlnInitQuery						' ���� ��ȸ���� 
	Dim iObjPS5G116

	Dim iLngRow, iLngSheetMaxRows
	Dim iArrWhereIn, iArrWhereOut, iArrRsOut

	Const C_PS5G116_PLANT_FOR_QUERY = 0              ' Plant
	Const C_PS5G116_FR_PROMISE_DT_FOR_QUERY = 1      ' Promise date(G/I) or Actual G/I date(Cancel G/I)
	Const C_PS5G116_TO_PROMISE_DT_FOR_QUERY = 2      ' Promise date(G/I) or Actual G/I date(Cancel G/I)
	Const C_PS5G116_MOVE_TYPE_FOR_QUERY = 3          ' Movement type
	Const C_PS5G116_SHIP_TO_PARTY_FOR_QUERY = 4      ' Ship to party

	Dim C_SHEETMAXROWS_D				' �ѹ��� Query�� Row�� 

	If Request("txtBatchQuery") = "Y" Then
		C_SHEETMAXROWS_D = -1			' ��ȸ���ǿ� �ش�Ǵ� ��� Row�� ��ȯ�Ѵ�.
	Else
		C_SHEETMAXROWS_D = 100
	End If
	
	'---------------------------------------------
    'next key���� �Ѱ��ش�.
    '---------------------------------------------
	iStrNextKey = Trim(Request("lgStrPrevKey"))

    '---------------------------------------------
    'Data manipulate  area(import view match)
    '---------------------------------------------
    Redim iArrWhereIn(C_PS5G116_SHIP_TO_PARTY_FOR_QUERY)
	iArrWhereIn(C_PS5G116_PLANT_FOR_QUERY)				= Trim(Request("txtConPlant"))
	iArrWhereIn(C_PS5G116_FR_PROMISE_DT_FOR_QUERY)		= UNIConvDate(Request("txtConFromDt"))
	iArrWhereIn(C_PS5G116_TO_PROMISE_DT_FOR_QUERY)		= UNIConvDate(Request("txtConToDt"))
	iArrWhereIn(C_PS5G116_MOVE_TYPE_FOR_QUERY)			= Trim(Request("txtConDnType"))
	iArrWhereIn(C_PS5G116_SHIP_TO_PARTY_FOR_QUERY)		= Trim(Request("txtConShipToParty"))
	    
	iStrPostFlag = Request("txtConPostFlag")		' Ȯ��(Y)/��ҿ���(N)

    Set iObjPS5G116 = Server.CreateObject("PS5G116.cListSDnHdrForGI")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

    Call iObjPS5G116.ListRows (gStrGlobalCollection, C_SHEETMAXROWS_D, iStrPostFlag, iArrWhereIn, iStrNextKey, iArrRsOut, iArrWhereOut)
	
	If CheckSYSTEMError(Err,True) = True Then
	   Set iObjPS5G116 = Nothing		                                                 '��: Unload Comproxy DLL
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "parent.frm1.txtConPlant.focus" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'��: Process End
	   Response.End 
	End If

	Set iObjPS5G116 = Nothing		                                                 '��: Unload Comproxy DLL

    ' Check Query Condition
    If iStrNextKey = "" Then
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
		' �����׷� 
		If iArrWhereIn(C_PS5G116_PLANT_FOR_QUERY) = iArrWhereOut(0, C_PS5G116_PLANT_FOR_QUERY) Then
			Response.Write "Parent.frm1.txtConPlantNm.value = """ & iArrWhereOut(1, C_PS5G116_PLANT_FOR_QUERY) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1.txtConPlant.alt, ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConPlantNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConPlant.focus " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If
		
		' �������� 
		If iArrWhereIn(C_PS5G116_MOVE_TYPE_FOR_QUERY) = iArrWhereOut(0, C_PS5G116_MOVE_TYPE_FOR_QUERY) Then
			Response.Write "Parent.frm1.txtConDnTypeNm.value = """ & iArrWhereOut(1, C_PS5G116_MOVE_TYPE_FOR_QUERY) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1.txtConDnType.alt, ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConDnTypeNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConDnType.focus " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If
		
		' ��ǰó 
		If iArrWhereIn(C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) = iArrWhereOut(0, C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) Then
			Response.Write "Parent.frm1.txtConShipToPartyNm.value = """ & iArrWhereOut(1, C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1.txtConShipToParty.alt, ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConShipToPartyNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConShipToParty.focus " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If
		
		' ó���� �ڷᰡ �����ϴ�.
		If UBound(iArrRsOut) < 0 Then
			Response.Write "Call Parent.DisplayMsgBox(""800161"", ""X"", ""X"", ""X"")" & vbCr
			Response.Write "parent.frm1.txtConPlant.focus " & vbCr   
			Response.Write "</SCRIPT> " & VbCr
			Response.End		
		Else
			Response.Write "</SCRIPT> " & VbCr
		End If
	End If

	' Client(MA)�� ���� ��ȸ�� ������ Row
	iLngLastRow = CLng(Request("txtLastRow")) + 1
	
	' Set Next key
	If C_SHEETMAXROWS_D > 0 And Ubound(iArrRsOut,2) = C_SHEETMAXROWS_D Then
		'����ȣ 
		iStrNextKey = iArrRsOut(0, C_SHEETMAXROWS_D)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(iArrRsOut,2)
	End If

	ReDim iArrCols(9)						' Column �� 
	Redim iArrRows(iLngSheetMaxRows)		' ��ȸ�� Row ����ŭ �迭 ������ 

	iArrCols(0) = ""
   	iArrCols(1) = "0"
		
   	For iLngRow = 0 To iLngSheetMaxRows
   		iArrCols(2) = ConvSPChars(iArrRsOut(0, iLngRow))						' ����ȣ 
   		iArrCols(3) = UNIDateClientFormat(iArrRsOut(1, iLngRow))				' ������� 
   		iArrCols(4) = ConvSPChars(iArrRsOut(2, iLngRow))						' ��ǰó 
   		iArrCols(5) = ConvSPChars(iArrRsOut(3, iLngRow)) 						' ��ǰó�� 
   		iArrCols(6) = ConvSPChars(iArrRsOut(4, iLngRow)) 						' �������� 
   		iArrCols(7) = ConvSPChars(iArrRsOut(5, iLngRow)) 						' �������¸� 
   		iArrCols(8) = ConvSPChars(iArrRsOut(6, iLngRow)) 						' ��������� 
   		iArrCols(9) = iLngLastRow + iLngRow 
   		
   		iArrRows(iLngRow) = Join(iArrCols, gColSep)
	Next
	
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
	Response.Write "With parent " & vbCr   
	
	' ���� Display
    Response.Write ".ggoSpread.Source = .frm1.vspdData " & vbCr
    Response.Write ".frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write ".ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write ".lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr  
    Response.Write ".DbQueryOk" & vbCr   
	Response.Write ".frm1.vspdData.Redraw = True  "       & vbCr
	
	' ���� Query�� ���� ��ȸ���� ���� 
	If iStrNextKey <> "" Then
		Response.Write ".frm1.txtHConPlant.value = """ & iArrWhereIn(C_PS5G116_PLANT_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConFromDt.value = """ & iArrWhereIn(C_PS5G116_FR_PROMISE_DT_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConToDt.value	= """ & iArrWhereIn(C_PS5G116_TO_PROMISE_DT_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConDnType.value = """ & iArrWhereIn(C_PS5G116_MOVE_TYPE_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConShipToParty.value = """ & iArrWhereIn(C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) & """" & vbCr
	End If
	' �۾����� - ����� ����ϹǷ� ó���� ��ȸ������ �����ϰ� �־�� �Ѵ�.
	Response.Write ".frm1.txtHConPostFlag.value = """ & iStrPostFlag & """" & vbCr
	
	Response.Write "End With " & vbCr   
	Response.Write "</SCRIPT> " & vbCr      	

	Response.End 
    
Case CStr(UID_M0002)						'��: ���� ��û�� ���� 
	'=========================================================================================
	' Post Goods Issue
	'=========================================================================================
	Dim iArrPostInfo
	Dim iIntLoop
	Dim iObjPS5G115
	Dim pvCB
	Dim iStrCommand			
    Dim iIntIndex, iCCount, itxtSpreadArr, itxtSpreadIns
	Dim iStrFrstDnNo, iStrLastDnNo, iIntLastRow

	Redim iArrPostInfo(5)
		
	' ��� Ȯ������ ���� ���� 
	iArrPostInfo(1) = UNIConvDate(Request("txtActualGIDt"))	' ���� ����� 
	iArrPostInfo(2) = Trim(Request("txtHArFlag"))			' ����������� 
	iArrPostInfo(3) = Trim(Request("txtHVatFlag"))			' ���ݰ�꼭 �������� 
	iArrPostInfo(5) = "ST"									' STO ���� 

	If Request("txtHConPostFlag") = "Y" Then	' Ȯ��(Y)/��ҿ���(N)
		iStrCommand = "POST"					' �׻� �빮�� 
	Else
		iStrCommand = "CANCEL"					' �׻� �빮�� 
	End If

	pvCB = "F" 	   

    iCCount = Request.Form("txtCSpread").Count

    ReDim itxtSpreadArr(iCCount)
    For iIntIndex = 1 To iCCount
        itxtSpreadArr(iIntIndex) = Request.Form("txtCSpread")(iIntIndex)
    Next
    
    itxtSpreadIns = Join(itxtSpreadArr,"")
	
	iArrRows = Split(itxtSpreadIns, gRowSep)
	
	iIntLastRow = UBound(iArrRows) - 1

	Set iObjPS5G115 = CreateObject("PS5G115.cSPOSTGISvr")

	For iIntLoop = 0 To iIntLastRow
		iArrCols = Split(iArrRows(iIntLoop), gColSep)
		
	    iArrPostInfo(0) = iArrCols(1)					' ����ȣ 
		iArrPostInfo(4) = iArrCols(2)					' ��������� 
	    
	    Call iObjPS5G115.S_POST_GOODS_ISSUE_SVR(pvCB, gStrGlobalCollection, iStrCommand, Array(""), iArrPostInfo)
		    
		If CheckSYSTEMError2(Err, True, "(����ȣ : " & iArrCols(1) & ")","","","","") = True Then
			Set iObjPS5G115 = Nothing
			' �Ϻθ� ó�� �� ��� ó���� ������ �����ش�.
			If iIntLoop > 0 Then
				iArrCols = Split(iArrRows(0), gColSep)
				iStrFrstDnNo = iArrCols(1)
				iArrCols = Split(iArrRows(iIntLoop - 1), gColSep)
				iStrLastDnNo = iArrCols(1)
	
				Call DisplayMsgBox("204267", vbOKOnly, iStrFrstDnNo & "~" & iStrLastDnNo & " (" & iIntLastRow & ")", "", I_MKSCRIPT)
				
				Response.Write "<Script language=vbs> " & vbCr   
				Response.Write "Call parent.DbSaveOk " & vbCr   
				Response.Write "</Script> "	& vbCr          

				Response.End
			Else
				Response.Write "<Script language=vbs> " & vbCr   
				Response.Write " Call parent.RemovedivTextArea " & vbCr   
				Response.Write "</Script> "																				         & vbCr          
				Response.End
			End If
		End If
	Next

	Set iObjPS5G115 = Nothing
	
	iArrCols = Split(iArrRows(0), gColSep)
	iStrFrstDnNo = iArrCols(1)
	iArrCols = Split(iArrRows(iIntLastRow), gColSep)
	iStrLastDnNo = iArrCols(1)
	
	Call DisplayMsgBox("204267", vbOKOnly, iStrFrstDnNo & "~" & iStrLastDnNo & " (" & iIntLastRow + 1 & ")", "", I_MKSCRIPT)

	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "Call parent.DbSaveOk " & vbCr   
	Response.Write "</Script> "	& vbCr          
End Select
%>

