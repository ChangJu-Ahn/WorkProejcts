<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111MB6
'*  4. Program Name         : 일괄출고처리 
'*  5. Program Desc         :
'*  6. DLL List				: PS5G116
'*  7. Modified date(First) : 2003/07/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

On Error Resume Next									

Call HideStatusWnd

Dim iStrMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iArrCols, iArrRows 

iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case iStrMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
	
	Dim iStrNextKey							' 다음 값 
	Dim iLngLastRow							' 현재 그리드의 최대Row
	Dim iStrPostFlag
	Dim iBlnInitQuery						' 최초 조회여부 
	Dim iObjPS5G116

	Dim iLngRow, iLngSheetMaxRows
	Dim iArrWhereIn, iArrWhereOut, iArrRsOut

	Const C_PS5G116_PLANT_FOR_QUERY = 0              ' Plant
	Const C_PS5G116_FR_PROMISE_DT_FOR_QUERY = 1      ' Promise date(G/I) or Actual G/I date(Cancel G/I)
	Const C_PS5G116_TO_PROMISE_DT_FOR_QUERY = 2      ' Promise date(G/I) or Actual G/I date(Cancel G/I)
	Const C_PS5G116_MOVE_TYPE_FOR_QUERY = 3          ' Movement type
	Const C_PS5G116_SHIP_TO_PARTY_FOR_QUERY = 4      ' Ship to party

	Dim C_SHEETMAXROWS_D				' 한번에 Query할 Row수 

	If Request("txtBatchQuery") = "Y" Then
		C_SHEETMAXROWS_D = -1			' 조회조건에 해당되는 모든 Row를 반환한다.
	Else
		C_SHEETMAXROWS_D = 100
	End If
	
	'---------------------------------------------
    'next key값을 넘겨준다.
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
	    
	iStrPostFlag = Request("txtConPostFlag")		' 확정(Y)/취소여부(N)

    Set iObjPS5G116 = Server.CreateObject("PS5G116.cListSDnHdrForGI")    
  
    If CheckSYSTEMError(Err,True) = True Then
		Response.End		
    End If

    Call iObjPS5G116.ListRows (gStrGlobalCollection, C_SHEETMAXROWS_D, iStrPostFlag, iArrWhereIn, iStrNextKey, iArrRsOut, iArrWhereOut)
	
	If CheckSYSTEMError(Err,True) = True Then
	   Set iObjPS5G116 = Nothing		                                                 '☜: Unload Comproxy DLL
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "parent.frm1.txtConPlant.focus" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'☜: Process End
	   Response.End 
	End If

	Set iObjPS5G116 = Nothing		                                                 '☜: Unload Comproxy DLL

    ' Check Query Condition
    If iStrNextKey = "" Then
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
		' 영업그룹 
		If iArrWhereIn(C_PS5G116_PLANT_FOR_QUERY) = iArrWhereOut(0, C_PS5G116_PLANT_FOR_QUERY) Then
			Response.Write "Parent.frm1.txtConPlantNm.value = """ & iArrWhereOut(1, C_PS5G116_PLANT_FOR_QUERY) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1.txtConPlant.alt, ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConPlantNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConPlant.focus " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If
		
		' 출하형태 
		If iArrWhereIn(C_PS5G116_MOVE_TYPE_FOR_QUERY) = iArrWhereOut(0, C_PS5G116_MOVE_TYPE_FOR_QUERY) Then
			Response.Write "Parent.frm1.txtConDnTypeNm.value = """ & iArrWhereOut(1, C_PS5G116_MOVE_TYPE_FOR_QUERY) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1.txtConDnType.alt, ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConDnTypeNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConDnType.focus " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If
		
		' 납품처 
		If iArrWhereIn(C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) = iArrWhereOut(0, C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) Then
			Response.Write "Parent.frm1.txtConShipToPartyNm.value = """ & iArrWhereOut(1, C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) & """" & vbCr
		Else
			Response.Write "Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1.txtConShipToParty.alt, ""X"")" & vbCr   
			Response.Write "parent.frm1.txtConShipToPartyNm.value = """"" & vbCr   
			Response.Write "parent.frm1.txtConShipToParty.focus " & vbCr   
			Response.Write "</SCRIPT> "
			Response.End
		End If
		
		' 처리할 자료가 없습니다.
		If UBound(iArrRsOut) < 0 Then
			Response.Write "Call Parent.DisplayMsgBox(""800161"", ""X"", ""X"", ""X"")" & vbCr
			Response.Write "parent.frm1.txtConPlant.focus " & vbCr   
			Response.Write "</SCRIPT> " & VbCr
			Response.End		
		Else
			Response.Write "</SCRIPT> " & VbCr
		End If
	End If

	' Client(MA)의 현재 조회된 마직막 Row
	iLngLastRow = CLng(Request("txtLastRow")) + 1
	
	' Set Next key
	If C_SHEETMAXROWS_D > 0 And Ubound(iArrRsOut,2) = C_SHEETMAXROWS_D Then
		'출고번호 
		iStrNextKey = iArrRsOut(0, C_SHEETMAXROWS_D)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(iArrRsOut,2)
	End If

	ReDim iArrCols(9)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	iArrCols(0) = ""
   	iArrCols(1) = "0"
		
   	For iLngRow = 0 To iLngSheetMaxRows
   		iArrCols(2) = ConvSPChars(iArrRsOut(0, iLngRow))						' 출고번호 
   		iArrCols(3) = UNIDateClientFormat(iArrRsOut(1, iLngRow))				' 출고예정일 
   		iArrCols(4) = ConvSPChars(iArrRsOut(2, iLngRow))						' 납품처 
   		iArrCols(5) = ConvSPChars(iArrRsOut(3, iLngRow)) 						' 납품처명 
   		iArrCols(6) = ConvSPChars(iArrRsOut(4, iLngRow)) 						' 출하형태 
   		iArrCols(7) = ConvSPChars(iArrRsOut(5, iLngRow)) 						' 출하형태명 
   		iArrCols(8) = ConvSPChars(iArrRsOut(6, iLngRow)) 						' 예외출고여부 
   		iArrCols(9) = iLngLastRow + iLngRow 
   		
   		iArrRows(iLngRow) = Join(iArrCols, gColSep)
	Next
	
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
	Response.Write "With parent " & vbCr   
	
	' 내역 Display
    Response.Write ".ggoSpread.Source = .frm1.vspdData " & vbCr
    Response.Write ".frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write ".ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write ".lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr  
    Response.Write ".DbQueryOk" & vbCr   
	Response.Write ".frm1.vspdData.Redraw = True  "       & vbCr
	
	' 다음 Query를 위한 조회조건 설정 
	If iStrNextKey <> "" Then
		Response.Write ".frm1.txtHConPlant.value = """ & iArrWhereIn(C_PS5G116_PLANT_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConFromDt.value = """ & iArrWhereIn(C_PS5G116_FR_PROMISE_DT_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConToDt.value	= """ & iArrWhereIn(C_PS5G116_TO_PROMISE_DT_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConDnType.value = """ & iArrWhereIn(C_PS5G116_MOVE_TYPE_FOR_QUERY) & """" & vbCr
		Response.Write ".frm1.txtHConShipToParty.value = """ & iArrWhereIn(C_PS5G116_SHIP_TO_PARTY_FOR_QUERY) & """" & vbCr
	End If
	' 작업유형 - 저장시 사용하므로 처음의 조회조건을 저장하고 있어야 한다.
	Response.Write ".frm1.txtHConPostFlag.value = """ & iStrPostFlag & """" & vbCr
	
	Response.Write "End With " & vbCr   
	Response.Write "</SCRIPT> " & vbCr      	

	Response.End 
    
Case CStr(UID_M0002)						'☜: 저장 요청을 받음 
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
		
	' 출고 확정관련 정보 설정 
	iArrPostInfo(1) = UNIConvDate(Request("txtActualGIDt"))	' 실제 출고일 
	iArrPostInfo(2) = Trim(Request("txtHArFlag"))			' 매출생성여부 
	iArrPostInfo(3) = Trim(Request("txtHVatFlag"))			' 세금계산서 생성여부 
	iArrPostInfo(5) = "ST"									' STO 여부 

	If Request("txtHConPostFlag") = "Y" Then	' 확정(Y)/취소여부(N)
		iStrCommand = "POST"					' 항상 대문자 
	Else
		iStrCommand = "CANCEL"					' 항상 대문자 
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
		
	    iArrPostInfo(0) = iArrCols(1)					' 출고번호 
		iArrPostInfo(4) = iArrCols(2)					' 예외출고여부 
	    
	    Call iObjPS5G115.S_POST_GOODS_ISSUE_SVR(pvCB, gStrGlobalCollection, iStrCommand, Array(""), iArrPostInfo)
		    
		If CheckSYSTEMError2(Err, True, "(출고번호 : " & iArrCols(1) & ")","","","","") = True Then
			Set iObjPS5G115 = Nothing
			' 일부만 처리 된 경우 처리된 정보를 보여준다.
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

