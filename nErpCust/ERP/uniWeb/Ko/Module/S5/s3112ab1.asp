<%
'********************************************************************************************************
'*  1. Module Name          : ����																		*
'*  2. Function Name        : ���ϰ���																	*
'*  3. Program ID           : iPS5G112A1																*
'*  4. Program Name         : ���ֳ�������																*
'*  5. Program Desc         : ���ϳ�������� ���� ���ֳ������� (Business Logic Asp)						*
'*  6. Comproxy List        : iPS5G112ListSoDtlForDnSvr													*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Cho Song Hyon																*
'* 10. Modifier (Last)      : Cho Song Hyon																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2001/12/19 : Date ǥ������												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
																				'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call LoadBasisGlobalInf()
Call HideStatusWnd

Dim strMode																		'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim GroupCount
Dim StrNextKey							' ���� �� 
Dim lgStrPrevKey						' ���� �� 

strMode = Request("txtMode")		

Select Case strMode
	Case CStr(UID_M0001)													
		Dim iPS5G212													
		Dim iLngRow
		
		Const C_SHEETMAXROWS_D  = 100
		
		' ��ȸ���� 
		Dim iArrWhere, iArrWhereOut
		
		Const S5G212_WH_SHIP_TO_PARTY = 0            ' ��ǰó 
		Const S5G212_WH_SL_CD = 1                    ' â�� 
		Const S5G212_WH_ITEM_CD = 2                  ' ǰ�� 
		Const S5G212_WH_FR_PROMISE_DT = 3            ' ������� 
		Const S5G212_WH_TO_PROMISE_DT = 4            ' ������� 
		Const S5G212_WH_SO_NO = 5                    ' ���ֹ�ȣ 
		Const S5G212_WH_TRACKING_NO = 6              ' Tracking ��ȣ 
		Const S5G212_WH_PLANT_CD = 7                 ' ���� 
		Const S5G212_WH_SO_TYPE = 8                  ' �������� 
		Const S5G212_WH_MOV_TYPE = 9                 ' �������� 
		Const S5G212_WH_RET_ITEM_FLAG = 10           ' ��ǰ���� 

		Redim iArrWhere(S5G212_WH_RET_ITEM_FLAG)
		
		' Next Key(Scroll ��ȸ��)
		Dim iArrNextKey
		
		' ��ȸ��� Index
		Const S5G212_RS_SHIP_TO_PARTY = 0
		Const S5G212_RS_SHIP_TO_PARTY_NM = 1
		Const S5G212_RS_PROMISE_DT = 2
		Const S5G212_RS_SO_NO = 3
		Const S5G212_RS_SO_SEQ = 4
		Const S5G212_RS_SO_SCHD_NO = 5
		Const S5G212_RS_TRACKING_NO = 6
		Const S5G212_RS_ITEM_CD = 7
		Const S5G212_RS_ITEM_NM = 8
		Const S5G212_RS_SPEC = 9
		Const S5G212_RS_PLANT_CD = 10
		Const S5G212_RS_PLANT_NM = 11
		Const S5G212_RS_SL_CD = 12
		Const S5G212_RS_SL_NM = 13
		Const S5G212_RS_LOT_NO = 14
		Const S5G212_RS_LOT_SEQ = 15
		Const S5G212_RS_LC_NO = 16
		Const S5G212_RS_LC_SEQ = 17
		Const S5G212_RS_REMN_QTY = 18
		Const S5G212_RS_REMN_BONUS_QTY = 19
		Const S5G212_RS_TOL_MORE_QTY = 20
		Const S5G212_RS_TOL_LESS_QTY = 21
		Const S5G212_RS_SO_UNIT = 22
		Const S5G212_RS_GOOD_ON_HAND_QTY = 23
		Const S5G212_RS_BASIC_UNIT = 24
		Const S5G212_RS_LOT_FLG = 25
		Const S5G212_RS_ITEM_ACCT = 26
		Const S5G212_RS_RET_ITEM_FLAG = 27
		Const S5G212_RS_RET_TYPE = 28
		Const S5G212_RS_RET_TYPE_NM = 29
		Const S5G212_RS_REMARK = 30

		' ��ȸ���� 
		iArrWhere(S5G212_WH_SHIP_TO_PARTY) = Trim(Request("txtShipToParty"))		' ��ǰó 
		iArrWhere(S5G212_WH_SL_CD) = Trim(Request("txtSlCd"))						' â�� 
		iArrWhere(S5G212_WH_ITEM_CD) = Trim(Request("txtItemCd"))					' ǰ�� 
		iArrWhere(S5G212_WH_FR_PROMISE_DT) = UNIConvDate(Request("txtFromDt"))		' ������� 
		If UNIConvDate(Request("txtToDt")) = "1900-01-01" Then
			iArrWhere(S5G212_WH_TO_PROMISE_DT) = ""									' ������� 
		Else
			iArrWhere(S5G212_WH_TO_PROMISE_DT) = UNIConvDate(Request("txtToDt"))	' ������� 
		End If
		iArrWhere(S5G212_WH_SO_NO) = Trim(Request("txtSoNo"))						' ���ֹ�ȣ 
		iArrWhere(S5G212_WH_TRACKING_NO) = Trim(Request("txtTrackingNo"))			' Tracking ��ȣ 
		iArrWhere(S5G212_WH_PLANT_CD) = Trim(Request("txtPlantCd"))					' ���� 
		iArrWhere(S5G212_WH_SO_TYPE) = Trim(Request("txtSoType"))					' �������� 
		iArrWhere(S5G212_WH_MOV_TYPE) = Trim(Request("txtMovType"))					' �������� 
		iArrWhere(S5G212_WH_RET_ITEM_FLAG) = Trim(Request("txtHRetFlag"))			' ��ǰ���� 
        
        ' Scroll ��ȸ���� 
		lgStrPrevKey = Trim(Request("lgStrPrevKey"))
		If lgStrPrevKey <> "" then
			iArrNextKey = Split(lgStrPrevKey, gColSep)
		Else
			Redim iArrNextKey(2)
		End if

		' �ڷ� ��ȸ 
		Set iPS5G212 = Server.CreateObject("PS5G212.cSListSSoSchdForDn")
      
        If CheckSYSTEMError(Err,True) = True Then
			Response.Write "<Script Language=vbscript> " & vbCr
			Response.Write "Call parent.SetFocusToDocument(""P"") " & vbCr
			Response.Write "parent.txtFromDt.focus " & vbCr
			Response.Write "</Script> " & vbCr
			Response.End
        End If

        Call iPS5G212.ListRows(gStrGlobalCollection, C_SHEETMAXROWS_D, iArrWhere, iArrRsOut, iArrNextKey, iArrWhereOut)

		'-----------------------
		'���Ǻ� ���Ǹ� - ó�� ��ȸ�ø� ó�� 
		'-----------------------
		If lgStrPrevKey = "" Then
			Response.Write "<Script Language=vbscript> " & vbCr
			Response.Write "With parent " & vbCr
			Response.Write		".txtShipToPartyNm.Value = """ & ConvSPChars(iArrWhereOut(S5G212_WH_SHIP_TO_PARTY)) & """" & vbCr
			Response.Write		".txtSlNm.Value = """ & ConvSPChars(iArrWhereOut(S5G212_WH_SL_CD)) & """" & vbCr
			Response.Write		".txtItemNm.Value = """ & ConvSPChars(iArrWhereOut(S5G212_WH_ITEM_CD)) & """" & vbCr
			Response.Write "End With " & vbCr
			Response.Write "</Script> " & vbCr
		End If

	    If CheckSYSTEMError(Err,True) = True Then
			Set iPS5G212 = Nothing
			Response.Write "<Script Language=vbscript> " & vbCr
			Response.Write "Call parent.SetFocusToDocument(""P"") " & vbCr
			Response.Write "parent.txtFromDt.focus " & vbCr
			Response.Write "</Script> " & vbCr
           Response.End
        End If  
        
        Set iPS5G212 = Nothing	        

		'-----------------------
		'Result data display area
		'-----------------------
		Dim iLngLastRow, iLngSheetMaxRows
		' Client(MA)�� ���� ��ȸ�� ������ Row
		iLngLastRow = CLng(Request("txtLastRow")) + 1
	
		' Set Next key
		If Ubound(iArrRsOut,2) = C_SHEETMAXROWS_D Then
			'������ 
			iStrNextKey = iArrRsOut(S5G212_RS_SO_NO, C_SHEETMAXROWS_D) & gColSep & _
						  iArrRsOut(S5G212_RS_SO_SEQ, C_SHEETMAXROWS_D) & gColSep & _
						  iArrRsOut(S5G212_RS_SO_SCHD_NO, C_SHEETMAXROWS_D)
			iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
		Else
			iStrNextKey = ""
			iLngSheetMaxRows = Ubound(iArrRsOut,2)
		End If

		ReDim iArrCols(31)						' Column �� 
		Redim iArrRows(iLngSheetMaxRows)		' ��ȸ�� Row ����ŭ �迭 ������ 

		iArrCols(0)  = ""		' Row Header

		For iLngRow = 0 To iLngSheetMaxRows
   			iArrCols(1) = UNIDateClientFormat(iArrRsOut(S5G212_RS_PROMISE_DT, iLngRow))							' ������� 
   			iArrCols(2) = ConvSPChars(iArrRsOut(S5G212_RS_ITEM_CD, iLngRow))					' ǰ���ڵ� 
   			iArrCols(3) = ConvSPChars(iArrRsOut(S5G212_RS_ITEM_NM, iLngRow))					' ǰ��� 
   			iArrCols(4) = UNINumClientFormat(iArrRsOut(S5G212_RS_REMN_QTY, iLngRow), ggQty.DecPoint, 0)			' �������� 
   			iArrCols(5) = UNINumClientFormat(iArrRsOut(S5G212_RS_REMN_BONUS_QTY, iLngRow), ggQty.DecPoint, 0)		' ���������� 
   			iArrCols(6) = ConvSPChars(iArrRsOut(S5G212_RS_SO_UNIT, iLngRow))					' ���� 
   			iArrCols(7) = UNINumClientFormat(iArrRsOut(S5G212_RS_GOOD_ON_HAND_QTY, iLngRow), ggQty.DecPoint, 0)	' ��� 
   			iArrCols(8) = ConvSPChars(iArrRsOut(S5G212_RS_BASIC_UNIT, iLngRow))					' ������ 
   			iArrCols(9) = ConvSPChars(iArrRsOut(S5G212_RS_SO_NO, iLngRow))					' ���ֹ�ȣ 
   			iArrCols(10) = UNINumClientFormat(iArrRsOut(S5G212_RS_SO_SEQ, iLngRow), 0, 0)		' ���ּ��� 
   			iArrCols(11) = UNINumClientFormat(iArrRsOut(S5G212_RS_SO_SCHD_NO, iLngRow), 0, 0)	' ������������ 
   			iArrCols(12) = ConvSPChars(iArrRsOut(S5G212_RS_TRACKING_NO, iLngRow))				' ���� 
   			iArrCols(13) = ConvSPChars(iArrRsOut(S5G212_RS_SHIP_TO_PARTY, iLngRow))			' ��ǰó 
   			iArrCols(14) = ConvSPChars(iArrRsOut(S5G212_RS_SHIP_TO_PARTY_NM, iLngRow))		' ��ǰó�� 
   			iArrCols(15) = ConvSPChars(iArrRsOut(S5G212_RS_PLANT_CD, iLngRow))				' �����ڵ� 
   			iArrCols(16) = ConvSPChars(iArrRsOut(S5G212_RS_PLANT_NM, iLngRow))				' ����� 
   			iArrCols(17) = ConvSPChars(iArrRsOut(S5G212_RS_SL_CD, iLngRow))					' â���ڵ� 
   			iArrCols(18) = ConvSPChars(iArrRsOut(S5G212_RS_SL_NM, iLngRow))					' â��� 
   			iArrCols(19) = UNINumClientFormat(iArrRsOut(S5G212_RS_TOL_MORE_QTY, iLngRow), ggQty.DecPoint, 0)	' ��������뷮(+) 
   			iArrCols(20) = UNINumClientFormat(iArrRsOut(S5G212_RS_TOL_LESS_QTY, iLngRow), ggQty.DecPoint, 0)	' ��������뷮(-) 
   			iArrCols(21) = ConvSPChars(iArrRsOut(S5G212_RS_LC_NO, iLngRow))					' L/C��ȣ 
   			iArrCols(22) = UNINumClientFormat(iArrRsOut(S5G212_RS_LC_SEQ, iLngRow), 0, 0)	' L/C���� 
   			iArrCols(23) = ConvSPChars(iArrRsOut(S5G212_RS_LOT_FLG, iLngRow))				' Lot �������� 
   			iArrCols(24) = ConvSPChars(iArrRsOut(S5G212_RS_LOT_NO, iLngRow))					' Lot No
   			iArrCols(25) = UNINumClientFormat(iArrRsOut(S5G212_RS_LOT_SEQ, iLngRow), 0, 0)	' Lot Seq
   			iArrCols(26) = ConvSPChars(iArrRsOut(S5G212_RS_RET_ITEM_FLAG, iLngRow))			' ��ǰ���� 
   			iArrCols(27) = ConvSPChars(iArrRsOut(S5G212_RS_RET_TYPE, iLngRow))				' ��ǰ ���� �ڵ� 
   			iArrCols(28) = ConvSPChars(iArrRsOut(S5G212_RS_RET_TYPE_NM, iLngRow))			' ��ǰ ���� �� 
   			iArrCols(29) = ConvSPChars(iArrRsOut(S5G212_RS_SPEC, iLngRow))					' ǰ��԰� 
   			iArrCols(30) = ConvSPChars(iArrRsOut(S5G212_RS_REMARK, iLngRow))					' ��� 
			iArrCols(31) = iLngLastRow + iLngRow
			
   			iArrRows(iLngRow) = Join(iArrCols, gColSep)
		Next
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write "With parent " & vbCr   
		Response.Write " .ggoSpread.Source = .vspdData" & vbCr
		Response.Write " .vspdData.Redraw = False  "      & vbCr      
		Response.Write " .ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
		Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  

		' Scroll Query�� ���� ��ȸ�� Hidden �ʵ忡 �Ҵ�.		
		Response.Write " .HFromDt.value = """ & Request("txtFromDt") & """" & vbCr   
		Response.Write " .HToDt.value = """ & Request("txtToDt") & """" & vbCr   
		Response.Write " .HShipToParty.value = """ & Request("txtShipToParty") & """" & vbCr   
		Response.Write " .HSlCd.value = """ & Request("txtSlCd") & """" & vbCr   
		Response.Write " .HItemCd.value = """ & Request("txtItemCd") & """" & vbCr   
		Response.Write " .HTrackingNo.value = """ & Request("txtTrackingNo") & """" & vbCr   
		Response.Write " .HSoNo.value = """ & Request("txtSoNo") & """" & vbCr   
		
		Response.Write " .DbQueryOk " & vbCr   
		Response.Write " .vspdData.Redraw = True  "       & vbCr
		Response.Write "End With " & vbCr   
		Response.Write "</Script> " & vbCr          

End Select
%>
