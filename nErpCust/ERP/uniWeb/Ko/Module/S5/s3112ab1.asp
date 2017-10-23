<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 출하관리																	*
'*  3. Program ID           : iPS5G112A1																*
'*  4. Program Name         : 수주내역참조																*
'*  5. Program Desc         : 출하내역등록을 위한 수주내역참조 (Business Logic Asp)						*
'*  6. Comproxy List        : iPS5G112ListSoDtlForDnSvr													*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2001/12/19																*
'*  9. Modifier (First)     : Cho Song Hyon																*
'* 10. Modifier (Last)      : Cho Song Hyon																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*				            : 2. 2000/09/21 : 4th Coding												*
'*				            : 3. 2001/12/19 : Date 표준적용												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()
Call HideStatusWnd

Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim GroupCount
Dim StrNextKey							' 다음 값 
Dim lgStrPrevKey						' 이전 값 

strMode = Request("txtMode")		

Select Case strMode
	Case CStr(UID_M0001)													
		Dim iPS5G212													
		Dim iLngRow
		
		Const C_SHEETMAXROWS_D  = 100
		
		' 조회조건 
		Dim iArrWhere, iArrWhereOut
		
		Const S5G212_WH_SHIP_TO_PARTY = 0            ' 납품처 
		Const S5G212_WH_SL_CD = 1                    ' 창고 
		Const S5G212_WH_ITEM_CD = 2                  ' 품목 
		Const S5G212_WH_FR_PROMISE_DT = 3            ' 출고예정일 
		Const S5G212_WH_TO_PROMISE_DT = 4            ' 출고예정일 
		Const S5G212_WH_SO_NO = 5                    ' 수주번호 
		Const S5G212_WH_TRACKING_NO = 6              ' Tracking 번호 
		Const S5G212_WH_PLANT_CD = 7                 ' 공장 
		Const S5G212_WH_SO_TYPE = 8                  ' 수주형태 
		Const S5G212_WH_MOV_TYPE = 9                 ' 출하형태 
		Const S5G212_WH_RET_ITEM_FLAG = 10           ' 반품여부 

		Redim iArrWhere(S5G212_WH_RET_ITEM_FLAG)
		
		' Next Key(Scroll 조회시)
		Dim iArrNextKey
		
		' 조회결과 Index
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

		' 조회조건 
		iArrWhere(S5G212_WH_SHIP_TO_PARTY) = Trim(Request("txtShipToParty"))		' 납품처 
		iArrWhere(S5G212_WH_SL_CD) = Trim(Request("txtSlCd"))						' 창고 
		iArrWhere(S5G212_WH_ITEM_CD) = Trim(Request("txtItemCd"))					' 품목 
		iArrWhere(S5G212_WH_FR_PROMISE_DT) = UNIConvDate(Request("txtFromDt"))		' 출고예정일 
		If UNIConvDate(Request("txtToDt")) = "1900-01-01" Then
			iArrWhere(S5G212_WH_TO_PROMISE_DT) = ""									' 출고예정일 
		Else
			iArrWhere(S5G212_WH_TO_PROMISE_DT) = UNIConvDate(Request("txtToDt"))	' 출고예정일 
		End If
		iArrWhere(S5G212_WH_SO_NO) = Trim(Request("txtSoNo"))						' 수주번호 
		iArrWhere(S5G212_WH_TRACKING_NO) = Trim(Request("txtTrackingNo"))			' Tracking 번호 
		iArrWhere(S5G212_WH_PLANT_CD) = Trim(Request("txtPlantCd"))					' 공장 
		iArrWhere(S5G212_WH_SO_TYPE) = Trim(Request("txtSoType"))					' 수주형태 
		iArrWhere(S5G212_WH_MOV_TYPE) = Trim(Request("txtMovType"))					' 출하형태 
		iArrWhere(S5G212_WH_RET_ITEM_FLAG) = Trim(Request("txtHRetFlag"))			' 반품여부 
        
        ' Scroll 조회조건 
		lgStrPrevKey = Trim(Request("lgStrPrevKey"))
		If lgStrPrevKey <> "" then
			iArrNextKey = Split(lgStrPrevKey, gColSep)
		Else
			Redim iArrNextKey(2)
		End if

		' 자료 조회 
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
		'조건부 조건명 - 처음 조회시만 처리 
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
		' Client(MA)의 현재 조회된 마직막 Row
		iLngLastRow = CLng(Request("txtLastRow")) + 1
	
		' Set Next key
		If Ubound(iArrRsOut,2) = C_SHEETMAXROWS_D Then
			'출고순번 
			iStrNextKey = iArrRsOut(S5G212_RS_SO_NO, C_SHEETMAXROWS_D) & gColSep & _
						  iArrRsOut(S5G212_RS_SO_SEQ, C_SHEETMAXROWS_D) & gColSep & _
						  iArrRsOut(S5G212_RS_SO_SCHD_NO, C_SHEETMAXROWS_D)
			iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
		Else
			iStrNextKey = ""
			iLngSheetMaxRows = Ubound(iArrRsOut,2)
		End If

		ReDim iArrCols(31)						' Column 수 
		Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

		iArrCols(0)  = ""		' Row Header

		For iLngRow = 0 To iLngSheetMaxRows
   			iArrCols(1) = UNIDateClientFormat(iArrRsOut(S5G212_RS_PROMISE_DT, iLngRow))							' 출고예정일 
   			iArrCols(2) = ConvSPChars(iArrRsOut(S5G212_RS_ITEM_CD, iLngRow))					' 품목코드 
   			iArrCols(3) = ConvSPChars(iArrRsOut(S5G212_RS_ITEM_NM, iLngRow))					' 품목명 
   			iArrCols(4) = UNINumClientFormat(iArrRsOut(S5G212_RS_REMN_QTY, iLngRow), ggQty.DecPoint, 0)			' 미출고수량 
   			iArrCols(5) = UNINumClientFormat(iArrRsOut(S5G212_RS_REMN_BONUS_QTY, iLngRow), ggQty.DecPoint, 0)		' 미출고덤수량 
   			iArrCols(6) = ConvSPChars(iArrRsOut(S5G212_RS_SO_UNIT, iLngRow))					' 단위 
   			iArrCols(7) = UNINumClientFormat(iArrRsOut(S5G212_RS_GOOD_ON_HAND_QTY, iLngRow), ggQty.DecPoint, 0)	' 재고량 
   			iArrCols(8) = ConvSPChars(iArrRsOut(S5G212_RS_BASIC_UNIT, iLngRow))					' 재고단위 
   			iArrCols(9) = ConvSPChars(iArrRsOut(S5G212_RS_SO_NO, iLngRow))					' 수주번호 
   			iArrCols(10) = UNINumClientFormat(iArrRsOut(S5G212_RS_SO_SEQ, iLngRow), 0, 0)		' 수주순번 
   			iArrCols(11) = UNINumClientFormat(iArrRsOut(S5G212_RS_SO_SCHD_NO, iLngRow), 0, 0)	' 수주일정순번 
   			iArrCols(12) = ConvSPChars(iArrRsOut(S5G212_RS_TRACKING_NO, iLngRow))				' 제번 
   			iArrCols(13) = ConvSPChars(iArrRsOut(S5G212_RS_SHIP_TO_PARTY, iLngRow))			' 납품처 
   			iArrCols(14) = ConvSPChars(iArrRsOut(S5G212_RS_SHIP_TO_PARTY_NM, iLngRow))		' 납품처명 
   			iArrCols(15) = ConvSPChars(iArrRsOut(S5G212_RS_PLANT_CD, iLngRow))				' 공장코드 
   			iArrCols(16) = ConvSPChars(iArrRsOut(S5G212_RS_PLANT_NM, iLngRow))				' 공장명 
   			iArrCols(17) = ConvSPChars(iArrRsOut(S5G212_RS_SL_CD, iLngRow))					' 창고코드 
   			iArrCols(18) = ConvSPChars(iArrRsOut(S5G212_RS_SL_NM, iLngRow))					' 창고명 
   			iArrCols(19) = UNINumClientFormat(iArrRsOut(S5G212_RS_TOL_MORE_QTY, iLngRow), ggQty.DecPoint, 0)	' 과부족허용량(+) 
   			iArrCols(20) = UNINumClientFormat(iArrRsOut(S5G212_RS_TOL_LESS_QTY, iLngRow), ggQty.DecPoint, 0)	' 과부족허용량(-) 
   			iArrCols(21) = ConvSPChars(iArrRsOut(S5G212_RS_LC_NO, iLngRow))					' L/C번호 
   			iArrCols(22) = UNINumClientFormat(iArrRsOut(S5G212_RS_LC_SEQ, iLngRow), 0, 0)	' L/C순번 
   			iArrCols(23) = ConvSPChars(iArrRsOut(S5G212_RS_LOT_FLG, iLngRow))				' Lot 관리여부 
   			iArrCols(24) = ConvSPChars(iArrRsOut(S5G212_RS_LOT_NO, iLngRow))					' Lot No
   			iArrCols(25) = UNINumClientFormat(iArrRsOut(S5G212_RS_LOT_SEQ, iLngRow), 0, 0)	' Lot Seq
   			iArrCols(26) = ConvSPChars(iArrRsOut(S5G212_RS_RET_ITEM_FLAG, iLngRow))			' 반품여부 
   			iArrCols(27) = ConvSPChars(iArrRsOut(S5G212_RS_RET_TYPE, iLngRow))				' 반품 사유 코드 
   			iArrCols(28) = ConvSPChars(iArrRsOut(S5G212_RS_RET_TYPE_NM, iLngRow))			' 반품 사유 명 
   			iArrCols(29) = ConvSPChars(iArrRsOut(S5G212_RS_SPEC, iLngRow))					' 품목규격 
   			iArrCols(30) = ConvSPChars(iArrRsOut(S5G212_RS_REMARK, iLngRow))					' 비고 
			iArrCols(31) = iLngLastRow + iLngRow
			
   			iArrRows(iLngRow) = Join(iArrCols, gColSep)
		Next
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write "With parent " & vbCr   
		Response.Write " .ggoSpread.Source = .vspdData" & vbCr
		Response.Write " .vspdData.Redraw = False  "      & vbCr      
		Response.Write " .ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
		Response.Write " .lgStrPrevKey = """ & iStrNextKey  & """" & vbCr  

		' Scroll Query를 위한 조회값 Hidden 필드에 할당.		
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
