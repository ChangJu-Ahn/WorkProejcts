<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>


<%
'**********************************************************************************************
'*  1. Module Name          : 원가 
'*  2. Function Name        : 품목그룹별매출이익분석 
'*  3. Program ID           : c4232ma1_KO441
'*  4. Program Name         : 품목그룹별매출이익분석 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2008/10/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : LSY
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     
     Call SubBizQuery()
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt
	Dim sItemGroupCd, sItemAcct, sItemCd, sSalesGrp, sBPCd, sTypeFlag, sGrid, arrTmp, TmpBuffer()

	Dim C_SHEET_MAX_CONT 
	
	C_SHEET_MAX_CONT = 10000

'    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")	
	sNextKey	= Request("lgStrPrevKey")
	
	sItemGroupCd	= Request("txtItem_Group_CD")	
	sItemAcct	= Request("txtITEM_ACCT")	
	sItemCd		= Request("txtITEM_CD")	
	sSalesGrp	= Request("txtSALES_GRP")	
	sBPCd		= Request("txtBP_CD")	
	'sTypeFlag	= Request("txtTAB")
	sTypeFlag	= 3
	sGrid		= Request("rdoTYPE")
		
	If sStartDt = "" And sItemGroupCd = "" And sItemAcct = "" And sItemCd = "" Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sItemGroupCd	= "" Then sItemGroupCd	= "%"
	If sItemAcct = "" Then sItemAcct = "%"
	If sItemCd = "" Then sItemCd = "%"
	If sSalesGrp = "" Then sSalesGrp = "%"
	If sBPCd = "" Then sBPCd = "%"
	
	If Instr(1, sItemCd, "%") = 0 Then sItemCd = sItemCd & "%"
	If Instr(1, sSalesGrp, "%") = 0 Then sSalesGrp = sSalesGrp & "%"
	If Instr(1, sBPCd, "%") = 0 Then sBPCd	= sBPCd '& "%"

	' -- sNextKey 변경사항 
	If Instr(1, sNextKey, "*") > 0 Then
		arrTmp = Split(sNextKey, gColSep)
		sNextKey = arrTmp(0)
		C_SHEET_MAX_CONT = 32000
	End If


    With lgObjComm
		.CommandTimeout = 0
		

		'.CommandText = "dbo.usp_C_C4232MA1_TYPE" & sTypeFlag & "_KO441"		' --  변경해야할 SP 명 
		.CommandText = "dbo.usp_C_C4232MA1_TYPE3_KO441"		' --  변경해야할 SP 명 
		.CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_GROUP_CD",	adVarXChar,	adParamInput, 25,Replace(sItemGroupCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_ACCT",	adVarXChar,	adParamInput, 2,Replace(sItemAcct, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 18,Replace(sItemCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SALES_GRP",	adVarXChar,	adParamInput, 4,Replace(sSalesGrp, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BP_CD",	adVarXChar,	adParamInput, 10,Replace(sBPCd, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TYPE",	adXChar,	adParamInput, 1,Replace(sGrid, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, C_SHEET_MAX_CONT)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 
			
		Set oRs = lgObjComm.Execute
				
    End With

    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If
		
    
    ' -- 리턴 레코드셋은 총 2-3종 
    If Not oRs.EOF Then

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
		Response.Write "	.frm1.hEND_DT.value = """ & sEndDt & """" & vbCr 	
		Response.Write "	.frm1.hITEM_GROUP_CD.value = """ & sItemGroupCd & """" & vbCr
		Response.Write "	.frm1.hITEM_ACCT.value = """ & sItemAcct & """" & vbCr
		Response.Write "	.frm1.hITEM_CD.value = """ & sItemCd & """" & vbCr
		Response.Write "	.frm1.hSALES_GRP.value = """ & sSalesGrp & """" & vbCr
		Response.Write "	.frm1.hBP_CD.value = """ & sBPCd & """" & vbCr
		
		If sTypeFlag = "1" Then	' -- All 또는 1번 그리드		
			' -- 1번 그리드 조회 
			If sNextKey = "" Then	' -- 헤더 레코드셋 포함되어 옴 
				' --- 헤더 출력(변경분) ----			
				Dim arrColRow, i, j, ColHeadRowNo, sTxt2
				
				ColHeadRowNo = -1000
				arrColRow = oRs.GetRows()
				iLngRowCnt	= UBound(arrColRow, 2) 
				iLngColCnt	= UBound(arrColRow, 1) 
				
				For j = 0 To iLngRowCnt
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt2 = sTxt2 & "매출수량" & gColSep
					sTxt2 = sTxt2 & "매출액" & gColSep
					sTxt2 = sTxt2 & "매출원가" & gColSep
					sTxt2 = sTxt2 & "매출이익" & gColSep
				Next								
			
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
				sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep
				
				' --- 각 화면별로 변경해야될 치환 문자열들..
				sTxt	= Replace(sTxt, "%CS", "합계")
						
				Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
				' ----------------------------------------
			End If

			' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
			arrRows		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
		
			Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
			arrDataSet = oRs.GetRows()

			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
			iLngGrColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
		
			iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	* 4		' 컬럼수 
			iMaxCols	= 8	+ iLngColCnt 	' 그룹바이 컬럼행수 
		
			' -- 데이타셋을 기초로 배열로 재구성한다.
			ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
			' -- 데이타셋 행수 
			iLngRowCnt	= UBound(arrDataSet, 2)
		
			For iLngRow = 0 To 	iLngRowCnt
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+1, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(3, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+2, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(4, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+3, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(5, iLngRow)	' -- 열, 행, 값(썸)

		Next		
			' ----------------------------------------------------------
			sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- 다은키값 (변경분)
			' -- 조회 Row가 최대행수와 일치할때만 다음데이타 존재함 
			If CInt(iLngGrRowCnt) < C_SHEET_MAX_CONT-1 Then
				sRowSeq = ""
			End If

			Redim TmpBuffer(iLngGrRowCnt)			
			' -- 그룹바이 행으로 루핑 
			For iLngRow = 0 To 	iLngGrRowCnt						
				iStrData = Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- PROJECT_NO
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))									' -- ITEM_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(6, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				
				' -- 데이타그리드 출력(수정분)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next

				If Trim(arrRows(7, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
					sGrpTxt = sGrpTxt & arrRows(7, iLngRow) & gColSep & arrRows(8, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
				End If
				'iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(7, iLngRow)))	' -- Row_Seq
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(8, iLngRow)))	' -- Row_Seq
				iStrData = iStrData & Chr(12)				
				TmpBuffer(iLngRow) = iStrData
			Next
			
			iStrData = Join(TmpBuffer, "")
			
			If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet(" & iMaxCols & ")" & vbCr
				Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			End If

			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 		
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr	 

			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
			End If
			
		End If	
		' ------------------------------------------------------------------
		If sTypeFlag = "2" Then	' -- All 또는 2번 그리드		
			' -- 1번 그리드 조회 
			If sNextKey = "" Then	' -- 헤더 레코드셋 포함되어 옴 
				ColHeadRowNo = -1000
				arrColRow = oRs.GetRows()
				iLngRowCnt	= UBound(arrColRow, 2) 
				iLngColCnt	= UBound(arrColRow, 1) 
				
				For j = 0 To iLngRowCnt
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt2 = sTxt2 & "매출수량" & gColSep
					sTxt2 = sTxt2 & "매출액" & gColSep
					sTxt2 = sTxt2 & "매출원가" & gColSep
					sTxt2 = sTxt2 & "매출이익" & gColSep
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
				sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep				
				' --- 각 화면별로 변경해야될 치환 문자열들..
				sTxt	= Replace(sTxt, "%CS", "합계")						
				
				Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
				' ----------------------------------------
			End If

			' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
			arrRows		= oRs.GetRows()		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			arrDataSet = oRs.GetRows()

			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
			iLngGrColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
		
			iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	* 4				' 컬럼수 
			iMaxCols	= 8	+ iLngColCnt	' 그룹바이 컬럼행수 

	
			' -- 데이타셋을 기초로 배열로 재구성한다.
			ReDim arrTemp(iLngColCnt, iLngGrRowCnt)		
			' -- 데이타셋 행수 
			iLngRowCnt	= UBound(arrDataSet, 2)
		
			For iLngRow = 0 To 	iLngRowCnt
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+1, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(3, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+2, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(4, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+3, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(5, iLngRow)	' -- 열, 행, 값(썸)

			Next		
			sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- 다은키값 (변경분)

			' -- 조회 Row가 최대행수와 일치할때만 다음데이타 존재함 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If

			Redim TmpBuffer(iLngGrRowCnt)
			
			' -- 그룹바이 행으로 루핑 
			For iLngRow = 0 To 	iLngGrRowCnt						
				iStrData = Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- PROJECT_NO
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(6, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT_NM

				' -- 데이타그리드 출력(수정분)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next

				If Trim(arrRows(7, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
					sGrpTxt = sGrpTxt & arrRows(7, iLngRow) & gColSep & arrRows(8, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
				End If
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(8, iLngRow)))	' -- Row_Seq
				iStrData = iStrData & Chr(12)
				
				TmpBuffer(iLngRow) = iStrData
			Next
			
			iStrData = Join(TmpBuffer, "")

			If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet2(" & iMaxCols & ")" & vbCr
				Response.Write  "   Call Parent.SetGridHead2(""" & sTxt & """)" & vbCr
			End If
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey2 = """ & sRowSeq & """" & vbCr
			
			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor2(""" & sGrpTxt & """)" & vbCr
			End If			
		End If
		If sTypeFlag = "3" Then	' -- All 또는 3번 그리드		
			' -- 1번 그리드 조회 
			If sNextKey = "" Then	' -- 헤더 레코드셋 포함되어 옴 
				ColHeadRowNo = -1000
				arrColRow = oRs.GetRows()
				iLngRowCnt	= UBound(arrColRow, 2) 
				iLngColCnt	= UBound(arrColRow, 1) 
				
				For j = 0 To iLngRowCnt
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt = sTxt & arrColRow(0, j) & gColSep
					sTxt2 = sTxt2 & "매출수량" & gColSep
					sTxt2 = sTxt2 & "매출액" & gColSep
					sTxt2 = sTxt2 & "매출원가" & gColSep
					sTxt2 = sTxt2 & "매출이익" & gColSep
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
				sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep				
				' --- 각 화면별로 변경해야될 치환 문자열들..
				sTxt	= Replace(sTxt, "%CS", "합계")						
				
				Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
				' ----------------------------------------
			End If

			' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
			arrRows		= oRs.GetRows()		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			arrDataSet = oRs.GetRows()

			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
			iLngGrColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
		
			iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	* 4				' 컬럼수 
			iMaxCols	= 6	+ iLngColCnt	' 그룹바이 컬럼행수 

	
			' -- 데이타셋을 기초로 배열로 재구성한다.
			ReDim arrTemp(iLngColCnt, iLngGrRowCnt)		
			' -- 데이타셋 행수 
			iLngRowCnt	= UBound(arrDataSet, 2)
		
			For iLngRow = 0 To 	iLngRowCnt
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+1, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(3, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+2, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(4, iLngRow)	' -- 열, 행, 값(썸)
				arrTemp((CLng(arrDataSet(1, iLngRow))-1)*4+3, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(5, iLngRow)	' -- 열, 행, 값(썸)

			Next		
			sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- 다은키값 (변경분)

			' -- 조회 Row가 최대행수와 일치할때만 다음데이타 존재함 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If

			Redim TmpBuffer(iLngGrRowCnt)
			
			' -- 그룹바이 행으로 루핑 
			For iLngRow = 0 To 	iLngGrRowCnt						
				iStrData = 	      Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- PROJECT_NO
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_NM

				' -- 데이타그리드 출력(수정분)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next

				If Trim(arrRows(5, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
					sGrpTxt = sGrpTxt & arrRows(5, iLngRow) & gColSep & arrRows(6, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
				End If
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(6, iLngRow)))	' -- Row_Seq
				iStrData = iStrData & Chr(12)
				
				TmpBuffer(iLngRow) = iStrData
			Next
			
			iStrData = Join(TmpBuffer, "")

			If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet3(" & iMaxCols & ")" & vbCr
				Response.Write  "   Call Parent.SetGridHead3(""" & sTxt & """)" & vbCr
			End If
			Response.Write "	.frm1.vspdData3.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData3              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData3.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			
			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor3(""" & sGrpTxt & """)" & vbCr
			End If			
		End If
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write " </Script>	                        " & vbCr    
    ElseIf sNextKey = "" Then
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If
End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "품목계정소계")
		pTmp = Replace(pTmp , "%3", "품목소계")
'		pTmp = Replace(pTmp , "%4", "품목소계")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function

%>

