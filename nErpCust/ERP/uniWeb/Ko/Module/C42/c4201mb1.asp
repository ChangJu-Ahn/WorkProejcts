<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 원가 
'*  2. Function Name        : 제조원가명세서조회 
'*  3. Program ID           : c4201mb1
'*  4. Program Name         : 제조원가명세서조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seong Bae
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
	Dim sPlantCd, sCostCd, sType, sTypeFlag, sOptionFlag, arrTmp

	Dim C_SHEET_MAX_CONT 
	
	C_SHEET_MAX_CONT = 1000
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")
	sNextKey	= Request("lgStrPrevKey")
	
	sPlantCd	= Request("txtPLANT_CD")	
	sCostCd		= Request("txtCOST_CD")	
	sType		= Request("rdoTYPE")	
	sTypeFlag	= Request("rdoTYPE_FLAG")
		
	If sStartDt = "" And sEndDt = ""  And sPlantCd = "" And sCostCd = "" And sTypeFlag = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sPlantCd = "" Then sPlantCd = "%"
	If sCostCd = "" Then sCostCd = "%"
	If sTypeFlag = "1" Or sTypeFlag = "2" Then
		sOptionFlag = "A"
	Else
		sOptionFlag = "S"
	End If

	' -- sNextKey 변경사항 
	If Instr(1, sNextKey, "*") > 0 Then
		arrTmp = Split(sNextKey, gColSep)
		sNextKey = arrTmp(0)
		C_SHEET_MAX_CONT = 32000
	End If

    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4201MA1_TYPE" & sType		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 10,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 10,Replace(sEndDt, "'", "''"))
		
		If sType = "2" Then
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 10,Replace(sPlantCd, "'", "''"))
		End If
		
		If sType = "3" Then
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 20,Replace(sCostCd, "'", "''"))
		End If
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TYPE_FLAG",	adVarXChar,	adParamInput, 10,Replace(sTypeFlag, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OPTION_FLAG",	adVarXChar,	adParamInput, 10,Replace(sOptionFlag, "'", "''"))
		
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
    If Not oRs is Nothing Then

		If sNextKey = "" And (sType = "2" Or sType = "3") Then	' -- 헤더 레코드셋 포함되어 옴 

			' --- 헤더 출력(변경분) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
			
			For j = 0 To iLngColCnt 
				For i = 0 To iLngRowCnt 
					sTxt = sTxt & arrColRow(j, i) & gColSep 
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
			Next
			
			' --- 각 화면별로 변경해야될 치환 문자열들..
			'sTxt	= Replace(sTxt, "%4", "합계")
					
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			' ----------------------------------------
		End If

		' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
		arrRows		= oRs.GetRows()
		
		If sType = "2" Or sType = "3" Then
		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
		
			Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
			arrDataSet = oRs.GetRows()

			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
			iLngGrColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
		
			iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' 컬럼수 

			' ----------------------------------------------------------
			If iLngGrRowCnt = 0 Then	' -- 데이타 없을 경우 
				Response.Write " <Script Language=vbscript>	                        " & vbCr
				Response.Write "	Parent.lgStrPrevKey = """"	                    " & vbCr
				Response.Write " </Script>	                        " & vbCr
			
				Exit Sub
			End If

		
			' -- 변경할곳 : 컬럼수(그룹바이 고정컬럼수)
			If sType = "2" Then
				If sTypeFlag = "1" Or sTypeFlag = "3" Then
					iMaxCols	= 7	+ iLngColCnt	' 그룹바이 컬럼행수 
				Else
					iMaxCols	= 9	+ iLngColCnt	' 그룹바이 컬럼행수 
				End If
			Else
				If sTypeFlag = "1" Or sTypeFlag = "3" Then
					iMaxCols	= 5	+ iLngColCnt	' 그룹바이 컬럼행수 
				Else
					iMaxCols	= 7	+ iLngColCnt	' 그룹바이 컬럼행수 
				End If
			End If
		
			' -- 데이타셋을 기초로 배열로 재구성한다.
			ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
			' -- 데이타셋 행수 
			iLngRowCnt	= UBound(arrDataSet, 2)
		
			For iLngRow = 0 To 	iLngRowCnt
				arrTemp((CLng(arrDataSet(1, iLngRow))-1), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- 열, 행, 값(썸)
			Next
		
			' ----------------------------------------------------------
			sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- 다은키값 (변경분)
			' -- 조회 Row가 최대행수와 일치할때만 다음데이타 존재함 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If
			
			' -- 그룹바이 행으로 루핑 
			For iLngRow = 0 To 	iLngGrRowCnt
					
				If sType = "2" Then
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(1, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(3, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(4, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(5, iLngRow)))
	
					If sTypeFlag = "2" Or sTypeFlag = "4" Then ' 집계 
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(6, iLngRow)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(7, iLngRow)))
					End If
					
					' -- 데이타그리드 출력(수정분)
					For iLngCol = 0 To iLngColCnt -1
						iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
					Next

					If sTypeFlag = "2" Or sTypeFlag = "4" Then ' 집계 
						iStrData = iStrData & Chr(11) & arrRows(8, iLngRow)
					Else
						iStrData = iStrData & Chr(11) & arrRows(6, iLngRow)
					End If
				Else	 ' -- B 인 경우 
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(1, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(3, iLngRow)))

					If sTypeFlag = "2" Or sTypeFlag = "4" Then ' 집계 
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(4, iLngRow)))
						iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(5, iLngRow)))
					End If

					' -- 데이타그리드 출력(수정분)
					For iLngCol = 0 To iLngColCnt -1
						iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
					Next

					If sTypeFlag = "2" Or sTypeFlag = "4" Then ' 집계 
						iStrData = iStrData & Chr(11) & arrRows(6, iLngRow)
					Else
						iStrData = iStrData & Chr(11) & arrRows(5, iLngRow)
					End If
				End If				
				
				iStrData = iStrData & Chr(12)
			Next

		Else
			' -- 집계단위가 Company 인 경우 ---------------
			iLngRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
			iLngColCnt = UBound(arrRows, 1)				' 그룹바이 열수 

			' ----------------------------------------------------------
			If iLngRowCnt = 0 Then	' -- 데이타 없을 경우 
				Response.Write " <Script Language=vbscript>	                        " & vbCr
				Response.Write "	Parent.lgStrPrevKey = """"	                    " & vbCr
				Response.Write " </Script>	                        " & vbCr
				Exit Sub
			End If

			sRowSeq = arrRows(iLngColCnt , iLngRowCnt)		' -- 다은키값 (변경분)

			' -- 조회 Row가 최대행수와 일치할때만 다음데이타 존재함 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If
			
			' -- 그룹바이 행으로 루핑 
			For iLngRow = 0 To 	iLngRowCnt
					
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))		
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(1, iLngRow)))		
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))	
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(3, iLngRow)))	
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(4, iLngRow)))	
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(5, iLngRow)))	
				
				If sTypeFlag = "1" Or sTypeFlag = "3" Then ' 집계 
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(6, iLngRow)))	' -- 금액 
					iStrData = iStrData & Chr(11) & arrRows(7, iLngRow) & gRowSep			' -- ROW_SEQ
				Else
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(6, iLngRow)))	' -- 계정 
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(7, iLngRow)))	' -- 계정명 
					iStrData = iStrData & Chr(11) & arrRows(8, iLngRow)						' -- 금액 
					iStrData = iStrData & Chr(11) & arrRows(9, iLngRow) & gRowSep			' -- ROW_SEQ
				End If
			Next	

		End if
			
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
		Response.Write "	.frm1.hEND_DT.value = """ & sEndDt & """" & vbCr 	
		Response.Write "	.frm1.hPLANT_CD.value = """ & sPlantCd & """" & vbCr
		Response.Write "	.frm1.hCOST_CD.value = """ & sCostCd & """" & vbCr
		Response.Write "	.frm1.hTYPE.value = """ & sType & """" & vbCr
		Response.Write "	.frm1.hTYPE_FLAG.value = """ & sTypeFlag & """" & vbCr 	 	 	 	

		If (sType = "2" Or sType = "3") And sNextKey = "" Then
			' --  화면 초기화 생성 
			Response.Write  "   Call Parent.InitSpreadSheet" & sType & "(" & iMaxCols & ")" & vbCr
			
			' -- 헤더 생성 
			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead" & sType & "(""" & sTxt & """)" & vbCr
			End If
		ElseIf sType = "1" And sNextKey = "" Then
			Response.Write  "   Call Parent.InitSpreadSheet()" & vbCr
		End If
		
		Response.Write "	.frm1.vspdData" & sType & ".ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData" & sType & "              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		Response.Write "	.frm1.vspdData" & sType & ".ReDraw = True					" & vbCr 		
		
		' -- 엑셀조회의 경우 
		If 	C_SHEET_MAX_CONT = 100 Then
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
		Else
			Response.Write "	.lgStrPrevKey = ""*""" & vbCr
		End If

		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    ElseIf sGrid = "A" And sNextKey = "" Then
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "원가요소별소계")
		pTmp = Replace(pTmp , "%3", "계정별소계")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function


Function ConvLang2(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "관리항목별 소계")
	Else
		pTmp = pLang
	End If
	ConvLang2 = pTmp
End Function
%>

