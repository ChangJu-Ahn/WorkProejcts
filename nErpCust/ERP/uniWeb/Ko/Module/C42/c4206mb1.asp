<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 원가 
'*  2. Function Name        : 공정별원가조회 
'*  3. Program ID           : c4206ma1
'*  4. Program Name         : 공정별원가조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : HJO
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
	Dim sPlantCd, sItemAcct, sProcType, sItemCd, sWcCd, sType, iBas, arrCol, iColDept, iiColSize
	
	iBas = 4	' -- 앞에 고정이 변할경우 대비 
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sYYYYMM, sEndDt
	
	sYYYYMM	= Request("txtYYYYMM")	
	sNextKey	= Request("lgStrPrevKey")
	
	sPlantCd	= Request("txtPlantCD")	

	sWcCd		= Request("txtWpCd")
	sType		= Request("txtOptionValue")				' -- 그리드 구분 
		
	If sYYYYMM = "" And sPlantCd = "" And sWcCd = "" And sType = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sPlantCd = ""	Then sPlantCd = "%"
	If sWcCd = ""		Then sWcCd = "%"
	If sType = ""		Then sType = "F"
	If sType = "A" Then 
		iiColSize = 20
	Else
		iiColSize = 28
	End If
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4206MA1" 		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 
		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6, sYYYYMM)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@OPTION_FLAG",	adVarXChar,	adParamInput, 1,sType)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@WC_CD",	adVarXChar,	adParamInput, 7,Replace(sWcCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 
		    
        Set oRs = lgObjComm.Execute
        
    End With
    'Response.Write "Err=" & Err.Description
    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If
		
    
    If Not oRs.EOF Then

		arrRows = oRs.GetRows()
		iLngRowCnt = UBound(arrRows, 2) 
		iLngColCnt	= UBound(arrRows, 1) 

		If iLngRowCnt =0 And sNextKey = "" Then 
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Exit Sub
		End If
		
		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
		
		For iLngRow = 0 To 	iLngRowCnt

			If sOptionFlag = "F" Then	' 28개씩 진행			
				For iColDept = 0 To 5	' -- 컬럼행뎁스가 6개 

					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- 공장 
					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- procur type
					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- 공정 
					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- 공정명				
					
					For iLngCol = (iBas +1) + (iiColSize * iColDept) To (iBas +iiColSize) + (iiColSize * iColDept) Step 4
						iStrData = iStrData & Chr(11) & Trim(arrRows(iLngCol, iLngRow))
						iStrData = iStrData & Chr(11) & arrRows(iLngCol+1, iLngRow)
						iStrData = iStrData & Chr(11) & arrRows(iLngCol+2, iLngRow)
						iStrData = iStrData & Chr(11) & arrRows(iLngCol+3, iLngRow)
					
					Next
					
					If arrRows(iLngColCnt - 1, iLngRow) <> "0" Then					' -- GroupNo|Row_Seq
						sGrpTxt = sGrpTxt & arrRows(iLngColCnt - 1, iLngRow) & gColSep & CStr(CDbl(arrRows(iLngColCnt, iLngRow))) & gRowSep
					End If
					iStrData = iStrData & Chr(11) & arrRows(iLngColCnt, iLngRowCnt)	' -- Row_Seq
					
					iStrData = iStrData & gRowSep
				Next
			Else 'T
				For iColDept = 0 To 5	' -- 컬럼행뎁스가 6개 

					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- 공장 
					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- procur type
					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- 공정 
					iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngColCnt-1, iLngRow))	' -- 공정명				
					
					For iLngCol = (iBas +1) + (iiColSize * iColDept) To (iBas +iiColSize) + (iiColSize * iColDept) Step 4
						iStrData = iStrData & Chr(11) & Trim(arrRows(iLngCol, iLngRow))
						iStrData = iStrData & Chr(11) & arrRows(iLngCol+1, iLngRow)
						iStrData = iStrData & Chr(11) & arrRows(iLngCol+2, iLngRow)
						iStrData = iStrData & Chr(11) & arrRows(iLngCol+3, iLngRow)					
					Next
					
					If arrRows(iLngColCnt - 1, iLngRow) <> "0" Then					' -- GroupNo|Row_Seq
						sGrpTxt = sGrpTxt & arrRows(iLngColCnt - 1, iLngRow) & gColSep & CStr(CDbl(arrRows(iLngColCnt, iLngRow))) & gRowSep
					End If
					iStrData = iStrData & Chr(11) & arrRows(iLngColCnt, iLngRowCnt)	' -- Row_Seq					
					iStrData = iStrData & gRowSep
				Next
			
			
			End If		
		Next
			
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
		
		If sNextKey <> "*" Then
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 	
		Else
			Response.Write "  parent.lgEOF = True                           " & vbCr
		End If

		If sGrpTxt <> "" Then
			Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
		End If
		
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    ElseIf sNextKey = "" Then 
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If
    
    If oRs.EOF And sNextKey <> "" Then
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write "  parent.lgEOF = True                           " & vbCr
		Response.Write  " </Script>                  " & vbCr
		
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp	
	pTmp = Replace(pLang , "%1", "합계")		
	pTmp = Replace(pTmp , "%2", "사내/외주소계")		
	pTmp = Replace(pTmp , "%3", "공정소계")
	pTmp = Replace(pTmp , "%M", "사내")
	pTmp = Replace(pTmp , "%O", "외주")

	ConvLang = pTmp
End Function
%>

