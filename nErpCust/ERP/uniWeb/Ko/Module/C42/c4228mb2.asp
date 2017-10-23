<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :월별단가추이 
'*  3. Program ID           : c4228mb1.asp
'*  4. Program Name         : 월별단가추이 
'*  5. Program Desc         : 월별단가추이 
'*  6. Modified date(First) : 2005-11-18
'*  7. Modified date(Last)  : 2005-11-18
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
	
	Call LoadBasisGlobalInf()	
	Call loadInfTB19029B("Q", "C", "NOCOOKIE","MB")
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim strFlag
	Dim sDeptCd,sWcCd, sEmpNo,tmpKey
	Dim sStartDt,sFrame
	
	sStartDt	= Request("txtYYYYMM")		
	
	'sWcCd	= Request("txtWc_cd")	
	sEmpNo	= Request("txtEmp_No")		
	sFrame=request("txtFrame")

	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)

	 If sFrame =1 THEN 
		Call SubBizQueryA()
     Else
		Call SubBizQueryB()
     
     End If
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryA()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey2, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	'sNextKey2	= Request("lgStrPrevKey2")	

	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4254MA1_T1_DTL"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))					
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@EMP_NO",	adVarXChar,	adParamInput, 13,Replace(sEmpNo, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace("*", "'", "''"))
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

	If oRs.EoF and oRs.Bof then
		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If

    If Not oRs is nothing Then

		' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
		arrRows		= oRs.GetRows()
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols

		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' 컬럼수 
		
		' -- 데이타셋 행수 
		iLngRowCnt	= UBound(arrRows, 2)
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngRowCnt < 0 Then 

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			'Response.Write  "	.lgStrPrevKey2=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)

		For iLngRow = 0 To 	iLngRowCnt			
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(2, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
					
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))					
				iStrData = iStrData & Chr(12)	
		Next
		
			
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr	
	Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 		 
	'Response.Write "	.lgStrPrevKey2 = """ & sRowSeq & """" & vbCr 		

	'Response.Write "	Call .DbDtlQueryOk()    " & vbcr

	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If
       
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryB()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey2, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4254MA1_T2_dtl"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@EMP_NO",	adVarXChar,	adParamInput, 13,Replace(sEmpNo, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace("*", "'", "''"))
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

	If oRs.EoF and oRs.Bof  then
		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If	

    If Not oRs is nothing Then
			' --- 헤더 출력(변경분) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
						
			For j = 0 To iLngColCnt
				For i = 0 To   iLngRowCnt 
					sTxt = sTxt & arrColRow(j, i) & gColSep 
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
			Next
			
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			' ----------------------------------------
		
		' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		iLngGrRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
		iLngGrColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
	
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' 컬럼수 
		
		
		iMaxCols	= 3	+ iLngColCnt	' 그룹바이 컬럼행수 
			
		' -- 데이타셋을 기초로 배열로 재구성한다.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)

		' -- 데이타셋 행수 
		iLngRowCnt	= UBound(arrDataSet, 2)
	
		For iLngRow = 0 To 	iLngRowCnt
			arrTemp(CLng(arrDataSet(1, iLngRow)), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(4, iLngRow)	' -- 행, 열, 값(썸)
		Next
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt < 0 Then 

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			'Response.Write  "	.lgStrPrevKey2=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			Exit Sub
		End If

		' -- 그룹바이 행으로 루핑 
		For iLngRow = 0 To 	iLngGrRowCnt	
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
					
				' -- 데이타그리드 출력(수정분)
				For iLngCol = 1 To iLngColCnt 
					Response.Write ILNGCOLCNT & "=LNGCOLCNT<BR>"
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next					
				iStrData = iStrData & Chr(11) & arrRows(2, iLngRow)
				iStrData = iStrData & Chr(12)				
		Next			
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.vspdData4.MaxCols = """ & (iMaxCols) & """" & vbCr 	 	
	Response.Write "	.frm1.vspdData4.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData4              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		
	Response.Write "	.frm1.vspdData4.ReDraw = True					" & vbCr 
	Response.Write  "   Call Parent.SetGridHead2(""" & sTxt & """)" & vbCr

	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hDept_Cd.value=""" & sDeptCd & """" & vbcr	
	Response.write "	Call parent.ReInitSpreadSheet2() " & vbcr
	Response.Write "	.frm1.vspdData4.style.display=block"  & vbcr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If			
		
End Sub	


%>

