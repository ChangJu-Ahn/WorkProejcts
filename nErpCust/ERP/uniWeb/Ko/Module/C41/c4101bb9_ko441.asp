<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3101bb9
'*  4. Program Name         : 실제원가 계산 
'*  5. Program Desc         : 실제원가 계산 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Cho Ig sung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
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
	Dim sCostCd,sOrderNo,sItemCd
	Dim tmpC1
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt
	
	sStartDt	= Request("txtYYYYMM")		
	sNextKey	= Request("lgStrPrevKey")	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4101bA1_list_ko441"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))		
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

	If oRs.EoF and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If
	
    If Not oRs is nothing Then

		arrRows = oRs.GetRows()
		iLngRowCnt = UBound(arrRows,2) 
		iLngColCnt	= UBound(arrRows, 1) 

		'Response.Write 		iLngRowCnt & "=iLngRowCnt," & iLngColCnt & "=iLngColCnt"
		If iLngRowCnt < 0 Then 		
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			Exit Sub
		End If
		tmpC1=""
		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
		For iLngRow = 0 To 	iLngRowCnt	

				iStrData = iStrData & Chr(11) & 0
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(2, iLngRow))		
				IF arrRows(3, iLngRow) <>"" THEN 
					iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))				
				ELSE
					iStrData = iStrData & Chr(11) & gUsrID	
				END IF
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(4, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(5, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(6, iLngRow))
				
				'경고,오류 건수 항목 추가...2009.10.12...kbs
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(7, iLngRow))
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(8, iLngRow))
				iStrData = iStrData & Chr(11) & ""

				iStrData = iStrData & Chr(11) & Chr(12)
				
	Next
			
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 

	Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 	

	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If
       
End Sub	


%>
