<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :공통재료비투입현황 
'*  3. Program ID           : c4237mb1.asp
'*  4. Program Name         : 공통재료비투입현황 
'*  5. Program Desc         : 공통재료비투입현황 
'*  6. Modified date(First) : 2005-12-30
'*  7. Modified date(Last)  : 2005-12-22
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : HJO
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

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt, TmpBuffer
	Dim sCostCd,sOrderNo,sItemCd, sWcCd, sItemAcct,sPlantCd, sMovType
	Dim tmpC1
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt
	
	sStartDt	= TRIM(Request("txtYYYYMM")		)
	sNextKey	= TRIM(Request("lgStrPrevKey")	)
	sCostCd	= TRIM(Request("txtCost_cd"))
	sItemCd	= TRIM(Request("txtItem_cd")		)
	sWcCd =TRIM(Request("txtWc_cd"))
	sItemAcct= TRIM(Request("txtItemAcct"))
	sPlantCd=trim(request("txtPlant_Cd"))
	sMovType=trim(request("txtMovType"))

	If sCostCd = "" Then sCostCd = "%"	
	If sItemCd = "" Then sItemCd = "%"
	If sWcCd = "" Then sWcCd = "%"
	If sItemAcct = "" Then sItemAcct = "%"
	If sPlantCd = "" Then sPlantCd = "%"
	If sMovType = "" Then sMovType = "%"
		
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4237MA1_LIST"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_ACCT",	adVarXChar,	adParamInput, 2,Replace(sItemAcct, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 18,Replace(sItemCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@WC_CD",	adVarXChar,	adParamInput, 7,Replace(sWCCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@MOV_TYPE",	adVarXChar,	adParamInput, 3,Replace(sMovType, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 1000)	
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

	If oRs.EoF and oRs.Bof then 
		 If sNextKey="" then
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			oRs.Close
			Set oRs = Nothing
			Exit Sub
		Else
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			Set oRs = Nothing
			Exit Sub
		End If
	End If
	
    If Not oRs is nothing Then

		arrRows = oRs.GetRows()
		iLngRowCnt = UBound(arrRows,2) 
		iLngColCnt	= UBound(arrRows, 1) 
		' -- 데이타셋을 기초로 배열로 재구성한다.
		ReDim TmpBuffer(iLngRowCnt)

		tmpC1=""
		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
		
		For iLngRow = 0 To 	iLngRowCnt	
				If ConvSPChars(arrRows(0, iLngRow))="%1" then 
					tmpC1=tmpC1 & ConvSPChars(arrRows(0, iLngRow)) & gColsep &  0 & gColsep & arrRows(13, iLngRow) & gRowSep
				elseIf  ConvSPChars(arrRows(1, iLngRow))="%2" then 
					tmpC1=tmpC1 &  ConvSPChars(arrRows(1, iLngRow)) & gColsep & 1 & gColsep & arrRows(13, iLngRow) & gRowSep
				elseIf  ConvSPChars(arrRows(3, iLngRow))="%3" then 
					tmpC1=tmpC1 &  ConvSPChars(arrRows(3, iLngRow)) & gColsep & 3 & gColsep & arrRows(13, iLngRow) & gRowSep
				elseIf  ConvSPChars(arrRows(5, iLngRow))="%4" then 
					tmpC1=tmpC1 &  ConvSPChars(arrRows(5, iLngRow)) & gColsep & 5 & gColsep & arrRows(13, iLngRow) & gRowSep
				elseIf  ConvSPChars(arrRows(7, iLngRow))="%5" then 
					tmpC1=tmpC1 &  ConvSPChars(arrRows(7, iLngRow)) & gColsep & 7 & gColsep & arrRows(13, iLngRow) & gRowSep	
				
				End If
				iStrData = ""
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(0, iLngRow)),"%1","합계")
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(1, iLngRow)),"%2","C/C소계")
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(2, iLngRow))			
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(3, iLngRow)),"%3","작업장소계")				
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(4, iLngRow))
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(5, iLngRow)),"%4","계정소계")
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(6, iLngRow))
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(7, iLngRow)),"%5","품목소계")
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(8, iLngRow))
				iStrData = iStrData & Chr(11) &  ConvSPChars(arrRows(9, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(10, iLngRow))				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(11, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(12, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(13, iLngRow))		
				iStrData = iStrData & Chr(11) & Chr(12)
				TmpBuffer(iLngRow) = iStrData				
		Next		
		iStrData = Join(TmpBuffer, "")		
		
	Response.Write " <Script Language=vbscript>							" & vbCr
	Response.Write " With parent										" & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData				" & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData			& """" & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """"			& vbCr 	
	Response.Write "	Call parent.SetQuerySpreadColor(""" & tmpC1 & """)" & vbCr
	Response.Write "	.frm1.hYYYYMM.value=	""" & sStartDt	& """" & vbcr
	Response.Write "	.frm1.hCost_cd.value=	""" & sCostCd	& """" & vbcr	
	Response.Write "	.frm1.hItem_cd.value=	""" & sItemCd	& """" & vbcr
	Response.Write "	.frm1.hWc_Cd.value=		""" & sWcCd		& """" & vbcr		
	Response.Write "	.frm1.hItemAcct.value=	""" & sItemAcct	& """" & vbcr
	Response.Write "	.frm1.hPlant_cd.value=	""" & sPlantCd	& """" & vbcr		
	Response.Write "	.frm1.hMovType.value=	""" & sMovType	& """" & vbcr		
	Response.Write  "   Call Parent.DbQueryOk()							" & vbCr
	Response.Write " End With											" & vbCr
	Response.Write  " </Script>											" & vbCr
   End If
       
End Sub	


%>

