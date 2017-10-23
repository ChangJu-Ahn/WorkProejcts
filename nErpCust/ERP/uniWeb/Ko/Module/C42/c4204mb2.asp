<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :품목/오더별 실제원가조회 
'*  3. Program ID           : c4204mb2.asp
'*  4. Program Name         : 품목/오더별 실제원가 조회 
'*  5. Program Desc         : 품목/오더별 실제원가 조회 
'*  6. Modified date(First) : 2005-10-04
'*  7. Modified date(Last)  : 2005-10-04
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
	Dim sPlantCd, sItemAcct, sItemCd, sType,sCostCd, sWorkCd
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")
		
	sPlantCd	= Request("txtPLANT_CD")	
	sItemCd		= Request("txtITEM_CD")
	sCostCd		= Request("txtCC_CD")
	sWorkCd		= Request("txtWork_CD")
	sType		= Request("rdoTYPE")			
		
	If sStartDt = "" And sEndDt = ""  And sPlantCd = "" And sItemAcct = "" And sItemCd = ""  Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sWorkCd = "" Then sWorkCd = "%"
	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4204MA1_DTL" ' & sType		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 10,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 10,Replace(sEndDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 18,Replace(sItemCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@WORK_CD",	adVarXChar,	adParamInput, 10,Replace(sWorkCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@GUBUN",	adVarXChar,	adParamInput, 2,Replace(sType, "'", "''"))
		
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
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If	
    
    If Not oRs is nothing Then

		arrRows = oRs.GetRows()
		iLngRowCnt = UBound(arrRows, 2) 
		iLngColCnt	= UBound(arrRows, 1) 
		
		If iLngRowCnt ="" Then 		Exit Sub
	
		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)

		For iLngRow = 0 To 	iLngRowCnt
			IF iLngRow<sRowSeq Then
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(0, iLngRow)),"%1","합계")
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
				If ConvSPChars(arrRows(2, iLngRow))="%11" Then 
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(2, iLngRow)),"%11","사내")
				Else
				iStrData = iStrData & Chr(11) & replace(ConvSPChars(arrRows(2, iLngRow)),"%12","외주")
				End IF
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(4, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(5, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(6, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(7, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(8, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(9, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(10, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				If sType= "A" Then				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(12, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(13, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(14, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(15, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(16, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(18, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)																				
				Else
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(11, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(12, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(13, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(14, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(15, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(16, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(17, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)						
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(18, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				End If
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(19, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(20, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(21, iLngRow))

				iStrData = iStrData & Chr(11) & Chr(12)
		End IF
	Next
	
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
	
	Response.Write "	Call parent.SetQuerySpreadColor2(""" & sRowSeq & """)" & vbCr
	Response.Write  "   Call Parent.DbDtlQueryOk		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If
       
End Sub	


%>

