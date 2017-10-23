<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :평균단가/재고평가내역 
'*  3. Program ID           : c4226mb1.asp
'*  4. Program Name         : 평균단가/재고평가내역 
'*  5. Program Desc         : 평균단가/재고평가내역 
'*  6. Modified date(First) : 2005-11-25
'*  7. Modified date(Last)  : 2005-11-25
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
	Dim sPlantCd,sItemAcct,sItemCd
	Dim tmpC1
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt
	
	sStartDt	= Request("txtYYYYMM")		
	sNextKey	= Request("lgStrPrevKey")	
	sPlantCd	= Request("txtPlant_cd")
	sItemAcct	= Request("txtItem_acct")	
	sItemCd	= Request("txtItem_cd")		

	If sPlantCd = "" Then sPlantCd = "%"
	If sItemAcct = "" Then sItemAcct = "%"
	If sItemCd = "" Then sItemCd = "%"

	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4226MA1_list"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_ACCT",	adVarXChar,	adParamInput, 2,Replace(sItemAcct, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 18,Replace(sItemCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 500)	
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
		
		If iLngRowCnt ="" Then 		
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			sRowSeq=""
			Exit Sub
		End If
		tmpC1=""
		sRowSeq = arrRows(UBound(arrRows, 1), iLngRowCnt)
		For iLngRow = 0 To 	iLngRowCnt	
				  
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(2, iLngRow))			
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(4, iLngRow))
				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(5, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(6, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(7, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)

				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(8, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(9, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(10, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)

				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(11, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(12, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(13, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)

				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(14, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)

				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(15, iLngRow),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(16, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(17, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(18, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(19, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(20, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(21, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
												
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(22, iLngRow))				
				iStrData = iStrData & Chr(11) & Chr(12)
								
	Next
			
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 
	
	Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 	
	Response.Write "	Call parent.SetQuerySpreadColor(""" & tmpC1 & """)" & vbCr
	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hplant_cd.value=""" & sPlantCd & """" & vbcr
	Response.Write "	.frm1.hItem_Acct.value=""" & sItemAcct & """" & vbcr
	Response.Write "	.frm1.hItem_cd.value=""" & sItemCd & """" & vbcr	
	
	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
	else	 		
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			sRowSeq=""
   End If
       
End Sub	


%>

