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
'*  6. Modified date(First) : 2005-12-01
'*  7. Modified date(Last)  : 2005-12-01
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
	Dim sPlantCd,sItemAcct,sItemCd,sEndDt
	Dim tmpC1
	Dim IntRetCD
	Dim strMsg_cd
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt
	
	sStartDt	= Request("txtFrom_YYYYMM")
	sEndDt = Request("txtTo_YYYYMM")		
	sNextKey	= Request("lgStrPrevKey")	
	sPlantCd	= Request("txtPlant_cd")
	sItemAcct	= Request("txtItem_acct")	
	sItemCd	= Request("txtItem_cd")		

	If sPlantCd = "" Then sPlantCd = "%"
	If sItemAcct = "" Then sItemAcct = "%"
	If sItemCd = "" Then sItemCd = "%"

	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4228MA1_list"		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		.Parameters.Append .CreateParameter("@FROM_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		.Parameters.Append .CreateParameter("@TO_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))				
		.Parameters.Append .CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		.Parameters.Append .CreateParameter("@ITEM_ACCT",	adVarXChar,	adParamInput, 2,Replace(sItemAcct, "'", "''"))
		.Parameters.Append .CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 18,Replace(sItemCd, "'", "''"))
		.Parameters.Append .CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 500)	
		.Parameters.Append .CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		.Parameters.Append .CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 
		.Parameters.Append .CreateParameter("@MSG_CD"     ,adVarChar,adParamOutput,6)
		    
        Set oRs = lgObjComm.Execute
        
    End With
    
     If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value

        If  IntRetCD <>0 Then
            strMsg_cd = lgObjComm.Parameters("@MSG_CD").Value
            Call DisplayMsgBox(strMsg_cd, vbInformation, "Batch Process Error", "", I_MKSCRIPT )  
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write "	.frm1.txtItem_Acct.Focus	" & vbCr 	
			Response.Write "	.lgStrPrevKey = """"" & vbCr 	
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
            Exit sub                                                           '☜: Protect system from crashing   
	'	Response.end
        End If
    Else
		If CheckSYSTEMError(Err, True) = True Then
			exit Sub
		End If
    End If
       
	If oRs.EoF and oRs.Bof  then
		If  sNextKey="" then
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
			oRs.Close
			Set oRs = Nothing
			Exit Sub
		End If
	End If
	
    'If Not oRs is nothing Then
    If Not oRs.BOF or Not oRs.Eof Then

		arrRows = oRs.GetRows()
		iLngRowCnt = UBound(arrRows,2) 
		iLngColCnt	= UBound(arrRows, 1) 
		
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
				  
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(0, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(1, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(2, iLngRow))			
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(3, iLngRow))
				
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(4, iLngRow))
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(5, iLngRow))
				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(6, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(7, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(8, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(9, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(10, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(11, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(12, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(13, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(14, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(15, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(16, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrRows(17, iLngRow),ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
							
				iStrData = iStrData & Chr(11) & ConvSPChars(arrRows(18, iLngRow))				
				iStrData = iStrData & Chr(11) & Chr(12)
				
	Next			
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 	
	
	Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 	
	Response.Write "	Call parent.SetQuerySpreadM()" & vbCr
	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hYYYYMM2.value=""" & sEndDt & """" & vbcr
	Response.Write "	.frm1.hplant_cd.value=""" & sPlantCd & """" & vbcr
	Response.Write "	.frm1.hItem_Acct.value=""" & sItemAcct & """" & vbcr
	Response.Write "	.frm1.hItem_cd.value=""" & sItemCd & """" & vbcr	
	
	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   Else
	Response.Write "	<script language=vbscript> " & vbcr
	Response.Write "		with parent	" & vbcr
	Response.Write "			.lgStrPrevKey = """"" & vbCr 	
	Response.Write "		end with" & vbcr
	Response.Write "	</script> "	& vbcr
   End If
       
End Sub	

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    
    If CheckSYSTEMError(pErr,True) = True Then
		ObjectContext.SetAbort
		'Call SetErrorStatus
    Else
		If CheckSQLError(pConn,True) = True Then
			ObjectContext.SetAbort
		'	Call SetErrorStatus
		End If
	End If
End Sub   
%>

