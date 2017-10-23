<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%Call LoadBasisGlobalInf()
Call loadInfTB19029B("I", "*","NOCOOKIE","BB")%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizBatch()

    Dim intRetCD
	Dim iStrWorkType	' 작업유형 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IntRetCD = 0

	iStrWorkType = Request("txtWorkType")

    With lgObjComm
		.CommandTimeout = 0
		' 등록 
		If iStrWorkType = "C" Then
			Call CreateTaxBill
		' 삭제 
		Else
			Call DeleteTaxBill
		End If

        .CommandType = adCmdStoredProc

        .Execute ,, adExecuteNoRecords
        
    End With
    
    If CheckSYSTEMError(Err,True) = True Then
       IntRetCD = -1
       Exit Sub
    End If
    
    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
    
    If CDbl(intRetCD) = 0 Then
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Parent.ExeReflectOk  " & vbCr
       Response.Write  " </Script>                  " & vbCr
    Else
       Call DisplayMsgBox(IntRetCd, vbInformation, "", "", I_MKSCRIPT)
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Call parent.SetFocusToDocument(""M"")  " & vbCr
       Response.Write  "       parent.frm1.txtFromDt.Focus  " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If

End Sub	

' 세금계산서 등록 
Sub CreateTaxBill()
	With lgObjComm
		.CommandText = "dbo.usp_s_CreateTaxbill_Batch"

	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
	    .Parameters.Append .CreateParameter("@fr_bill_dt",  adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtFromDt")))
	    .Parameters.Append .CreateParameter("@to_bill_dt",  adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtToDt")))

	    ' 발행처 
	    IF Len(Trim(Request("txtBillToParty"))) Then
		    .Parameters.Append .CreateParameter("@bill_to_party", adVarXChar,adParamInput,10,Replace(Request("txtBillToParty"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@bill_to_party", adVarXChar,adParamInput,10)
	    End If
	    ' 세금신고사업장 
	    IF Len(Trim(Request("txtTaxbizArea"))) Then
		    .Parameters.Append .CreateParameter("@tax_biz_area", adVarXChar,adParamInput,10,Replace(Request("txtTaxbizArea"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@tax_biz_area", adVarXChar,adParamInput,10)
	    End If
	    ' 영업그룹 
	    IF Len(Trim(Request("txtSalesGrp"))) Then
		    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,Replace(Request("txtSalesGrp"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4)
	    End If
	    ' 매출채권유형 
	    IF Len(Trim(Request("txtBillType"))) Then
		    .Parameters.Append .CreateParameter("@bill_type", adVarXChar,adParamInput,4,Replace(Request("txtBillType"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@bill_type", adVarXChar,adParamInput,4)
	    End If
		' B/L 포함여부	    
	    .Parameters.Append .CreateParameter("@bl_flag",   adXChar,adParamInput,1,Request("txtBLFlag"))
	    ' 발행여부 
	    .Parameters.Append .CreateParameter("@post_flag", adXChar,adParamInput,1,Request("txtPostFlag"))
	    ' 발행일 
	    .Parameters.Append .CreateParameter("@issued_dt",	adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtIssuedDt")))
	    ' VAT 적용기준 
	    .Parameters.Append .CreateParameter("@vat_calc_type", adXChar,adParamInput,1,Request("txtVatCalcType"))
	    ' 영업그룹별 집계 
	    .Parameters.Append .CreateParameter("@by_SalesGrp",   adXChar,adParamInput,1,Request("txtBySalesGrp"))
	    ' 매출채권유형별 집계 
	    .Parameters.Append .CreateParameter("@by_BillType",   adXChar,adParamInput,1,Request("txtByBillType"))
	    ' 매출번호별 집계 
	    .Parameters.Append .CreateParameter("@by_BillNo",     adXChar,adParamInput,1,Request("txtByBillNo"))
	    ' 부가세 유형 혼합여부 
	    .Parameters.Append .CreateParameter("@by_OnlyBillNo", adXChar,adParamInput,1,Request("txtByOnlyBillNo"))
		' 호출여부 
	    .Parameters.Append .CreateParameter("@CalledFlag",    adXChar,adParamInput,1, "N")
	    ' User ID
	    .Parameters.Append .CreateParameter("@user_id",       adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
    End With
   
End Sub

' 세금계산서 삭제 
Sub DeleteTaxBill()
	With lgObjComm
		.CommandText = "dbo.usp_s_DeleteTaxBillByBatch"

	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
	    .Parameters.Append .CreateParameter("@fr_issued_dt",  adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtFromDt")))
	    .Parameters.Append .CreateParameter("@to_issued_dt",  adDBTimeStamp,adParamInput,, UNIConvDate(Request("txtToDt")))

	    ' 발행처 
	    IF Len(Trim(Request("txtBillToParty"))) Then
		    .Parameters.Append .CreateParameter("@bp_cd", adVarXChar,adParamInput,10,Replace(Request("txtBillToParty"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@bp_cd", adVarXChar,adParamInput,10)
	    End If
	    ' 세금신고사업장 
	    IF Len(Trim(Request("txtTaxbizArea"))) Then
		    .Parameters.Append .CreateParameter("@report_biz_area", adVarXChar,adParamInput,10,Replace(Request("txtTaxbizArea"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@report_biz_area", adVarXChar,adParamInput,10)
	    End If
	    ' 영업그룹 
	    IF Len(Trim(Request("txtSalesGrp"))) Then
		    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,Replace(Request("txtSalesGrp"), "'", "''"))
		Else
		    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4)
	    End If

	    ' 발행여부 
	    .Parameters.Append .CreateParameter("@post_flag", adXChar,adParamInput,1,Request("txtPostFlag"))
	    ' User ID
	    .Parameters.Append .CreateParameter("@user_id",   adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
    End With
   
End Sub

'============================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

