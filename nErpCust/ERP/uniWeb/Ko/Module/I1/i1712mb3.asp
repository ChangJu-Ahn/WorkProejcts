<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Batch Posting µî·Ï asp
'*  2. Function Name        : 
'*  3. Program ID           : i1712mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 

'*  7. Modified date(First) : 2004/10/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"
'*                            this mark(¢Á) Means that "may  change"
'*                            this mark(¡Ù) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->

<%
	Call LoadBasisGlobalInf()

	On Error Resume Next
    Call HideStatusWnd
	Dim PI1G201			
	
	Dim LngRow
	Dim iMaxRow
	
	Dim arrRowVal		
	Dim arrColVal		
	Dim strStatus		

    '-----------------------
    'IMPORTS View
    '-----------------------
	Dim I1_ief_supplied_select_char
	Dim I2_i_goods_mvmt_header
	Const I136_I1_biz_area_cd = 0
	Const I136_I1_mvmt_yymm = 1
	Const I136_I1_cost_mvmt_flag = 2
	ReDim I2_i_goods_mvmt_header(I136_I1_cost_mvmt_flag)
	
	Dim I3_gl_info 
    Const I136_I3_org_change_id = 0
    Const I136_I3_dept_Cd = 1
    Const I136_I3_post_dt = 2
	ReDim I3_gl_info(I136_I3_post_dt)

	'-----------------------
	'EXPORTS View
	'-----------------------
	Dim iErrorPosition
 	Dim itxtSpread
   

    itxtSpread = ""
             
    itxtSpread = Request("txtSpread")

	'-----------------------
	'Data manipulate area
	'-----------------------											
    I1_ief_supplied_select_char = "C"
	I2_i_goods_mvmt_header(I136_I1_biz_area_cd) = Trim(Request("txtBizCd"))
	I2_i_goods_mvmt_header(I136_I1_mvmt_yymm) = Request("hYyMm")
	I2_i_goods_mvmt_header(I136_I1_cost_mvmt_flag) = Request("hCostFlag")
	
	I3_gl_info(I136_I3_org_change_id) = Trim(Request("hOrgChangeId"))
	I3_gl_info(I136_I3_dept_Cd) = Trim(Request("txtDeptCd"))
	I3_gl_info(I136_I3_post_dt)	= UniconvDate(Trim(Request("GIDate")))
	
	If itxtSpread <> "" Then
	
		Set PI1G201 = Server.CreateObject("PI1G201.cIMonthlyBchPostSvr")
    
		If CheckSYSTEMError(Err, True) = True Then
			Response.End
		End If
		
			
		Call PI1G201.I_MONTHLY_BATCH_POST_SVR(gStrGlobalCollection, _
									I1_ief_supplied_select_char, _
									I2_i_goods_mvmt_header, _
									I3_gl_info, _
									itxtSpread, _
									, _
									iErrorPosition)
			
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If CheckSYSTEMError(Err, True) = True Then
			Set PI1G201 = Nothing

			Response.Write "<Script Language=vbscript> " & vbCrlf
			Response.Write "Parent.DbSaveOk " & vbCrlf
			Response.Write "</Script>" & vbCrlf
			Response.End
		End If
			
		Set PI1G201 = Nothing
	End If
	
	Response.Write "<Script Language=vbscript> " & vbCrlf
    Response.Write "Parent.DbSaveOk " & vbCrlf
    Response.Write "</Script>" & vbCrlf
	Response.End
%>