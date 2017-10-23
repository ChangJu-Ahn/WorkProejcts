<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'********************************************************************************************************
'*  1. Module Name          : Inventory           
'*  2. Function Name        : DocumentNo Popup Business Part         
'*  3. Program ID              : i1111bp1.asp            
'*  4. Program Name         :               
'*  5. Program Desc         : 수불번호팝업             
'*  7. Modified date(First) : 2000/04/18            
'*  8. Modified date(Last)  : 2005/08/05          
'*  9. Modifier (First)     :           
'* 10. Modifier (Last)      : Lee Seung Wook          
'* 11. Comment              :                   
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"         
'*                            this mark(⊙) Means that "may  change"         
'*                            this mark(☆) Means that "must change"         
'* 13. History              :                   
'*                            2000/04/18 : Coding Start             
'********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%        
Call LoadBasisGlobalInf()
Call HideStatusWnd 

err.clear
On Error Resume Next     

DIm strData
Dim lgStrPrevKey 
Dim LngMaxRow  
Dim LngRow
Dim pPI1G150
Dim PvArr

Const C_SHEETMAXROWS_D = 100

    '-----------------------
    'IMPORTS View
    '-----------------------
    Dim I1_fr_prod_work_set_temp_timestamp
    Dim I2_to_prod_work_set_temp_timestamp
    Dim I3_i_goods_movement_header
		Const I125_I3_mov_type			= 0
		Const I125_I3_item_document_no	= 1
		Const I125_I3_document_year		= 2
		Const I125_I3_trns_type			= 3
		Const I125_I3_plant_cd			= 4    
	ReDim I3_i_goods_movement_header(I125_I3_plant_cd)
 '-----------------------
 'EXPORTS View
 '-----------------------
    Dim E3_i_goods_movement_header
		Const I125_E3_item_document_no	= 0
		Const I125_E3_document_year		= 1
		Const I125_E3_mov_type			= 2
		Const I125_E3_trns_type			= 3
    Dim EG1_group_export
		Const I125_EG1_E1_i_goods_movement_header_item_document_no	= 0
		Const I125_EG1_E1_i_goods_movement_header_document_year		= 1
		Const I125_EG1_E1_i_goods_movement_header_trns_type			= 2
		Const I125_EG1_E1_i_goods_movement_header_mov_type			= 3
		Const I125_EG1_E1_i_goods_movement_header_document_dt		= 4
		Const I125_EG1_E1_i_goods_movement_header_pos_dt			= 5
		Const I125_EG1_E1_i_goods_movement_header_document_text		= 6
		Const I125_EG1_E1_i_goods_movement_header_plant_cd			= 7
		Const I125_EG1_E1_i_goods_movement_header_biz_area_cd		= 8
		Const I125_EG1_E1_i_goods_movement_header_post_flag			= 9
		Const I125_EG1_E1_i_goods_movement_header_cost_cd			= 10
		Const I125_EG1_E2_b_plant_plant_nm							= 11
		Const I125_EG1_E3_b_minor_minor_nm							= 12


	lgStrPrevKey = Request("lgStrPrevKey")
	 
	I3_i_goods_movement_header(I125_I3_item_document_no) = Request("txtDocumentNo")
	I3_i_goods_movement_header(I125_I3_document_year)    = Request("txtYear") 
	I3_i_goods_movement_header(I125_I3_mov_type)         = Request("txtMovType")
	I3_i_goods_movement_header(I125_I3_trns_type)        = Request("txtTrnsType")
	I3_i_goods_movement_header(I125_I3_plant_cd)         = Request("txtPlantCd")
	I1_fr_prod_work_set_temp_timestamp                   = UNIConvDate(Request("txtFromDt"))  
	I2_to_prod_work_set_temp_timestamp = ""
	
	if Request("txtToDt") <> "" then I2_to_prod_work_set_temp_timestamp = UNIConvDate(Request("txtToDt"))
	if lgStrPrevKey <> "" then I3_i_goods_movement_header(I125_I3_item_document_no) = lgStrPrevKey

	Set pPI1G150 = Server.CreateObject("PI1G150.cIListGoodsMvmtSvr")
	    
	If CheckSYSTEMError(Err, True) = True Then
	   Response.End          
	End If    
 
	Call pPI1G150.CAB_I_LIST_GOODS_MVMT_HEADER (gStrGlobalCollection, _
												C_SHEETMAXROWS_D, _
												I1_fr_prod_work_set_temp_timestamp, _
												I2_to_prod_work_set_temp_timestamp, _
												I3_i_goods_movement_header, _
												E3_i_goods_movement_header, _
												EG1_group_export)
	'-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
		Set pPI1G150 = Nothing          
		Response.End             
	End If

	Set pPI1G150 = Nothing

	if isEmpty(EG1_group_export) then
		Response.End             
	End If
 
 
	LngMaxRow = CLng(Request("txtMaxRows")) + 1
	
	ReDim PvArr(ubound(EG1_group_export,1))
	
	For LngRow = 0 To ubound(EG1_group_export,1)
		strData = Chr(11) & ConvSPChars(EG1_group_export(LngRow, I125_EG1_E1_i_goods_movement_header_item_document_no)) & _
				Chr(11) & EG1_group_export(LngRow, I125_EG1_E1_i_goods_movement_header_document_year) & _
				Chr(11) & UNIDateClientFormat(EG1_group_export(LngRow, I125_EG1_E1_i_goods_movement_header_document_dt)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I125_EG1_E1_i_goods_movement_header_mov_type)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I125_EG1_E3_b_minor_minor_nm)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I125_EG1_E1_i_goods_movement_header_plant_cd)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I125_EG1_E1_i_goods_movement_header_document_text)) & _
				Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		PvArr(LngRow) = strData
	Next
	strData = Join(PvArr, "")
	
	If EG1_group_export(ubound(EG1_group_export,1), I125_EG1_E1_i_goods_movement_header_item_document_no) _
		= E3_i_goods_movement_header(I125_E3_item_document_no) Then  
  
		lgStrPrevKey = ""
	Else
		lgStrPrevKey = E3_i_goods_movement_header(I125_E3_item_document_no)
	End if

    Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr

    Response.Write "   .ggoSpread.Source     = .vspdData "                          & vbCr
    Response.Write "   .ggoSpread.SSShowData """ & strData  & """"                  & vbCr
    Response.Write "   .vspdData.focus "                                            & vbCr
    Response.Write "    .hlgDocumentNo = """ & ConvSPChars(Request("txtDocumentNo"))     & """" & vbCr  
    Response.Write "    .hlgYear = """       & Request("txtYear")                        & """" & vbCr  
    Response.Write "    .hlgFromDt = """     & Request("txtFromDt")                      & """" & vbCr  
    Response.Write "    .hlgToDt = """       & Request("txtToDt")                        & """" & vbCr  
    Response.Write "    .hlgMovType = """    & ConvSPChars(Request("txtMovType"))        & """" & vbCr  
    Response.Write "    .hlgTrnsType = """   & ConvSPChars(Request("txtTrnsType"))       & """" & vbCr  
    Response.Write "    .hlgPlantCd = """    & ConvSPChars(Request("txtPlantCd"))        & """" & vbCr  

    Response.Write "   .lgStrPrevKey  = """ & ConvSPChars(lgStrPrevKey) & """" & vbCr  
    Response.Write " if .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) and .lgStrPrevKey <> """" Then "    & vbCr  
    Response.Write "    .DbQuery "                                                              & vbCr  
    Response.Write " else "                                                                     & vbCr  
    Response.Write "    .DbQueryOk "                                                            & vbCr  
    Response.Write " end if  "                                                                  & vbCr  
    
    Response.Write "End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr   
    
    Response.End 

%>

