<!--'********************************************************************************************************
'*  1. Module Name          : Inventory           *
'*  2. Function Name        : LOT Popup
'*  3. Program ID           : i2511pb1.asp            *
'*  4. Program Name         :               *
'*  5. Program Desc         : 
'*  7. Modified date(First) : 2000/10/09     
'*  8. Modified date(Last)  : 2000/10/09     
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :                   *
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"         *
'*                            this mark(¢Á) Means that "may  change"         *
'*                            this mark(¡Ù) Means that "must change"         *
'* 13. History              :                   *
'*                            2000/10/09 : 4th Iteration
'********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%                  
Call LoadBasisGlobalInf()
Call HideStatusWnd 
          
On Error Resume Next

Dim i25118
Dim strData

Dim lgStrPrevKey
Dim lgStrPrevSubKey
Dim LngRow
Dim LngMaxRow
Dim intGroupCount
Dim PvArr 

	Const C_SHEETMAXROWS_D = 100

    '-----------------------
    'IMPORTS View
    '-----------------------
 Dim I1_b_item_cd
 Dim I2_b_plant_cd
 Dim I3_i_lot_master
  Const I419_I3_lot_no = 0
  Const I419_I3_lot_sub_no = 1
 ReDim I3_i_lot_master(I419_I3_lot_sub_no)
 '-----------------------
 'EXPORTS View
 '-----------------------
 Dim E1_i_lot_master
  Const I419_E1_lot_no = 0
  Const I419_E1_lot_sub_no = 1
 DIm EG1_group_export
  Const I419_EG1_E1_i_lot_master_lot_no = 0
  Const I419_EG1_E1_i_lot_master_lot_sub_no = 1
  Const I419_EG1_E1_i_lot_master_order_type = 2
  Const I419_EG1_E1_i_lot_master_lot_genrtd_dt = 3
  Const I419_EG1_E1_i_lot_master_unit = 4
  Const I419_EG1_E1_i_lot_master_qty = 5
  Const I419_EG1_E1_i_lot_master_tracking_no = 6
  Const I419_EG1_E1_i_lot_master_po_no = 7
  Const I419_EG1_E1_i_lot_master_po_seq_no = 8
  Const I419_EG1_E1_i_lot_master_prodt_order_no = 9
  Const I419_EG1_E1_i_lot_master_insp_req_no = 10
  Const I419_EG1_E1_i_lot_master_insp_req_seq_no = 11

 lgStrPrevKey		= Request("lgStrPrevKey")
 lgStrPrevSubKey	= Request("lgStrPrevSubKey")
 
 I1_b_item_cd		= Request("txtItemCd")
 I2_b_plant_cd		= Request("txtPlantCd")
 I3_i_lot_master(I419_I3_lot_no)		= Request("txtLotNo")   
 I3_i_lot_master(I419_I3_lot_sub_no)	= Request("txtLotSubNo")
 
 if lgStrPrevKey <> "" then
     I3_i_lot_master(I419_I3_lot_no)     = lgStrPrevKey
     I3_i_lot_master(I419_I3_lot_sub_no) = lgStrPrevSubKey
 end if

 if I3_i_lot_master(I419_I3_lot_sub_no) <> "" then 
  I3_i_lot_master(I419_I3_lot_sub_no) = 0   
 end if


	If CheckSYSTEMError(Err, True) = True Then
		Response.End            
	End If    
 
 
	Set i25118 = Server.CreateObject("PI4G060.cIListLotMasterSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End           
	End If    

	Call i25118.I_LIST_LOT_MASTER(gStrGlobalCollection, C_SHEETMAXROWS_D, _
							I1_b_item_cd, _
							I2_b_plant_cd, _
							I3_i_lot_master, _
							E1_i_lot_master, _
							EG1_group_export)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set i25118 = Nothing           
		Response.End             
	End If

	Set i25118 = Nothing

 if isEmpty(EG1_group_export) then
	Response.End             
 END IF
 
 strData		= ""
 intGroupCount	= UBOUND(EG1_group_export,1)
 LngMaxRow		= Request("txtMaxRows")
 ReDim	PvArr(UBOUND(EG1_group_export,1))
 
 For LngRow = 0 To intGroupCount
  strData =		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I419_EG1_E1_i_lot_master_lot_no)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I419_EG1_E1_i_lot_master_lot_sub_no)) & _
				Chr(11) & UNIDateClientFormat(EG1_group_export(LngRow, I419_EG1_E1_i_lot_master_lot_genrtd_dt)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I419_EG1_E1_i_lot_master_order_type)) & _
				Chr(11) & ConvSPChars(EG1_group_export(LngRow, I419_EG1_E1_i_lot_master_tracking_no)) & _
				Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
  
  PvArr(LngRow) = strData    
 Next
  strData = Join(PvArr, "")
  
 If CheckSYSTEMError(Err, True) = True Then
  Response.End
 End If

 If EG1_group_export(intGroupCount, I419_EG1_E1_i_lot_master_lot_no) = E1_i_lot_master(I419_E1_lot_no) AND _
    EG1_group_export(intGroupCount, I419_EG1_E1_i_lot_master_lot_sub_no) = E1_i_lot_master(I419_E1_lot_sub_no) then    

  lgStrPrevKey    = ""
  lgStrPrevSubKey = ""
 Else
  lgStrPrevKey    = E1_i_lot_master(I419_E1_lot_no)
  lgStrPrevSubKey = E1_i_lot_master(I419_E1_lot_sub_no)
 End If


 
 
    Response.Write "<Script Language=vbscript> " & vbCr
    Response.Write "With Parent "                & vbCr
 
    Response.Write " .ggoSpread.Source = .vspdData "             & vbCr 
    Response.Write " .ggoSpread.SSShowData """ & strData  & """" & vbCr
    Response.Write " .vspdData.focus "                           & vbCr
  
    Response.Write " .lgStrPrevKey = """    & ConvSPChars(lgStrPrevKey)    & """" & vbCr 
    Response.Write " .lgStrPrevSubKey = """ & ConvSPChars(lgStrPrevSubKey) & """" & vbCr 
    
    Response.Write " If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> """" Then " & vbCr
    Response.Write "   .DbQuery "                                                         & vbCr
    Response.Write " Else "                                                                  & vbCr
    Response.Write "   .DbQueryOk "                                                       & vbCr
    Response.Write " End If "                                                                & vbCr
  
    Response.Write "End with " & vbCr
    Response.Write "</Script> " & vbCr
%>

