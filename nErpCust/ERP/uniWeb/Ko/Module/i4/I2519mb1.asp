<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : i2519mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         : lot hdr Lookup(Tree View)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Nam hoon
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  :                             
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%            
Response.Flush 

Call LoadBasisGlobalInf()
Call HideStatusWnd              

On Error Resume Next            

Const C_Open  = "Open"
Const C_PROD  = "Folder"
Const C_MATL  = "URL"

Dim pI25119                 

    '-----------------------
    'IMPORTS View
    '-----------------------
 Dim I1_i_lot_explosion_no
 Dim I2_b_plant_CD
 Dim I3_b_item_CD
 Dim I4_i_lot_master
  Const I416_I4_lot_no = 0
  Const I416_I4_lot_sub_no = 1
 ReDim I4_i_lot_master(I416_I4_lot_sub_no)
    '-----------------------
 'EXPORTS View
 '-----------------------
 Dim E1_b_plant
  Const I416_E1_plant_cd = 0
  Const I416_E1_plant_nm = 1 
 Dim E2_b_item
  Const I416_E2_item_cd = 0
  Const I416_E2_item_nm = 1
 Dim E3_i_lot_master
  Const I416_E3_lot_no = 0
  Const I416_E3_lot_sub_no = 1
  Const I416_E3_order_type = 2
  Const I416_E3_lot_genrtd_dt = 3
  Const I416_E3_unit = 4
  Const I416_E3_qty = 5
  Const I416_E3_tracking_no = 6
  Const I416_E3_po_no = 7
  Const I416_E3_po_seq_no = 8
  Const I416_E3_prodt_order_no = 9
  Const I416_E3_insp_req_no = 10
  Const I416_E3_insp_req_seq_no = 11
  Const I416_E3_delete_flag = 12

 

 If CInt(Request("txtMode")) <> UID_M0001 Then
  Response.End 
 End If

 '-----------------------
 'Data manipulate  area(import view match)
 '-----------------------
 I2_b_plant_CD = Request("txtPlantCd")
 I3_b_item_CD  = Request("txtItemCd")  
 I4_i_lot_master(I416_I4_lot_no)     = Request("txtLotNo")
 I4_i_lot_master(I416_I4_lot_sub_no) = Request("txtLotSubNo")


 If CheckSYSTEMError(Err, True) = True Then
  Response.End            '☜: 비지니스 로직 처리를 종료함 
 End If    
 

 Set pI25119 = Server.CreateObject("PI4S050.cILookUpLotMasterSvr")
    
 If CheckSYSTEMError(Err, True) = True Then
  Response.End            '☜: 비지니스 로직 처리를 종료함 
 End If    
 
 Call pI25119.I_LOOKUP_LOT_MASTER(gStrGlobalCollection, _
         I1_i_lot_explosion_no, _
         I2_b_plant_CD, _
         I3_b_item_CD, _
         I4_i_lot_master, _
         E1_b_plant, _
         E2_b_item, _
         E3_i_lot_master)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
     Set pI25119 = Nothing            '☜: ComProxy Unload
  Response.End              '☜: 비지니스 로직 처리를 종료함 
 End If

 Set pI25119 = Nothing
  
    Response.Write "<Script Language=vbscript> " & vbCr
    Response.Write " With parent.frm1 "       & vbCr
    
    Response.Write "  .txtItemCd2.value = """ & ConvSPChars(E2_b_item(I416_E2_item_cd))   & """" & vbCr
    Response.Write "  .txtItemNm2.value = """ & ConvSPChars(E2_b_item(I416_E2_item_nm))   & """" & vbCr
    Response.Write "  .txtPlantNm.value = """ & ConvSPChars(E1_b_plant(I416_E1_plant_nm)) & """" & vbCr

 if E3_i_lot_master(I416_E3_order_type) = "PR" then
  Response.Write "  .txtOrdNo.value    = """ & ConvSPChars(E3_i_lot_master(I416_E3_po_no))     & """" & vbCr
  Response.Write "  .txtOrdSubNo.value = """ & ConvSPChars(E3_i_lot_master(I416_E3_po_seq_no)) & """" & vbCr
 else
  Response.Write "  .txtOrdNo.value    = """ & ConvSPChars(E3_i_lot_master(I416_E3_prodt_order_no)) & """" & vbCr
  Response.Write "  .txtOrdSubNo.value = """" " & vbCr
 end if
   
    Response.Write "  .txtOrdType.value  = """ & ConvSPChars(E3_i_lot_master(I416_E3_order_type))            & """" & vbCr
    Response.Write "  .txtLotGenDt.value   = """ & UNIDateClientFormat(E3_i_lot_master(I416_E3_lot_genrtd_dt)) & """" & vbCr
    Response.Write "  .txtItemUnit.value  = """ & ConvSPChars(E3_i_lot_master(I416_E3_unit))                  & """" & vbCr
    Response.Write "  .txtRcptQty.value  = """ & E3_i_lot_master(I416_E3_qty)                                & """" & vbCr
    Response.Write "  .txtTrackingNo.value = """ & ConvSPChars(E3_i_lot_master(I416_E3_tracking_no))           & """" & vbCr
      
    Response.Write " End With "                & vbCr
    Response.Write " Call parent.DbQueryOk() " & vbCr
    Response.Write "</Script> "                   & vbCr

%>
