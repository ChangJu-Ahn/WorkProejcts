<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : i2511mb1.asp
'*  4. Program Name         :
'*  5. Program Desc         : lot tracing(Tree View)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/28
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  :                             
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%            
Response.Flush 

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")  

Call HideStatusWnd               

On Error Resume Next             

Const C_Open  = "Open"
Const C_PROD  = "Folder"
Const C_MATL  = "URL"

Dim pI25128                
Dim iCommandSent
	Const C_SHEETMAXROWS_D = 100

Dim strCode                
Dim strMode                
Dim LngRow

Dim blnLoopFlg
Dim IntNextKey
Dim strPlantCd
Dim strItemCd
Dim strItemNm
Dim strBomNo
Dim strExpFlg
Dim strItemAcct
Dim strLotNo
Dim strLotSubNo

    '-----------------------
    'IMPORTS View
    '-----------------------
 Dim I1_i_lot_explosion_no
 Dim I2_ief_supplied_action_entry
 Dim I3_good_mvmt_workset_temp_flag
 Dim I4_i_lot_master
	Const I407_I4_lot_no = 0
	Const I407_I4_lot_sub_no = 1 
 ReDim I4_i_lot_master(I407_I4_lot_sub_no)
 Dim I5_b_item_cd
 Dim I6_b_plant_cd
 '-----------------------
 'EXPORTS View
 '-----------------------    
 Dim E1_i_lot_master
  Const I407_E1_lot_no = 0
  Const I407_E1_lot_sub_no = 1
  Const I407_E1_order_type = 2
  Const I407_E1_lot_genrtd_dt = 3
  Const I407_E1_unit = 4
  Const I407_E1_qty = 5
  Const I407_E1_tracking_no = 6
  Const I407_E1_po_no = 7
  Const I407_E1_po_seq_no = 8
  Const I407_E1_prodt_order_no = 9
  Const I407_E1_insp_req_no = 10
  Const I407_E1_insp_req_seq_no = 11 
 Dim EG1_group_export
  Const I407_EG1_E1_i_lot_explosion_plant_cd = 0
  Const I407_EG1_E1_i_lot_explosion_explosion_no = 1
  Const I407_EG1_E1_i_lot_explosion_prnt_node = 2
  Const I407_EG1_E1_i_lot_explosion_own_node = 3
  Const I407_EG1_E1_i_lot_explosion_level_cd = 4
  Const I407_EG1_E1_i_lot_explosion_prnt_item_cd = 5
  Const I407_EG1_E1_i_lot_explosion_prnt_lot_no = 6
  Const I407_EG1_E1_i_lot_explosion_prnt_lot_sub_no = 7
  Const I407_EG1_E1_i_lot_explosion_chld_lot_no = 8
  Const I407_EG1_E1_i_lot_explosion_chld_lot_sub_no = 9
  Const I407_EG1_E1_i_lot_explosion_child_item_cd = 10
  Const I407_EG1_E1_i_lot_explosion_material_flag = 11 
 Dim E2_i_lot_explosion_no


If CInt(Request("txtMode")) <> UID_M0001 Then
 Response.End 
End If

 strPlantCd  = Request("txtPlantCd")         ' 조회할 키 
 strItemCd  = Request("txtItemCd")         ' 조회할 상위키 
 strExpFlg   = Request("rdoSrchType")
 strItemAcct = Request("txtHdnItemAcct")
 strLotNo    = Request("txtLotNo")
 strLotSubNo = Request("txtLotSubNo")

    '----------------------------------------------------
    '- Parent Node를 Setting하고 Header Data를 가져온다.
    '---------------------------------------------------    
    

 Response.Write "<Script Language=vbscript> " &  vbcr
 Response.Write "Dim PrntKey "                &  vbcr
 Response.Write "Dim NodX "                   &  vbcr
 Response.Write "With parent.frm1 "           &  vbcr
         
 Response.Write " PrntKey = """ & ConvSPChars(UCase(Trim(strItemCd))) & "|^|^|" & ConvSPChars(UCase(Trim(strLotNo))) & "|^|^|" & ConvSPChars(UCase(Right("0000000000000" & Trim(strLotSubNo),13))) & """" &  vbcr
 Response.Write " Set NodX = .uniTree1.Nodes.Add(,,PrntKey,""" & ConvSPChars(UCase(Trim(strLotNo) & "-" & Cint(Trim(strLotSubNo)))) & """,parent.C_PROD, parent.C_PROD) "                                 &  vbcr
 
 Response.Write " NodX.Expanded = True " &  vbcr
 Response.Write " Set NodX = Nothing "   &  vbcr
 Response.Write "End With "                 &  vbcr
 Response.Write "</Script> "                &  vbcr

 
 IntNextKey = 0

 Do Until blnLoopFlg = True


  '-----------------------
  'Data manipulate  area(import view match)
  '-----------------------
  I6_b_plant_cd = strPlantCd
  I5_b_item_cd = strItemCd
  I1_i_lot_explosion_no  = IntNextKey

  If IntNextKey <> 0 Then  
   I3_good_mvmt_workset_temp_flag = "NX"   'Next
  Else
   iCommandSent = "LOOKUP"
  End If
  
  I4_i_lot_master(I407_I4_lot_no)     = strLotNo
  I4_i_lot_master(I407_I4_lot_sub_no) = strLotSubNo
  
  I2_ief_supplied_action_entry = strExpFlg
  
  If CheckSYSTEMError(Err, True) = True Then
   Response.End            '☜: 비지니스 로직 처리를 종료함 
  End If    


  Set pI25128 = Server.CreateObject("PI4G040.cILstLotInfoDtlSvr")

    
  If CheckSYSTEMError(Err, True) = True Then
   Response.End           
  End If    
 
  Call pI25128.I_LIST_LOT_INFO_DETAIL_SVR(gStrGlobalCollection, iCommandSent, C_SHEETMAXROWS_D, _
           I1_i_lot_explosion_no , _
           I2_ief_supplied_action_entry , _
           I3_good_mvmt_workset_temp_flag , _
           I4_i_lot_master , _
           I5_b_item_cd , _
           I6_b_plant_cd , _
           E1_i_lot_master , _
           EG1_group_export , _
           E2_i_lot_explosion_no)  
  '-----------------------
  'Com action result check area(OS,internal)
  '-----------------------
  If CheckSYSTEMError(Err, True) = True Then
   Set pI25128 = Nothing            
   Response.End             
  End If

  Set pI25128 = Nothing

  '--------------------------------
  'Next가 아니면 Header정보 Setting
  '--------------------------------
  If IntNextKey = 0 Then

   Response.Write "<Script Language=vbscript> "                                                      &  vbcr
   Response.Write "With parent.frm1 "                                                                &  vbcr
   Response.Write " .txtItemCd2.value    = .txtItemCd.value " &  vbcr
   Response.Write " .txtItemNm2.value    = .txtItemNm.value " &  vbcr

   if Trim(E1_i_lot_master(I407_E1_order_type)) <> "MR" then

    Response.Write " .txtOrdNo.value     = """ & ConvSPChars(E1_i_lot_master(I407_E1_po_no))     & """" &  vbcr
    Response.Write " .txtOrdSubNo.value  = """ & ConvSPChars(E1_i_lot_master(I407_E1_po_seq_no)) & """" &  vbcr

   else

    Response.Write " .txtOrdNo.value     = """ & ConvSPChars(E1_i_lot_master(I407_E1_prodt_order_no)) & """" &  vbcr
    Response.Write " .txtOrdSubNo.value  = """" "                                                            &  vbcr

   end if

   Response.Write " .txtOrdType.value  = """ & ConvSPChars(E1_i_lot_master(I407_E1_order_type))                  & """" &  vbcr
   Response.Write " .txtLotGenDt.value   = """ & UNIDateClientFormat(E1_i_lot_master(I407_E1_lot_genrtd_dt))       & """" &  vbcr
   Response.Write " .txtItemUnit.value  = """ & ConvSPChars(E1_i_lot_master(I407_E1_unit))                        & """" &  vbcr
   Response.Write " .txtRcptQty.value  = """ & UniConvNumberDBToCompany(E1_i_lot_master(I407_E1_qty),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & """" &  vbcr
   Response.Write " .txtTrackingNo.value = """ & ConvSPChars(E1_i_lot_master(I407_E1_tracking_no))                 & """" &  vbcr
      
   Response.Write "End With " &  vbcr
   Response.Write "</Script> " &  vbcr

  End If

'  If pI25128.OperationStatusMessage = "161601" Then
'   
'   Call DisplayMsgBox(pI25128.OperationStatusMessage, vbOKOnly, "", "", I_MKSCRIPT)
'
'   Response.Write "<Script Language=vbScript> "                         &  vbcr
'   Response.Write "parent.lgPrntItemCd        = """ & strItemCd  & """" &  vbcr
'   Response.Write "parent.frm1.hPlantCd.value = """ & strPlantCd & """" &  vbcr
'   Response.Write "Call parent.DbQueryOk()  "                           &  vbcr
'   Response.Write "</Script> "                                          &  vbcr
'
'   Response.End              '☜: 비지니스 로직 처리를 종료함 
'  End If

  if isEmpty(EG1_group_export) then
   Call ServerMesgBox("ERR", vbCritical, I_MKSCRIPT)
   Response.End
  end if

  iGroupCnt = ubound(EG1_group_export,1)  

  If E2_i_lot_explosion_no = EG1_group_export(iGroupCnt, I407_EG1_E1_i_lot_explosion_explosion_no) Then
   blnLoopFlg = True 
  Else
   IntNextKey = E2_i_lot_explosion_no
  End If

  Response.Write "<Script Language=vbscript> " &  vbcr
  Response.Write "Dim Node "                   &  vbcr
  Response.Write "With parent.frm1.uniTree1 " &  vbcr
  Response.Write " .MousePointer = 11 "     &  vbcr        '⊙: 마우스 포인트 변화 
  Response.Write " .Indentation = 50 "      &  vbcr        '⊙: 부모트리와 자식트리 사이의 간격 

  For LngRow = 0 to iGroupCnt
   If EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_material_flag) = "N" Then  ' 폴더일 경우   
    
		Response.Write " Set Node = .Nodes.Add(""" & ConvSPChars(UCase(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_prnt_node))) & """, " 
		Response.Write " parent.tvwChild, " 
		Response.Write " """ & ConvSPChars(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_own_node)) & """, "
		Response.Write " """ & ConvSPChars(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_chld_lot_no)) & "-" & ConvSPChars(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_chld_lot_sub_no)) & """, "
		Response.Write " parent.C_PROD, "
		Response.Write " parent.C_PROD) " &  vbcr
                                              
		Response.Write " Node.Expanded = True " &  vbcr
   Else
		Response.Write " Set Node = .Nodes.Add(""" & ConvSPChars(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_prnt_node)) & """, parent.tvwChild, """ & ConvSPChars(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_own_node)) & """, """ & ConvSPChars(EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_chld_lot_no)) & "-" & EG1_group_export(LngRow, I407_EG1_E1_i_lot_explosion_chld_lot_sub_no) & """, parent.C_MATL, parent.C_MATL) " &  vbcr
   End If
  Next

  Response.Write " .MousePointer = 1 " &  vbcr
  Response.Write " Set Node = Nothing " &  vbcr
  Response.Write "End With "               &  vbcr
  Response.Write "</Script> "              &  vbcr

 Loop
 
Response.Write "<Script Language=vbscript> "	&  vbcr
Response.Write "Call parent.DbQueryOk()	"		&  vbcr
Response.Write "</Script> "						&  vbcr
%>
<!--<Script Language=vbscript>
 Call parent.DbQueryOk()
</Script>-->

