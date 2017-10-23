<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : 실사선별관리  저장 업무 처리 ASP
'*  2. Function Name        : 
'*  3. Program ID           : i2161mb2.asp
'*  4. Program Name         : 실사선별Manual 등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : PI2S020.cIMaintPhyInv

'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2002/03/08
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%Call LoadBasisGlobalInf

    On Error Resume Next												
	Err.Clear

	Call HideStatusWnd 

	Dim pPI2S020															

	Dim I2_b_storage_location_sl_cd
	Dim I3_b_plant_plant_cd 

	Dim I4_i_physical_inventory_header
	    Const I205_I4_phy_inv_no = 0 
	    Const I205_I4_real_insp_dt = 1
	Redim I4_i_physical_inventory_header(I205_I4_real_insp_dt)    

	Dim E3_i_physical_inventory_header_phy_inv_no
	Dim prErrorPosition
	
	Dim iCommandSent

	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim i
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For i = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(i)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")

	iCommandSent 					    = "CREATE"
	I3_b_plant_plant_cd 				= Request("txtPlantCd")
	I2_b_storage_location_sl_cd			= Request("txtSLCd")
	I4_i_physical_inventory_header(I205_I4_phy_inv_no)	= Request("txtCondPhyInvNo")
	I4_i_physical_inventory_header(I205_I4_real_insp_dt)	= UNIConvDate(Request("txtInspDt"))
	

	Set pPI2S020 = Server.CreateObject("PI2S020.cIMaintPhyInv")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
   		Response.End
	End If

	Call pPI2S020.I_MAINT_PHY_INV(gStrGlobalCollection, _
								iCommandSent, _
								I2_b_storage_location_sl_cd, I3_b_plant_plant_cd, _
								I4_i_physical_inventory_header, _
								itxtSpread, _
								E3_i_physical_inventory_header_phy_inv_no, _
								prErrorPosition)

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set pPI2S020 = Nothing														
		If prErrorPosition <> "" Then
			Call SheetFocus(prErrorPosition, 2)
		End If
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
		Response.End
	End If

	Set pPI2S020 = Nothing														

	
	Call SubBizBatchMulti()
	
	Response.Write "<Script Language=vbscript> " & vbCr   
    Response.Write " With Parent "               & vbCr
	Response.Write "	.frm1.txtCondPhyInvNo.value = """ & ConvSPChars(E3_i_physical_inventory_header_phy_inv_no) & """" & vbCr  	   	  
  	Response.Write "    .RemovedivTextArea	"	& vbCr
	Response.Write "    .DbSaveOk			"	& vbCr
	Response.Write " End with				"	& vbCr
    Response.Write "</Script>				"	& vbCr   
	Response.End     

Sub SheetFocus(ByVal lRow, ByVal lCol)
	Response.Write " <Script Language=VBScript> "                    & vbCrLF
	Response.Write " With parent.frm1 "                              & vbCrlf
	Response.Write "	.vspdData.focus "                           & vbCrlf
	Response.Write "	.vspdData.Row = """ & lRow & """" & vbCr
	Response.Write "	.vspdData.Col = """ & lCol & """" & vbCr
	Response.Write "	.vspdData.Action = 0 "                      & vbCrlf
	Response.Write "	.vspdData.SelStart = 0 "                    & vbCrlf
	Response.Write "	.vspdData.SelLength = len(.vspdData.Text) " & vbCrlf
	Response.Write " End With"                                       & vbCrlf
	Response.Write " </Script>"                                      & vbCrLF
End Sub


Sub SubBizBatchMulti()

	Dim IntRetCD
	Dim strMsg_cd, strMsg_text
	Dim strPhyInvNo,strInspODt, strCurrDt, strInspDt
		   
	strPhyInvNo	=	E3_i_physical_inventory_header_phy_inv_no
	strInspODt	=	UNIConvDate(Request("txtInspDt"))		
		
	strInspDt	=	UNIConvDateAToB(strInspODt, gAPDateFormat, gServerDateFormat)
	strCurrDt	=   GetSvrDate	

	If strInspDt < strCurrDt Then
	    
		Call SubOpenDB(lgObjConn)                    
		Call SubCreateCommandObject(lgObjComm)
	
		With lgObjComm
			.CommandText = "usp_i_update_phy_inv_qty"
			.CommandType = adCmdStoredProc

			lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@phy_inv_no"     ,advarxchar,adParamInput,Len(Trim(strPhyInvNo)), strPhyInvNo)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@real_count_dt",adDate,adParamInput,Len(Trim(strInspODt)), strInspODt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@updt_user_id"     ,advarxchar,adParamInput,Len(Trim(gUsrID)), gUsrID)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"   ,advarxchar ,adParamOutput,6)

			lgObjComm.Execute ,, adExecuteNoRecords
				
		End With

		If  Err.number = 0 Then
			IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
					 
			If  IntRetCD <> 1 then
		            
				strMsg_cd   = lgObjComm.Parameters("@msg_cd").Value
				strSpId     = FilterVar(lgObjComm.Parameters("@updt_user_id").Value, "''", "S")
				            
				If strMsg_cd <> MSG_OK_STR Then
					Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
				End If
	
				IntRetCD = -1
				Exit Sub
			Else
				IntRetCD = 1
			End if
		Else           
			Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
			Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
			IntRetCD = -1
		End if

		Call SubCloseCommandObject(lgObjComm)
		Call SubCloseDB(lgObjConn)       
	Else 
		Exit Sub
	End if
End Sub		
%>
