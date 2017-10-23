<!--'**********************************************************************************************
'*  1. Module Name          : ROP품목  저장 업무 처리 ASP
'*  2. Function Name        : 
'*  3. Program ID           : i2411mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : i2411Maint Rop Item Svr

'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2000/04/14
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Mr  Kim Nam Hoon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%          

On Error Resume Next              

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")  
	Call HideStatusWnd
	    
	Dim pi24111                
	 
	Dim IntRows
	Dim IntCols
	Dim vbIntRet
	Dim lEndRow
	Dim boolCheck
	Dim lgIntFlgMode      

	Dim LngMaxRow        
	 
	Dim arrRowVal        
	Dim arrColVal       
	Dim strStatus        
	Dim lGrpCnt, lGrpCnt1      
	Dim intCount

	'-----------------------
	'IMPORTS View
	'-----------------------
	Dim IG1_import_group
		Const I404_IG1_I1_m_pur_req_req_qty = 0
		Const I404_IG1_I1_m_pur_req_req_unit = 1
		Const I404_IG1_I1_m_pur_req_req_dt = 2
		Const I404_IG1_I1_m_pur_req_req_prsn = 3
		Const I404_IG1_I1_m_pur_req_dlvy_dt = 4
		Const I404_IG1_I1_m_pur_req_sl_cd = 5
		Const I404_IG1_I1_m_pur_req_tracking_no = 6
		Const I404_IG1_I2_b_item_item_cd = 7    
	Dim I1_b_plant_cd
	        
	'-----------------------
	'EXPORTS View
	'-----------------------    
	Dim E2_m_pur_req 
		Const   M090_E2_pr_no = 0
		Const   M090_E2_sppl_cd = 1

	LngMaxRow = CInt(Request("txtMaxRows")) -1      
	  
	If LngMaxRow < 0 Then            
		Response.End  
	End if
	 
	lgIntFlgMode = Request("txtFlgMode")        
	 
	'-----------------------
	'Data manipulate area
	'-----------------------          
	'헤더부분 
	I1_b_plant_cd = Request("txtPlantCd")
	 
	lGrpCnt  = 0
	intCount = 0

	If Request("txtSpread") <> "" Then
	 
		arrRowVal = Split(Request("txtSpread"), gRowSep)
		 
		ReDim IG1_import_group(LngMaxRow, I404_IG1_I2_b_item_item_cd)
		  
		For LngRow = 0 To LngMaxRow
		      
		'      Call ServerMesgBox("테스트 ROW : " & arrRowVal(LngRow) , vbCritical, I_MKSCRIPT)
		      
			IF arrRowVal(LngRow) <> "" THEN
				arrColVal = Split(arrRowVal(LngRow), gColSep)
				  
				strStatus = arrColVal(0)              '☜: Row 의 상태 
				 
				Select Case strStatus
				 
				Case "U"     
					IG1_import_group(LngRow,I404_IG1_I2_b_item_item_cd)   = arrColVal(2)     
					IG1_import_group(LngRow,I404_IG1_I1_m_pur_req_req_qty)  = UNIConvNum(arrColVal(3),0)
					IG1_import_group(LngRow,I404_IG1_I1_m_pur_req_req_unit)  = arrColVal(4) 
					IG1_import_group(LngRow,I404_IG1_I1_m_pur_req_dlvy_dt)  = UNIConvDate(arrColVal(5))
					IG1_import_group(LngRow,I404_IG1_I1_m_pur_req_req_dt)  = UNIConvDate(arrColVal(6))
					IG1_import_group(LngRow,I404_IG1_I1_m_pur_req_tracking_no)  = arrColVal(7)
					IG1_import_group(LngRow,I404_IG1_I1_m_pur_req_req_prsn)  = arrColVal(8) 
				End Select
			END IF
		Next

		If CheckSYSTEMError(Err, True) = True Then
			Response.End            '☜: 비지니스 로직 처리를 종료함 
		End If    
		  
		Set pi24111 = Server.CreateObject("PI4G020.cIMaintRopItemSvr")
		    
		If CheckSYSTEMError(Err, True) = True Then
			Response.End           
		End If    
		 
		E2_m_pur_req  = pi24111.I_MAINT_ROP_ITEM_SVR(gStrGlobalCollection, _
														IG1_import_group, _
														I1_b_plant_cd)
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If CheckSYSTEMError(Err, True) = True Then
			Set pi24111 = Nothing           
			Response.Write "<Script Language=vbscript> " & vbCr   
			Response.Write "    Parent.frm1.btnRun.Disabled  = False  " & vbCr
			Response.Write " </Script> " & vbCr   
			Response.End 
		End If
		Set pi24111 = Nothing
	  
	End If

Response.Write "<Script Language=vbscript> " & vbCr   
Response.Write "    Parent.DbSaveOk " & vbCr  
Response.Write " </Script> " & vbCr   
    
Response.End 
%>

