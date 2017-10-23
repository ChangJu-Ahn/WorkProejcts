<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 품목기준정보 
'*  3. Program ID           : B1B05MB1
'*  4. Program Name         : 품목별공장배분비등록 
'*  5. Program Desc         : 품목별공장배분비등록 
'*  6. Comproxy List        : PB3CS90.dll, PB3CS91.dll               
'*  7. Modified date(First) : 2002/06/17
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Ahn Tae Hee
'* 11. Comment              :																			
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									
'*                            this mark(⊙) Means that "may  change"									
'*                            this mark(☆) Means that "must change"									
'* 13. History              : 2002/11/15 : UI성능 적용						                            
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call loadInfTB19029B( "I", "*", "NOCOOKIE", "MB") 

Dim lgOpModeCRUD
On Error Resume Next                                                            '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
'---------------------------------------Common-----------------------------------------------------------
	
lgOpModeCRUD = Request("txtMode")                                                '☜: Read Operation Mode (CRUD)
    
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)                                                         '☜: Delete
End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
    Dim LngRow	
	Dim LngMaxRow
	
	Dim lgstrData

	Dim StrPrevKey
    Dim StrNext
    Dim StrNextKey  	
    
    Const C_SHEETMAXROWS_D  = 100
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	LngMaxRow       = CLng(Request("txtMaxRows"))                                  '☜: Fetechd Count
	StrPrevKey      =TRIM(Request("lgStrPrevKey"))                                 '☜: Next Key
    
    '--------------
	'Interface 정의 
	'--------------
    'View Name : export_item b_plant
    Const B411_EG1_E1_plant_cd = 0
    Const B411_EG1_E1_plant_nm = 1
    'View Name : export_item b_plant_rate_by_item
    Const B411_EG1_E2_rate = 2
    'View Name : export b_item
    Const B411_E2_item_cd = 0
    Const B411_E2_item_nm = 1
    Const B411_E2_plant_cd = 2
    Const B411_E2_plant_nm = 3
    'View Name : export_next b_plant
    Const B411_E1_plant_cd = 0
    
    '--------
	'View선언 
	'--------
    'export next view
    Dim E1_b_plant
    'View Name : export b_item    
    Dim E2_b_item
    'export group view
    Dim EG1_export_group
    'import view
    Dim I2_b_item_item_cd 
    'import next view
    Dim I1_b_plant_plant_cd 
	
	'-------------
	'comproxy 선언 
	'-------------
	Dim iB1b058	    
    
    '-----------
	'import view
	'-----------
    I2_b_item_item_cd = FilterVar(trim(Request("txtConItemCode")), "", "SNM")
	 
	If  StrPrevKey <> "" Then
		I1_b_plant_plant_cd = trim(StrPrevKey)
	Else 
	    I1_b_plant_plant_cd = ""	           
    End If
   
    '-------------
	'call comproxy
	'-------------
    Set iB1b058 = Server.CreateObject("PB3CS91.cBListPlantItemSvr")
  
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If 
    
    Call iB1b058.B_LIST_PLANT_RATE_ITEM_SVR( gStrGlobalCollection, C_SHEETMAXROWS_D, Cstr(I1_b_plant_plant_cd),Cstr(I2_b_item_item_cd), _
                                              EG1_export_group ,E1_b_plant, E2_b_item )

   	If CheckSYSTEMError(Err,True) = True Then
       Set iB1b058 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   
   
     Set iB1b058 = Nothing	
	
	'-----------------------------------
	'품목별 공장배분비가 없는 경우(신규) 
	'-----------------------------------
	If Isempty(EG1_export_group) Then 

 		For LngRow = 0 To Ubound(E2_b_item,1)
			'--------------------
			'화면에 뿌려지는 data 
			'--------------------
			'공장 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(E2_b_item(LngRow,B411_E2_plant_cd))
			'공장팝업버튼 
			lgstrData = lgstrData & Chr(11)
			'공장명 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(E2_b_item(LngRow,B411_E2_plant_nm))
			'배분비 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(0,ggExchRate.DecPoint,0)
		    lgstrData = lgstrData & Chr(11) & LngMaxRow + LngRow
		    lgstrData = lgstrData & Chr(11) & Chr(12)  
		Next
		Response.Write "<Script language=vbs> " & vbCr   
		'------------------------------------
		' 품목별공장배분비 헤더의 내용을 표시 
		'------------------------------------ 
		Response.Write " Parent.frm1.txtConItemName.value   = """ & ConvSPChars(E2_b_item(0,B411_E2_item_nm))        & """" & vbCr
		Response.Write " Parent.frm1.txtItemCode.value   = """ & ConvSPChars(E2_b_item(0,B411_E2_item_cd))           & """" & vbCr
		Response.Write " Parent.frm1.txtItemName.value   = """ & ConvSPChars(E2_b_item(0,B411_E2_item_nm))           & """" & vbCr
		'------------
		'Hidden Input
		'------------
		Response.Write " Parent.frm1.txtHConItemCode.value    = """ & Request("txtConItemCode")                  	 & """" & vbCr  
		'------------------------
		'Result data display area
		'------------------------    
		Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData		                              " & vbCr
		Response.Write " Parent.ggoSpread.SSShowData        """ & lgstrData											& """" & vbCr
			    
		Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey								   		 & """" & vbCr  
		Response.Write " Parent.ItemByPlantOK "																   		    	& vbCr   
		Response.Write "</Script> "
	'----------------------------
	'품목별 공장배분비가 있는 경우 
	'----------------------------
	Else
		'----------------
		'export next view
		'----------------	   
		StrNext = E1_b_plant(B411_E1_plant_cd)
			  
		For LngRow = 0 To Ubound(EG1_export_group,1)
			If  LngRow < C_SHEETMAXROWS_D  Then
		 	Else
		       StrNextKey = StrNext 
		          Exit For
		    End If  

			'--------------------
			'화면에 뿌려지는 data 
			'--------------------
			'공장 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,B411_EG1_E1_plant_cd))
			'공장팝업버튼 
			lgstrData = lgstrData & Chr(11)
			'공장명 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(EG1_export_group(LngRow,B411_EG1_E1_plant_nm))
			'배분비 
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(EG1_export_group(LngRow,B411_EG1_E2_rate),ggExchRate.DecPoint,0)
		    lgstrData = lgstrData & Chr(11) & LngMaxRow + LngRow 
		    lgstrData = lgstrData & Chr(11) & Chr(12)  
		Next
		Response.Write "<Script language=vbs> " & vbCr   
		'------------------------------------
		' 품목별공장배분비 헤더의 내용을 표시 
		'------------------------------------ 
		Response.Write " Parent.frm1.txtConItemName.value   = """ & ConvSPChars(E2_b_item(0,B411_E2_item_nm))        & """" & vbCr
		Response.Write " Parent.frm1.txtItemCode.value   = """ & ConvSPChars(E2_b_item(0,B411_E2_item_cd))           & """" & vbCr
		Response.Write " Parent.frm1.txtItemName.value   = """ & ConvSPChars(E2_b_item(0,B411_E2_item_nm))           & """" & vbCr
		'------------
		'Hidden Input
		'------------
		Response.Write " Parent.frm1.txtHConItemCode.value    = """ & Request("txtConItemCode")     	        	 & """" & vbCr  
		'------------------------
		'Result data display area
		'------------------------    
		Response.Write " Parent.ggoSpread.Source          = Parent.frm1.vspdData				                      " & vbCr
		Response.Write " Parent.ggoSpread.SSShowData        """ & lgstrData						 		     & """" & vbCr
			    
		Response.Write " Parent.lgStrPrevKey              = """ & StrNextKey							    		 & """" & vbCr  
		Response.Write " Parent.DbQueryOk "																    		    	& vbCr   
		Response.Write "</Script> "
   End If    
   
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   
		                                                                    
	Dim iPB3CS90
	Dim iErrorPosition
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status                                                            
     
	Set iPB3CS90 = Server.CreateObject("PB3CS90.cBPlantRtItemSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
	Dim reqtxtSpread
	reqtxtSpread = Request("txtSpread")
    Call iPB3CS90.B_MAINT_PLANT_RATE_ITEM_SVR(gStrGlobalCollection, Trim(reqtxtSpread), iErrorPosition) 
          
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPB3CS90 = Nothing
       Exit Sub
	End If
	
    Set iPB3CS90 = Nothing
    
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr  
    Response.Write " Parent.frm1.txtConItemCode.value    = """ & Request("txtItemCode")   & """" & vbCr  
    Response.Write "</Script> "           
 
End Sub    

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>
