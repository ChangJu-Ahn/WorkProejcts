<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    Dim lgOpModeCRUD
    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()	
	Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")	
	
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query            
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update            
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete            
    End Select

'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
    Dim iStrNextKey  	
	
    Dim iD2GS69   
    Dim iarrValue
    Dim imp_next_s_bill
    
    Const C_SHEETMAXROWS_D  = 100
    
    Dim   IG1_InportGroup
    
    Const S111_I2_bill_type	             = 0
    Const S111_I2_usage_flag             = 1
    
    Dim   EG1_ExportGroup                                                           ' EG1_ExportGroup 저장 
    
    Const S111_E2_bill_type              = 0
    Const S111_E2_bill_type_nm           = 1
    
    Dim   EG2_ExportGroup                                                                ' EG2_ExportGroup 저장 

    Const S111_EG1_E2_bill_type			= 0
    Const S111_EG1_E2_bill_type_nm		= 1
    Const S111_EG1_E2_except_flag		= 2
    Const S111_EG1_E2_export_flag		= 3
    Const S111_EG1_E2_ref_dn_flag		= 4
    Const S111_EG1_E2_trans_type		= 5
    Const S111_EG1_E1_trans_nm			= 6    
    Const S111_EG1_E2_usage_flag		= 7
    Const S111_EG1_E2_ext1_qty			= 8
    Const S111_EG1_E2_ext2_qty			= 9
    Const S111_EG1_E2_ext1_amt			= 10
    Const S111_EG1_E2_ext2_amt			= 11
    Const S111_EG1_E2_ext1_cd			= 12
    Const S111_EG1_E2_ext2_cd			= 13
    Const S111_EG1_E2_as_flag			= 14

    On Error Resume Next                                                             
    Err.Clear                                                                        '☜: Clear Error status
	
	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	
	
	Redim IG1_InportGroup(1)
	IG1_InportGroup(S111_I2_bill_type) = FilterVar(Trim(Request("txtBilltype")),"","SNM")
	IG1_InportGroup(S111_I2_usage_flag) = Trim(Request("rdoUsageFlg"))
	
	If iStrPrevKey <> "" then					
		iarrValue = Split(iStrPrevKey, gColSep)
		imp_next_s_bill = Trim(iarrValue(0))					
	else			
		imp_next_s_bill = Trim(Request("txtBilltype"))					
	End If		        

	Set iD2GS69 = Server.CreateObject("PD2GS69.cListbilltypesvr")	

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
 
	Call iD2GS69.S_LIST_BILL_TYPE_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D , IG1_InportGroup, FilterVar(imp_next_s_bill,"","SNM"),  _
											   EG1_ExportGroup , EG2_ExportGroup)	

		
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script language=vbs>  " & vbCr
	    Response.Write " Parent.frm1.txtBillTypeNm.value = """"" & vbCr
	    Response.Write "</Script>  " & vbCr    
       Set iD2GS69 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If   

    Set iD2GS69 = Nothing	
    
    istrData = ""
    iLngMaxRow  = CLng(Request("txtMaxRows"))										 '☜: Fetechd Count      
	For iLngRow = 0 To UBound(EG2_ExportGroup,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_bill_type)) 
           Exit For
        End If 

        istrData = istrData & Chr(11) & ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_bill_type)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_bill_type_nm))
        															
		If ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_except_flag)) = "Y" then
			istrData = istrData & Chr(11) & "1"
		else
			istrData = istrData & Chr(11) & "0"					
		End if
			
		If  ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_export_flag)) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"					
		End if

		If  ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_ref_dn_flag)) = "Y" then
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"					
		End if
        istrData = istrData & Chr(11) & ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_trans_type))
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E1_trans_nm))
		
		If ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_usage_flag)) = "Y" then									
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"					
		End if

		If ConvSPChars(EG2_ExportGroup(iLngRow, S111_EG1_E2_as_flag)) = "Y" then									
			istrData = istrData & Chr(11) & "1"
		Else
			istrData = istrData & Chr(11) & "0"					
		End if

        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12) 
    
    Next    
        
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent	       " & vbCr
    Response.Write "      .frm1.txtBillTypeNm.value = """ & ConvSPChars(EG1_ExportGroup(S111_E2_bill_type_nm)) & """" & vbCr    
    Response.Write "   .SetQuerySpreadColor -1	              	                    " & vbCr 
    Response.Write "   .ggoSpread.Source          = .frm1.vspdData		        " & vbCr
    Response.Write "   .ggoSpread.SSShowData        """ & istrData	       & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & iStrNextKey	   & """" & vbCr  
    Response.Write "   .DbQueryOk  " & vbCr   
    Response.Write "End With       " & vbCr															    	
    Response.Write "</Script>      " & vbCr      
           	
End Sub    

'============================================================================================================
Sub SubBizSaveMulti()   
		                                                                    
	Dim iD2GS68
	Dim iErrorPosition
	Dim strtxtSpread
	
    On Error Resume Next                                                                 
    Err.Clear																			 '☜: Clear Error status                                                            
    
    strtxtSpread = Trim(Request("txtSpread"))
	Set iD2GS68 = CreateObject("PD2GS68.cMaintbilltypesvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
   
    Call iD2GS68.S_MAINT_BILL_TYPE_SVR(gStrGlobalCollection, strtxtSpread ,iErrorPosition)

    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iD2GS68 = Nothing
       Exit Sub
	End If
	
    Set iD2GS68 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr   
              
End Sub    


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
