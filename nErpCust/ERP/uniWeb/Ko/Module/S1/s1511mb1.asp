<%@ LANGUAGE=VBSCript%>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1511MB1
'*  4. Program Name         : 품목그룹별구성비등록 
'*  5. Program Desc         : 품목그룹별구성비등록 
'*  6. Comproxy List        : PS1G117.dll, PS1G118.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/03/26
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : choinkuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTB19029.asp" -->

<%
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "*", "NOCOOKIE", "MB") 
	Call HideStatusWnd                                                                 '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
	Dim iLngRow	
	Dim iLngMaxRow
	Dim istrData
	Dim iStrPrevKey
    Dim iStrNextKey  	
	
    Dim iS15118    
    Dim imp_b_item_group
    Dim imp_next_b_item
    
    Dim intGroupCount
    Dim iarrValue
    
    Const C_SHEETMAXROWS_D  = 100
    
    Dim   EG1_ExportGroup                                                           ' EG1_ExportGroup 저장 
    Const EG1_E1_B_ITEM_GROUP_group_cd = 0
    Const EG1_E1_B_ITEM_GROUP_group_nm = 1
    
    Dim   EG2_ExportGroup                                                                ' EG2_ExportGroup 저장 
    Const EG2_E1_S_ITEM_GRP_RATE_item_cd   = 0
    Const EG2_E1_S_ITEM_GRP_RATE_item_nm   = 1
    Const EG2_E1_S_ITEM_GRP_RATE_item_rate = 2
    Const EG2_E1_S_ITEM_GRP_RATE_item_spec = 3

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status    
    	
	imp_b_item_group = Trim(Request("txtItemGroup"))

	iStrPrevKey      = Trim(Request("lgStrPrevKey"))                                        '☜: Next Key	
	
	If iStrPrevKey <> "" then					
		iarrValue = Split(iStrPrevKey, gColSep)
		imp_next_b_item = Trim(iarrValue(0))					
	else			
		imp_next_b_item = ""					
	End If		        

	Set iS15118 = Server.CreateObject("PS1G118.cListItemGrpRateSvr")	

	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If   
	  		
	Call iS15118.LIST_ITEM_GRP_RATE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, imp_b_item_group, imp_next_b_item, _
								    	EG1_ExportGroup, EG2_ExportGroup)	
								    	
    IF cStr(Err.Description) = "B_MESSAGE" & Chr(11) & "127400" Then
       Response.Write "<Script language=vbs>  " & vbCr   
       Response.Write " With Parent	       " & vbCr
       Response.Write "   .frm1.txtItemGroupNm.value = """ & "" & """" & vbCr   
       Response.Write " End With       " & vbCr															    	
       Response.Write "</Script>      " & vbCr    
    End If								    	
    
	If CheckSYSTEMError(Err,True) = True Then
       Set iS15118 = Nothing    
       Response.Write "<Script language=vbs>  " & vbCr   
       Response.Write "   Parent.frm1.txtItemGroup.focus " & vbCr    
       Response.Write "</Script>      " & vbCr
       Exit Sub
    End If   

    Set iS15118 = Nothing	
    
    iLngMaxRow  = CLng(Request("txtMaxRows"))										 '☜: Fetechd Count      
    
	For iLngRow = 0 To UBound(EG2_ExportGroup,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   iStrNextKey = ConvSPChars(EG2_ExportGroup(iLngRow, EG2_E1_S_ITEM_GRP_RATE_item_cd)) 
           Exit For
        End If  		
		istrData = istrData & Chr(11) &    ConvSPChars(EG2_ExportGroup(iLngRow, EG2_E1_S_ITEM_GRP_RATE_item_cd))
        istrData = istrData & Chr(11) &    ConvSPChars(EG2_ExportGroup(iLngRow, EG2_E1_S_ITEM_GRP_RATE_item_nm))
        istrData = istrData & Chr(11) &    ConvSPChars(EG2_ExportGroup(iLngRow, EG2_E1_S_ITEM_GRP_RATE_item_spec))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG2_ExportGroup(iLngRow, EG2_E1_S_ITEM_GRP_RATE_item_rate), ggExchRate.DecPoint, 0)
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)     
    Next    
    
    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent	       " & vbCr
    Response.Write "   .frm1.txtItemGroup.value   = """ & ConvSPChars(EG1_ExportGroup(EG1_E1_B_ITEM_GROUP_group_cd)) & """" & vbCr    
    Response.Write "   .frm1.txtItemGroupNm.value = """ & ConvSPChars(EG1_ExportGroup(EG1_E1_B_ITEM_GROUP_group_nm)) & """" & vbCr    
    Response.Write "   .SetSpreadColor -1, -1	              	                    " & vbCr
    Response.Write "   .frm1.txtHItemGroup.value  = """ & ConvSPChars(imp_b_item_group) & """" & vbCr   
    Response.Write "   .ggoSpread.Source          = .frm1.vspdData		        " & vbCr
    Response.Write "   .ggoSpread.SSShowDataByClip        """ & istrData	       & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & iStrNextKey	   & """" & vbCr  
    Response.Write "   .DbQueryOk  " & vbCr   
    Response.Write "End With       " & vbCr															    	
    Response.Write "</Script>      " & vbCr      
           	
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   
		                                                                    
	Dim iS15111	
	Dim iErrorPosition
	
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear																			 '☜: Clear Error status                                                            
     
	Set iS15111 = Server.CreateObject("PS1G117.cMaintItemGrpRateSvr")
	
    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    Dim reqtxtItemGroup
    Dim reqtxtSpread
    reqtxtItemGroup = Request("txtItemGroup")
    reqtxtSpread = Request("txtSpread")
    Call iS15111.MAINT_ITEM_GRP_RATE_SVR(gStrGlobalCollection, Trim(reqtxtItemGroup), _
								    	 Trim(reqtxtSpread), iErrorPosition)    
												      
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iS15111 = Nothing
       Exit Sub
	End If
	
    Set iS15111 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr   
              
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
