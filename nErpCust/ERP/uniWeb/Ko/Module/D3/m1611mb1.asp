<%@ LANGUAGE=VBSCript%>
<% Option explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->

<%Call LoadBasisGlobalInf()%>
<%
     
Dim lgOpModeCRUD
 
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message
	
lgOpModeCRUD  = Request("txtMode")                                               '☜: Read Operation Mode (CRUD)
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
    
    Dim TmpBuffer
    Dim iMax
    Dim iIntLoopCount
    Dim iTotalStr
    
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	dim istrData
	Dim iLngRow
	Dim GroupCount

	Dim I1_m_iv_type
	Dim I2_m_iv_type  
	Dim E1_m_iv_type
	Dim E2_m_iv_type
	Dim EG1_export_group
	Dim EG2_export_group
	Dim PM1G428
	Dim rdoUsageFlg

	Const M356_EG2_E1_m_iv_type_iv_type_cd = 0
	Const M356_EG2_E1_m_iv_type_iv_type_nm = 1
	Const M356_EG2_E1_m_iv_type_trans_cd = 2
	Const M356_EG2_E1_m_iv_type_import_flg = 3
	Const M356_EG2_E1_m_iv_type_except_flg = 4
	Const M356_EG2_E1_m_iv_type_ret_flg = 5
	Const M356_EG2_E1_m_iv_type_usage_flg = 6
	Const M356_EG2_E1_m_iv_type_ext1_cd = 7
	Const M356_EG2_E1_m_iv_type_ext2_cd = 8
	Const M356_EG2_E1_m_iv_type_ext3_cd = 9
	Const M356_EG2_E1_m_iv_type_ext4_cd = 10
	Const M356_EG2_E1_m_iv_type_trans_nm = 11
	Const M356_EG2_E1_m_iv_type_stock_flg = 12

	Const C_SHEETMAXROWS_D  = 100        
       
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                      '☜: Clear Error status

	lgStrPrevKey = Request("lgStrPrevKey")
    
	Redim I1_m_iv_type(1)	
	
	I1_m_iv_type(0) =  Request("txtIvType")
	I1_m_iv_type(1) =  Request("rdoUsageFlg")
	
    Set PM1G428 = CreateObject("PM1G428.cMListIvTypeS")

	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if
    
    
   call PM1G428.M_LIST_IV_TYPE_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_iv_type, lgStrPrevKey, E1_m_iv_type, E2_m_iv_type, EG1_export_group, EG2_export_group)  

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 	
		Set PM1G428 = Nothing												'☜: ComProxy Unload
        Exit Sub
	End if 
	if isempty(EG1_export_group) then exit sub
	
	iLngMaxRow = CLng(Request("txtMaxRows"))

	iIntLoopCount = 0
	iMax = UBound(EG1_export_group , 1)
	ReDim TmpBuffer(iMax)
	
	For iLngRow = 0 To iMax
	
	    If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow,M356_EG2_E1_m_iv_type_iv_type_cd))
           Exit For
        End If  
        
        istrData = ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M356_EG2_E1_m_iv_type_iv_type_cd))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M356_EG2_E1_m_iv_type_iv_type_nm))
       
        If EG1_export_group(iLngRow,M356_EG2_E1_m_iv_type_import_flg) = "Y" then
           istrData = istrData & Chr(11) & "1"
        Else
        	istrData = istrData & Chr(11) & "0"
        End If
       	If EG1_export_group(iLngRow,M356_EG2_E1_m_iv_type_except_flg) = "Y" then
           istrData = istrData & Chr(11) & "1"
        Else
        	istrData = istrData & Chr(11) & "0"
        End If
        If EG1_export_group(iLngRow, M356_EG2_E1_m_iv_type_ret_flg ) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        Else
        	istrData = istrData & Chr(11) & "0"
        End If
        '==== 2005.06.22 재고반영여부 추가 ==========
        If EG1_export_group(iLngRow, M356_EG2_E1_m_iv_type_stock_flg ) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        Else
        	istrData = istrData & Chr(11) & "0"
        End If
        '==== 2005.06.22 재고반영여부 추가 ==========
       	If EG1_export_group(iLngRow,M356_EG2_E1_m_iv_type_usage_flg ) = "Y" then
        	istrData = istrData & Chr(11) & "1"
        Else
        	istrData = istrData & Chr(11) & "0"
        End If        
   
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow ,M356_EG2_E1_m_iv_type_trans_cd))
        istrData = istrData & Chr(11) & ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow ,M356_EG2_E1_m_iv_type_trans_nm))
        
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)
        
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next 
    iTotalStr = Join(TmpBuffer, "")
      
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With Parent "               & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData"        & vbCr
	Response.Write " .lgQuery = true	"                       & vbCr
	Response.Write " .ggoSpread.SSShowData     """ & iTotalStr   & """" & vbCr
	Response.Write " .lgStrPrevKey           = """ & StrNextKey & """" & vbCr	
	if E1_m_iv_type(0) <> "*" Then
		Response.Write " .frm1.txtIvTypeNm.value = """ & ConvSPChars(E1_m_iv_type(1)) & """" & vbCr
	end if
	Response.Write " .frm1.hdnUseflg.value   = """ & UCase(Request("rdoUsageFlg"))   & """" & vbCr
	Response.Write " .DbQueryOk "		    	  & vbCr 
	Response.Write " .frm1.vspdData.focus "		  & vbCr 
	Response.Write "End With" & vbCr  
    Response.Write "</Script>" & vbCr
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
 
    Dim M14121
    Dim iErrorPosition
	Dim txtSpread
								
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                      '☜: Clear Error status

    Set M14121 = Server.CreateObject("PM1G421.cMMaintIvTypeS") 
     
    txtSpread = Trim(Request("txtSpread"))
    Call  M14121.M_MAINT_IV_TYPE_SVR(gStrGlobalCollection, txtSpread, iErrorPosition)	  
       
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set M14121 = Nothing												'☜: ComProxy Unload
		exit sub
 	end if
        
    Set M14121 = Nothing	
                 
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
    Response.Write "</Script>"                  & vbCr 
               
End Sub    

%>
