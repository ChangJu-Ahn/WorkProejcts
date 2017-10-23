<%@ LANGUAGE=VBSCript%>
<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매관리	
'*  3. Program ID           : S2211MB3
'*  4. Program Name         : 판매계획기간정보수정 
'*  5. Program Desc         : 판매계획기간정보수정 
'*  6. Comproxy List        : PS2G213.dll
'*  7. Modified date(First) : 2003/01/13
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    Dim lgOpModeCRUD
    
	Call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "*", "NOCOOKIE", "MB")     
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
 
    lgOpModeCRUD = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)    
    
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
	Dim iPvStrWhere
	Dim iPvStrNextKey
	Dim iPrArrRsOut
	Dim iPrArrWhereOut
	Dim iStrPrevKey
	Dim iStrNextKey
	Dim iLngSheetMaxRows	' Row numbers to be displayed in the spread.
	
	
	Dim iPS2G213
	Dim iarrValue
	 
    Const C_PS2G213_SP_TYPE_FOR_LIST = 0
    Const C_PS2G213_FR_SP_PERIOD_FOR_LIST = 1
    Const C_PS2G213_TO_SP_PERIOD_FOR_LIST = 2
    
	Const C_PS2G213_sp_type = 0
	Const C_PS2G213_sp_period = 1
	Const C_PS2G213_sp_period_desc = 2
	Const C_PS2G213_sp_period_seq = 3
	Const C_PS2G213_from_dt = 4
	Const C_PS2G213_to_dt = 5
	Const C_PS2G213_sp_year = 6
	Const C_PS2G213_sp_quarter = 7
	Const C_PS2G213_sp_month = 8
	Const C_PS2G213_sp_week = 9
	Const C_PS2G213_sp_create_method = 10
	    
    On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
 
	iPvStrWhere		 = Trim(Request("txtWhere"))
	iPvStrNextKey	 = Request("lgStrPrevKey")
	iLngSheetMaxRows = CLng(100)
				
	Set iPS2G213 = Server.CreateObject("PS2G213.cListSSpPeriodInfo") 

	If CheckSYSTEMError(Err,True) = True Then
		Set iPS2G213 = Nothing   
		Response.Write "<Script language=vbs>    " & vbCr  
		Response.Write "   Parent.SetDefaultVal " & vbCr 
		Response.Write "</Script>      " & vbCr 
		Exit Sub
	End If

	Call iPS2G213.ListRows(gStrGlobalCollection, iLngSheetMaxRows , iPvStrWhere, iPvStrNextKey, iPrArrRsOut, iPrArrWhereOut) 

	If CheckSYSTEMError(Err,True) = True Then
		Set iPS2G213 = Nothing                                                   '☜: Unload Comproxy DLL
		Response.Write "<Script language=vbs>  " & vbCr  
		Response.Write "   Parent.SetDefaultVal " & vbCr 
		Response.Write "</Script>      " & vbCr 
		Exit Sub
	End If   

    Set iPS2G213 = Nothing 
    
    iLngMaxRow  = CLng(Request("txtMaxRows"))           '☜: Fetechd Count      
    
	For iLngRow = 0 To UBound(iPrArrRsOut,1)
		If  iLngRow < iLngSheetMaxRows  Then
		Else
			iStrNextKey = ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_period)) 
			Exit For
		End If 
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_type))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_period))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_period_desc))
		istrData = istrData & Chr(11) & UNIDateClientFormat(iPrArrRsOut(iLngRow, C_PS2G213_from_dt))
		istrData = istrData & Chr(11) & UNIDateClientFormat(iPrArrRsOut(iLngRow, C_PS2G213_to_dt))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_year))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_quarter))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_month))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_week))
		istrData = istrData & Chr(11) & ConvSPChars(iPrArrRsOut(iLngRow, C_PS2G213_sp_create_method))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12) 
    Next    

    Response.Write "<Script language=vbs>  " & vbCr   
    Response.Write " With Parent        " & vbCr

    Response.Write "   .frm1.txtConFrSpPeriod.value = """ & ConvSPChars(iPrArrWhereOut(0,C_PS2G213_FR_SP_PERIOD_FOR_LIST)) & """" & vbCr
    Response.Write "   .frm1.txtConFrSpPeriodDesc.value = """ & ConvSPChars(iPrArrWhereOut(1,C_PS2G213_FR_SP_PERIOD_FOR_LIST)) & """" & vbCr
    Response.Write "   .frm1.txtConToSpPeriod.value = """ & ConvSPChars(iPrArrWhereOut(0,C_PS2G213_TO_SP_PERIOD_FOR_LIST)) & """" & vbCr
    Response.Write "   .frm1.txtConToSpPeriodDesc.value = """ & ConvSPChars(iPrArrWhereOut(1,C_PS2G213_TO_SP_PERIOD_FOR_LIST)) & """" & vbCr    

    Response.Write "   .ggoSpread.Source          =   .frm1.vspdData           " & vbCr
    Response.Write "   .ggoSpread.SSShowDataByClip        """ & istrData        & """" & vbCr
    Response.Write "   .lgStrPrevKey              = """ & iStrNextKey    & """" & vbCr
   	Response.Write "   .DbQueryOk  " & vbCr
    Response.Write " End With       " & vbCr                    
    Response.Write "</Script>      " & vbCr      
               
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()   
                                                                      
	Dim iPS2G213
	Dim itxtSpread
	Dim iErrorPosition
 
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                    '☜: Clear Error status                                                            

	Set iPS2G213 = Server.CreateObject("PS2G213.cMaintSSpPeriodInfo")

    If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If
    
	itxtSpread = Trim(Request("txtSpread"))
	
    Call iPS2G213.Maintain (gStrGlobalCollection, "", itxtSpread , "", iErrorPosition )
                 
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPS2G213 = Nothing
       Exit Sub
	End If
 
    Set iPS2G213 = Nothing
                                                       
    Response.Write "<Script language=vbs> " & vbCr  
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "             & vbCr                                                                        
              
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
