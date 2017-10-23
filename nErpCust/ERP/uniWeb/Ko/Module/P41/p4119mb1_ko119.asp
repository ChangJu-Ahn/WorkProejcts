<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4119mb1_ko119.asp
'*  4. Program Name         : 시간대별 작업지시확정(S)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2006/04/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")

	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD 
    Dim lgLngMaxRow

	On Error Resume Next								'☜: 
	Err.Clear
 

	Call HideStatusWnd
   
'---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'   lgMaxCount        = CInt(Request("lgMaxCount"))
   
'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status


	Dim PP4G172_ko119_Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrPrevKey1
    Dim iStrPrevKey2
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd
	Dim iJobLine   
	Dim lgMaxCount
	Dim lgStrPrevKey	' 이전 값 
    Dim lgStrPrevKey1   ' 이전 값 
    Dim lgStrPrevKey2   ' 이전 값 
    Dim iJobPlanDt      ' 작업계획일자
    Dim iProdOrderNo    ' 제조오더번호
    Dim iItemCd 
    Dim iJobPlanTime    '작업계획시간
    Dim iConfirmFlg
 		

	Const C_MaxFetchRc		= 0
    Const C_NextKey			= 1
    Const C_NextKey1		= 2
    Const C_NextKey2		= 3
    Const C_PlantCd			= 4
    Const C_JobPlanDt		= 5
    Const C_ProdtOrderNo	= 6
    Const C_ItemCd			= 7
    Const C_JobLine			= 8
    Const C_JobPlanTime		= 9
    Const C_ConfirmFlg		= 10
  
	Const C_SHEETMAXROWS_D  = 100 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
	'Key 값을 읽어온다 
	iPlantCd      = Trim(Request("txtPlantCd"))
	iStrPrevKey   = Trim(Request("lgStrPrevKey"))           '☜: Next Key Value
	iStrPrevKey1  = Trim(Request("lgStrPrevKey1"))
	iStrPrevKey2  = Trim(Request("lgStrPrevKey2"))
	iJobPlanDt	  = UNIConvDate(Trim(Request("txtProdFromDt")))
	iJobLine	  = Trim(Request("cboLine"))
	iProdOrderNo  = Trim(Request("txtProdOrderNo"))
	iItemCd		  = Trim(Request("txtItemCd"))
	iJobPlanTime  = Trim(Request("cboTime"))
	iConfirmFlg	  = Trim(Request("txtRadio"))
	

    'Component 입력변수        
    ReDim importArray(10)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
    importArray(C_NextKey)		= iStrPrevKey
    importArray(C_NextKey1)		= iStrPrevKey1
    importArray(C_NextKey2)		= iStrPrevKey2
    importArray(C_PlantCd)		= iPlantCd
    importArray(C_JobPlanDt)	= iJobPlanDt   
	importArray(C_ProdtOrderNo)	= iProdOrderNo   
	importArray(C_ItemCd)		= iItemCd   
	importArray(C_JobLine)		= iJobLine
	importArray(C_JobPlanTime)	= iJobPlanTime     
	importArray(C_ConfirmFlg)	= iConfirmFlg   
   
    Set PP4G172_ko119_Data = Server.CreateObject("PP4G172_KO119.cPListTWOrdConf")
    
	If CheckSYSTEMError(Err, True) = True Then
	   Call SetErrorStatus					
       Exit Sub
    End If    
   
    Call PP4G172_ko119_Data.C_LIST_TWOrd_Conf_Svr(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PP4G172_ko119_Data = Nothing
        Response.Write " <Script Language=vbscript>	                         " & vbCr
        Response.Write " parent.frm1.txtPlantNm.value = """ & ConvSPChars(exportData)			& """" & vbCr
        Response.Write "</Script>  " & vbCr 
        Call SetErrorStatus
       Exit Sub
    End If    
        
    Set PP4G172_ko119_Data = nothing    
	
	Const E_CheckBox	 = 0		
	Const E_ItemCd		 = 1
	Const E_ItemNm		 = 2
	Const E_Spec		 = 3
	Const E_JobPlanDt	 = 4
	Const E_JobLine      = 5
	Const E_JobPlanTime  = 6
	Const E_JobQty		 = 7
	Const E_JobSeq		 = 8
	Const E_JobOrderNo	 = 9
	Const E_SecItemCd	 = 10
	Const E_ProdtOrderNo = 11
	Const E_RoutNo		 = 12
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_CheckBox)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ItemCd)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ItemNm)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_Spec)))
   			iStrData = iStrData & Chr(11) & UNIDateClientFormat(Trim(exportData1(iLngRow, E_JobPlanDt)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_JobLine)))
   			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_JobLine)))
'			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_JobPlanTime)))
			iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(Trim(exportData1(iLngRow, E_JobQty)),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_JobSeq)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_JobOrderNo)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_SecItemCd)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_ProdtOrderNo)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_RoutNo)))
            iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_ItemCd)
			iStrPrevKey1 = exportData1(UBound(exportData1, 1), E_JobLine)
			IStrPrevKey2 = exportData1(UBound(exportData1, 1), E_JobPlanTime)
			Exit For
		End If
	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
		iStrPrevKey1 = ""
	End If

	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData)			& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & iPlantCd    & """" & vbCr
    Response.Write " .frm1.hProdFromDt.value = """ & iJobPlanDt			& """" & vbCr
    Response.Write " .lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)		& """" & vbCr
    Response.Write " .lgStrPrevKey1 = """ & ConvSPChars(iStrPrevKey1)    	& """" & vbCr  
    Response.Write " .lgStrPrevKey2 = """ & ConvSPChars(iStrPrevKey2)    	& """" & vbCr   
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr

End Sub    	 



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PP4G172_ko119_Data
    Dim importString
    Dim txtSpread
    Dim iErrPosition 

   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                        '☜: Clear Error status

    
   importString = Trim(Request("txtPlantCd"))
   txtSpread    = Trim(Request("txtSpread"))

    Set PP4G172_ko119_Data = Server.CreateObject("PP4G172_KO119.cPMngTWOrdConf")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PP4G172_ko119_Data.P_MANAGE_TWOrd_Conf(gStrGlobalCollection, importString, txtSpread, iErrPosition)			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
       Call SetErrorStatus
       Set PP4G172_ko119_Data = Nothing
       Exit Sub
    End If    
    
    Set PP4G172_ko119_Data = Nothing
	
    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
