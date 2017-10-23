<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : Costing
'*  2. Function Name        : 원가계산 기준정보 
'*  3. Program ID           : c1210mb1
'*  4. Program Name         : 간접비 배부율 등록 
'*  5. Program Desc         : 공장별 표준계산시 간접비에 대한 배부율을 등록한다.
'*  6. Modified date(First) : 2000/09/02
'*  7. Modified date(Last)  : 2002/08/09
'*  8. Modifier (First)     : Ig Sung, Cho
'*  9. Modifier (Last)      : Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<% Call LoadBasisGlobalInf() 
   Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

   Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD
   Dim lgLngMaxRow

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            


  Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    
  
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


	Dim PAAG090Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey 
    Dim importArray
    Dim iIntLoopCount
	Dim iDeprCd
    Dim lgMaxCount
	
	
	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_DeprMthCd    = 2

	Const C_SHEETMAXROWS_D  = 100
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time
    
    iStrPrevKey   = Trim(Request("lgStrPrevKey"))           '☜: Next Key Value

	
	'Key 값을 읽어온다 
	iDeprCd= Trim(Request("txtDeprCd"))

    'Component 입력변수        
    ReDim importArray(2)
            
    importArray(C_MaxFetchRc)	= lgMaxCount
    importArray(C_NextKey)		= iStrPrevKey
    importArray(C_DeprMthCd)		= iDeprCd
	            
    Set PAAG090Data = Server.CreateObject("PAAG090.cAListDeprRateSvr")
	
	If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus				
       Exit Sub
    End If    
   
    Call PAAG090Data.A_LIST_DEPR_RATE_SVR(gStrGlobalCollection,importArray, exportData, exportData1)

	If CheckSYSTEMError(Err, True) = True Then					
       Set PAAG090Data = Nothing
        Response.Write " <Script Language=vbscript>	                         " & vbCr
        Response.Write " parent.frm1.txtDeprNm.value = """ & ConvSPChars(exportData)		& """" & vbCr
        Response.Write "</Script>  " & vbCr
        Call SetErrorStatus
       Exit Sub
    End If    
        
    Set PAAG090Data = nothing  
    
    Const E_DeprMthCd	= 0
	Const E_DeprMthNm	= 1
	Const E_DurYrs		= 2
	Const E_DeprRate	= 3

	
    iStrData = ""
    iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
	    
	    If  iIntLoopCount < (lgMaxCount + 1) Then
	        iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DeprMthCd)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DeprMthNm)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DurYrs)))
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_DeprRate),ggExchRate.DecPoint,0)			
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_DurYrs)))
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
        Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_DurYrs)
			Exit For
			  
		End If
	Next
	
	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
	End If
	
	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData			& """" & vbCr
    Response.Write " .frm1.txtDeprNm.value = """ & ConvSPChars(exportData(0))		& """" & vbCr
    Response.Write " .frm1.hDeprCd.value = """ & iDeprCd    & """" & vbCr
    Response.Write " .lgStrPrevKey          = """ & ConvSPChars(iStrPrevKey)		& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr
    
End Sub    	 


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PAAG090Data
    Dim importString 
    Dim txtSpread 
    Dim iErrPosition

   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                        '☜: Clear Error status

   importString  = Trim(Request("txtDeprCd"))   
   txtSpread     = Trim(Request("txtSpread"))
   
    Set PAAG090Data = Server.CreateObject("PAAG090.cAMngDeprRateSvr")

    If CheckSYSTEMError(Err, True) = True Then	
	   Call SetErrorStatus				
       Exit Sub
    End If    
	
    Call PAAG090Data.A_MANAGE_DEPR_RATE_SVR(gStrGloBalCollection, importString, txtSpread, iErrPosition)		
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "행","","","","") = True Then					
       Call SetErrorStatus
       Set PAAG090Data = Nothing
       Exit Sub
    End If    
    
    Set PAAG090Data = Nothing

    
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
