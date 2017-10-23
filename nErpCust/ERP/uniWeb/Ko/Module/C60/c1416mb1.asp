<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>


<%'**********************************************************************************************
'*  1. Module��          : ���� 
'*  2. Function��        : C_cost_Element_by_Resource
'*  3. Program ID        : c1416ma
'*  4. Program �̸�      : ������ ������� ��� 
'*  5. Program ����      : ������ ������� ��� ���� ���� ��ȸ 
'*  6. Comproxy ����Ʈ   : c11021, c11028 , ...
'*  7. ���� �ۼ������   : 2000/09/04
'*  8. ���� ���������   : 2002/08/09
'*  9. ���� �ۼ���       : �ۺ��� 
'* 10. ���� �ۼ���       : Cho Ig sung  / Park, Joon-Won
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         -2000/08/17 : ..........
'**********************************************************************************************

Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%

	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB") 													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	
	Dim lgErrorStatus, lgErrorPos, lgOpModeCRUD
	Dim lgLngMaxRow


	On Error Resume Next								'��: 
	Err.Clear

	Call HideStatusWnd
	
'---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    
    'Multi SpreadSheet
	lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
'   lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData

'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             'Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status


	Dim PC1G020Data		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey,iStrPrevKey1 
    Dim importArray
    Dim iIntLoopCount
	Dim iPlantCd, iRscGRpCD
	Dim lgMaxCount

	
	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_NextKey1   = 2
	Const C_PlantCd    = 3
	Const C_RscGRpCD   = 4 	
    
	Const C_SHEETMAXROWS_D  = 100                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)
    
	'Key ���� �о�´� 
	iPlantCd        = Trim(Request("txtPlantCd"))
	iRscGRpCD       = Trim(Request("txtRscGRpCD"))
	iStrPrevKey		= Trim(Request("lgStrPrevKey"))         '��: Next Key Value
	iStrPrevKey1	= Trim(Request("lgStrPrevKey1"))         '��: Next Key Value

	
    'Component �Էº���        
    ReDim importArray(4)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	importArray(C_NextKey1)		= iStrPrevKey1
    importArray(C_PlantCd)		= iPlantCd
	importArray(C_RscGRpCD)     = iRscGRpCD 
   
	ReDim exportData(1)
	
    Set PC1G020Data = Server.CreateObject("PC1G020.cCListCeByRsrcSvr")
    
	If CheckSYSTEMError(Err, True) = True Then
	   Call SetErrorStatus					
       Exit Sub
    End If    
   
    Call PC1G020Data.C_LIST_CE_BY_RESOURCE_SVR(gStrGlobalCollection,importArray, exportData, exportData1)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PC1G020Data = Nothing
       Response.Write " <Script Language=vbscript>	                         " & vbCr 
       Response.Write " parent.frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))	& """" & vbCr
       Response.Write " parent.frm1.txtRscGRpNM.value = """ & ConvSPChars(exportData(1))	& """" & vbCr
       Response.Write "</Script>  " & vbCr 
       Call SetErrorStatus
       Exit Sub
    End If    
        
    Set PC1G020Data = nothing 
    
    const E_RESOURCE_GRP_CD = 0
	const E_RESOURCE_GRP_NM = 1
	Const E_CE_CD = 2
	Const E_CE_NM = 3
	Const E_COMPOSITE_RATE	= 4
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_RESOURCE_GRP_CD)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_RESOURCE_GRP_NM)))
			iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData1(iLngRow, E_CE_CD)))
			iStrData = iStrData & Chr(11) & ""
			iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData1(iLngRow, E_COMPOSITE_RATE),ggExchRate.DecPoint,0)			
			iStrData = iStrData & Chr(11) & Cstr(lgLngMaxRow + iLngRow + 1) 
			iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), E_RESOURCE_GRP_CD)
			iStrPrevKey1 = exportData1(UBound(exportData1, 1), E_CE_CD)
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
    Response.Write " .frm1.txtPlantNm.value = """ & ConvSPChars(exportData(0))	& """" & vbCr
    Response.Write " .frm1.txtRscGRpNM.value = """ & ConvSPChars(exportData(1))	& """" & vbCr
    Response.Write " .frm1.hPlantCd.value = """ & iPlantCd    & """" & vbCr
    Response.Write " .frm1.hRscGrpCd.value = """ & iRscGRpCD  & """" & vbCr
    Response.Write " .lgStrPrevKey        = """ & ConvSPChars(iStrPrevKey)		& """" & vbCr
    Response.Write " .lgStrPrevKey1        = """ & ConvSPChars(iStrPrevKey1)	& """" & vbCr
'   Response.Write " .DbQueryOk " & vbCr
    Response.Write "End With   " & vbCr
    Response.Write "</Script>  " & vbCr

End Sub    	 



'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim PC1G020Data
    Dim importString 
	Dim txtSpread
	Dim iErrPosition 

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    
    importString = Trim(Request("txtPlantCd"))
    txtSpread    = Trim(Request("txtSpread"))
    

    Set PC1G020Data = Server.CreateObject("PC1G020.cCMngCeByRsrcGp")

    If CheckSYSTEMError(Err, True) = True Then					
       Call SetErrorStatus
       Exit Sub
    End If    

    Call PC1G020Data.C_MANAGE_CE_BY_RESOURCE_SVR(gStrGlobalCollection, importString, txtSpread, iErrPosition)			
		
    If CheckSYSTEMError2(Err, True ,iErrPosition & "��","","","","") = True Then					
	   Call SetErrorStatus
       Set PC1G020Data = Nothing
       Exit Sub
    End If    
    
    Set PC1G020Data = Nothing
	
   
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
   On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
'          Else
'            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
</Script>	
