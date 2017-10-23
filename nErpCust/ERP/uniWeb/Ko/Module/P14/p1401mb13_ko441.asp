<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd                                                               '☜: Hide Processing message

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim strPlantCd
Dim strItemCd
Dim strBomNo
Dim strBaseDt
Dim strBaseQty
Dim strExpFlg
Dim strSpId
Dim intRetCD

Dim TmpBuffer
Dim iTotalStr

'---------------------------------------Common-----------------------------------------------------------
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space

'------ Developer Coding part (Start ) ------------------------------------------------------------------

lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount        = 100							                                 '☜: Fetch count at a time for VspdData
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
ReDim TmpBuffer(0)
    
'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
Call SubOpenDB(lgObjConn)                    
Call SubCreateCommandObject(lgObjComm)
     
Call SubMakeParameter()
Call SubBomExplode()
     
Call SubBizQuery()

Call SubCloseCommandObject(lgObjComm)
Call SubCloseDB(lgObjConn)       


Response.Write "<Script Language = VBScript>" & vbCrLf
	If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
	    Response.Write "With Parent" & vbCrLf
	       Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
	       Response.Write ".lgStrPrevKeyIndex = """ & lgStrPrevKeyIndex & """" & vbCrLf
	       Response.Write ".ggoSpread.SSShowDataByClip """ & ConvSPChars(iTotalStr) & """" & vbCrLf
	'                Response.Write ".lgStrPrevKey = """ & lgStrPrevKey & """" & vbCrLf
	       Response.Write ".DBQueryOk()" & vbCrLf        
		Response.Write "End with" & vbCrLf
	End If
Response.Write "</Script>" & vbCrLf

Response.End

'============================================================================================================
' Name : SubMakeParameter
' Desc : Make SP Parameter
'============================================================================================================
Sub SubMakeParameter()
	
	If CInt(Request("txtMode")) <> UID_M0001 Then
		Response.End 
	End If

	strPlantCd = Trim(Request("txtPlantCd"))									' 조회할 키 
	strItemCd = Trim(Request("txtItemCd"))									' 조회할 상위키 
	strBomNo = Trim(Request("txtBomNo"))
	strBaseDt = UniConvDate(Request("txtBaseDt"))
	strBaseQty = UniConvNum((Request("txtBaseQty")),0)
	strExpFlg = 5
    
End Sub     
'============================================================================================================
' Name : SubBomExplode
' Desc : Query Data from Db
'============================================================================================================
Sub SubBomExplode()

    Dim strMsg_cd
    Dim strMsg_text
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc

	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, strExpFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, strItemCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4,strBomNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	advarXchar,adParamInput,10,strBaseDt)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_qty",	adInteger,adParamInput,15,strBaseQty)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	advarXchar,adParamOutput,60)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	advarXchar,adParamOutput,13)
		
        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        
        If  IntRetCD <> 0 then
            
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            strSpId = FilterVar(lgObjComm.Parameters("@user_id").Value, "''", "S")
            
            If strMsg_cd <> MSG_OK_STR Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
			End If

			IntRetCD = -1
            Response.End 
        Else
			strSpId = FilterVar(lgObjComm.Parameters("@user_id").Value, "''", "S")
			IntRetCD = 1
        End If
    Else           
		strSpId = "''"
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
    End if
    
End Sub	

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iDx
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                          
	
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	strBomNo = FilterVar(Trim(Request("txtBomNo"))	, "''", "S")
	
	'--------------
	'공장 체크		
	'--------------	
	lgStrSQL = ""
	Call SubMakeSQLStatements("P_CK",strPlantCd,"","")           '☜ : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
		IntRetCD = -1
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtPlantNm.Value  = """"" & vbCrLf   'Set condition area
			Response.Write "parent.Frm1.txtPlantCd.focus" & vbCrLf   'Set condition area
		Response.Write "</Script>" & vbcRLf
		Response.End
	Else
		IntRetCD = 1
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf 'Set condition area
		Response.Write "</Script>" & vbcRLf
	End If
		
	Call SubCloseRs(lgObjRs) 
	
	'------------------
	'품목체크 
	'------------------
	lgStrSQL = ""
	Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"")           '☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
		IntRetCD = -1
				
		Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()

		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtItemNm.Value  = """"" & vbCrLf   'Set condition area
			Response.Write "parent.Frm1.txtItemCd.Focus" & vbCrLf 
		Response.Write "</Script>" & vbcRLf
		Response.End
	Else
		IntRetCD = 1
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.Frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(0))) & """" & vbCrLf
		Response.Write "</Script>" & vbcRLf
	End If
		
	Call SubCloseRs(lgObjRs) 
	
	'------------------
	' bom type 체크 
	'------------------
	lgStrSQL = ""
			
	Call SubMakeSQLStatements("BT_CK",strBomNo,"","")           '☜ : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
		IntRetCD = -1
				
		Call DisplayMsgBox("182622", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBomNo.focus" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End							
	Else
		IntRetCD = 1
	End If
		
	Call SubCloseRs(lgObjRs) 
	
	'---------------------------
	' Header Single 조회 
	'---------------------------    
	lgStrSQL = ""
			
	Call SubMakeSQLStatements("H",strPlantCd,strItemCd,strBomNo)           '☜ : Make sql statements
     	 
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
	
		IntRetCD = -1
		Call DisplayMsgBox("182600", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.End 
	Else
		IntRetCD = 1

		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "With Parent" & vbCrLf
				Response.Write ".frm1.txtBomNo1.value = """ & ConvSPChars(lgObjRs(0)) & """" & vbCrLf
				Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf
				Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(lgObjRs(5)) & """" & vbCrLf
				Response.Write ".frm1.txtBasicUnit.value = """ & ConvSPChars(lgObjRs(12)) & """" & vbCrLf
		    Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
		    
	Call SubCloseRs(lgObjRs) 
	
	'---------------------------
	' 집약정전개 결과 조회 
	'---------------------------    
	lgStrSQL = ""
			
	Call SubMakeSQLStatements("E",strPlantCd,strSpId,"")           '☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
	
		IntRetCD = -1
				
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.End 
	Else
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        iDx       = 1

        Do While Not lgObjRs.EOF
			lgstrData = ""
            lgstrData = lgstrData & Chr(11) & lgObjRs(0)
			lgstrData = lgstrData & Chr(11) & lgObjRs(1)
			lgstrData = lgstrData & Chr(11) & lgObjRs(2)
            lgstrData = lgstrData & Chr(11) & lgObjRs(3)
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(4), 6, 3, "", 0)
            lgstrData = lgstrData & Chr(11) & lgObjRs(5)
            lgstrData = lgstrData & Chr(11) & lgObjRs(7)
            
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
			
			lgObjRs.MoveNext
			
			ReDim Preserve TmpBuffer(iDx-1)
			
			TmpBuffer(iDx-1) = lgstrData
            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
        Loop 
    End If
	
	iTotalStr = Join(TmpBuffer, "")
	
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                             
		
	lgStrSQL = ""
	'-------------------------
	' 생성된 temp table 삭제 
	'-------------------------
    lgStrSQL = "DELETE FROM p_bom_for_explosion "
	lgStrSQL = lgStrSQL & " WHERE plant_cd = " & FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	lgStrSQL = lgStrSQL & " AND user_id = " & strSpId
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
  '  lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub	    
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType
	
		Case "H"	
			lgStrSQL = "SELECT a.BOM_NO, a.DESCRIPTION, b.* FROM P_BOM_HEADER a, B_ITEM b " 
			lgStrSQL = lgStrSQL & " WHERE a.ITEM_CD = b.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.BOM_NO = " & pCode2
			
		Case "E"
			lgStrSQL = " SELECT a.CHILD_ITEM_CD, b.ITEM_NM, b.SPEC, a.CHILD_BOM_NO, SUM(a.CHILD_ITEM_QTY) AS SUM_CHILD_ITEM_QTY, a.CHILD_ITEM_UNIT, a.MATERIAL_FLG, c.MINOR_NM "
			lgStrSQL = lgStrSQL & " FROM P_BOM_FOR_EXPLOSION a, B_ITEM b, B_MINOR c "
			lgStrSQL = lgStrSQL & " WHERE a.child_item_cd = b.item_cd and a.material_flg = c.minor_cd and c.major_cd = " & FilterVar("p1003", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND a.user_id = " & pCode1
			lgStrSQL = lgStrSQL & " AND b.phantom_flg <> " & FilterVar("Y", "''", "S") & "  "
			lgStrSQL = lgStrSQL & " GROUP BY a.child_item_cd, a.child_bom_no, b.item_nm, b.SPEC, a.child_item_unit, a.material_flg, c.minor_nm "
			
		Case "BT_CK"
			lgStrSQL = "SELECT * FROM b_minor WHERE major_cd = " & FilterVar("P1401", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND minor_cd = " & pCode 
			
		Case "I_CK"
			lgStrSQL = "SELECT b.item_nm FROM b_item_by_plant a, b_item b, b_minor c, b_minor d "
			lgStrSQL = lgStrSQL & " WHERE a.item_cd =b.item_cd and c.minor_cd = a.item_acct and d.minor_cd = a.procur_type and c.major_cd =" & FilterVar("p1001", "''", "S") & " and d.major_cd =" & FilterVar("p1003", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1
			
		Case "P_CK"
			lgStrSQL = "SELECT * FROM b_plant where plant_cd = " & pCode 

	End Select    
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
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub
%>
