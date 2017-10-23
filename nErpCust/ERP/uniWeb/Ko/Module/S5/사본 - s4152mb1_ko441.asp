<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "*","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey      = UNICInt(Trim(Request("lgStrPrevKey")),0)                    '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
   
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim strWhere 
	Dim strBp_cd, strPlant_cd, strItem_cd, strOutType, strFrDt, strToDt
  
    On Error Resume Next    
    Err.Clear                                                               '¢Ð: Clear Error status

	strBp_cd = Trim(Request("txtconBp_cd"))
	strPlant_cd = Trim(Request("txtPlantCode"))
	strItem_cd = Trim(Request("txtconItem_cd"))
	strOutType = Trim(Request("txtconOutType"))
	strFrDt = Trim(Request("txtconFr_dt"))
	strToDt = Trim(Request("txtconTo_dt"))

	If strBp_cd <> "" Then
		strWhere = " And (SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=SHIP_TO_PARTY and USAGE_FLAG = 'y') = " & FilterVar(strBp_cd, "''", "S")
	End If
	If strPlant_cd <> "" Then
		strWhere = strWhere & " And PLANT_CD = " & FilterVar(strPlant_cd, "''", "S")
	End If

	If strItem_cd <> "" Then
		strWhere = strWhere & " And DBO.UFN_GETITEMCD(MES_ITEM_CD) = " & FilterVar(strItem_cd, "''", "S")
	End If
	If strOutType <> "" Then
		strWhere = strWhere & " And OUT_TYPE = " & FilterVar(strOutType, "''", "S")
	End If

	If strFrDt <> "" Then
		strWhere = strWhere & " And convert(varchar(10),ACTUAL_GI_DT,121) >= " & FilterVar(strFrDt, "''", "S")
	End If
	If strToDt <> "" Then
		strWhere = strWhere & " And convert(varchar(10),ACTUAL_GI_DT,121) <= " & FilterVar(strToDt, "''", "S")
	End If

   	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '¢Ð : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD") )                 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACTUAL_GI_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CUST_LOT_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_TYPE_Nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GI_QTY"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GI_UNIT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TOT_ISSUE_PRICE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TOT_ISSUE_AMT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TOT_ISSUE_AMT_LOC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("XCHG_RATE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CURRENCY"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT1"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE3"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT3"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE4"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT4"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE5"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT5"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE6"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT6"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE7"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT7"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE8"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT8"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE9"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT9"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE10"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT10"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE11"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT11"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE12"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT12"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE13"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT13"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE14"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT14"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_PRICE15"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ISSUE_AMT15"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pgm_name"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRANS_TIME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_TYPE_SUB"))

            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

    	    lgObjRs.MoveNext
          
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
                       
        Loop 
    End If
      If iDx <= C_SHEETMAXROWS_D Then
         lgStrPrevKey = ""
      Else
         if lgStrPrevKey = 1 Then
%>
<Script Language=vbscript>
       With Parent	
		  .frm1.txtHconBp_cd.value = .frm1.txtconBp_cd.value
		  .frm1.txtHconItem_cd.value = .frm1.txtconItem_cd.value
		  .frm1.txtHconFr_dt.value = .frm1.txtconFr_dt.text
		  .frm1.txtHconTo_dt.value = .frm1.txtconTo_dt.text
		  .frm1.txtHconOutType.value = .frm1.txtconOutType.value 
		  .frm1.txtHPlantCode.value = .frm1.txtPlantCode.value

       End With          
</Script>       
<%     
         End If   
    
      End If   

      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear                                                                        '¢Ð: Clear Error status

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
    Next

End Sub      

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    on error resume next
    Err.Clear  
    

' Call svrmsgbox (lgstrsql, vbinformation, i_mkscript)
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    on error resume next
     Err.Clear  
    
    lgStrSQL = "UPDATE  T_IF_RCV_PART_OUT_KO441 "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " ISSUE_PRICE1	=       " & UNIConvNum(arrColVal(5),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE2	=       " & UNIConvNum(arrColVal(6),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE3	=       " & UNIConvNum(arrColVal(7),0)	& ","
	lgStrSQL = lgStrSQL & " ISSUE_PRICE4	=       " & UNIConvNum(arrColVal(8),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE5	=       " & UNIConvNum(arrColVal(9),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE6	=       " & UNIConvNum(arrColVal(10),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE7	=       " & UNIConvNum(arrColVal(11),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE8	=       " & UNIConvNum(arrColVal(12),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE9	=       " & UNIConvNum(arrColVal(13),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE10	=       " & UNIConvNum(arrColVal(14),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE11	=       " & UNIConvNum(arrColVal(15),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE12	=       " & UNIConvNum(arrColVal(16),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE13	=       " & UNIConvNum(arrColVal(17),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE14	=       " & UNIConvNum(arrColVal(18),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_PRICE15	=       " & UNIConvNum(arrColVal(19),0)	& ","

	lgStrSQL = lgStrSQL & " ISSUE_AMT1		=       " & UNIConvNum(arrColVal(20),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT2		=       " & UNIConvNum(arrColVal(21),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT3		=       " & UNIConvNum(arrColVal(22),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT4		=       " & UNIConvNum(arrColVal(23),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT5		=       " & UNIConvNum(arrColVal(24),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT6		=       " & UNIConvNum(arrColVal(25),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT7		=       " & UNIConvNum(arrColVal(26),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT8		=       " & UNIConvNum(arrColVal(27),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT9		=       " & UNIConvNum(arrColVal(28),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT10		=       " & UNIConvNum(arrColVal(29),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT11		=       " & UNIConvNum(arrColVal(30),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT12		=       " & UNIConvNum(arrColVal(31),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT13		=       " & UNIConvNum(arrColVal(32),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT14		=       " & UNIConvNum(arrColVal(33),0)	& ","
    lgStrSQL = lgStrSQL & " ISSUE_AMT15		=       " & UNIConvNum(arrColVal(34),0)	& ","

    lgStrSQL = lgStrSQL & " TOT_ISSUE_PRICE =       " & UNIConvNum(arrColVal(35),0)	& ","
    lgStrSQL = lgStrSQL & " TOT_ISSUE_AMT	=       " & UNIConvNum(arrColVal(36),0)	& ","
    lgStrSQL = lgStrSQL & " TOT_ISSUE_AMT_LOC	=   " & UNIConvNum(arrColVal(37),0)	& ","
        
    lgStrSQL = lgStrSQL & " UPDT_DT			=       " & FilterVar(lgSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID	=       " & FilterVar(gUsrId, "''", "S")  
    lgStrSQL = lgStrSQL & " WHERE					"
    lgStrSQL = lgStrSQL & " OUT_NO			=       " & FilterVar(UCase(arrColVal(2)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " TRANS_TIME		=       " & FilterVar(UCase(arrColVal(3)), "''", "S") & " AND "
    lgStrSQL = lgStrSQL & " OUT_TYPE_SUB	=       " & FilterVar(UCase(arrColVal(4)), "''", "S") 



'Response.Write lgStrSQL
'Response.End 
  
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status


    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
     Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
        
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
           
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount  
                       lgStrSQL = lgStrSQL & " SHIP_TO_PARTY,ITEM_CD = DBO.UFN_GETITEMCD(MES_ITEM_CD),ACTUAL_GI_DT=convert(varchar(10),ACTUAL_GI_DT,121), OUT_NO, CUST_LOT_NO, LOT_NO, OUT_TYPE, "   
                       lgStrSQL = lgStrSQL & "    GI_QTY,GI_UNIT, CURRENCY, PO_NO,  rtrim(PGM_NAME) as pgm_name , TOT_ISSUE_PRICE, TOT_ISSUE_AMT, TOT_ISSUE_AMT_LOC, XCHG_RATE, TRANS_TIME, OUT_TYPE_SUB, "  
                       lgStrSQL = lgStrSQL & "    ISSUE_PRICE1, ISSUE_PRICE2, ISSUE_PRICE3, ISSUE_PRICE4, ISSUE_PRICE5, ISSUE_PRICE6, ISSUE_PRICE7, ISSUE_PRICE8,  "  
                       lgStrSQL = lgStrSQL & "    ISSUE_PRICE9, ISSUE_PRICE10, ISSUE_PRICE11, ISSUE_PRICE12, ISSUE_PRICE13, ISSUE_PRICE14, ISSUE_PRICE15,  "  
                       lgStrSQL = lgStrSQL & "    ISSUE_AMT1, ISSUE_AMT2, ISSUE_AMT3, ISSUE_AMT4, ISSUE_AMT5, ISSUE_AMT6, ISSUE_AMT7, ISSUE_AMT8, "  
                       lgStrSQL = lgStrSQL & "    ISSUE_AMT9, ISSUE_AMT10, ISSUE_AMT11, ISSUE_AMT12, ISSUE_AMT13, ISSUE_AMT14, ISSUE_AMT15,  "  
                       lgStrSQL = lgStrSQL & "    ITEM_NM = dbo.ufn_x_getcodename('B_ITEM',DBO.UFN_GETITEMCD(MES_ITEM_CD),''), "  
                       lgStrSQL = lgStrSQL & "    PLANT_NM = dbo.ufn_x_getcodename('B_PLANT',PLANT_CD,''), "  
                       lgStrSQL = lgStrSQL & "    OUT_TYPE_NM=dbo.ufn_x_getcodename('B_USER_MINOR', OUT_TYPE, 'ZZ002') "  
                       lgStrSQL = lgStrSQL & " FROM T_IF_RCV_PART_OUT_KO441 "
                       lgStrSQL = lgStrSQL & " WHERE PLANT_CD in ('P01','P02') "  & pCode 
                       lgStrSQL = lgStrSQL & " ORDER BY OUT_NO, TRANS_TIME, OUT_TYPE_SUB " 


'Call svrmsgbox (lgstrsql, vbinformation, i_mkscript)

'Response.Write lgStrSQL
'Response.End 
          End Select 
    End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
    Response.Write "<BR> Commit Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    Response.Write "<BR> Abort Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MS"
                 Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MX"
                 Call DisplayMsgBox("800350", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MY"
                 Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
		  Else
                Parent.DBQueryFail  		  	         
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
    End Select       
       
</Script>	
