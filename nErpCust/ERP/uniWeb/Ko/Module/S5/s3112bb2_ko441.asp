
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% 

    Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
    Call LoadBasisGlobalInf()

    Dim lgStrPrevKey
    On Error Resume Next                                                   '☜: Protect prorgram from crashing

    Err.Clear                                                              '☜: Clear Error status
    
    Call HideStatusWnd                                                     '☜: Hide Processing message

    lgErrorStatus  = ""
    lgKeyStream    = Split(Request("txtKeyStream"),gColSep) 
    lgStrPrevKey   = UNICInt(Trim(Request("lgStrPrevKey2")), 0)                   '☜: Next Key

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    If lgStrPrevKey = 0 Then

       lgStrSQL = "Select plant_cd,plant_nm " 
       lgStrSQL = lgStrSQL & " From B_Plant (Nolock) " 
       lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(lgKeyStream(0),"''", "S")

       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
'20080220::hanc           Response.Write  " <Script Language=vbscript>            " & vbCr
'20080220::hanc           Response.Write  "   Parent.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
'20080220::hanc           Response.Write  "   Parent.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
'20080220::hanc           Response.Write  "   Parent.htxtPlantCd.Value = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set next key data
'20080220::hanc           Response.Write  " </Script> " & vbCr
          Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
          Call SubBizQueryMulti()
       Else
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
       End If
    Else
       Call SubBizQueryMulti()
    End If 
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim iSelCount
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
        
    lgStrSQL = " SELECT OUT_NO,                                " & vbcrlf
    lgStrSQL = lgStrSQL & "       SHIP_TO_PARTY,               " & vbcrlf
    lgStrSQL = lgStrSQL & "       SHIP_TO_PARTY_NM,            " & vbcrlf
    lgStrSQL = lgStrSQL & "       ITEM_CD,                     " & vbcrlf
    lgStrSQL = lgStrSQL & "       PLANT_CD,                    " & vbcrlf
    lgStrSQL = lgStrSQL & "       OUT_TYPE,                    " & vbcrlf
    lgStrSQL = lgStrSQL & "       OUT_TYPE_NM,                 " & vbcrlf
    lgStrSQL = lgStrSQL & "       GI_QTY,                      " & vbcrlf
    lgStrSQL = lgStrSQL & "       GI_UNIT,                     " & vbcrlf
    lgStrSQL = lgStrSQL & "       LOT_NO,                      " & vbcrlf
    lgStrSQL = lgStrSQL & "       LOT_SEQ,                     " & vbcrlf
    lgStrSQL = lgStrSQL & "       ACTUAL_GI_DT,                " & vbcrlf
    lgStrSQL = lgStrSQL & "       ITEM_NM,                     " & vbcrlf
    lgStrSQL = lgStrSQL & "       SPEC,                        " & vbcrlf
    lgStrSQL = lgStrSQL & "	      FLAG, TRANS_TIME, CREATE_TYPE,OUT_TYPE_SUB  " & vbcrlf
    lgStrSQL = lgStrSQL & "FROM                                " & vbcrlf
    lgStrSQL = lgStrSQL & "(                                   " & vbcrlf
    lgStrSQL = lgStrSQL & " SELECT	A.OUT_NO,                  " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B.BP_CD SHIP_TO_PARTY,         " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B.BP_NM SHIP_TO_PARTY_NM,      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		D.ITEM_CD,                      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.PLANT_CD,                     " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.OUT_TYPE,                     " & vbcrlf
    lgStrSQL = lgStrSQL & " 		C.UD_MINOR_NM OUT_TYPE_NM,      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.GI_QTY,                       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.GI_UNIT,                      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.LOT_NO,                       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.LOT_SEQ,                      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.ACTUAL_GI_DT,                 " & vbcrlf
    lgStrSQL = lgStrSQL & " 		D.ITEM_NM,              				" & vbcrlf
    lgStrSQL = lgStrSQL & " 		D.SPEC,                    			" & vbcrlf
    lgStrSQL = lgStrSQL & " 		'A' FLAG, A.TRANS_TIME, A.CREATE_TYPE,A.OUT_TYPE_SUB        " & vbcrlf
    lgStrSQL = lgStrSQL & " FROM	T_IF_RCV_PART_OUT_KO441 A,       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B_BIZ_PARTNER B,       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B_USER_DEFINED_MINOR C,      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B_ITEM   D    " & vbcrlf
    lgStrSQL = lgStrSQL & " WHERE   A.SHIP_TO_PARTY *= B.BP_ALIAS_NM         " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     A.MES_ITEM_CD = D.CBM_DESCRIPTION            " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     A.OUT_TYPE = C.UD_MINOR_CD               " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     C.UD_MAJOR_CD = 'ZZ002'                  " & vbcrlf
    lgStrSQL = lgStrSQL & " AND		(SELECT BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=A.SHIP_TO_PARTY and USAGE_FLAG = 'y') LIKE " & FilterVar(Trim(Request("txtShipToParty"))&"%","''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & " AND		A.PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd")),"''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & " AND		D.ITEM_CD = " & FilterVar(Trim(Request("txtItemCd")),"''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     C.UD_REFERENCE = 'Y' " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     ISNULL(A.ERP_APPLY_FLAG,'N') <> 'Y' " & vbcrlf
    lgStrSQL = lgStrSQL & " UNION ALL	" & vbcrlf
    lgStrSQL = lgStrSQL & " SELECT	A.OUT_NO,                   " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B.BP_CD SHIP_TO_PARTY,          " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B.BP_NM SHIP_TO_PARTY_NM,       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		D.ITEM_CD,                      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.PLANT_CD,                     " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.OUT_TYPE,                     " & vbcrlf
    lgStrSQL = lgStrSQL & " 		C.UD_MINOR_NM OUT_TYPE_NM,      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.GI_QTY,                       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.GI_UNIT,                      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.LOT_NO,                       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.LOT_SEQ,                      " & vbcrlf
    lgStrSQL = lgStrSQL & " 		A.ACTUAL_GI_DT,                 " & vbcrlf
    lgStrSQL = lgStrSQL & " 		D.ITEM_NM,              				" & vbcrlf
    lgStrSQL = lgStrSQL & " 		D.SPEC,                    			" & vbcrlf
    lgStrSQL = lgStrSQL & " 		'B' FLAG, A.TRANS_TIME, A.CREATE_TYPE,A.OUT_TYPE_SUB        " & vbcrlf
    lgStrSQL = lgStrSQL & " FROM	T_IF_RCV_PART_OUT_KO441 A,       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B_BIZ_PARTNER B,       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B_USER_DEFINED_MINOR C,       " & vbcrlf
    lgStrSQL = lgStrSQL & " 		B_ITEM   D    " & vbcrlf
    lgStrSQL = lgStrSQL & " WHERE   A.SHIP_TO_PARTY *= B.BP_ALIAS_NM         " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     A.MES_ITEM_CD = D.CBM_DESCRIPTION            " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     A.OUT_TYPE = C.UD_MINOR_CD               " & vbcrlf
    lgStrSQL = lgStrSQL & " AND     C.UD_MAJOR_CD = 'ZZ002'                  " & vbcrlf
    LGSTRSQL = LGSTRSQL & " AND     ISNULL(A.ERP_APPLY_FLAG,'N') <> 'Y' " & vbcrlf
    lgStrSQL = lgStrSQL & " AND		(SELECT BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=A.SHIP_TO_PARTY and USAGE_FLAG = 'y') LIKE " & FilterVar(Trim(Request("txtShipToParty"))&"%","''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & " AND		A.PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd")),"''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & " AND		D.ITEM_CD = " & FilterVar(Trim(Request("txtItemCd")),"''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & " AND C.UD_REFERENCE = " & FilterVar(Trim(Request("txtDnType")),"''", "S") & " " & vbcrlf
    lgStrSQL = lgStrSQL & "    ) X " & vbcrlf
    lgStrSQL = lgStrSQL & "WHERE X.FLAG = " & FilterVar(Trim(Request("txtFlag")),"''", "S") & " " & vbcrlf

'<textarea id="txtSQL" name="txtSQL" rows=10 cols=10>lgStrSQL</textarea>

        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
        lgStrPrevKey  = ""
        lgErrorStatus = "YES"
        Exit Sub 
    Else    
      
	   If CDbl(lgStrPrevKey) > 0 Then
		  lgObjRs.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgStrPrevKey)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	   End If   
       iDx = 1		
       
       lgstrData = ""
       lgLngMaxRow       = CLng(Request("txtMaxRows"))

       Do While Not lgObjRs.EOF
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_NO"))                           
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SHIP_TO_PARTY"))                           
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SHIP_TO_PARTY_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_TYPE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_TYPE_NM"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("GI_QTY"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GI_UNIT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_SEQ"))
          lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ACTUAL_GI_DT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRANS_TIME"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CREATE_TYPE"))
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
    End If            
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
       Exit Sub
    End If   
       
    If lgErrorStatus = "" Then
       Response.Write  " <Script Language=vbscript>                                  " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.vspdData2       " & vbCr
       Response.Write  "    Parent.lgStrPrevKey2         = """ & lgStrPrevKey    & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr       
       Response.Write  "    Parent.DbDtlQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim itxtSpread
    Dim arrRowVal
    Dim arrColVal
    Dim lgErrorPos
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgErrorPos        = ""                                                           '☜: Set to space

    itxtSpread = Trim(Request("txtSpread"))
    
    If itxtSpread = "" Then
       Exit Sub
    End If   
    
	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data
	
    For iDx = 0 To UBound(arrRowVal,1) - 1
        arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C" :  Call SubBizSaveMultiCreate(arrColVal)                        '☜: Create
            Case "U" :  Call SubBizSaveMultiUpdate(arrColVal)                        '☜: Update
            Case "D" :  Call SubBizSaveMultiDelete(arrColVal)                        '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If

    Next
    
    If lgErrorStatus = "YES" Then
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " Parent.SubSetErrPos(Trim(""" & lgErrorPos & """))" & vbCr
       Response.Write  " </Script>                  " & vbCr
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  " Parent.DBSaveOk            " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If
    
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "INSERT INTO student("
    lgStrSQL = lgStrSQL & " SchoolCD     , StudentID    ,"    '3
    lgStrSQL = lgStrSQL & " StudentNM    , Grade        ,"    '5
    lgStrSQL = lgStrSQL & " Phone        , ZipCd        ,"    '7
    lgStrSQL = lgStrSQL & " StudyOnOff   , EnrollDT     ,"    '9
    lgStrSQL = lgStrSQL & " GraduatedDT  , SMoney       ,"    '11
    lgStrSQL = lgStrSQL & " SMoneyCnt    , INSRT_UID    ,"    '13
    lgStrSQL = lgStrSQL & " INSRT_DT     , UPDT_UID     ,"    '15
    lgStrSQL = lgStrSQL & " UPDT_DT      )"    '16
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(02)),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(03)),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(04) ,"","S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(05) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(06) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(07) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(08) ,"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(09)),"","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(10)),"","S") & ","
    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(11),0)         & ","
    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(12),0)         & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate()," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
    lgStrSQL = lgStrSQL & " GetDate())" 
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    Dim lgStrSQL
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE STUDENT SET "
    lgStrSQL = lgStrSQL & " StudentNM   = " & FilterVar(            arrColVal(04) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " Grade       = " & FilterVar(            arrColVal(05) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " Phone       = " & FilterVar(            arrColVal(06) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " ZipCd       = " & FilterVar(            arrColVal(07) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " StudyOnOff  = " & FilterVar(            arrColVal(08) ,Null,"S")  & ","
    lgStrSQL = lgStrSQL & " EnrollDT    = " & FilterVar(UniConvDate(arrColVal(09)),Null,"S")  & ","
    lgStrSQL = lgStrSQL & " GraduatedDT = " & FilterVar(UniConvDate(arrColVal(10)),Null,"S")  & ","
    lgStrSQL = lgStrSQL & " SMoney      = " &            UNIConvNum(arrColVal(11),0)          & ","
    lgStrSQL = lgStrSQL & " SMoneyCnt   = " &            UNIConvNum(arrColVal(12),0)          & ","          
    lgStrSQL = lgStrSQL & " UPDT_UID    = " & FilterVar(gUsrId,"","S")                        & ","             
    lgStrSQL = lgStrSQL & " UPDT_DT     = GetDate() " 
    lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    Dim lgStrSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "DELETE  FROM STUDENT"
    lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
    
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>


