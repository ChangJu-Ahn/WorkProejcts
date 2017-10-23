
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
    lgStrPrevKey   = UNICInt(Trim(Request("lgStrPrevKey")), 0)                   '☜: Next Key

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

Call ServerMesgBox("HANC : 100 "  , vbInformation, I_MKSCRIPT)
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

'    '---------- Developer Coding part (Start) ---------------------------------------------------------------
'    If lgStrPrevKey = 0 Then
'
'       lgStrSQL = "Select plant_cd,plant_nm " 
'       lgStrSQL = lgStrSQL & " From B_Plant (Nolock) " 
'       lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(lgKeyStream(0),"''", "S")
'
'       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
'          Response.Write  " <Script Language=vbscript>            " & vbCr
'          Response.Write  "   Parent.Frm1.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
'          Response.Write  "   Parent.Frm1.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
'          Response.Write  "   Parent.Frm1.htxtPlantCd.Value = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set next key data
'          Response.Write  " </Script> " & vbCr
'          Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
'          Call SubBizQueryMulti()
'       Else
'          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
'       End If
'    Else
       Call SubBizQueryMulti()
'    End If 
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

    lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " BBP.BP_CD,                                      "
    lgStrSQL = lgStrSQL & "                   BBP.BP_NM,                                                 "
    lgStrSQL = lgStrSQL & "                   BI.ITEM_CD,                                                "
    lgStrSQL = lgStrSQL & "                   BI.ITEM_NM,                                                "
    lgStrSQL = lgStrSQL & "                   BI.SPEC,                                                   "
    lgStrSQL = lgStrSQL & "                   SBIP.DEAL_TYPE,                                            "
    lgStrSQL = lgStrSQL & "                   SBIP.PAY_METH,                                             "
    lgStrSQL = lgStrSQL & "                   SBIP.SALES_UNIT,                                           "
    lgStrSQL = lgStrSQL & "                   SBIP.CURRENCY,                                             "
    lgStrSQL = lgStrSQL & "                   CONVERT(CHAR(20),SBIP.VALID_FROM_DT,20) valid_from_dt,     "
    lgStrSQL = lgStrSQL & "                   SBIP.ITEM_PRICE,                                           "
    lgStrSQL = lgStrSQL & "                   SBIP.EXT1_QTY,                                             "
    lgStrSQL = lgStrSQL & "                   SBIP.EXT2_QTY,                                             "
    lgStrSQL = lgStrSQL & "                   SBIP.EXT1_AMT,                                             "
    lgStrSQL = lgStrSQL & "                   SBIP.EXT2_AMT,                                             "
    lgStrSQL = lgStrSQL & "                   SBIP.EXT1_CD,                                              "
    lgStrSQL = lgStrSQL & "                   SBIP.EXT2_CD,                                              "
    lgStrSQL = lgStrSQL & "                   DBO.UFN_GETCODENAME('S0001',SBIP.DEAL_TYPE) deal_type_nm,               "
    lgStrSQL = lgStrSQL & "                   DBO.UFN_GETCODENAME('B9004',SBIP.PAY_METH) pay_meth_nm,                "
    lgStrSQL = lgStrSQL & "                   SBIP.PRC_FLG,                                              "
    lgStrSQL = lgStrSQL & "                   SBIP.REMRK                                                 "
    lgStrSQL = lgStrSQL & " FROM     B_BIZ_PARTNER BBP,                                                  "
    lgStrSQL = lgStrSQL & "          B_ITEM BI,                                                          "
    lgStrSQL = lgStrSQL & "          S_BP_ITEM_PRICE_KO441 SBIP                                          "
    lgStrSQL = lgStrSQL & " WHERE    SBIP.BP_CD = BBP.BP_CD                                              "
    lgStrSQL = lgStrSQL & "          AND SBIP.ITEM_CD = BI.ITEM_CD                                       "
    lgStrSQL = lgStrSQL & "          AND BBP.BP_CD >= 'D000000206'                                       "
    lgStrSQL = lgStrSQL & "          AND BBP.BP_CD <= 'D000000206'                                       "
    lgStrSQL = lgStrSQL & "          AND BI.ITEM_CD >= ''                                                "
    lgStrSQL = lgStrSQL & "          AND BI.ITEM_CD <= 'zzzzzzzzzzzzzzzzzz'                              "
    lgStrSQL = lgStrSQL & "          AND SBIP.CURRENCY >= ''                                             "
    lgStrSQL = lgStrSQL & "          AND SBIP.CURRENCY <= 'zzz'                                          "
    lgStrSQL = lgStrSQL & "          AND SBIP.DEAL_TYPE >= ''                                            "
    lgStrSQL = lgStrSQL & "          AND SBIP.DEAL_TYPE <= 'zzzzz'                                       "
    lgStrSQL = lgStrSQL & "          AND SBIP.PAY_METH >= ''                                             "
    lgStrSQL = lgStrSQL & "          AND SBIP.PAY_METH <= 'zzzzz'                                        "
    lgStrSQL = lgStrSQL & "          AND SBIP.SALES_UNIT >= ''                                           "
    lgStrSQL = lgStrSQL & "          AND SBIP.SALES_UNIT <= 'zzz'                                        "
    lgStrSQL = lgStrSQL & "          AND SBIP.VALID_FROM_DT >= '1900-01-01'                              "
    lgStrSQL = lgStrSQL & "          AND ((BBP.BP_CD > '')                                               "
    lgStrSQL = lgStrSQL & "                OR (BBP.BP_CD = ''                                            "
    lgStrSQL = lgStrSQL & "                    AND BI.ITEM_CD > '')                                      "
    lgStrSQL = lgStrSQL & "                OR (BBP.BP_CD = ''                                            "
    lgStrSQL = lgStrSQL & "                    AND BI.ITEM_CD = ''                                       "
    lgStrSQL = lgStrSQL & "                    AND SBIP.DEAL_TYPE > '')                                  "
    lgStrSQL = lgStrSQL & "                OR (BBP.BP_CD = ''                                            "
    lgStrSQL = lgStrSQL & "                    AND BI.ITEM_CD = ''                                       "
    lgStrSQL = lgStrSQL & "                    AND SBIP.DEAL_TYPE = ''                                   "
    lgStrSQL = lgStrSQL & "                    AND SBIP.PAY_METH > '')                                   "
    lgStrSQL = lgStrSQL & "                OR (BBP.BP_CD = ''                                            "
    lgStrSQL = lgStrSQL & "                    AND BI.ITEM_CD = ''                                       "
    lgStrSQL = lgStrSQL & "                    AND SBIP.DEAL_TYPE = ''                                   "
    lgStrSQL = lgStrSQL & "                    AND SBIP.PAY_METH = ''                                    "
    lgStrSQL = lgStrSQL & "                    AND SBIP.VALID_FROM_DT > '1900-01-01')                    "
    lgStrSQL = lgStrSQL & "                OR (BBP.BP_CD = ''                                            "
    lgStrSQL = lgStrSQL & "                    AND BI.ITEM_CD = ''                                       "
    lgStrSQL = lgStrSQL & "                    AND SBIP.DEAL_TYPE = ''                                   "
    lgStrSQL = lgStrSQL & "                    AND SBIP.PAY_METH = ''                                    "
    lgStrSQL = lgStrSQL & "                    AND SBIP.VALID_FROM_DT = '1900-01-01'                     "
    lgStrSQL = lgStrSQL & "                    AND SBIP.SALES_UNIT > '')                                 "
    lgStrSQL = lgStrSQL & "                OR (BBP.BP_CD = ''                                            "
    lgStrSQL = lgStrSQL & "                    AND BI.ITEM_CD = ''                                       "
    lgStrSQL = lgStrSQL & "                    AND SBIP.DEAL_TYPE = ''                                   "
    lgStrSQL = lgStrSQL & "                    AND SBIP.PAY_METH = ''                                    "
    lgStrSQL = lgStrSQL & "                    AND SBIP.VALID_FROM_DT = '1900-01-01'                     "
    lgStrSQL = lgStrSQL & "                    AND SBIP.SALES_UNIT = ''                                  "
    lgStrSQL = lgStrSQL & "                    AND SBIP.CURRENCY >= ''))                                 "
    lgStrSQL = lgStrSQL & " ORDER BY BBP.BP_CD ASC,                                                      "
    lgStrSQL = lgStrSQL & "          BI.ITEM_CD ASC,                                                     "
    lgStrSQL = lgStrSQL & "          SBIP.DEAL_TYPE ASC,                                                 "
    lgStrSQL = lgStrSQL & "          SBIP.PAY_METH ASC,                                                  "
    lgStrSQL = lgStrSQL & "          SBIP.VALID_FROM_DT ASC,                                             "
    lgStrSQL = lgStrSQL & "          SBIP.SALES_UNIT ASC,                                                "
    lgStrSQL = lgStrSQL & "          SBIP.CURRENCY ASC                                                   "

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
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bp_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bp_nm"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm" ))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC" ))
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("deal_type"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("deal_type_nm"))				
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_meth"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pay_meth_nm"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("valid_from_dt" ))				
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sales_unit"))
            lgstrData = lgstrData & Chr(11) & ""	
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("currency"))
            lgstrData = lgstrData & Chr(11) & ""        
            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("item_price"), ggUnitCost.DecPoint, 0 )				
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRC_FLG"))
            if ConvSPChars(lgObjRs("PRC_FLG"))="T" then
            	lgstrData = lgstrData & Chr(11) & ConvSPChars("진단가")
            Else
            	lgstrData = lgstrData & Chr(11) & ConvSPChars("가단가")		
            End if
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMRK"))
            lgstrData = lgstrData & Chr(11) & iLngMaxRow + iLngRow
            lgstrData = lgstrData & Chr(11) & Chr(12)
            

          lgObjRs.MoveNext

          iDx =  iDx + 1
         If iDx > C_SHEETMAXROWS_D Then
			 lgStrPrevKey = lgStrPrevKey + 1
             Exit Do
         End If        
      Loop 
    End If
Call ServerMesgBox("정상적으로 조회되었습니다. "  , vbInformation, I_MKSCRIPT)

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
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData       " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
'Call ServerMesgBox("SubBizSaveMulti : 100"  , vbInformation, I_MKSCRIPT)

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
'Call ServerMesgBox("SubBizSaveMultiCreate : 200"  , vbInformation, I_MKSCRIPT)
    Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = " INSERT INTO S_BP_ITEM_PRICE_KO441                                          "
    lgStrSQL = lgStrSQL & "           (BP_CD,                                               "                
    lgStrSQL = lgStrSQL & "            ITEM_CD,                                             "
    lgStrSQL = lgStrSQL & "            DEAL_TYPE,                                           "
    lgStrSQL = lgStrSQL & "            PAY_METH,                                            "
    lgStrSQL = lgStrSQL & "            SALES_UNIT,                                          "
    lgStrSQL = lgStrSQL & "            CURRENCY,                                            "
    lgStrSQL = lgStrSQL & "            VALID_FROM_DT,                                       "
    lgStrSQL = lgStrSQL & "            ITEM_PRICE,                                          "
    lgStrSQL = lgStrSQL & "            DEPOSIT_PRICE,                                       "
    lgStrSQL = lgStrSQL & "            INSRT_USER_ID,                                       "
    lgStrSQL = lgStrSQL & "            INSRT_DT,                                            "
    lgStrSQL = lgStrSQL & "            UPDT_USER_ID,                                        "
    lgStrSQL = lgStrSQL & "            UPDT_DT,                                             "
    lgStrSQL = lgStrSQL & "            EXT1_QTY,                                            "
    lgStrSQL = lgStrSQL & "            EXT2_QTY,                                            "
    lgStrSQL = lgStrSQL & "            EXT1_AMT,                                            "
    lgStrSQL = lgStrSQL & "            EXT2_AMT,                                            "
    lgStrSQL = lgStrSQL & "            EXT1_CD,                                             "
    lgStrSQL = lgStrSQL & "            EXT2_CD,                                             "
    lgStrSQL = lgStrSQL & "            PRC_FLG,                                             "
    lgStrSQL = lgStrSQL & "            REMRK)                                               "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & "            '999999',                                            "
    lgStrSQL = lgStrSQL & "            '1250-10001',                                        "
    lgStrSQL = lgStrSQL & "            'BACK_GRIND',                                        "
    lgStrSQL = lgStrSQL & "            'CH',                                                "
    lgStrSQL = lgStrSQL & "            'EA',                                                "
    lgStrSQL = lgStrSQL & "            'ARS',                                               "
    lgStrSQL = lgStrSQL & "            CONVERT(CHAR(23),'2008-06-18',20),                   "
    lgStrSQL = lgStrSQL & "            350.0000,                                            "
    lgStrSQL = lgStrSQL & "            0,                                                   "
    lgStrSQL = lgStrSQL & "            'unierp',                                            "
    lgStrSQL = lgStrSQL & "            GETDATE(),                                           "
    lgStrSQL = lgStrSQL & "            'unierp',                                            "
    lgStrSQL = lgStrSQL & "            GETDATE(),                                           "
    lgStrSQL = lgStrSQL & "            0,                                                   "
    lgStrSQL = lgStrSQL & "            0,                                                   "
    lgStrSQL = lgStrSQL & "            0,                                                   "
    lgStrSQL = lgStrSQL & "            0,                                                   "
    lgStrSQL = lgStrSQL & "            '',                                                  "
    lgStrSQL = lgStrSQL & "            '',                                                  "
    lgStrSQL = lgStrSQL & "            'T',                                                 "
    lgStrSQL = lgStrSQL & "            '')                                                  "



'    lgStrSQL = "INSERT INTO student("
'    lgStrSQL = lgStrSQL & " SchoolCD     , StudentID    ,"    '3
'    lgStrSQL = lgStrSQL & " StudentNM    , Grade        ,"    '5
'    lgStrSQL = lgStrSQL & " Phone        , ZipCd        ,"    '7
'    lgStrSQL = lgStrSQL & " StudyOnOff   , EnrollDT     ,"    '9
'    lgStrSQL = lgStrSQL & " GraduatedDT  , SMoney       ,"    '11
'    lgStrSQL = lgStrSQL & " SMoneyCnt    , INSRT_UID    ,"    '13
'    lgStrSQL = lgStrSQL & " INSRT_DT     , UPDT_UID     ,"    '15
'    lgStrSQL = lgStrSQL & " UPDT_DT      )"    '16
'    lgStrSQL = lgStrSQL & " VALUES(" 
'    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(02)),"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(03)),"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(04) ,"","S") & "," 
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(05) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(06) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(07) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(            arrColVal(08) ,"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(09)),"","S") & ","
'    lgStrSQL = lgStrSQL & FilterVar(UniConvDate(arrColVal(10)),"","S") & ","
'    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(11),0)         & ","
'    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(12),0)         & ","
'    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
'    lgStrSQL = lgStrSQL & " GetDate()," 
'    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                       & "," 
'    lgStrSQL = lgStrSQL & " GetDate())" 
    
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
    lgStrSQL = lgStrSQL & " UPDT_DT     = GetDate() "     lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
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


