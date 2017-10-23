
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

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

'Call ServerMesgBox("33333 HANC : 10"  , vbInformation, I_MKSCRIPT)
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
'Call ServerMesgBox("mb3 - close : 10"  , vbInformation, I_MKSCRIPT)
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
          Response.Write  " <Script Language=vbscript>            " & vbCr
          Response.Write  "   Parent.Frm1.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
          Response.Write  "   Parent.Frm1.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
          Response.Write  "   Parent.Frm1.htxtPlantCd.Value = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set next key data
          Response.Write  " </Script> " & vbCr
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
        
    lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " d.wc_cd,d.wc_nm,a.item_cd,e.item_nm,e.spec,b.lot_no,b.lot_sub_no,b.good_on_hand_qty "
	lgStrSQL = lgStrSQL & " from b_item_by_plant a, "
	lgStrSQL = lgStrSQL & "         i_onhand_stock_detail b, "
	lgStrSQL = lgStrSQL & "         i_goods_movement_detail c, "
	lgStrSQL = lgStrSQL & "         p_work_center d, "
	lgStrSQL = lgStrSQL & "         b_item e	 "
	lgStrSQL = lgStrSQL & " where a.plant_cd =  " &  FilterVar(lgKeyStream(0),"''", "S")
	lgStrSQL = lgStrSQL & " and a.plant_cd = b.plant_cd  "
	lgStrSQL = lgStrSQL & " and a.procur_type = 'M'  "
	lgStrSQL = lgStrSQL & " and a.lot_flg = 'Y' "
	lgStrSQL = lgStrSQL & " and b.lot_no <> '*'  "
	lgStrSQL = lgStrSQL & " and b.good_on_hand_qty > 0 "
	lgStrSQL = lgStrSQL & " and a.item_cd = b.item_cd "
	lgStrSQL = lgStrSQL & " and a.plant_cd = c.plant_cd "
	lgStrSQL = lgStrSQL & " and a.item_cd = c.item_cd "
'	lgStrSQL = lgStrSQL & " and c.trns_type = 'MR' "
'	lgStrSQL = lgStrSQL & " and a.plant_cd = d.plant_cd "
'	lgStrSQL = lgStrSQL & " and c.wc_cd = d.wc_cd "
	lgStrSQL = lgStrSQL & " and b.plant_cd = c.plant_cd  "
	lgStrSQL = lgStrSQL & " and b.item_cd = c.item_cd  "
'	lgStrSQL = lgStrSQL & " and b.lot_no = c.lot_no " 
'	lgStrSQL = lgStrSQL & " and b.lot_sub_no = c.lot_sub_no "
'	lgStrSQL = lgStrSQL & " and a.item_cd = e.item_cd "
	lgStrSQL = lgStrSQL & " order by d.wc_cd,b.lot_sub_no "
        
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
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_no"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_sub_no"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("good_on_hand_qty"), ggQty.DecPoint, 0)
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
'Call ServerMesgBox("33333 SubBizSaveMulti : 700-1"  , vbInformation, I_MKSCRIPT)
    Dim itxtSpread
    Dim arrRowVal
    Dim arrColVal
    Dim lgErrorPos
    Dim iDx, seq_cnt
'Call ServerMesgBox("33333 SubBizSaveMulti : 700-2"  , vbInformation, I_MKSCRIPT)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgErrorPos        = ""                                                           '☜: Set to space
'Call ServerMesgBox("33333 SubBizSaveMulti : 700-3"  , vbInformation, I_MKSCRIPT)
    itxtSpread = Trim(Request("txtSpread"))
    
    If itxtSpread = "" Then
       Exit Sub
    End If   
 'Call ServerMesgBox("33333 SubBizSaveMulti : 700-4"  , vbInformation, I_MKSCRIPT)   
	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data
	
'Call ServerMesgBox("33333 SubBizSaveMulti : UBound(arrRowVal,1) : " & UBound(arrRowVal,1) , vbInformation, I_MKSCRIPT)

 '20080305::hanc::begin++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 ' 아래 select 문을 쓴 이유 : item_seq를 자동적으로 채번하기 위해서 아래와 같이 하였다.
   lgStrSQL = "Select (ISNULL(MAX(ITEM_SEQ), 0) +1) as seq_cnt " 
   lgStrSQL = lgStrSQL & " FROM M_ISSUE_REQ_DTL_KO441 (Nolock) " 
   lgStrSQL = lgStrSQL & " WHERE plant_cd  = " & FilterVar(Trim(Request("txtPlantCd")) ,"","S")
   lgStrSQL = lgStrSQL & " AND ISSUE_REQ_NO  = " & FilterVar(Trim(Request("txtPoNo1")) ,"","S")

   If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
      seq_cnt = lgObjRs("seq_cnt")
   end if
 '20080305::hanc::end++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


    For iDx = 0 To UBound(arrRowVal,1) 
        arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data
'Call ServerMesgBox("arrColVal(0) : " & arrColVal(0)  , vbInformation, I_MKSCRIPT)        
        Select Case arrColVal(0)
            Case "C" :  Call SubBizSaveMultiCreate(arrColVal, iDx, seq_cnt)          '☜: Create   20080305::hanc :: 파라미터를 2개 더 사용하였다. 이유는 item_seq를 자동적으로 +1 하기 위해 ...
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
Sub SubBizSaveMultiCreate(arrColVal, iDx, seq_cnt)
'Call ServerMesgBox("33333 SubBizSaveMultiCreate : 800"  , vbInformation, I_MKSCRIPT)
    Dim lgStrSQL
    Dim seq_cnt_1
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

   seq_cnt_1 = Cint(lgObjRs("seq_cnt"))   + Cint(iDx)       '20080305::hanc :: item_seq를 자동으로 +1 하기 위함.
   
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "INSERT INTO M_ISSUE_REQ_DTL_KO441          "
    lgStrSQL = lgStrSQL & " (                              "
    lgStrSQL = lgStrSQL & " 	PLANT_CD,                  "
    lgStrSQL = lgStrSQL & " 	ISSUE_REQ_NO,              "
    lgStrSQL = lgStrSQL & " 	ITEM_SEQ,                  "
    lgStrSQL = lgStrSQL & " 	PRODT_ORDER_NO,            "
    lgStrSQL = lgStrSQL & " 	ITEM_CD,                   "
    lgStrSQL = lgStrSQL & " 	REQ_QTY,                   "
    lgStrSQL = lgStrSQL & " 	ISSUE_QTY,                 "
    lgStrSQL = lgStrSQL & " 	REMARK,                    "
    lgStrSQL = lgStrSQL & " 	LIMIT_DT,                    "
    lgStrSQL = lgStrSQL & " 	INSRT_USER_ID,             "
    lgStrSQL = lgStrSQL & " 	INSRT_DT,                  "
    lgStrSQL = lgStrSQL & " 	UPDT_USER_ID,              "
    lgStrSQL = lgStrSQL & " 	UPDT_DT                    "
    lgStrSQL = lgStrSQL & " )                              "
    lgStrSQL = lgStrSQL & " VALUES                         "
    lgStrSQL = lgStrSQL & " (                              "
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtPlantCd")) ,"","S")    & ", "
    lgStrSQL = lgStrSQL & FilterVar(Trim(Request("txtPoNo1")) ,"","S")      & ", "
    lgStrSQL = lgStrSQL & FilterVar(      UCase(seq_cnt_1),0,"D")      & ","
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(03)),"","S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(04)),"","S")      & ","
    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(05),0)            & ","
    lgStrSQL = lgStrSQL &           UNIConvNum (arrColVal(06),0)            & ","
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(07)),"","S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(      UCase(arrColVal(08)),"","S")      & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                          & "," 
    lgStrSQL = lgStrSQL & " GetDate(),                                      " 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"","S")                          & "," 
    lgStrSQL = lgStrSQL & " GetDate()                                      " 
    lgStrSQL = lgStrSQL & " )                                               "
    
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

    lgStrSQL = "UPDATE  M_ISSUE_REQ_DTL_KO441 "
    lgStrSQL = lgStrSQL & "SET     PRODT_ORDER_NO    =  " & FilterVar(      UCase(arrColVal(03)),"","S")    & ", "
    lgStrSQL = lgStrSQL & "        ITEM_CD           =  " & FilterVar(      UCase(arrColVal(04)),"","S")    & ", "
    lgStrSQL = lgStrSQL & "        REQ_QTY           =  " & FilterVar(      UCase(arrColVal(05)),"","S")    & ", "
    lgStrSQL = lgStrSQL & "        ISSUE_QTY         =  " & FilterVar(      UCase(arrColVal(06)),"","S")    & ", "
    lgStrSQL = lgStrSQL & "        REMARK            =  " & FilterVar(      UCase(arrColVal(07)),"","S")    & ", "
    lgStrSQL = lgStrSQL & "        LIMIT_DT          =  " & FilterVar(      UCase(arrColVal(08)),"","S")    & ", "
    lgStrSQL = lgStrSQL & "        INSRT_USER_ID     =  " & FilterVar(gUsrId,"","S")                        & "," 
    lgStrSQL = lgStrSQL & "        INSRT_DT          =   GetDate(),                                        " 
    lgStrSQL = lgStrSQL & "        UPDT_USER_ID      =  " & FilterVar(gUsrId,"","S")                        & "," 
    lgStrSQL = lgStrSQL & "        UPDT_DT           =   GetDate()                                        " 
    lgStrSQL = lgStrSQL & "WHERE   PLANT_CD          =  " & FilterVar(Trim(Request("txtPlantCd")) ,"","S")  & " "
    lgStrSQL = lgStrSQL & "AND     ISSUE_REQ_NO      =	" & FilterVar(Trim(Request("txtPoNo1")) ,"","S")    & " "
    lgStrSQL = lgStrSQL & "AND     ITEM_SEQ          =  " & FilterVar(      UCase(arrColVal(02)),"","S")    & " "
           

'    lgStrSQL = "UPDATE STUDENT SET "
'    lgStrSQL = lgStrSQL & " StudentNM   = " & FilterVar(            arrColVal(04) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " Grade       = " & FilterVar(            arrColVal(05) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " Phone       = " & FilterVar(            arrColVal(06) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " ZipCd       = " & FilterVar(            arrColVal(07) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " StudyOnOff  = " & FilterVar(            arrColVal(08) ,Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " EnrollDT    = " & FilterVar(UniConvDate(arrColVal(09)),Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " GraduatedDT = " & FilterVar(UniConvDate(arrColVal(10)),Null,"S")  & ","
'    lgStrSQL = lgStrSQL & " SMoney      = " &            UNIConvNum(arrColVal(11),0)          & ","
'    lgStrSQL = lgStrSQL & " SMoneyCnt   = " &            UNIConvNum(arrColVal(12),0)          & ","          
'    lgStrSQL = lgStrSQL & " UPDT_UID    = " & FilterVar(gUsrId,"","S")                        & ","             
'    lgStrSQL = lgStrSQL & " UPDT_DT     = GetDate() " 
'    lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
'    lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")
    
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

'Call ServerMesgBox("PLANT_CD : " & FilterVar(Trim(Request("txtPlantCd")) ,"","S") , vbInformation, I_MKSCRIPT)
'Call ServerMesgBox("ISSUE_REQ_NO : " & FilterVar(Trim(Request("txtPoNo1")) ,"","S") ,"","S")  , vbInformation, I_MKSCRIPT)
'Call ServerMesgBox("ITEM_SEQ : " & FilterVar(      UCase(arrColVal(02)),"","S") ,"","S")  , vbInformation, I_MKSCRIPT)


    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = " DELETE  FROM M_ISSUE_REQ_DTL_KO441 "
    lgStrSQL = lgStrSQL & " WHERE   PLANT_CD          = " & FilterVar(Trim(Request("txtPlantCd")) ,""  ,"S")  & " "
    lgStrSQL = lgStrSQL & " AND     ISSUE_REQ_NO      =	" & FilterVar(Trim(Request("txtPoNo1"))   ,""  ,"S")  & " "
    lgStrSQL = lgStrSQL & " AND     ITEM_SEQ          = " & FilterVar(UCase(arrColVal(02))        ,""  ,"S")  & " "
'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)
    'lgStrSQL = "DELETE  FROM STUDENT"
    'lgStrSQL = lgStrSQL & " WHERE SchoolCD  = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    'lgStrSQL = lgStrSQL & " AND   StudentID = " &  FilterVar(Trim(UCase(arrColVal(3))),"''","S")

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


