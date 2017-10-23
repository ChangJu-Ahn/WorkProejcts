
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

'Call ServerMesgBox("HANC : " & cstr(lgStrPrevKey) , vbInformation, I_MKSCRIPT)

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
    
    Const C_SHEETMAXROWS_D  = 1000                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

		lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " b.OUT_NO " & vbcrlf
		lgStrSQL = lgStrSQL & " ,(SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=b.SHIP_TO_PARTY and USAGE_FLAG = 'y') BP_CD " & vbcrlf
		lgStrSQL = lgStrSQL & " ,c.BP_NM " & vbcrlf
		lgStrSQL = lgStrSQL & " ,a.ITEM_CD " & vbcrlf
		lgStrSQL = lgStrSQL & " ,d.ITEM_NM " & vbcrlf
		lgStrSQL = lgStrSQL & " ,d.SPEC " & vbcrlf
		lgStrSQL = lgStrSQL & " ,a.PLANT_CD " & vbcrlf
		lgStrSQL = lgStrSQL & " ,b.OUT_TYPE " & vbcrlf
		lgStrSQL = lgStrSQL & " ,e.UD_MINOR_NM " & vbcrlf
		lgStrSQL = lgStrSQL & " ,a.GOOD_ON_HAND_QTY " & vbcrlf
		lgStrSQL = lgStrSQL & " ,b.GI_QTY " & vbcrlf
		lgStrSQL = lgStrSQL & " ,b.GI_UNIT " & vbcrlf
		lgStrSQL = lgStrSQL & " ,a.LOT_NO " & vbcrlf
		lgStrSQL = lgStrSQL & " ,a.LOT_SUB_NO " & vbcrlf
		lgStrSQL = lgStrSQL & " ,b.ACTUAL_GI_DT,b.CUST_LOT_NO1 cust_lot_no ,a.SL_CD,b.TRANS_TIME,b.CREATE_TYPE,b.cust_lot_no rcpt_lot_no " & vbcrlf
		lgStrSQL = lgStrSQL & " ,b.PGM_NAME as pgm_name " & vbcrlf
		lgStrSQL = lgStrSQL & " , [DBO].[ufn_GetPrice]((SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=b.SHIP_TO_PARTY and USAGE_FLAG = 'y'), a.ITEM_CD, " & FilterVar(Request("txtSoNo"),"''","S") & ", b.GI_UNIT, b.PGM_NAME) pgm_price " & vbcrlf
		lgStrSQL = lgStrSQL & " from i_onhand_stock_detail a (nolock) " & vbcrlf
		lgStrSQL = lgStrSQL & " inner join T_IF_RCV_VIRTURE_OUT_KO441 b (nolock) on (a.PLANT_CD=b.PLANT_CD and a.ITEM_CD=[DBO].[UFN_GETITEMCD](b.MES_ITEM_CD) " & vbcrlf
		lgStrSQL = lgStrSQL & " 										 AND a.TRACKING_NO='*' AND a.LOT_NO=b.LOT_NO AND a.LOT_SUB_NO=0) " & vbcrlf
		lgStrSQL = lgStrSQL & " inner join ( " & vbcrlf
		lgStrSQL = lgStrSQL & " 					select OUT_NO,TRANS_TIME  " & vbcrlf
		lgStrSQL = lgStrSQL & " 					from T_IF_RCV_VIRTURE_OUT_KO441  " & vbcrlf
		lgStrSQL = lgStrSQL & " 					group by OUT_NO,TRANS_TIME having count(*) <> 2 " & vbcrlf
		lgStrSQL = lgStrSQL & " 					) b2 on (b.OUT_NO=b2.OUT_NO and b.TRANS_TIME=b2.TRANS_TIME) " & vbcrlf		
		lgStrSQL = lgStrSQL & " inner join B_BIZ_PARTNER c (nolock) on ((SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=b.SHIP_TO_PARTY and USAGE_FLAG = 'y')=c.BP_CD) " & vbcrlf
		lgStrSQL = lgStrSQL & " inner join B_ITEM d (nolock) on (a.ITEM_CD=d.ITEM_CD) " & vbcrlf
		lgStrSQL = lgStrSQL & " inner join B_USER_DEFINED_MINOR e (nolock)  on (e.UD_MAJOR_CD='ZZ002' and b.OUT_TYPE=e.UD_MINOR_CD) " & vbcrlf		
		lgStrSQL = lgStrSQL & " where a.PLANT_CD=" & FilterVar(Request("txtPlantCd"),"''","S") & vbcrlf
		lgStrSQL = lgStrSQL & " AND a.ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		
		If Trim(Request("txtShipToParty")) <> "" Then
			lgStrSQL = lgStrSQL & " AND (SELECT top 1 BP_CD FROM B_BIZ_PARTNER WHERE BP_ALIAS_NM=b.SHIP_TO_PARTY and USAGE_FLAG = 'y')=" & FilterVar(Request("txtShipToParty"),"''","S") & vbcrlf
		End If
						
		If Ucase(Trim(Request("txtPlantCd"))) = "P01" Then
				lgStrSQL = lgStrSQL & " AND a.SL_CD='011001'" & vbcrlf
		ElseIf Ucase(Trim(Request("txtPlantCd"))) = "P02" Then
				lgStrSQL = lgStrSQL & " AND a.SL_CD='021001'" & vbcrlf
		Else 
				'T_IF_RCV_PART_OUT_KO441.SHIP_TO_PARTY_LINE 컬럼있음. 조인 테이블을 T_IF_RCV_VIRTURE_OUT_KO441 변경 후 관련컬럼이 없어 주석처리함.
				'lgStrSQL = lgStrSQL & " AND a.SL_CD in (select UD_REFERENCE from B_USER_DEFINED_MINOR where ud_major_cd='zz005' and UD_MINOR_CD=b.SHIP_TO_PARTY_LINE) " & vbcrlf	
		End If

		lgStrSQL = lgStrSQL & " AND a.GOOD_ON_HAND_QTY>0 " & vbcrlf


		lgStrSQL = lgStrSQL & " AND e.UD_REFERENCE = 'Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " AND ISNULL(b.ERP_APPLY_FLAG,'N') <> 'Y' " & vbcrlf              


    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
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
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("OUT_TYPE"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UD_MINOR_NM"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("GOOD_ON_HAND_QTY"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("GI_QTY"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("GI_UNIT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_SUB_NO"))
          lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("ACTUAL_GI_DT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("rcpt_lot_no"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CUST_LOT_NO"))          
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("pgm_name"))          '2008-06-16 5:58오후 :: hanc
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("pgm_price"), ggQty.DecPoint, 0)         '2008-06-16 5:58오후 :: hanc
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRANS_TIME"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CREATE_TYPE"))          
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


