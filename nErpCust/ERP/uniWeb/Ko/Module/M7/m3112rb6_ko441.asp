
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
    Call SubBizQueryMulti() 
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
    Dim dBillDt, dBillDtFirst
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

    dBillDt     = Request("txtBillFrDt")  '20071226::hanc
    dBillDtFirst     = left(dBillDt,8) & "01"  '20071226::hanc
	
    '---------- Developer Coding part (Start) ---------------------------------------------------------------

lgStrSQL = "  SELECT         BP_NM,                                                                                                                                         "
lgStrSQL = lgStrSQL & "          		(SELECT IO_TYPE_NM FROM M_MVMT_TYPE WHERE IO_TYPE_CD = NEPES.IO_TYPE_CD ) IO_TYPE_NM,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		IO_TYPE_CD,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		BP_CD,                                                                                                                              "
lgStrSQL = lgStrSQL & "          		ISNULL(PO_NO, '') PO_NO,                                                                                                                              "
lgStrSQL = lgStrSQL & "          		[DBO].[ufn_GetPOSEQFromPARTIN](ISNULL(PO_NO, ''), ITEM_CD) PO_SEQ_NO,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		PLANT_CD,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		SL_CD,  " '--입출고창고(MVMT_SL_CD), 입고창고(MVMT_RCPT_SL_CD) 어떤것인지 확인 필요                                                    "
lgStrSQL = lgStrSQL & "          		ITEM_CD,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		ITEM_NM,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		SPEC,                                                                                                                               "
lgStrSQL = lgStrSQL & "          		TRACKING_NO,                                                                                                                        "
lgStrSQL = lgStrSQL & "          		PO_QTY,	" '--발주수량 : 좀더 확인해 보도록 한다.                                                                                       "
lgStrSQL = lgStrSQL & "          		PO_UNIT,			     " '--발주단위                                                                                                 "
lgStrSQL = lgStrSQL & "          		PO_PRC,		                                                                                                                        "
lgStrSQL = lgStrSQL & "          		PO_DOC_AMT,	" '--발주금액 : 좀더 확인해 보도록 한다.                                                                                   "
lgStrSQL = lgStrSQL & "          		PO_CUR,			     " '--발주화폐단위                                                                                                 "
lgStrSQL = lgStrSQL & "          		DLVY_DT,	" '--입출고일(MVMT_DT), 입고일(MVMT_RCPT_DT) 어떤것인지 확인 필요                                                          "
lgStrSQL = lgStrSQL & "          		RCPT_QTY,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		LC_QTY,                                                                                                                             "
lgStrSQL = lgStrSQL & "          		PRE_IV_QTY,		  " '--참조매입수량                                                                                                    "
lgStrSQL = lgStrSQL & "          		INSPECT_QTY,		      " '--검사수량                                                                                                "
lgStrSQL = lgStrSQL & "          		IV_QTY,                                                                                                                             "
lgStrSQL = lgStrSQL & "          		RECV_INSPEC_FLG,                                                                                                                    "
lgStrSQL = lgStrSQL & "          		MINOR_NM,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		INSPECT_METHOD,                                                                                                                     "
lgStrSQL = lgStrSQL & "          		PLANT_NM,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		SL_NM,                                                                                                                              "
lgStrSQL = lgStrSQL & "                 PUR_GRP,			  " '--구매그룹                                                                                                    "
lgStrSQL = lgStrSQL & "                 LC_RCPT_QTY,                                                                                                                        "
lgStrSQL = lgStrSQL & "                 LOT_FLG,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		LOT_GEN_MTHD,                                                                                                                       "
lgStrSQL = lgStrSQL & "          		FLAG,                                                                                                                       "
lgStrSQL = lgStrSQL & "          		MVMT_NO,                                                                                                                       "
lgStrSQL = lgStrSQL & "          		TRANS_TIME ,                                                                                                                       "
lgStrSQL = lgStrSQL & "          		MAIN_LOT   ,                                                                                                                       "
lgStrSQL = lgStrSQL & "          		IMPORT_TIME,                                                                                                                       "
lgStrSQL = lgStrSQL & "          		CREATE_TYPE,                                                                                                                       "
lgStrSQL = lgStrSQL & "                  1                                                                                                                                  "
lgStrSQL = lgStrSQL & "  FROM                                                                                                                                               "
lgStrSQL = lgStrSQL & "         ( select	B.BP_NM,                                                                                                                        "
lgStrSQL = lgStrSQL & "          		'' IO_TYPE_NM,                                                                                      "
lgStrSQL = lgStrSQL & " 				CASE WHEN ISNULL(A.IO_TYPE_CD,' ') = 'I33' AND A.MVMT_RCPT_QTY < 0 THEN 'RGI'                          "
lgStrSQL = lgStrSQL & " 					 WHEN ISNULL(A.IO_TYPE_CD,' ') = 'I33' AND A.MVMT_RCPT_QTY >= 0 THEN 'RGR'                         "
lgStrSQL = lgStrSQL & " 					 ELSE       		'DGR'                                                                       "
lgStrSQL = lgStrSQL & " 			    END IO_TYPE_CD,                                                                                     "
lgStrSQL = lgStrSQL & "          		B.BP_CD,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		A.PO_NO,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		A.PO_SEQ_NO,                                                                                                                        "
lgStrSQL = lgStrSQL & "          		A.PLANT_CD,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		H.MAJOR_SL_CD SL_CD,  " '--입출고창고(MVMT_SL_CD), 입고창고(MVMT_RCPT_SL_CD) 어떤것인지 확인 필요                                       "
lgStrSQL = lgStrSQL & "          		D.ITEM_CD,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		D.ITEM_NM,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		D.SPEC,                                                                                                                             "
lgStrSQL = lgStrSQL & "          		'*' TRACKING_NO,                                                                                                                    "
lgStrSQL = lgStrSQL & "          		0 PO_QTY,	" '--발주수량 : 좀더 확인해 보도록 한다.                                                                                   "
lgStrSQL = lgStrSQL & "          		A.MVMT_RCPT_UNIT PO_UNIT,			     " '--발주단위                                                                                     "
lgStrSQL = lgStrSQL & "          		0 PO_PRC,		                                                                                                            "
lgStrSQL = lgStrSQL & "          		0 PO_DOC_AMT,	" '--발주금액 : 좀더 확인해 보도록 한다.                                                                               "
lgStrSQL = lgStrSQL & "          		'' PO_CUR,			     " '--발주화폐단위                                                                                     "
lgStrSQL = lgStrSQL & "          		A.MVMT_RCPT_DT DLVY_DT,	" '--입출고일(MVMT_DT), 입고일(MVMT_RCPT_DT) 어떤것인지 확인 필요                                                  "
lgStrSQL = lgStrSQL & "          		ISNULL(MVMT_RCPT_QTY,0) RCPT_QTY,                                                                                                   "
lgStrSQL = lgStrSQL & "          		0 LC_QTY,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		0 PRE_IV_QTY,		  " '--참조매입수량                                                                                                "
lgStrSQL = lgStrSQL & "          		0 INSPECT_QTY,		      " '--검사수량                                                                                                "
lgStrSQL = lgStrSQL & "          		0 IV_QTY,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		[dbo].[ufn_GetRecvInspecFlg](A.PLANT_CD, D.ITEM_CD)  RECV_INSPEC_FLG,                                                                                                        "
lgStrSQL = lgStrSQL & "          		'입고전 검사' MINOR_NM,                                                                                                    "
lgStrSQL = lgStrSQL & "          		'B' INSPECT_METHOD,                                                                                                       "
lgStrSQL = lgStrSQL & "          		F.PLANT_NM,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		'' SL_NM,                                                                                                                            "
lgStrSQL = lgStrSQL & "                  '' PUR_GRP,			  " '--구매그룹                                                                                                "
lgStrSQL = lgStrSQL & "                  0 LC_RCPT_QTY,                                                                                                                     "
lgStrSQL = lgStrSQL & "                  H.LOT_FLG,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		ISNULL((SELECT LOT_GEN_MTHD FROM B_LOT_CONTROL WHERE PLANT_CD = A.PLANT_CD AND ITEM_CD = D.ITEM_CD),'') LOT_GEN_MTHD,               "
lgStrSQL = lgStrSQL & "                  '' MVMT_NO,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  A.TRANS_TIME  ,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  A.MAIN_LOT    ,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  A.IMPORT_TIME ,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  CREATE_TYPE ,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		'Y' FLAG       " '--무상                                                                                                               "
lgStrSQL = lgStrSQL & "          from	T_IF_RCV_PART_IN_KO441	A,                                                                                                          "
lgStrSQL = lgStrSQL & "                  B_BIZ_PARTNER B,                                                                                                                   "
lgStrSQL = lgStrSQL & "          		B_ITEM	D,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		B_PLANT F,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		B_ITEM_BY_PLANT H   ,                                                                                                               "
lgStrSQL = lgStrSQL & "         		M_PUR_ORD_HDR I ,                                                                                                                     "
lgStrSQL = lgStrSQL & "         		(SELECT COUNT(*) CNT, TRANS_TIME, MAIN_LOT, IMPORT_TIME                                                                                                                                            "
lgStrSQL = lgStrSQL & "         		FROM  T_IF_RCV_PART_IN_KO441                                                                                                                                                                       "
lgStrSQL = lgStrSQL & "         		WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlant")), " " , "S") & " "
lgStrSQL = lgStrSQL & "                 AND   CONVERT(CHAR(10), MVMT_RCPT_DT, 120) BETWEEN " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & " AND " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & " "      
lgStrSQL = lgStrSQL & "         		GROUP BY TRANS_TIME, MAIN_LOT, IMPORT_TIME                                                                                                                                                         "
lgStrSQL = lgStrSQL & "         		HAVING COUNT(*) = 1) J                                                                                                                                                                             "
lgStrSQL = lgStrSQL & "          WHERE	A.CUST_CD = B.BP_ALIAS_NM                                                                                                 "
lgStrSQL = lgStrSQL & "          AND		A.MES_ITEM_CD = D.CBM_DESCRIPTION   AND D.ITEM_ACCT <> '10'     AND D.ITEM_ACCT <> '20'                                                                                            "
lgStrSQL = lgStrSQL & "          AND		A.PLANT_CD			=	F.PLANT_CD                                                                                              "
lgStrSQL = lgStrSQL & "          AND		A.PLANT_CD			=	H.PLANT_CD                                                                                              "
lgStrSQL = lgStrSQL & "          AND		D.ITEM_CD = H.ITEM_CD                                                                                             "
lgStrSQL = lgStrSQL & "          AND		A.PO_NO				*=	I.PO_NO                                                                                                 "
lgStrSQL = lgStrSQL & "          AND	A.TRANS_TIME = J.TRANS_TIME   "
lgStrSQL = lgStrSQL & "          AND A.MAIN_LOT  = J.MAIN_LOT         "
lgStrSQL = lgStrSQL & "          AND A.IMPORT_TIME = J.IMPORT_TIME    "
lgStrSQL = lgStrSQL & "          AND        ISNULL(A.PO_NO, 'N') <> 'N'                                                                                                     "
lgStrSQL = lgStrSQL & "          AND        ISNULL(A.ERP_APPLY_FLAG, 'N') <> 'Y'                                                                                                      "
lgStrSQL = lgStrSQL & "          AND   CONVERT(CHAR(10), A.MVMT_RCPT_DT, 120) BETWEEN " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & " AND " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & " "           
lgStrSQL = lgStrSQL & "                                                                                                                                                     "
lgStrSQL = lgStrSQL & "         UNION ALL                                                                                                                                   "
lgStrSQL = lgStrSQL & "                                                                                                                                                     "
lgStrSQL = lgStrSQL & "          select	B.BP_NM,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		'' IO_TYPE_NM,                                                                                      "
lgStrSQL = lgStrSQL & " 				CASE WHEN ISNULL(A.IO_TYPE_CD,' ') = 'I33' AND A.MVMT_RCPT_QTY < 0 THEN 'RGI'                          "
lgStrSQL = lgStrSQL & " 					 WHEN ISNULL(A.IO_TYPE_CD,' ') = 'I33' AND A.MVMT_RCPT_QTY >= 0 THEN 'RGR'                         "
lgStrSQL = lgStrSQL & " 					 ELSE       		'FGR'                                                                       "
lgStrSQL = lgStrSQL & " 			    END IO_TYPE_CD,                                                                                     "
lgStrSQL = lgStrSQL & "          		B.BP_CD,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		A.PO_NO,                                                                                                                            "
lgStrSQL = lgStrSQL & "          		A.PO_SEQ_NO,                                                                                                                        "
lgStrSQL = lgStrSQL & "          		A.PLANT_CD,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		H.MAJOR_SL_CD SL_CD,  " '--입출고창고(MVMT_SL_CD), 입고창고(MVMT_RCPT_SL_CD) 어떤것인지 확인 필요                                       "
lgStrSQL = lgStrSQL & "          		D.ITEM_CD,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		D.ITEM_NM,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		D.SPEC,                                                                                                                             "
lgStrSQL = lgStrSQL & "          		'*' TRACKING_NO,                                                                                                                    "
lgStrSQL = lgStrSQL & "          		0 PO_QTY,	" '--발주수량 : 좀더 확인해 보도록 한다.                                                                                   "
lgStrSQL = lgStrSQL & "          		A.MVMT_RCPT_UNIT PO_UNIT,			     " '--발주단위                                                                                     "
lgStrSQL = lgStrSQL & "          		0 PO_PRC,		                                                                                                            "
lgStrSQL = lgStrSQL & "          		0 PO_DOC_AMT,	" '--발주금액 : 좀더 확인해 보도록 한다.                                                                               "
lgStrSQL = lgStrSQL & "          		'' PO_CUR,			     " '--발주화폐단위                                                                                     "
lgStrSQL = lgStrSQL & "          		A.MVMT_RCPT_DT DLVY_DT,	" '--입출고일(MVMT_DT), 입고일(MVMT_RCPT_DT) 어떤것인지 확인 필요                                                  "
lgStrSQL = lgStrSQL & "          		ISNULL(MVMT_RCPT_QTY,0) RCPT_QTY,                                                                                                   "
lgStrSQL = lgStrSQL & "          		0 LC_QTY,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		0 PRE_IV_QTY,		  " '--참조매입수량                                                                                                "
lgStrSQL = lgStrSQL & "          		0 INSPECT_QTY,		      " '--검사수량                                                                                                "
lgStrSQL = lgStrSQL & "          		0 IV_QTY,                                                                                                                           "
lgStrSQL = lgStrSQL & "          		[dbo].[ufn_GetRecvInspecFlg](A.PLANT_CD, D.ITEM_CD)  RECV_INSPEC_FLG,                                                                                                        "
lgStrSQL = lgStrSQL & "          		'입고전 검사' MINOR_NM,                                                                                                    "
lgStrSQL = lgStrSQL & "          		'B' INSPECT_METHOD,                                                                                                       "
lgStrSQL = lgStrSQL & "          		F.PLANT_NM,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		'' SL_NM,                                                                                                                            "
lgStrSQL = lgStrSQL & "                  '' PUR_GRP,			  " '--구매그룹                                                                                                "
lgStrSQL = lgStrSQL & "                  0 LC_RCPT_QTY,                                                                                                                     "
lgStrSQL = lgStrSQL & "                  H.LOT_FLG,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		ISNULL((SELECT LOT_GEN_MTHD FROM B_LOT_CONTROL WHERE PLANT_CD = A.PLANT_CD AND ITEM_CD = D.ITEM_CD),'') LOT_GEN_MTHD,               "
lgStrSQL = lgStrSQL & "                  '' MVMT_NO,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  A.TRANS_TIME  ,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  A.MAIN_LOT    ,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  A.IMPORT_TIME ,                                                                                                                         "
lgStrSQL = lgStrSQL & "                  CREATE_TYPE ,                                                                                                                         "
lgStrSQL = lgStrSQL & "          		'N' FLAG       " '--무상                                                                                                               "
lgStrSQL = lgStrSQL & "          from	T_IF_RCV_PART_IN_KO441	A,                                                                                                          "
lgStrSQL = lgStrSQL & "                  B_BIZ_PARTNER B,                                                                                                                   "
lgStrSQL = lgStrSQL & "          		B_ITEM	D,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		B_PLANT F,                                                                                                                          "
lgStrSQL = lgStrSQL & "          		B_ITEM_BY_PLANT H   ,                                                                                                               "
lgStrSQL = lgStrSQL & "         		M_PUR_ORD_HDR I ,                                                                                                                     "
lgStrSQL = lgStrSQL & "         		(SELECT COUNT(*) CNT, TRANS_TIME, MAIN_LOT, IMPORT_TIME                                                                                                                                            "
lgStrSQL = lgStrSQL & "         		FROM  T_IF_RCV_PART_IN_KO441                                                                                                                                                                       "
lgStrSQL = lgStrSQL & "         		WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlant")), " " , "S") & " "
lgStrSQL = lgStrSQL & "                 AND   CONVERT(CHAR(10), MVMT_RCPT_DT, 120) BETWEEN " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & " AND " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & " "      
lgStrSQL = lgStrSQL & "         		GROUP BY TRANS_TIME, MAIN_LOT, IMPORT_TIME                                                                                                                                                         "
lgStrSQL = lgStrSQL & "         		HAVING COUNT(*) = 1) J                                                                                                                                                                             "
lgStrSQL = lgStrSQL & "          WHERE	A.CUST_CD = B.BP_ALIAS_NM                                                                                                 "
lgStrSQL = lgStrSQL & "          AND		A.MES_ITEM_CD = D.CBM_DESCRIPTION     AND D.ITEM_ACCT <> '10'     AND D.ITEM_ACCT <> '20'                                                                                           "
lgStrSQL = lgStrSQL & "          AND		A.PLANT_CD			=	F.PLANT_CD                                                                                              "
lgStrSQL = lgStrSQL & "          AND		A.PLANT_CD			=	H.PLANT_CD                                                                                              "
lgStrSQL = lgStrSQL & "          AND		D.ITEM_CD = H.ITEM_CD                                                                                               "
lgStrSQL = lgStrSQL & "          AND		A.PO_NO				*=	I.PO_NO                                                                                                 "
lgStrSQL = lgStrSQL & "          AND	A.TRANS_TIME = J.TRANS_TIME   "
lgStrSQL = lgStrSQL & "          AND A.MAIN_LOT  = J.MAIN_LOT         "
lgStrSQL = lgStrSQL & "          AND A.IMPORT_TIME = J.IMPORT_TIME    "
lgStrSQL = lgStrSQL & "          AND        ISNULL(A.PO_NO, 'N') = 'N'                                                                                                      "
lgStrSQL = lgStrSQL & "          AND        ISNULL(A.ERP_APPLY_FLAG, 'N') <> 'Y'                                                                                                      "
lgStrSQL = lgStrSQL & "          AND   CONVERT(CHAR(10), A.MVMT_RCPT_DT, 120) BETWEEN " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & " AND " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & " "           
lgStrSQL = lgStrSQL & "         ) NEPES                                                                                                                                     "
lgStrSQL = lgStrSQL & " WHERE ISNULL(NEPES.PO_NO,'') LIKE " & FilterVar(Trim(Request("txtPoNo"))&"%", " " , "S") & " "
If Len(Request("txtPlant")) Then
	lgStrSQL = lgStrSQL & " AND NEPES.PLANT_CD = " & FilterVar(Trim(Request("txtPlant")), " " , "S") & " "
End If	
If Len(Request("txtSupplier")) Then
	lgStrSQL = lgStrSQL & " AND NEPES.BP_CD LIKE " & FilterVar(Trim(Request("txtSupplier")) & "%", " ", "S") & " "
End If
If Len(Trim(Request("rdoClsFlg"))) Then
	lgStrSQL = lgStrSQL & " AND NEPES.FLAG LIKE  " & FilterVar(Trim(UCase(Request("rdoClsFlg")))& "%", " " , "S") & " "
End If

        
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

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IO_TYPE_NM"))         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IO_TYPE_CD"))         
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))          
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_CD"))              
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))               
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))        
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PO_QTY"), ggQty.DecPoint, 0)

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))  
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PO_PRC"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PO_DOC_AMT"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_CUR"))               
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DLVY_DT"))            

            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("RCPT_QTY"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("LC_QTY"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PRE_IV_QTY"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("INSPECT_QTY"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("IV_QTY"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("RECV_INSPEC_FLG"))    
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM"))           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INSPECT_METHOD"))     
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))           
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))              
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_GRP"))            
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("LC_RCPT_QTY"), ggQty.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_FLG"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_GEN_MTHD")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FLAG")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MVMT_NO")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRANS_TIME")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MAIN_LOT")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IMPORT_TIME")) 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CREATE_TYPE")) 

'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MVMT_RCPT_NO"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IO_TYPE_CD"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IO_TYPE_NM"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MVMT_RCPT_DT"))
'         lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_GRP"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PUR_GRP_NM"))
'          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))

'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_cd"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("wc_nm"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_no"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("lot_sub_no"))
'20071221::hanc          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("good_on_hand_qty"), ggQty.DecPoint, 0)
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


