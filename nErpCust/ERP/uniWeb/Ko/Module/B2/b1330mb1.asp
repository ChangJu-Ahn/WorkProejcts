<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->

<%

    On Error Resume Next
    Err.Clear

    Call HideStatusWnd

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")

    lgKeyStream  = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext   = Request("txtPrevNext")                                            '¢Ð: "P"(Prev search) "N"(Next search)


    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next
    Err.Clear

    Call SubMakeSQLStatements("SR")                                                  '¢Ð : Make sql statements ,SR : Single Read
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then               'R(Read) X(CursorType) X(LockType) 
       If lgPrevNext = "Q" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : No data is found. 
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : This is the starting data. 
          lgPrevNext = "Q"
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)            '¢Ð : This is the ending data.
          lgPrevNext = "Q"
          Call SubBizQuery()
       End If
    Else
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
       Response.Write  " <Script Language=vbscript>" & vbCr
       Response.Write  " With Parent						 " & vbCr
	   Response.Write  "       .Frm1.txtCardCoCdQ.Value		= """ & ConvSPChars(lgObjRs("CARD_CO_CD"))		& """" & vbCr
       Response.Write  "       .Frm1.txtCardCoCd.Value			= """ & ConvSPChars(lgObjRs("CARD_CO_CD"))		& """" & vbCr
       Response.Write  "       .Frm1.txtCardCoNm.Value			= """ & ConvSPChars(lgObjRs("CARD_CO_NM"))		& """" & vbCr
       Response.Write  "       .Frm1.txtBankAcctNo.Value		= """ & ConvSPChars(lgObjRs("BANK_ACCT_NO"))	& """" & vbCr
       Response.Write  "       .Frm1.txtBankCd.Value			= """ & ConvSPChars(lgObjRs("BANK_CD"))				& """" & vbCr
       Response.Write  "       .Frm1.txtBankNm.Value			= """ & ConvSPChars(lgObjRs("BANK_NM"))				& """" & vbCr
       If ConvSPChars(lgObjRs("RCPT_CARD_FG"))	 = "Y" and ConvSPChars(lgObjRs("PAY_CARD_FG"))	 = "Y" Then
			Response.Write  "       .Frm1.ChkRcptCard.Checked		= """ & ConvSPChars(True)		& """" & vbCr
			Response.Write  "       .Frm1.ChkPayCard.Checked		= """ & ConvSPChars(True)		& """" & vbCr
		ElseIf ConvSPChars(lgObjRs("RCPT_CARD_FG"))	 = "Y" and ConvSPChars(lgObjRs("PAY_CARD_FG"))	 = "N" Then
			Response.Write  "       .Frm1.ChkRcptCard.Checked		= """ & ConvSPChars(True)		& """" & vbCr
			Response.Write  "       .Frm1.ChkPayCard.Checked		= """ & ConvSPChars(False)		& """" & vbCr
		ElseIf ConvSPChars(lgObjRs("RCPT_CARD_FG"))	 = "N" and ConvSPChars(lgObjRs("PAY_CARD_FG"))	 = "Y" Then
			Response.Write  "       .Frm1.ChkRcptCard.Checked		= """ & ConvSPChars(False)		& """" & vbCr
			Response.Write  "       .Frm1.ChkPayCard.Checked		= """ & ConvSPChars(True)		& """" & vbCr
		Else
			Response.Write  "       .Frm1.ChkRcptCard.Checked		= """ & ConvSPChars(False)		& """" & vbCr
			Response.Write  "       .Frm1.ChkPayCard.Checked		= """ & ConvSPChars(False)		& """" & vbCr
		End If

       Response.Write  "       .Frm1.txtRcptCard.value			= """ & ConvSPChars(lgObjRs("RCPT_CARD_FG"))		& """" & vbCr
       Response.Write  "       .Frm1.txtPayCard.value			= """ & ConvSPChars(lgObjRs("PAY_CARD_FG"))		& """" & vbCr
       Response.Write  "       .Frm1.txtZipCd.Value				= """ & ConvSPChars(lgObjRs("ZIP_CD"))				& """" & vbCr
       Response.Write  "       .Frm1.txtAddr1.Value				= """ & ConvSPChars(lgObjRs("ADDR1"))				& """" & vbCr
       Response.Write  "       .Frm1.txtAddr2.Value				= """ & ConvSPChars(lgObjRs("ADDR2"))				& """" & vbCr
       Response.Write  "       .Frm1.txtAddr3.Value				= """ & ConvSPChars(lgObjRs("ADDR3"))				& """" & vbCr
       Response.Write  "       .Frm1.txtTelNo1.Value			= """ & ConvSPChars(lgObjRs("TEL_NO1"))				& """" & vbCr
       Response.Write  "       .Frm1.txtTelNo2.Value			= """ & ConvSPChars(lgObjRs("TEL_NO2"))				& """" & vbCr
       Response.Write  "       .Frm1.txtFaxNo.Value				= """ & ConvSPChars(lgObjRs("FAX_NO"))				& """" & vbCr
       Response.Write  "       .Frm1.txtUrl.Value				= """ & ConvSPChars(lgObjRs("HOME_URL"))			& """" & vbCr
       Response.Write  "       .Frm1.txtDesc.Value				= """ & ConvSPChars(lgObjRs("CARD_CO_DESC"))	& """" & vbCr
       Response.Write  "       .DBQueryOk " & vbCr
       Response.Write  " End With         " & vbCr
       Response.Write  " </Script>        " & vbCr
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
    End If

    Call SubCloseRs(lgObjRs)                                                    '¢Ð : Release RecordSSet
	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    Dim lgIntFlgMode

    On Error Resume Next
    Err.Clear

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)


    Select Case lgIntFlgMode
        Case  OPMD_CMODE  : Call SubBizSaveSingleCreate()                            '¢Ð : Create
        Case  OPMD_UMODE  : Call SubBizSaveSingleUpdate()                            '¢Ð : Update
    End Select

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    Dim lgStrSQL

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  B_CARD_CO"
    lgStrSQL = lgStrSQL & " WHERE CARD_CO_CD   = " & FilterVar(lgKeyStream(0), "''", "S")    

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       Response.Write  " <Script Language=vbscript>	" & vbCr
       Response.Write  "       Parent.DbDeleteOk			" & vbCr
       Response.Write  " </Script>							" & vbCr
    End If   

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim lgStrSQL
    Dim tmpDate

    On Error Resume Next
    Err.Clear
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------   
    
    lgStrSQL = "INSERT INTO B_CARD_CO"
    lgStrSQL = lgStrSQL & " ( CARD_CO_CD      , "
    lgStrSQL = lgStrSQL & "   CARD_CO_NM      , "
    lgStrSQL = lgStrSQL & "   RCPT_CARD_FG    , "
    lgStrSQL = lgStrSQL & "   PAY_CARD_FG      , "
    lgStrSQL = lgStrSQL & "   ZIP_CD       , "
    lgStrSQL = lgStrSQL & "   ADDR1         , "
    lgStrSQL = lgStrSQL & "   ADDR2         , "
    lgStrSQL = lgStrSQL & "   ADDR3         , "
    lgStrSQL = lgStrSQL & "   TEL_NO1      , "
    lgStrSQL = lgStrSQL & "   TEL_NO2      , "
    lgStrSQL = lgStrSQL & "   FAX_NO       , "
    lgStrSQL = lgStrSQL & "   HOME_URL    , "
    lgStrSQL = lgStrSQL & "   BANK_CD    , "
    lgStrSQL = lgStrSQL & "   BANK_ACCT_NO    , "
    lgStrSQL = lgStrSQL & "   CARD_CO_DESC    , "
    lgStrSQL = lgStrSQL & "   INSRT_USER_ID    , "
    lgStrSQL = lgStrSQL & "   INSRT_DT    , "
    lgStrSQL = lgStrSQL & "   UPDT_USER_ID    , "
    lgStrSQL = lgStrSQL & "   UPDT_DT      ) "
    lgStrSQL = lgStrSQL & " VALUES(" & FilterVar(       UCase(Request("txtCardCoCd"))		, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtCardCoNm")	, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtRcptCard")		, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtPayCard")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtZipCd")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtAddr1")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtAddr2")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtAddr3")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtTELNo1")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtTELNo2")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtFaxNo")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtUrl")				, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtBankCd")			, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtBankAcctNo")		, "''", "S") & ","
    lgStrSQL = lgStrSQL &					 FilterVar(       Request("txtDesc")			, "''", "S") & ","
	lgStrSQL = lgStrSQL &					 FilterVar(		  gUsrId						, "''", "S") & ", "
	lgStrSQL = lgStrSQL &					 "getdate(), " 					
	lgStrSQL = lgStrSQL &					 FilterVar(		  gUsrId						, "''", "S") & ", "
	lgStrSQL = lgStrSQL &					 "getdate())" 		
	
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       Response.Write  " <Script Language=vbscript>	" & vbCr
       Response.Write  "       Parent.DBSaveOk			" & vbCr
       Response.Write  " </Script>							" & vbCr
    End If   
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    Dim lgStrSQL
    Dim tmpDate

    On Error Resume Next
    Err.Clear
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    tmpDate = FilterVar(UniConvDateAToB(Trim(Request("fpdtCloseDt")),gDateFormatYYYYMM,gServerDateFormat),null,"S")
    
    lgStrSQL = "UPDATE  B_CARD_CO"
    lgStrSQL = lgStrSQL & " SET "  
    lgStrSQL = lgStrSQL & " CARD_CO_NM			= " & FilterVar(           Request("txtCardCoNm")	, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " RCPT_CARD_FG		= " & FilterVar(           Request("txtRcptCard")			, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " PAY_CARD_FG			= " & FilterVar(           Request("txtPayCard")			, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " ZIP_CD				= " & FilterVar(           Request("txtZipCd")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " ADDR1				= " & FilterVar(           Request("txtAddr1")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " ADDR2				= " & FilterVar(           Request("txtAddr2")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " ADDR3				= " & FilterVar(           Request("txtAddr3")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " TEL_NO1				= " & FilterVar(           Request("txtTELNo1")			, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " TEL_NO2				= " & FilterVar(           Request("txtTELNo2")			, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " FAX_NO				= " & FilterVar(           Request("txtFaxNo")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " HOME_URL			= " & FilterVar(           Request("txtUrl")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " BANK_CD				= " & FilterVar(           Request("txtBankCd")			, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " BANK_ACCT_NO		= " & FilterVar(           Request("txtBankAcctNo")		, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " CARD_CO_DESC		= " & FilterVar(           Request("txtDesc")				, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " UPDT_USER_ID		= " & FilterVar(			gUsrId							, "''", "S")  & ","
    lgStrSQL = lgStrSQL & " UPDT_DT				= getdate() "      
    lgStrSQL = lgStrSQL & " WHERE CARD_CO_CD= " & FilterVar(           Request("txtCardCoCd")     , "''", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    If CheckSYSTEMError(Err,True) = True Then
       ObjectContext.SetAbort
    Else
       Response.Write  " <Script Language=vbscript> " & vbCr
       Response.Write  "       Parent.DBSaveOk      " & vbCr
       Response.Write  " </Script>                  " & vbCr
    End If   

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(ByVal pMode)

    On Error Resume Next
    Err.Clear

    Select Case Mid(pMode,2,1)
      Case "R"
             Select Case  lgPrevNext 
	                 Case "Q"
						   lgStrSQL = "Select TOP 1 A.*,B.ZIP_CD, C.BANK_CD , C.BANK_NM, D.BANK_ACCT_NO " 
                           lgStrSQL = lgStrSQL & " From  B_CARD_CO			A,	"
                           lgStrSQL = lgStrSQL & "			B_ZIP_CODE		B, " 
                           lgStrSQL = lgStrSQL & " 			B_BANK				C, " 
                           lgStrSQL = lgStrSQL & "  		B_BANK_ACCT		D " 
                           lgStrSQL = lgStrSQL & " WHERE	A.ZIP_CD *= B.ZIP_CD "
                           lgStrSQL = lgStrSQL & " AND		A.BANK_CD *= C.BANK_CD "
                           lgStrSQL = lgStrSQL & " AND		A.BANK_ACCT_NO *= D.BANK_ACCT_NO "  
                           lgStrSQL = lgStrSQL & " AND		A.CARD_CO_CD = " & FilterVar(lgKeyStream(0), "''", "S")
                           lgStrSQL = lgStrSQL & " ORDER BY CARD_CO_CD ASC "

                     Case "P"
						   lgStrSQL = "Select TOP 1 A.*,B.ZIP_CD, C.BANK_CD , C.BANK_NM, D.BANK_ACCT_NO " 
                           lgStrSQL = lgStrSQL & " From  B_CARD_CO			A,	"
                           lgStrSQL = lgStrSQL & "			B_ZIP_CODE		B, " 
                           lgStrSQL = lgStrSQL & " 			B_BANK				C, " 
                           lgStrSQL = lgStrSQL & "  		B_BANK_ACCT		D " 
                           lgStrSQL = lgStrSQL & " WHERE	A.ZIP_CD *= B.ZIP_CD "
                           lgStrSQL = lgStrSQL & " AND		A.BANK_CD *= C.BANK_CD "
                           lgStrSQL = lgStrSQL & " AND		A.BANK_ACCT_NO *= D.BANK_ACCT_NO "  
                           lgStrSQL = lgStrSQL & " AND		A.CARD_CO_CD < " & FilterVar(lgKeyStream(0), "''", "S")
                           lgStrSQL = lgStrSQL & " ORDER BY CARD_CO_CD DESC "
                     Case "N"
                            lgStrSQL = "Select TOP 1 A.*,B.ZIP_CD, C.BANK_CD , C.BANK_NM, D.BANK_ACCT_NO " 
                           lgStrSQL = lgStrSQL & " From  B_CARD_CO			A,	"
                           lgStrSQL = lgStrSQL & "			B_ZIP_CODE		B, " 
                           lgStrSQL = lgStrSQL & " 			B_BANK				C, " 
                           lgStrSQL = lgStrSQL & "  		B_BANK_ACCT		D " 
                           lgStrSQL = lgStrSQL & " WHERE	A.ZIP_CD *= B.ZIP_CD "
                           lgStrSQL = lgStrSQL & " AND		A.BANK_CD *= C.BANK_CD "
                           lgStrSQL = lgStrSQL & " AND		A.BANK_ACCT_NO *= D.BANK_ACCT_NO "  
                           lgStrSQL = lgStrSQL & " AND		A.CARD_CO_CD > " & FilterVar(lgKeyStream(0), "''", "S")
                           lgStrSQL = lgStrSQL & " ORDER BY CARD_CO_CD ASC "                           
             End Select
    End Select
    
End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

%>
