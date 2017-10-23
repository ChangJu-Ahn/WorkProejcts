<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
    Dim lgStrPrevKey_m
    On Error Resume Next                                                   '☜: Protect prorgram from crashing

    Err.Clear                                                              '☜: Clear Error status
    
    Call HideStatusWnd                                                     '☜: Hide Processing message
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

    lgErrorStatus  = "NO"
    lgKeyStream    = Split(Request("txtKeyStream_m"),gColSep) 
    lgStrPrevKey_m   = Request("lgStrPrevKey_m")                '☜: Next Key

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case CStr(Request("txtMode"))
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection


'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()


    Dim lgStrSQL
    Dim lgstrData
    Dim lgLngMaxRow
    Dim iDx
    
    Const C_SHEETMAXROWS_D  = 30                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	IF lgStrPrevKey_m = "" Then
		lgStrPrevKey_m = 0
	End IF

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    lgStrSQL = "Select TOP " & C_SHEETMAXROWS_D + 1 & "	Seq,	Item_desc, 	Dtl_amt,	Dtl_loc_amt"
    lgStrSQL = lgStrSQL & " From	a_asset_item" 
    lgStrSQL = lgStrSQL & " WHERE	asst_no = " & FilterVar(lgKeyStream(0), "''", "S")  
    lgStrSQL = lgStrSQL & " AND		seq >= " & lgStrPrevKey_m
    lgStrSQL = lgStrSQL & " ORDER BY	seq"

 '       Call DisplayMsgBox(lgStrSQL, vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 



   If 	FncOpenRs("U",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey_m = 0
'        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)              '☜ : No data is found. 
        lgErrorStatus     = "YES"
        Exit Sub
    Else 
       
       iDx = 1
       lgstrData = ""
       lgLngMaxRow       = Request("txtMaxRows_m")                                        '☜: Read Operation Mode (CRUD)
       
       Do While Not lgObjRs.EOF
          lgstrData = lgstrData & Chr(11) & Cint(lgObjRs("Seq"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Item_desc"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Dtl_amt"),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgKeyStream(0))			'Asset_no를 Hidden으로 가지고 간다.
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("Dtl_loc_amt"),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)
          
          lgObjRs.MoveNext

          iDx =  iDx + 1
          If iDx > C_SHEETMAXROWS_D Then
             Exit Do
         End If   
      Loop 
    End If
    
    If Not lgObjRs.EOF Then
       lgStrPrevKey_m = UNINumClientFormat(lgObjRs("Seq"),ggAmtOfMoney.DecPoint, 0)
    Else
       lgStrPrevKey_m = 0
    End If
    
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

    If lgErrorStatus  = "NO" Then
       Response.Write  " <Script Language=vbscript>                            " & vbCr
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData2 " & vbCr
       Response.Write  "    Parent.lgStrPrevKey_m         = " & lgStrPrevKey_m    & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData      & """" & vbCr
       Response.Write  "    Parent.DbQueryOk2   " & vbCr      
       Response.Write  " </Script>             " & vbCr
    End If
    
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

