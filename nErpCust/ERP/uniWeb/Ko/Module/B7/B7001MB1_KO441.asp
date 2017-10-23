
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
    If lgStrPrevKey = 0 Then
            lgStrSQL = "                    SELECT	Z.PLANT_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.PLANT_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.ITEM_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.ITEM_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.SPEC,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT1) AS CNT1,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT2) AS CNT2,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT3) AS CNT3,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT4) AS CNT4	"
            lgStrSQL = lgStrSQL & vbCrLf & " FROM	"
            lgStrSQL = lgStrSQL & vbCrLf & " 		("
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PLANT_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.ITEM_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT1,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT2,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT3,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT4"
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	B_ITEM_BY_PLANT		A"
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B"
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_cD	=	A.ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		A.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PLANT_CD,	C.PLANT_NM,A.ITEM_CD, B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 				UNION	ALL "
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PLANT_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.ITEM_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT1, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT2, " 
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT3, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT4 "
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	P_COMPNT_REQ_BY_ROUTING	A "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_cD	=	A.ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PLANT_CD,	C.PLANT_NM,A.ITEM_CD,B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 				UNION	ALL "
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PRNT_PLANT_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.PRNT_ITEM_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT1, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT2, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT3, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT4 "
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	P_BOM_DETAIL		A "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_cD	=	A.PRNT_ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PRNT_PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PRNT_PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PRNT_PLANT_CD,C.PLANT_NM,A.PRNT_ITEM_CD,B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 				UNION	ALL "
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PLANT_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.ITEM_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT1, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT2, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT3, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT4 "
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	P_ROUTING_DETAIL	A "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_CD	=	A.ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PLANT_CD,	C.PLANT_NM,A.ITEM_CD,B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 		)		Z "
            lgStrSQL = lgStrSQL & vbCrLf & " GROUP	BY	Z.PLANT_CD,Z.PLANT_NM,Z.ITEM_CD,Z.ITEM_NM,Z.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " ORDER	BY	Z.PLANT_CD,Z.PLANT_NM,Z.ITEM_CD,Z.ITEM_NM,Z.SPEC "


       If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
          Response.Write  " <Script Language=vbscript>            " & vbCr
          Response.Write  "   Parent.Frm1.txtPlantCd.Value  = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set condition area
          Response.Write  "   Parent.Frm1.txtPlantNm.Value  = """ & lgObjRs("plant_nm") & """" & vbCr 
          Response.Write  "   Parent.Frm1.htxtPlantCd.Value = """ & lgObjRs("plant_cd") & """" & vbCr             ' Set next key data
          'Response.Write  "   Parent.Frm1.txtItemCd.Value  = """ & lgObjRs("item_cd") & """" & vbCr             ' Set condition area
          'Response.Write  "   Parent.Frm1.txtItemNm.Value  = """ & lgObjRs("item_nm") & """" & vbCr 
          'Response.Write  "   Parent.Frm1.htxtItemCd.Value = """ & lgObjRs("item_cd") & """" & vbCr             ' Set next key data
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

            lgStrSQL = lgStrSQL & " select TOP " & iSelCount & " 	Z.PLANT_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.PLANT_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.ITEM_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.ITEM_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		Z.SPEC,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT1) AS CNT1,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT2) AS CNT2,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT3) AS CNT3,"
            lgStrSQL = lgStrSQL & vbCrLf & " 		SUM(Z.CNT4) AS CNT4	"
            lgStrSQL = lgStrSQL & vbCrLf & " FROM	"
            lgStrSQL = lgStrSQL & vbCrLf & " 		("
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PLANT_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.ITEM_CD,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT1,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT2,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT3,"
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT4"
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	B_ITEM_BY_PLANT		A"
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B"
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_cD	=	A.ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		A.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PLANT_CD,	C.PLANT_NM,A.ITEM_CD, B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 				UNION	ALL "
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PLANT_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.ITEM_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT1, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT2, " 
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT3, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT4 "
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	P_COMPNT_REQ_BY_ROUTING	A "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_cD	=	A.ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PLANT_CD,	C.PLANT_NM,A.ITEM_CD,B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 				UNION	ALL "
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PRNT_PLANT_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.PRNT_ITEM_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT1, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT2, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT3, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT4 "
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	P_BOM_DETAIL		A "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_cD	=	A.PRNT_ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PRNT_PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PRNT_PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")            
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PRNT_PLANT_CD,C.PLANT_NM,A.PRNT_ITEM_CD,B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 				UNION	ALL "
            lgStrSQL = lgStrSQL & vbCrLf & " 				SELECT	A.PLANT_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						C.PLANT_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						A.ITEM_CD, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.ITEM_NM, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						B.SPEC, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT1, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT2, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						0			CNT3, "
            lgStrSQL = lgStrSQL & vbCrLf & " 						COUNT(*)	CNT4 "
            lgStrSQL = lgStrSQL & vbCrLf & " 				FROM	P_ROUTING_DETAIL	A "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_ITEM		B "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		B.ITEM_CD	=	A.ITEM_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 						JOIN	B_PLANT		C "
            lgStrSQL = lgStrSQL & vbCrLf & " 						ON		C.PLANT_CD	=	A.PLANT_CD "
            lgStrSQL = lgStrSQL & vbCrLf & " 				WHERE	A.PLANT_CD	=	" & FilterVar(lgKeyStream(0),"''", "S")
			lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_CD LIKE " & FilterVar(Trim(lgKeyStream(1)) & "%","''", "S")                        
            lgStrSQL = lgStrSQL & vbCrLf & " 				AND		B.ITEM_ACCT	IN	('10','20') "
            lgStrSQL = lgStrSQL & vbCrLf & " 				GROUP	BY	A.PLANT_CD,	C.PLANT_NM,A.ITEM_CD,B.ITEM_NM,B.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " 		)		Z "
            lgStrSQL = lgStrSQL & vbCrLf & " GROUP	BY	Z.PLANT_CD,Z.PLANT_NM,Z.ITEM_CD,Z.ITEM_NM,Z.SPEC "
            lgStrSQL = lgStrSQL & vbCrLf & " ORDER	BY	Z.PLANT_CD,Z.PLANT_NM,Z.ITEM_CD,Z.ITEM_NM,Z.SPEC "
        
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
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_cd"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("item_nm"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("spec"))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CNT1"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CNT2"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CNT3"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("CNT4"), ggQty.DecPoint, 0)

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


