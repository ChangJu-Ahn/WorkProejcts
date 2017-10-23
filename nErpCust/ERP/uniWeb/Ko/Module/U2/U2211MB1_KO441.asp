<% 
'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 품의서 문서관리(S) 
'*  3. Program ID           : S3322MB1_KO412.asp
'*  4. Program Name         : S3322MB1_KO412.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005/01/25
'*  7. Modified date(Last)  : 2007/07/06
'*  8. Modifier (First)     : Lee Wol san
'*  9. Modifier (Last)      : Lee Ho Jun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 



Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow         
Dim strSpread


	Call loadInfTB19029B("I", "*","NOCOOKIE","MB") 
	Call LoadBasisGlobalInf()
	Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = ""
    lgErrorPos        = ""                                                           '☜: Set to space
    lgLngMaxRow       = Request("txtMaxRows")    
 	strSpread		  = Request("txtSpread")

    Select Case CStr(Request("txtMode"))
        Case "head" 
        	Call SubBizQuery()     
        Case CStr(UID_M0001)         
        	Call SubBizQueryMulti() 
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()        	
    End Select


'-----------------------------------------------------------------------------------------
Sub SubBizQuery()


	Dim strBpCd
	Dim strDlvyNo
	Dim strItemCd
	Dim strDvFrDt
	Dim strDvToDt	


    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim iSelCount
   
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    'on Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

		iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1

	 
		strBpCd		= FilterVar(Request("txtBpCd")&"%", "''", "S")
		strDlvyNo	= FilterVar(Request("txtDlvyNo")&"%", "''", "S")
		strItemCd	= FilterVar(Request("txtItemCd")&"%", "''", "S")	
		strDvFrDt	= FilterVar(UniConvDate(Request("txtDvFrDt")), "''", "S")	
		strDvToDt	= FilterVar(UniConvDate(Request("txtDvToDt")), "''", "S")			


		lgStrSQL =  ""		
		lgStrSQL = lgStrSQL & " SELECT	TOP " & iSelCount & " A.DLVY_NO, A.PO_NO, A.PO_SEQ_NO, E.BP_CD,	"
		lgStrSQL = lgStrSQL & " 		B.ITEM_CD, C.ITEM_NM, C.SPEC, C.BASIC_UNIT,	"
		lgStrSQL = lgStrSQL & " 		A.PLAN_DVRY_DT, A.PLAN_DVRY_QTY, A.D_BP_CD, D.SL_NM, "
		lgStrSQL = lgStrSQL & " 		A.SPLIT_SEQ_NO, B.PO_UNIT, B.TRACKING_NO	"
		lgStrSQL = lgStrSQL & " FROM	M_SCM_FIRM_PUR_RCPT	A(NOLOCK),				"
		lgStrSQL = lgStrSQL & " 		M_PUR_ORD_DTL		B(NOLOCK),              "
		lgStrSQL = lgStrSQL & " 		B_ITEM				C(NOLOCK),              "
		lgStrSQL = lgStrSQL & " 		B_STORAGE_LOCATION	D(NOLOCK),              "
		lgStrSQL = lgStrSQL & " 		M_SCM_DLVY_PUR_RCPT	E(NOLOCK)               "
		lgStrSQL = lgStrSQL & " WHERE	A.PO_NO		= B.PO_NO                       "
		lgStrSQL = lgStrSQL & " AND		A.PO_SEQ_NO	= B.PO_SEQ_NO                   "
		lgStrSQL = lgStrSQL & " AND		B.ITEM_CD	= C.ITEM_CD                     "
		lgStrSQL = lgStrSQL & " AND		A.D_BP_CD	= D.SL_CD                       "
		lgStrSQL = lgStrSQL & " AND		A.DLVY_NO	= E.DLVY_NO                     "
		lgStrSQL = lgStrSQL & " AND		A.INSRT_USER_ID = E.BP_CD                   "
		lgStrSQL = lgStrSQL & " AND		E.BP_CD LIKE " & strBpCd
		lgStrSQL = lgStrSQL & " AND		E.DLVY_NO LIKE " & strDlvyNo
		lgStrSQL = lgStrSQL & " AND		B.ITEM_CD LIKE " & strItemCd
		lgStrSQL = lgStrSQL & " AND		A.PLAN_DVRY_DT >= " & strDvFrDt
		lgStrSQL = lgStrSQL & " AND		A.PLAN_DVRY_DT <= " & strDvToDt
		

	Call SubOpenDB(lgObjConn)

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
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DLVY_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BASIC_UNIT"))  
          lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))                    
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("PLAN_DVRY_QTY"), ggQty.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("D_BP_CD"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPLIT_SEQ_NO"))  
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRACKING_NO"))                    
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)

          lgObjRs.MoveNext

          iDx =  iDx + 1
         If iDx > C_SHEETMAXROWS_D Then
			 lgStrPrevKey = lgStrPrevKey + 1
             Exit Do
         End If        
      Loop 
'Call ServerMesgBox(lgstrData , vbInformation, I_MKSCRIPT)	      
	 
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
       Response.Write  "    Parent.ggoSpread.Source     = Parent.frm1.vspdData1      " & vbCr
       Response.Write  "    Parent.lgStrPrevKey         = """ & lgStrPrevKey    & """" & vbCr
       Response.Write  "    Parent.ggoSpread.SSShowData   """ & lgstrData       & """" & vbCr
       Response.Write  "    Parent.DBQueryOk   " & vbCr       
       Response.Write  " </Script>             " & vbCr
    End If


End Sub  




'-----------------------------------------------------------------------------------------
Sub SubBizQueryMulti()
'-----------------------------------------------------------------------------------------
'Call ServerMesgBox("(10)SubMakeSQLStatements", vbInformation, I_MKSCRIPT)  
    Dim strData
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    DIm arr,arrCnt
    Dim i,j


	Err.Clear                                                                  '☜: Clear Error status
		 
	LngRow = 0



	Dim strBpCd
	Dim strDlvyNo
	Dim strItemCd
	Dim strDvFrDt
	Dim strDvToDt	


    Dim lgStrSQL
    Dim lgstrData
    Dim iDx
    Dim iSelCount
    
    Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수        
    
    'on Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
	
	strBpCd		= FilterVar(Request("txtBpCd")&"%", "''", "S")
	strDlvyNo	= FilterVar(Request("txtDlvyNo")&"%", "''", "S")
	strItemCd	= FilterVar(Request("txtItemCd")&"%", "''", "S")	
	strDvFrDt	= FilterVar(Request("txtDvFrDt"), "''", "S")	
	strDvToDt	= FilterVar(Request("txtDvToDt"), "''", "S")			


		lgStrSQL =  ""
		lgStrSQL = lgStrSQL & " SELECT	BP_CD, "
		lgStrSQL = lgStrSQL & " 		DLVY_NO, "
		lgStrSQL = lgStrSQL & "			DOCUMENT_NO, "
		lgStrSQL = lgStrSQL & "			TITLE, "
		lgStrSQL = lgStrSQL & "			INS_USER, "
		lgStrSQL = lgStrSQL & "			INS_DT,	"
		lgStrSQL = lgStrSQL & "			DOCUMENT_ABBR	"	
		'lgStrSQL = lgStrSQL & "			DOCUMENT_TEXT	"
		lgStrSQL = lgStrSQL & "	FROM	M_SCM_DOCUMENT_HDR_KO441(NOLOCK) "
		lgStrSQL = lgStrSQL & "	WHERE	BP_CD LIKE " & strBpCd
		lgStrSQL = lgStrSQL & "	AND 	DLVY_NO LIKE " & strDlvyNo
		lgStrSQL = lgStrSQL & "	ORDER BY BP_CD, DLVY_NO, DOCUMENT_NO"	



	
	Call SubOpenDB(lgObjConn)   
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		arrCnt = adoRec.RecordCount 
	If arrCnt > 0 then 		arr=adoRec.GetRows
	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

	If arrCnt>0 then
		For i=0 to arrCnt-1
		   LngRow = LngRow + 1
			For j=0 to uBound(arr,1)
				strData = strData & Chr(11) & arr(j,i)
			Next 
			strData =  strData & Chr(11) & LngRow &  Chr(11) & Chr(12) 
		Next 

	Else
		Call DisplayMsgBox("900014", vbOKOnly, "자료", "", I_MKSCRIPT)
    End if
%>

<Script Language=vbscript>
    Dim LngRow          
    Dim strTemp
    Dim strData
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
	.frm1.vspdData.ReDraw = False
	 strData = "<%=ConvSPChars(strData)%>"
    .ggoSpread.Source = .frm1.vspdData 
  	.ggoSpread.SSShowData strData
	.frm1.vspdData.ReDraw = True
	.DbQueryDtlOk
	End With
</Script>	
<% 

End Sub    


'========================================================================
'SubMakeSQLStatements
'========================================================================
Sub SubMakeSQLStatements(pvSpdNo)
End Sub



'========================================================================
'SubBizDelete(삭제)
'========================================================================

Sub SubBizSaveMulti()

	'on Error Resume Next
	
	Dim strSql
	Dim strBpCd
	Dim strDlvyNo
	Dim strDocumentNo
	
	Dim Temp,i
	Dim arrFile_id
	
	temp = split(strSpread,chr(12))
	
	Call SubOpenDB(lgObjConn)

	lgObjConn.beginTrans()

	for i = 0 to  UBound(temp)-1
	
		strBpCd = split(temp(i),chr(11))(2)
		strDlvyNo = split(temp(i),chr(11))(3)
		strDocumentNo = split(temp(i),chr(11))(4)
		
		lgStrSQL="SELECT DOCUMENT_ID FROM M_SCM_DOCUMENT_DTL_KO441 WHERE BP_CD ='"&strBpCd&"' and DLVY_NO = '"&strDlvyNo &"' and DOCUMENT_NO = '"&strDocumentNo &"' "
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		'================
		'FILE DELETE ADD
		'================
		if not adoRec.eof then 'file list가 있으면 file_Id배열에 담은후 B_CIS_FILE_DETAIL delete
			arrFile_id= adoRec.getRows()
			Call FileDelete(arrFile_id)
		end if
		adoRec.Close 
		strSql="DELETE FROM M_SCM_DOCUMENT_HDR_KO441 WHERE BP_CD='" & strBpCd & "' and DLVY_NO = '" & strDlvyNo & "' and DOCUMENT_NO = '" & strDocumentNo & "' "
		'Response.Write strSQL
		lgObjConn.execute strSql
		strSql="DELETE FROM M_SCM_DOCUMENT_DTL_KO441 WHERE BP_CD='" & strBpCd & "' and DLVY_NO = '" & strDlvyNo & "' and DOCUMENT_NO = '" & strDocumentNo & "' "
		lgObjConn.execute strSql
		'Response.Write strSQL		
	next

	If CheckSYSTEMError(Err,True) = True Then                                              
		lgObjConn.rollbacktrans()
		Call SubCloseDB(lgObjConn) 
		Response.End 
	else
		lgObjConn.committrans()
		Call SubCloseDB(lgObjConn) 
		Call DisplayMsgBox("210032", vbOKOnly, "", "", I_MKSCRIPT)  '삭제되었습니다!
    End If
	
%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		CALL .DbSaveOk()
		parent.MyBizASP1.location.reload
	End With
</Script>
<%					

End Sub

%>

<%


'========================================================================
'FileDelete
'========================================================================
Function FileDelete(byVal pArr )
 	'on Error Resume Next
	
	Dim filePath
 	Dim i
	
	filePath=server.MapPath (".")&"\files\"

	For i=0 to uBound(pArr,2)
		Call pfile.fileDelete(replace(filePath & pArr(0,i),"\","/"))   
	Next
End Function



Sub SubBizSave()
    'on Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

Sub SubBizDelete()
    'on Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
%>

<OBJECT RUNAT=server PROGID="ADODB.Recordset" id=adoRec></OBJECT>
<OBJECT RUNAT=server PROGID="PuniFile.CTransfer" id=pfile></OBJECT>

