<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : S1921MA1_KO441
'*  4. Program Name         : 업체별적용환율등록(KO441) 
'*  5. Program Desc         : 업체별적용환율등록(KO441)
'*  6. Component List       : 
'*  7. Modified date(First) : 2008/06
'*  8. Modified date(Last)  : 2008/06
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : ajc
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->



<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->



<%	
Call HideStatusWnd
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

Dim lgSvrDateTime
    lgSvrDateTime = GetSvrDateTime      

Dim CheckFlg , Grid2Key1, Grid2Key2
	Grid2Key1 = ""
	Grid2Key2 = ""
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgOpModeCRUD  = Request("txtMode") 
	lgErrorStatus     = "NO"
	
	Call SubOpenDB(lgObjConn) 
	 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)        
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
             Call SubBizSaveMulti2()
        Case CStr("Grid2")     
			 Call SubBizQueryMulti2()
             
    End Select
    
    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    Dim iLngRow, iDx
    Const MaxCnt = 9999
    
	lgStrSQL = " SELECT	EXCH_YYYYMM, CURRENCY, EXCH_RATE_F, EXCH_RATE_L, REMARK "
	lgStrSQL = lgStrSQL &  " FROM S_BP_APPLY_EXCH_HDR_KO441(nolock) " 
	lgStrSQL = lgStrSQL &  " WHERE 1=1 " 

	If Len(Request("txtConDt")) Then
		lgStrSQL = lgStrSQL &  " and EXCH_YYYYMM >= " & FilterVar(Replace(Request("txtConDt"), "-", ""), "''", "S")
	End If	

	If Len(Request("lgStrPrevKey")) Then
		lgStrSQL = lgStrSQL &  " and EXCH_YYYYMM >= " & FilterVar(Request("lgStrPrevKey"), "''", "S")
		lgStrSQL = lgStrSQL &  " and CURRENCY > " & FilterVar(Request("lgStrPrevKey2"), "''", "S")
	End If	
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
'			Call svrmsgbox("STA", vbinformation, i_mkscript)
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		lgErrorStatus = "Yes"
	Else
		lgstrData = ""
		iDx       = 1

		Do While Not lgObjRs.EOF
			
			If iDx = MaxCnt Then
			%>
				<Script Language=vbscript>
					With parent
					    .frm1.HlgStrPrevKey.value = "<%=ConvSPChars(lgObjRs("EXCH_YYYYMM"))%>"
					    .frm1.HlgStrPrevKey2.value = "<%=ConvSPChars(lgObjRs("1"))%>"
					End With
				</Script>
			<%	
				Exit Do
			Else
			
				lgstrData = lgstrData & Chr(11) & ConvSPChars(Left(lgObjRs("EXCH_YYYYMM"),4) & "-" & Right(lgObjRs("EXCH_YYYYMM"),2))																		'	C_Select_Fg
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CURRENCY"))											'	C_Pay_Ym
				lgstrData = lgstrData & Chr(11) & ""																		'	C_Pay_No
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXCH_RATE_F"))										'	C_Use_Dt
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXCH_RATE_L"))										'	C_Store_Nm
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))											'	C_Card_No
				
				lgstrData = lgstrData & Chr(11) & ConvSPChars(Left(lgObjRs("EXCH_YYYYMM"),4) & "-" & Right(lgObjRs("EXCH_YYYYMM"),2))
				lgstrData = lgstrData & Chr(11) & iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
				lgObjRs.MoveNext				
				
				iDx =  iDx + 1
			End If
		Loop 
		
	End If
	
End Sub

'============================================================================================================
' Name : SubBizQuery2
' Desc : Date data 
'============================================================================================================
Sub SubBizQueryMulti2()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    Dim iLngRow, iDx
    Const MaxCnt = 9999
    
	lgStrSQL = " SELECT	a.EXCH_YYYYMM, a.CURRENCY, a.BP_CD, a.EXCH_APPLY, a.REMARK, b.BP_NM "
	lgStrSQL = lgStrSQL &  " FROM S_BP_APPLY_EXCH_DTL_KO441 a(nolock), B_BIZ_PARTNER B(nolock) " 
	lgStrSQL = lgStrSQL &  " WHERE a.EXCH_YYYYMM = " & FilterVar(Request("Key1"), "''", "S")
	lgStrSQL = lgStrSQL &  "   AND a.CURRENCY = " & FilterVar(Request("Key2"), "''", "S")
	lgStrSQL = lgStrSQL &  "   AND a.BP_CD = b.BP_CD " 
	
	If Len(Request("lgStrPrevKey3")) Then
		lgStrSQL = lgStrSQL &  " and a.BP_CD > " & FilterVar(Request("lgStrPrevKey3"), "''", "S")
	End If	
	
'call svrmsgbox(lgStrSQL, vbinformation, i_mkscript)
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
	Else
		lgstrData = ""
		iDx       = 1

		Do While Not lgObjRs.EOF
			
			If iDx = MaxCnt Then
			%>
				<Script Language=vbscript>
					With parent
					    .frm1.HlgStrPrevKey3.value = "<%=ConvSPChars(lgObjRs("BP_CD"))%>"
					End With
				</Script>
			<%	
				Exit Do
			Else
'	Call svrmsgbox(ConvSPChars(lgObjRs("EXCH_YYYYMM")), vbinformation, i_mkscript)		
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))					'	C_BpCd
				lgstrData = lgstrData & Chr(11) & ""											'	C_BpCdPopup
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))					'	C_BpNm
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXCH_APPLY"))			'	C_ExchApplyCd
				
				If ConvSPChars(lgObjRs("EXCH_APPLY")) = "F" Then
					lgstrData = lgstrData & Chr(11) & "최초고시환율"							'	C_ExchApply
				Else 
					lgstrData = lgstrData & Chr(11) & "최종고시환율"
				End If
				
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))				'	C_Remark2
				
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXCH_YYYYMM"))			'	C_ConDt2				
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CURRENCY"))				'	C_DocCur2
				lgstrData = lgstrData & Chr(11) & iDx
				lgstrData = lgstrData & Chr(11) & Chr(12)
				lgObjRs.MoveNext				
				
				iDx =  iDx + 1
			End If
		Loop 
		
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
	Err.Clear	

	Dim lgIntFlgMode	
    Dim arrVal, arrTemp	
	Dim LngMaxRow,LngRow
    Dim iRowsep,iColsep
	Dim ii
	
	Dim itxtSpread, iCUCount, itxtSpreadArrCount
	Dim itxtSpreadArr
	
    lgIntFlgMode = CInt(Request("txtFlgMode"))

    itxtSpread = ""
    iCUCount = Request.Form("txtCUSpread").Count
    itxtSpreadArrCount = -1

    ReDim itxtSpreadArr(iCUCount)

    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next

    itxtSpread = Join(itxtSpreadArr,"")
    
    Response.Write "<Script language=vbs> " & vbCr
    Response.Write "Parent.RemovedivTextArea "      & vbCr
    Response.Write "</Script> "      & vbCr
    
    
	arrTemp = itxtSpread
	LngMaxRow = CLng(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

	arrTemp = Split(arrTemp, gRowSep)
	
	CheckFlg = True
		
	If Ubound(arrTemp) <= 0 Then 
		CheckFlg = False
		Exit Sub
	End If
	
    For LngRow = 1 To Ubound(arrTemp)													'☜: Group Count
		 
		arrVal = Split(arrTemp(LngRow-1), gColSep)									' 컬럼 넘어 온다
		
		Select Case arrVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrVal)                            '☜: Delete
        End Select

    Next 
End Sub   
'============================================================================================================
' Name : SubBizSaveMulti2
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti2()
	On Error Resume Next
	Err.Clear	
	
	Dim itxtSpread
	Dim lgIntFlgMode	
    Dim arrVal, arrTemp	
	Dim LngMaxRow,LngRow
    Dim iRowsep,iColsep
	Dim ii
 	
    itxtSpread = Request("txtSpread")
'Call ServerMesgBox("STA", vbInformation, I_MKSCRIPT)        
	arrTemp = itxtSpread

	arrTemp = Split(arrTemp, gRowSep)
	
    For LngRow = 1 To Ubound(arrTemp)													'☜: Group Count
		 
		arrVal = Split(arrTemp(LngRow-1), gColSep)									' 컬럼 넘어 온다
	
		If CheckFlg = False And Grid2Key1 = "" Then
			Grid2Key1 = arrVal(5)
			Grid2Key2 = arrVal(6)
		End If
		
		Select Case arrVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate2(arrVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate2(arrVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete2(arrVal)                            '☜: Delete
        End Select

    Next 
    				
End Sub    
'==================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
        
'    Call ServerMesgBox(FilterVar(UCase(arrColVal(6)), "''", "S"), vbInformation, I_MKSCRIPT)
	
    lgStrSQL = "INSERT INTO S_BP_APPLY_EXCH_HDR_KO441( "    
    lgStrSQL = lgStrSQL & " EXCH_YYYYMM,	CURRENCY,	EXCH_RATE_F,	EXCH_RATE_L, REMARK,		"    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID,	INSRT_DT,  UPDT_USER_ID,	UPDT_DT	)	"

    lgStrSQL = lgStrSQL & " VALUES("       
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")		& ","							'년월			2
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")		& ","							'통화			3	    
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(4),0)					& ","							'최초고시환율	4    
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)					& ","							'최종고시환율	5
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")     & ","							'비고			6    

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")					& "," 
    lgStrSQL = lgStrSQL & "getdate()"									& "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")					& "," 
    lgStrSQL = lgStrSQL & "getdate()"
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	    
End Sub
'==================================================
Sub SubBizSaveMultiCreate2(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
        
'    Call ServerMesgBox(arrColVal(2), vbInformation, I_MKSCRIPT)
	
    lgStrSQL = "INSERT INTO S_BP_APPLY_EXCH_DTL_KO441( "    
    lgStrSQL = lgStrSQL & " EXCH_YYYYMM,	CURRENCY,	BP_CD,	EXCH_APPLY, REMARK,		"    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID,	INSRT_DT,  UPDT_USER_ID,	UPDT_DT	)	"

    lgStrSQL = lgStrSQL & " VALUES("       
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(5)), "''", "S")		& ","							'년월			5
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(6)), "''", "S")		& ","							'통화			6	    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S")		& ","							'업체			2    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")		& ","							'적용고시		3
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","							'비고사항		4    

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")					& "," 
    lgStrSQL = lgStrSQL & "getdate()"									& "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")					& "," 
    lgStrSQL = lgStrSQL & "getdate()"
    lgStrSQL = lgStrSQL & ")"
'Call svrmsgbox(lgstrsql, vbinformation, i_mkscript)

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	    
End Sub
'==================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
        
'    Call ServerMesgBox(FilterVar(UCase(arrColVal(6)), "''", "S"), vbInformation, I_MKSCRIPT)
	
    lgStrSQL = "UPDATE S_BP_APPLY_EXCH_HDR_KO441 SET "    
    lgStrSQL = lgStrSQL & "  EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & " ,EXCH_RATE_F = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    lgStrSQL = lgStrSQL & " ,EXCH_RATE_L = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & " ,REMARK = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    
    lgStrSQL = lgStrSQL & " ,UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & " ,UPDT_DT = getdate() "
    
    lgStrSQL = lgStrSQL & " WHERE EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(7)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND CURRENCY	  = " & FilterVar(UCase(arrColVal(3)), "''", "S")    
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
	lgStrSQL = "UPDATE S_BP_APPLY_EXCH_DTL_KO441 SET "    
    lgStrSQL = lgStrSQL & "  EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(2)), "''", "S")
  
    lgStrSQL = lgStrSQL & " ,UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & " ,UPDT_DT = getdate() "
    
    lgStrSQL = lgStrSQL & " WHERE EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(7)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND CURRENCY	  = " & FilterVar(UCase(arrColVal(3)), "''", "S")    
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	    
End Sub
'==================================================
Sub SubBizSaveMultiUpdate2(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
        
'    Call ServerMesgBox(FilterVar(UCase(arrColVal(6)), "''", "S"), vbInformation, I_MKSCRIPT)
	
    lgStrSQL = "UPDATE S_BP_APPLY_EXCH_DTL_KO441 SET "    
    lgStrSQL = lgStrSQL & "  EXCH_APPLY = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " ,REMARK = " & FilterVar(UCase(arrColVal(4)), "''", "S")
       
    lgStrSQL = lgStrSQL & " ,INSRT_USER_ID = " & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & " ,INSRT_DT = getdate() "
    lgStrSQL = lgStrSQL & " ,UPDT_USER_ID = " & FilterVar(gUsrId, "''", "S")
    lgStrSQL = lgStrSQL & " ,UPDT_DT = getdate() "
    
    lgStrSQL = lgStrSQL & " WHERE EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(5)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND CURRENCY	  = " & FilterVar(UCase(arrColVal(6)), "''", "S") 
    lgStrSQL = lgStrSQL & "   AND BP_CD		  = " & FilterVar(UCase(arrColVal(2)), "''", "S")    
   
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  S_BP_APPLY_EXCH_DTL_KO441"
    lgStrSQL = lgStrSQL & " WHERE EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND CURRENCY	  = " & FilterVar(UCase(arrColVal(3)), "''", "S")  
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
	lgStrSQL = "DELETE  S_BP_APPLY_EXCH_HDR_KO441"
    lgStrSQL = lgStrSQL & " WHERE EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND CURRENCY	  = " & FilterVar(UCase(arrColVal(3)), "''", "S")  
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete2(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  S_BP_APPLY_EXCH_DTL_KO441"
    lgStrSQL = lgStrSQL & " WHERE EXCH_YYYYMM = " & FilterVar(UCase(arrColVal(2)), "''", "S")
    lgStrSQL = lgStrSQL & "   AND CURRENCY	  = " & FilterVar(UCase(arrColVal(3)), "''", "S")  
    lgStrSQL = lgStrSQL & "   AND BP_CD		  = " & FilterVar(UCase(arrColVal(4)), "''", "S")  
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
		
End Sub

'==================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
        Case "MV"
                 Call DisplayMsgBox("800446", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MX"
                 Call DisplayMsgBox("800350", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MY"
                 Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "MZ"
                 Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)     '이 기간에 대해 이미 입력된 기간근태사항이 있습니다 
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

%>

<Script Language="VBScript">
	With parent
	
	    Select Case "<%=lgOpModeCRUD %>" 
		   Case "<%=UID_M0001%>"       
			
			    If Len("<%=lgstrData%>") > 0 then
				.ggoSpread.Source     = .frm1.vspdData
				.ggoSpread.SSShowData "<%=lgstrData%>"

	'			.frm1.HGrid3.value = "<%=Request("Grid3SlipNo")%>"
			
				.SetSpreadLockAfterQuery -1, -1
				
				.DBQueryOk
				End if
				
			Case "Grid2"       
			
			    If Len("<%=lgstrData%>") > 0 then
				.ggoSpread.Source     = .frm1.vspdData2
				.ggoSpread.SSShowData "<%=lgstrData%>"

	'			.frm1.HGrid3.value = "<%=Request("Grid3SlipNo")%>"
			
				.SetSpreadLockAfterQuery2 -1, -1
				
'				.DBQuery2Ok
				End if
					   
	       Case "<%=UID_M0002%>"                                                         '☜ : Save
	          If Trim("<%=lgErrorStatus%>") = "NO" Then
	          
					If "<%=CheckFlg%>" = False Then
						parent.DbQuery2 "<%=Grid2Key1%>","<%=Grid2Key2%>"
					Else
						Parent.DBSaveOk
					End If
	             
	          Else
	             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
	          End If     
	    End Select    
	    
    End With
       
</Script>	

