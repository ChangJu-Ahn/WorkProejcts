<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%	
Call HideStatusWnd
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111mb201
'*  4. Program Name         : 업체지정 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'							  
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/03/03
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'* 14. Business Logic of m2111ma2(업체지정)
'**********************************************************************************************
    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0						'☜ : DBAgent Parameter 선언 
    Dim istrData
	Dim iStrPrNo, iStrPoNo, iStrPoSeq	
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim index     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수     
    Dim lgDataExist
    Dim lgPageNo1
    Dim sRow
    Dim lglngHiddenRows
    DIM MaxRow2
    Dim MaxCount
    Dim iStrOgrCd
    Dim lgStrNextFlag
    Dim lgStrNextKey
	
    Dim arrRsVal(11)
	Const C_SHEETMAXROWS_D  = 100										'☜ : MA단의 C_SHEETMAXROWS_D 와 꼭 일치.
 
    On Error Resume Next  
    Err.Clear             

    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                      
             Call SubBizSaveMulti()
        Case "LookUpItemByPlant"			
			 Call LookUpItemByPlant
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	Err.Clear
	
	lgPageNo1       = UNICInt(Trim(Request("lgPageNo1")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist     = "No"
	iLngMaxRow      = CLng(Request("txtMaxRows"))
	sRow            = CLng(Request("lRow"))
	lglngHiddenRows = CLng(Request("lglngHiddenRows"))
	lgStrNextFlag	= Trim(Request("txtNextFlag"))
	
	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,3)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M3111MA602" 											' header
 
    iStrPrNo = Trim(Request("txtPrNo"))
    iStrPoNo = Trim(Request("txtPoNo"))
    iStrPoSeq = Trim(Request("txtPoSeq"))
    
    UNIValue(0,0) = "  " & FilterVar(UCase(Request("txtPrNo")), "''", "S") & "  "
	UNIValue(0,1) = "  " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & "  "
	UNIValue(0,2) = "  " & FilterVar(Request("txtPoSeq"), "''", "S") & "  "
	UNIValue(0,3) = "  " & FilterVar(Request("lgStrResvdSeqNo"), "''", "S") & "  "
	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    Dim FalsechkFlg
    
    FalsechkFlg = False    

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write "Call parent.ResetToolBar(parent.frm1.vspdData.activeRow,1)" & vbCr
	    Response.Write "</Script>"		& vbCr        
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData1 "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     iStrCopyData" & vbCr	
    Response.Write "	.lgPageNo1  = """ & lgPageNo1   & """" & vbCr  
    Response.Write "    .lgStrPrevKeyM(" & sRow - 1 & ") = """ & lgStrNextKey & """" & vbCr  
    Response.Write "    .lglngHiddenRows(" & sRow - 1 & ") = """ & MaxRow2 & """" & vbCr  
    'Response.Write "msgbox """ & MaxRow2 & """" & vbCr  
    Response.Write "    .DbQueryOk2(" & MaxCount & ")" & vbCr
	Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr        

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

	Dim iLoopCount                                                                     
	Dim iRowStr
	Dim ColCnt
	DIM i
	Dim iStrTmp
	
	lgDataExist    = "Yes"
	
	If CLng(lgPageNo1) > 0 Then
		'주의 : 이프로그램에 한해 레코드셋 MoveNext 처리를 안하고 코드상의 NextKey처리를 함.
	    'rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If
	
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	Dim iStrCopyData" & vbCr
	iLoopCount = 0
    Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("item_cd"))
        iRowStr = iRowStr & Chr(11) & ""                         
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("item_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("spec"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sppl_type"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sl_cd"))
        iRowStr = iRowStr & Chr(11) & ""                     
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sl_nm"))
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("resrv_dt"))
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("resrv_qty"),ggExchRate.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & ""                     
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("issue_qty"),ggQty.DecPoint,0)	
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("tot_bk_flush_qty"),ggQty.DecPoint,0)	
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("reqmt_unit"))
        
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("resvd_seq_no"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("resrv_sts"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sub_seq_no"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("reqmt_no"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPrNo)	       
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPoNo)	       
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPoSeq)
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sppl_type"))
	%>
        

			iStrCopyData	   = iStrCopyData & "<%=iRowStr%>"	
			.frm1.vspdData.row =   .frm1.vspdData.ActiveRow
			
			.frm1.vspdData.col = .C_PoQty 
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text
			
			.frm1.vspdData.col = .C_PoUnit 
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text 
			
			.frm1.vspdData.col =  .C_PoDt	                
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text 
				
			.frm1.vspdData.col =  .C_RcptQty	                
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text
			
			.frm1.vspdData.col = .C_TrackingNo 
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text 
			
			.frm1.vspdData.col = .C_PlantCd
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text 
			
			.frm1.vspdData.col = .C_SpplCd
			iStrCopyData = iStrCopyData  & Chr(11) & .frm1.vspdData.text 
	
	<%
		iRowStr = Chr(11) & ConvSPChars(rs0("item_cd"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sppl_type"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sl_cd"))
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("resrv_qty"),ggExchRate.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("resrv_dt"))
        iRowStr = iRowStr & Chr(11) & sRow
		iRowStr = iRowStr & Chr(11) & CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1) + iLoopCount
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             

        If iLoopCount - 1 < C_SHEETMAXROWS_D Then
           istrData = iRowStr & Chr(11) & Chr(12)
        Else
           lgStrNextKey = ConvSPChars(rs0("resvd_seq_no")) 
           lgPageNo1 = lgPageNo1 + 1
           Exit Do
        End If
		
		Response.Write "	iStrCopyData = iStrCopyData & """ & istrData & """" & vbCr
		
        rs0.MoveNext
   Loop
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr
	
    If CLng(lgPageNo1) > 0 Then
		If Trim(lgStrNextFlag) Then
	 		MaxRow2 = (CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1)) + iLoopCount 
	 	Else
	 		MaxRow2 = (CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1)) 
	 	End If	
	Else
		MaxRow2 = CLng(iLoopCount)
	End If
	
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo1 = 0
    End If
    
    MaxCount = iLoopCount
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
	Err.Clear	

	Dim iOBJ_PM2G521
	Dim lgIntFlgMode
	
    Dim arrTemp	
    Dim LngMaxRow,LngRow,lGrpCnt
	Dim iErrorPosition
	Dim iStrSpread
    Dim lRow
    Dim iRowsep,iColsep
	Dim lgTransSep
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii
             
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

    lgIntFlgMode = CInt(Request("txtFlgMode"))
	
	lgTransSep = "Ð"
	LngMaxRow = CLng(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    arrTemp = Split(itxtSpread, lgTransSep)

    lGrpCnt = 0
	
	If ubound(arrTemp,1) > 0 Then
	
		Set iOBJ_PM2G521 = Server.CreateObject("PM2G521.cMMaintAdjustChildS")
			    
		If CheckSYSTEMError(Err,True) = true then 
		    Exit Sub
		End If
		
		For LngRow = 0 To UBound(arrTemp,1) -1
		
			Call iOBJ_PM2G521.M_MAINT_ADJUST_CHILD_SVR(gStrGlobalCollection, arrTemp(LngRow), iErrorPosition) 
			    
			If CheckSYSTEMError2(Err, True, iErrorPosition & " - " & LngRow+1 & "행" ,"","","","") = True Then
	     %>	
		       <Script Language=vbscript>
	            Dim msgCreditlimit
	            With parent	
					
	               msgCreditlimit = .Parent.DisplayMsgBox("17A016", .Parent.VB_YES_NO,"X", "X")
		           
		           If msgCreditlimit = vbYes Then 
						Err.Clear
	               Else
	                 .DbSaveOk
	               End if
	           End With
	        </Script> 
	    <%                           
			Err.Clear
			End If
			
		Next 
	
		If Not(iOBJ_PM2G521 is Nothing) Then
			Set iOBJ_PM2G521 = Nothing                                                   '☜: Unload Comproxy
		End If	
	
   End IF
   		
   Response.Write "<Script language=vbs> " & vbCr         
   Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
   Response.Write "</Script> "   & vbCr	  
       
End Sub    

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

%>
