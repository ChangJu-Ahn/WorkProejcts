<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111mb201
'*  4. Program Name         : 업체지정 
'*  5. Program Desc         : 업체지정 
'*  6. Component List       : PM1G131.cMAssignQuotaRatebySppl
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%	
Call HideStatusWnd
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("*","*","NOCOOKIE","MB")

    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim istrData
	Dim iStrPrNo
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
    
    Dim arrRsVal(11)
	Const C_SHEETMAXROWS_D  = 100
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	lgPageNo1       = UNICInt(Trim(Request("lgPageNo1")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist    = "No"
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	sRow           = CLng(Request("lRow"))
	lglngHiddenRows = CLng(Request("lglngHiddenRows"))
	iStrOgrCd	=  Request("txtOgrCd")
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
    Redim UNIValue(0,0)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M2111MA201" 											' header
 
    iStrPrNo = Trim(Request("txtPrNo"))
    
	UNIValue(0,0) = "  " & FilterVar(Trim(UCase(Request("txtPrNo"))), " " , "S") & "  "
	
	    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
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
		'Call DisplayMsgBox("172400", vbOKOnly, iStrPrNo, "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
    Response.Write "	.lgPageNo1  = """ & lgPageNo1   & """" & vbCr  
    Response.Write "    .lglngHiddenRows(" & sRow - 1 & ") = """ & MaxRow2 & """" & vbCr  

    'Response.Write "	.frm1.vspddata2.Row  =  .frm1.vspddata2.ActiveRow "			  & vbCr
    'Response.Write "	.frm1.vspddata2.Col  = Parent.C_PlanDt "       & vbCr
    'Response.Write "	.checkdt(.frm1.vspddata2.ActiveRow) "       & vbCr

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
    Dim PvArr

	lgDataExist    = "Yes"
	
	If CLng(lgPageNo1) > 0 Then
	    rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If
	iLoopCount = 0

   ReDim PvArr(C_SHEETMAXROWS_D - 1)
   Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(0))		                        '1 공급처 
        iRowStr = iRowStr & Chr(11) & ""                                                '2 공급처 팝업 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(1))								'3 공급처명 
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(2),ggExchRate.DecPoint,0)	'4 배분비율 
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(3),ggQty.DecPoint,0)		'5 배부량 
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(4))						'6 발주예정일 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(5))								'7 구매그룹 
        iRowStr = iRowStr & Chr(11) & ""												'8 구매그룹팝업 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(6))		                        '9 구매그룹명 
        iRowStr = iRowStr & Chr(11) & iStrOgrCd												'10 구매조직 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPrNo)	                            '13 상위 prno
        iRowStr = iRowStr & Chr(11) & sRow
		iRowStr = iRowStr & Chr(11) & CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1) + iLoopCount
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             

        If iLoopCount - 1 < C_SHEETMAXROWS_D Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount-1) = istrData	
		   istrData = ""
        Else
           lgPageNo1 = lgPageNo1 + 1
           Exit Do
        End If
        rs0.MoveNext
   Loop
	istrData = Join(PvArr, "")

    If CLng(lgPageNo1) > 0 Then
	 	MaxRow2 = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo1) + iLoopCount
	 	
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

	Const C_CommandSent = 0										
	Const C_PlantCd     = 1                              
	Const C_ItemCd      = 2
    Const C_SpplCd      = 3
    Const C_Quota_Rate  = 4                                   
	Const C_Row         = 5
	
	Dim iPM1G131
	Dim lgIntFlgMode
	Dim iStrCommandSent
	Dim I1_b_company
	Dim I2_m_config_process
	Dim I3_b_PlantCd
	Dim I4_b_ItemCd
	
	Dim I5_m_supplier_item_by_plant
	
    Dim arrVal, arrTemp	
	Dim LngMaxRow,LngRow,lGrpCnt
	Dim iErrorPosition
	Dim iStrSpread
    Dim oldplant,olditem   
    Dim newplant,newitem  
	Dim lRow
    Dim iRowsep,iColsep
	
    lgIntFlgMode = CInt(Request("txtFlgMode"))

	LngMaxRow = CLng(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
    arrTemp = Split(Request("txtSpread"), gRowSep)

    lGrpCnt = 0

    For LngRow = 1 To LngMaxRow

		lGrpCnt = lGrpCnt +1														'☜: Group Count
		lRow = lRow + 1
		 
		arrVal = Split(arrTemp(LngRow-1), gColSep)
    
    	newplant = Trim(arrVal(C_PlantCd)) 
                
        newitem =  Trim(arrVal(C_ItemCd)) 
    
   
        if LngRow = 1 then
            oldplant = newplant
            olditem  = newitem
            I3_b_PlantCd = Trim(arrVal(C_PlantCd)) 
	        I4_b_ItemCd  = Trim(arrVal(C_ItemCd))
        end if

		if Trim(UCase(newplant)) <> Trim(UCase(oldplant)) or Trim(UCase(olditem)) <> Trim(UCase(newitem)) or LngRow = LngMaxRow then
			
		    if LngRow = LngMaxRow then
		        iStrSpread = iStrSpread & Trim(arrVal(C_CommandSent)) & gColSep
		        iStrSpread = iStrSpread & Trim(arrVal(C_SpplCd)) & gColSep	
		        iStrSpread = iStrSpread & Trim(arrVal(C_Quota_Rate)) & gColSep	
		        iStrSpread = iStrSpread & Trim(arrVal(C_Row)) & gRowSep
		    end if
		    
		    Set iPM1G131 = Server.CreateObject("PM1G131.cMAssignQuotaRatebySppl")
		    
		    If CheckSYSTEMError(Err,True) = true then 
			    Set iPM1G131 = Nothing		
			    Exit Sub
		    End If
		    
		    Call iPM1G131.m_Maint_Quota_by_Sppl(gStrGlobalCollection, _
											    I3_b_PlantCd, _
											    I4_b_ItemCd, _
						                        iStrSpread, _
						                        iErrorPosition)
		    
		    
		    If CheckSYSTEMError2(Err, True, iErrorPosition ,"","","","") = True Then
        %>	
	       <Script Language=vbscript>
               Dim msgCreditlimit
               With parent	

                   msgCreditlimit = .Parent.DisplayMsgBox("17A016", .Parent.VB_YES_NO,"X", "X")
	           
	               If msgCreditlimit = vbYes Then 

                   else
                   	   .DbSaveOk
                  end if
              End With
           </Script> 
       <%                           
		Else
		   Set iPM1G131 = Nothing
		End If
		    
		    
		    oldplant = newplant
            olditem  = newitem
            I3_b_PlantCd = Trim(arrVal(C_PlantCd)) 
	        I4_b_ItemCd  = Trim(arrVal(C_ItemCd))
			iStrSpread = ""
	    end if       
	    
        iStrSpread = iStrSpread & Trim(arrVal(C_CommandSent)) & gColSep
        iStrSpread = iStrSpread & Trim(arrVal(C_SpplCd)) & gColSep	
		iStrSpread = iStrSpread & Trim(arrVal(C_Quota_Rate)) & gColSep	
		iStrSpread = iStrSpread & Trim(arrVal(C_Row)) & gRowSep
    next 
   
   Set iPM1G131 = Nothing 

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
