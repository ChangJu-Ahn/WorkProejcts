<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1211MB1
'*  4. Program Name         : ����ó����к��� 
'*  5. Program Desc         : ����ó����к��� 
'*  6. Component List       : PM1G131.cMAssignQuotaRatebySppl
'*  7. Modified date(First) : 2003/01/09
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%	
call LoadBasisGlobalInf()

    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag                  '�� : DBAgent Parameter ���� 
    Dim rs0, rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
    Dim lgPageNo
	Dim intARows
	Dim intTRows
    Dim arrRsVal(11)
	intARows=0
	intTRows=0
 
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message

    lgOpModeCRUD  = Request("txtMode") 
							                                              '��: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	lgStrPrevKey   = Request("lgStrPrevKey")
	
	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
	Dim arrVal(3)
	Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(2,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                        '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "M1211MA201" 											' header
    UNISqlId(1) = "M2111QA302"								              '����� 
	UNISqlId(2) = "M2111QA303"											  'ǰ���     
     
    UNIValue(1,0) = "" & FilterVar("zzzzz", "''", "S") & ""
    UNIValue(2,0) = "" & FilterVar("zzzzzzzzzz", "''", "S") & ""
    UNIValue(2,1) = "" & FilterVar("zzzzz", "''", "S") & ""
    
    UNIValue(0,0) = "^" 
    
    '����                    
    If Trim(Request("txtPlantCd")) <> "" Then
		UNIValue(0,1) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	    UNIValue(1,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	Else 
	    UNIValue(0,1) = "|"
	End If
	
	'ǰ��                    
    If Trim(Request("txtitemcd")) <> "" Then
		UNIValue(0,2) = "  " & FilterVar(Trim(UCase(Request("txtitemcd"))), " " , "S") & "  "
	    UNIValue(2,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	    UNIValue(2,1) = "  " & FilterVar(Trim(UCase(Request("txtitemcd"))), " " , "S") & "  "
	Else 
	    UNIValue(0,2) = "|"
	End If
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    Dim FalsechkFlg
    
    FalsechkFlg = False    
    
    '============================= �߰��� �κ� =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If

    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtitemcd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		   rs0.Close
		   Set rs0 = Nothing
		   Exit Sub
    	End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
'		Call DisplayMsgBox("173132", vbOKOnly, "����ó����к�", "", I_MKSCRIPT)
		Call DisplayMsgBox("970000", vbOKOnly, "����ó����к�", "", I_MKSCRIPT)


		rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
   
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
    Response.Write "	.frm1.txtPlantNm.value = """ & Trim(UCase(ConvSPChars(arrRsVal(1))))              	& """" & vbCr
	Response.Write "	.frm1.txtitemNm.value = """ & Trim(UCase(ConvSPChars(arrRsVal(3))))              	& """" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
    Response.Write "	.lgPageNo  = """ & lgPageNo   & """" & vbCr  
	Response.Write "	.frm1.hdnPlant.value = """ & Trim(UCase(ConvSPChars(Request("txtPlantCd"))))              	& """" & vbCr
	Response.Write "	.frm1.hdnItem.value = """ & Trim(UCase(ConvSPChars(Request("txtItemCd"))))              	& """" & vbCr
    Response.Write "    .DbQueryOk " & intARows & "," & intTRows & vbCr 
	Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr        

End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If

   iLoopCount = -1
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
   Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(0))	    '1 ���� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(1))		'2 ����� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(2))		'3 ǰ�� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(3))		'4 ǰ��� 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(4))		'5 ǰ��԰� 
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount + 1                             

        If iLoopCount < C_SHEETMAXROWS_D Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = istrData	
		   istrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
	Loop
	intARows = iLoopCount+1
	istrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

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
	Const C_ParentRowNo   = 5
	Const C_ChildRowNo   = 6
	
	Dim OBJ_PM1G131
	Dim lgIntFlgMode
	Dim iStrCommandSent
	Dim I2_ParentRowNo
	Dim I3_b_PlantCd
	Dim I4_b_ItemCd
	Dim I5_m_supplier_item_by_plant
	Dim I6_childRowNo
    Dim arrVal, arrTemp	, arrTemp1
	Dim LngMaxRow,LngRow,LngRow2, lGrpCnt
	Dim iErrorPosition
    Dim lRow
    Dim iRowsep,iColsep
	Dim Zsep
	Dim iStrSpread

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
	Zsep = "@"
   arrTemp1 = Split(itxtSpread, Zsep)

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
	
    '1�Ǿ� ó���Ѵ� 
    For LngRow = 1 To UBound(arrTemp1,1)
        
		iStrSpread = ""
		arrTemp = Split(arrTemp1(LngRow-1), gRowSep)
		Response.Write UBound(arrTemp,1) & vbCr
		
		For LngRow2 = 1 To UBound(arrTemp,1)
			arrVal = Split(arrTemp(LngRow2-1), gColSep)
	    
			I2_ParentRowNo = arrVal(C_ParentRowNo)
			I3_b_PlantCd = arrVal(C_PlantCd)
			I4_b_ItemCd = arrVal(C_ItemCd)
			I6_childRowNo = arrVal(C_ChildRowNo)
	    
			iStrSpread = iStrSpread & Trim(arrVal(C_CommandSent)) & gColSep & Trim(arrVal(C_SpplCd)) & gColSep	& _
						 Trim(arrVal(C_Quota_Rate)) & gColSep  & Trim(arrVal(C_ChildRowNo)) & gRowSep
		Next
		    
	    Set OBJ_PM1G131 = Server.CreateObject("PM1G131.cMAssignQuotaRatebySppl")

		If CheckSYSTEMError(Err,True) = true then 
		    Set OBJ_PM1G131 = Nothing		
		    Exit Sub
		End If
		
		Call OBJ_PM1G131.m_Maint_Quota_by_Sppl(gStrGlobalCollection, _
										    I3_b_PlantCd, _
										    I4_b_ItemCd, _
					                        iStrSpread, _
					                        iErrorPosition)

		If CheckSYSTEMError2(Err, True, LngRow & "-" & iErrorPosition & "��","","","","") = True Then
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
		   Set OBJ_PM1G131 = Nothing
		End If


    Next
    
	Set OBJ_PM1G131 = Nothing 

	Response.Write "<Script language=vbs> " & vbCr         
	Response.Write " Parent.DbSaveOk "      & vbCr							'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write "</Script> "   & vbCr	  
        
    
End Sub    

%>
