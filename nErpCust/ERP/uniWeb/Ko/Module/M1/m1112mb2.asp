<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1112MB2
'*  4. Program Name         : ����ó���ܰ����(Multi)
'*  5. Program Desc         : ����ó���ܰ����(Multi)
'*  6. Component List       : PM1G121.cMMntSpplItemPriceS
'*  7. Modified date(First) : 2002/12/061
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
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

    Dim lgOpModeCRUD
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim GroupCount  
    Dim lgPageNo
	Dim iErrorPosition
	Dim arrRsVal(11)
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
	
    lgOpModeCRUD  = Request("txtMode") 
											                                              '��: Read Operation 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
   On Error Resume Next

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

	Call FixUNISQLData()
	Call QueryData()	
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
    Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "    .frm1.vspdData.Redraw = False   "                  & vbCr   
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & ",""F""" & vbCr
    Response.Write "     Call .ReFormatSpreadCellByCellByCurrency(.Frm1.vspdData,-1,-1,.C_Curr,.C_Cost,""C"" ,""I"",""X"",""X"")" & vbCr
    Response.Write "	.frm1.hdnPlantCd.value		= """ & ConvSPChars(Request("txtPlantCd"))		& """" & vbCr  
    Response.Write "	.frm1.hdnitemcd.value		= """ & ConvSPChars(Request("txtitemcd"))		& """" & vbCr  
    Response.Write "	.frm1.hdnSuppliercd.value	= """ & ConvSPChars(Request("txtSuppliercd"))   & """" & vbCr  
    Response.Write "	.frm1.hdnAppFrDt.value		= """ & ConvSPChars(Request("txtAppFrDt"))		& """" & vbCr  
    Response.Write "	.frm1.hdnAppToDt.value		= """ & ConvSPChars(Request("txtAppToDt"))		& """" & vbCr  
    Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr  
    Response.Write "	.DbQueryOk " & vbCr 
    Response.Write  "   .frm1.vspdData.Redraw = True " & vbCr   
    Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr    
End Sub    
	    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
	Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(3,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "m1112ma201" 											' header
	UNISqlId(1) = "M2111QA302"								              '����� 
	UNISqlId(2) = "M2111QA303"											  'ǰ���  
	UNISqlId(3) = "M3111QA102"								              '�ŷ�ó�� 
	
	UNIValue(1,0) = "" & FilterVar("zzzzz", "''", "S") & ""
    UNIValue(2,0) = "" & FilterVar("zzzzzzzzzz", "''", "S") & ""
    UNIValue(2,1) = "" & FilterVar("zzzzz", "''", "S") & ""
    UNIValue(3,0) = "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    
    If Trim(Request("txtPlantCd")) <> ""  Then
		strVal =  " AND A.PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd")), " " , "S") 
		UNIValue(1,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	End If
	
	If Trim(Request("txtitemcd")) <> "" Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Trim(Request("txtitemcd")), " " , "S") 
		UNIValue(2,0) = "  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & "  "
	    UNIValue(2,1) = "  " & FilterVar(Trim(UCase(Request("txtitemcd"))), " " , "S") & "  "
	End If
	
	If Trim(Request("txtSuppliercd")) <> "" Then
		strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(Request("txtSuppliercd")), " " , "S") 
		UNIValue(3,0)	= " " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & " "
	End If
	
	If Len(Trim(Request("txtAppFrDt"))) Then
		strVal = strVal & " AND A.VALID_FR_DT >= " & FilterVar(UNIConvDate(Request("txtAppFrDt")), "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtAppToDt"))) Then
		strVal = strVal & " AND A.VALID_FR_DT <= " & FilterVar(UNIConvDate(Request("txtAppToDt")), "''", "S") & ""		
	End If
	
	strVal = strVal & " ORDER BY A.PLANT_CD,  A.ITEM_CD, A.PUR_CUR, A.VALID_FR_DT, A.BP_CD"

    UNIValue(0,0) = strval			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2,rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Response.end
    End If 
    
    '============================= �߰��� �κ� =====================================================================
    Dim FalsechkFlg
    FalsechkFlg = False    
    
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
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       rs0.Close
		   Set rs0 = Nothing
		   Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "With parent.frm1" & vbCr
			Response.Write "	.txtItemCd.focus" & vbCr
			Response.Write "End With" & vbCr
			Response.Write "</Script>" & vbCr
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtSupplierCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("171200", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call parent.SetToolBar(""111011010011111"") " & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.end
    ELSE
        Call  MakeSpreadSheetData()
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
	Dim PvArr
	
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
   
   iLoopCount = -1
   ReDim PvArr(C_SHEETMAXROWS_D - 1)

   Do while Not (rs0.EOF Or rs0.BOF)
		
        iLoopCount =  iLoopCount + 1
        iRowStr = ""

		iRowStr = Chr(11) & ConvSPChars(Trim(rs0(0)))
		iRowStr = iRowStr &	Chr(11) & ""
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(1)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(2)))
		iRowStr = iRowStr &	Chr(11) & ""
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(3)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(4)))
		iRowStr = iRowStr &	Chr(11) & ""
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(5)))
		iRowStr = iRowStr &	Chr(11) & ""
		iRowStr = iRowStr &	Chr(11) & UNIDateClientFormat(Trim(rs0(6)))
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(7)))
		iRowStr = iRowStr &	Chr(11) & ""
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(8)))
		iRowStr = iRowStr &	Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0(9), 0)
		'�ܰ����� �߰� 
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(10)))
		iRowStr = iRowStr &	Chr(11) & ""
		iRowStr = iRowStr &	Chr(11) & ConvSPChars(Trim(rs0(11)))
		iRowStr = iRowStr &	Chr(11) & iLngMaxRow + iLoopCount + 1                            
		iRowStr = iRowStr &	Chr(11) & Chr(12)                          
        
        If iLoopCount < C_SHEETMAXROWS_D Then
	        PvArr(iLoopCount) = iRowStr
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
	Loop
	
	istrData = Join(PvArr, "")
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
Sub subBizSaveMulti()															'��: ���� ��û�� ���� 
    On Error Resume Next 		
    Err.Clear														'��: Protect system from crashing

    Dim iPM1G121
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim iDCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")

	'Call ServerMesgBox(itxtSpread , vbInformation, I_MKSCRIPT)
    
    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   

	Set iPM1G121 = Server.CreateObject("PM1G121.cMMntSpplItemPriceS")    

    If CheckSYSTEMError(Err,True) = True Then Exit Sub
    
	Call iPM1G121.M_MAINT_MULTI_SPPL_ITEM_PRICE_SVR(gStrGlobalCollection, _
													itxtSpread, _
													iErrorPosition)
	
	If CheckSYSTEMError2(Err,True, iErrorPosition(0) & "��:" ,"","","","") = true then 		

	   Set iPM1G121 = Nothing
  		If iErrorPosition(0) <> "" Then
			Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call parent.SheetFocus(" & iErrorPosition(0) & ")" & vbCrLF
			Response.Write "Call parent.SetToolBar(""111011110011111"") " & vbCrLF

			Response.Write "</Script>" & vbCrLF
		End If
		Response.End
	End If		
   
    Set iPM1G121 = Nothing                                                   '��: Unload Comproxy  
        
	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
    Response.Write ".DbSaveOk "      & vbCr						
    Response.Write "End With " & vbCr
    Response.Write "</Script> "    
End Sub	

%>
