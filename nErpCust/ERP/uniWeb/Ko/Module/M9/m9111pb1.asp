<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M9111PB1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2002/12/10
'*  7. Modified date(Last)  : 
'*                            
'*  8. Modifier (First)     : Oh Chang Won
'*  9. Modifier (Last)      : 
'*                            
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncServer.asp" -->
<%
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3      '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort 종류 
Dim istrData


Dim PotypeNm														'☜ : 발주형태명 저장 
Dim GroupNm										   				    '☜ : 구매그룹명 저장 
Dim SupplierNm														'☜ : 공급처명 저장 

    Call HideStatusWnd 
    Call LoadBasisGlobalInf
    
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))             '☜ : 한번에 가져올수 있는 데이타 건수 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

   iLoopCount = 0
   Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
        iRowStr = ""
 
 		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(0))           '이동요청번호 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(1))		    '이동유형 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(2))		    '이동유형명 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(8))			'확정여부 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(3))		    '공급창고 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(4))		    '공급창고명 
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0(7))   '등록일 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(5))	        '구매그룹                                '품목규격 '8	
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(6))	        '구매그룹명 
		'iRowStr = iRowStr & Chr(11) & ""							'14								'27
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             
      
        If iLoopCount - 1 < lgMaxCount Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs0.MoveNext
  Loop
    
    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    SetConditionData = True
    
	If Not(rs1.EOF Or rs1.BOF) Then
		PotypeNm = rs1("PO_TYPE_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtPotypeCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "이동유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    SetConditionData = False
		End If
	End If   	
	
	If Not(rs2.EOF Or rs2.BOF) Then
		SupplierNm = rs2("BP_NM")
		Set rs2 = Nothing
	Else
		Set rs2 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공급창고", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    SetConditionData = False
		End If
	End If   	
	
	If Not(rs3.EOF Or rs3.BOF) Then
		GroupNm = rs3("PUR_GRP_NM")
		Set rs3 = Nothing
	Else
		Set rs3 = Nothing
		If Len(Request("txtGroupCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    SetConditionData = False
		End If
	End If   	

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	dim sTemp
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,6)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
	strVal = ""
    UNISqlId(0) = "M9111PA101"
    UNISqlId(1) = "s0000qa020"
    UNISqlId(2) = "s0000qa002"
    UNISqlId(3) = "s0000qa019"
    
    '--- 2004-08-19 by Byun Jee Hyun for UNICODE
    UNIValue(1,0) = FilterVar("zzzzz", "''", "S")
    UNIValue(2,0) = FilterVar("zzzzzzzzzz", "''", "S")
    UNIValue(3,0) = FilterVar("zzzz", "''", "S")
    
    sTemp = "1"
    
	If Len(Trim(Request("txtFrPoDt"))) Then
		If UNIConvDate(Request("txtFrPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrPoDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If
	
	If Len(Trim(Request("txtToPoDt"))) Then
		If UNIConvDate(Request("txtToPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToPoDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If

    
    UNIValue(0,0) = "^"
    '이동형태                    
    If Trim(Request("txtPotypeCd")) <> "" Then
		UNIValue(0,1) = "  " & FilterVar(UCase(Request("txtPotypeCd")), "''", "S") & "  "
	    UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtPotypeCd"))), " " , "S")
	Else 
	    UNIValue(0,1) = "|"
	End If


	'공급창고 
    If Trim(Request("txtSupplierCd"))  <> "" Then
		UNIValue(0,2) = "  " & FilterVar(UCase(Request("txtSupplierCd")), "''", "S") & "  "
	    UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S")
	Else
	    UNIValue(0,2) = "|"
	End If

    '등록일 
    If Trim(Request("txtFrPoDt")) <> "" Then
		UNIValue(0,3) =  "  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
    Else
        UNIValue(0,3) = "|"
	End If
			
    If Trim(Request("txtToPoDt")) <> "" Then
		UNIValue(0,4) =  "  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
    Else
        UNIValue(0,4) = "|"
	End If

	'구매구룹 
	If Trim(Request("txtGroupCd")) <> "" Then
		UNIValue(0,5) =  "  " & FilterVar(UCase(Request("txtGroupCd")), "''", "S") & "  "
	    UNIValue(3,0) = FilterVar(Trim(UCase(Request("txtGroupCd"))), " " , "S")
	Else
	    UNIValue(0,5) = "|"
	End If

    '확정여부			
    If Trim(Request("txtRadio")) = "Y" then
	    UNIValue(0,6) =  " " & FilterVar("Y", "''", "S") & "  "
	ElseIf Trim(Request("txtRadio")) = "N" then
	    UNIValue(0,6) =  " " & FilterVar("N", "''", "S") & "  "
	Else
	    UNIValue(0,6) =  "|"
	End If


'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
'	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
'	UNIValue(0,1) = strVal & " ORDER BY A.PO_NO DESC"

    'UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 



     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
	If SetConditionData = False Then Exit Sub

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()

    End If  
End Sub

%>
<Script Language=vbscript>
    With parent
		.frm1.txtPotypeNm.value = "<%=ConvSPChars(PotypeNm)%>"
		.frm1.txtSupplierNm.value = "<%=ConvSPChars(SupplierNm)%>"
		.frm1.txtGroupNm.value = "<%=ConvSPChars(GroupNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnPotype.value	= "<%=ConvSPChars(Request("txtPotypeCd"))%>"
				.frm1.hdnSupplier.value	= "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.frm1.hdnFrDt.value 	= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
				.frm1.hdnToDt.value 	= "<%=ConvSPChars(Request("txtToPoDt"))%>"
				.frm1.hdnGroup.value	= "<%=ConvSPChars(Request("txtGroupCd"))%>"
				.frm1.hdtxtRadio.value	= "<%=ConvSPChars(Request("txtRadio"))%>"
				.frm1.hdnRetFlg.value	= "<%=ConvSPChars(Request("hdnRetFlg"))%>"					
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data 
			
			.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
