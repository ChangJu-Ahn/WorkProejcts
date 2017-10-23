<%'======================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출관리 
'*  3. Program ID           : s5111pb1
'*  4. Program Name         : 매출채권번호 Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/08/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next														'☜: 
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 30							            '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
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
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim iStrWhere2,iStrWhere3
	
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,5)

	If Len(Trim(Request("txtSoNo"))) OR Len(Trim(Request("txtDnNo"))) Then
	    UNISqlId(0) = "S5111PA102"
	Else
	    UNISqlId(0) = "S5111PA101"
	End If
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	iStrWhere2 = ""
	iStrWhere3 = ""
	iStrWhere4 = ""
	
	UNIValue(0,1) = Request("txtExceptflag")
	
    If Len(Trim(Request("txtSoldToParty"))) Then
		iStrWhere2 = iStrWhere2 & " AND BH.sold_to_party =  " & FilterVar(Request("txtSoldToParty"), "''", "S") & ""		'주문처 
	End If

    If Len(Trim(Request("txtFromDt"))) Then
		iStrWhere2 = iStrWhere2 & " AND BH.bill_dt >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""						'시작일 
	End If

    If Len(Trim(Request("txtToDt"))) Then
		iStrWhere2 = iStrWhere2 & " AND BH.bill_dt <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""							'종료일 
	End If

    If Len(Trim(Request("txtBillType"))) Then
		iStrWhere2 = iStrWhere2 & " AND BH.bill_type =  " & FilterVar(Request("txtBillType"), "''", "S") & ""				'매출채권유형 
	End If

    If Len(Trim(Request("txtSalesGrp"))) Then
		iStrWhere2 = iStrWhere2 & " AND BH.sales_grp =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""				'영업그룹 
	End If

    If Len(Trim(Request("txtPostFlag"))) Then
		iStrWhere2 = iStrWhere2 & " AND BH.post_flag =  " & FilterVar(Request("txtPostFlag"), "''", "S") & ""				'확정여부 
	End If
	
    If Len(Trim(Request("txtDnNo"))) Then
		iStrWhere3 = iStrWhere3 & " AND BD.dn_no =  " & FilterVar(Request("txtDnNo"), "''", "S") & ""				'출고번호 
    End If

    If Len(Trim(Request("txtSoNo"))) Then
		iStrWhere3 = iStrWhere3 & " AND BD.so_no =  " & FilterVar(Request("txtSoNo"), "''", "S") & ""				'수주번호 
		iStrWhere4 = iStrWhere4 & " OR BH.so_no =  " & FilterVar(Request("txtSoNo"), "''", "S") & ""				'수주번호 
    End If
    
	UNIValue(0,2) = iStrWhere2
	UNIValue(0,3) = iStrWhere3
	UNIValue(0,4) = iStrWhere4
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.DbQueryOk" & vbCr
		Response.Write "</Script>" & vbCr
    Else    
        Call  MakeSpreadSheetData()
		If lgPageNo = "1" Then Call SetConditionData()
        Call WriteResult()
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".txtHSoldToParty.value	= """ & ConvSPChars(Request("txtSoldToParty")) & """" & vbCr
	Response.Write ".txtHFromDt.value = """ & Request("txtFromDT") & """" & vbCr
	Response.Write ".txtHToDt.value	= """ & Request("txtToDT") & """" & vbCr
	Response.Write ".txtHBillType.value	= """ & ConvSPChars(Request("txtBillType")) & """" & vbCr
	Response.Write ".txtHSalesGrp.value	= """ & ConvSPChars(Request("txtSalesGrp")) & """" & vbCr
	Response.Write ".txtHSoNo.value	= """ & ConvSPChars(Request("txtSoNo")) & """" & vbCr
	Response.Write ".txtHDnNo.value	= """ & ConvSPChars(Request("txtDnNo")) & """" & vbCr
	Response.Write ".txtHPostFlag.value	= """ & Request("txtPostFlag") & """" & vbCr
	Response.Write ".txtHExceptFlag.value	= """ & Request("txtExceptFlag") & """" & vbCr
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Parent.ggoSpread.Source	= .vspdData" & vbCr
 	Response.Write ".vspdData.Redraw = False " & vbCr      
	Response.Write "parent.ggoSpread.SSShowDataByClip """ & lgstrData  & """ ,""F""" & vbCr
	Response.Write "parent.lgPageNo	= """ & lgPageNo & """" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr
 	Response.Write ".vspdData.Redraw = True " & vbCr      
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr
End Sub
%>
