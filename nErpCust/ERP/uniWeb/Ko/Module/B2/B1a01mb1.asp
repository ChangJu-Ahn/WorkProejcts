<% 


'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Major Code)
'*  3. Program ID           : b1a01mb1.asp
'*  4. Program Name         : b1a01mb1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 1999/09/10
'*  7. Modified date(Last)  : 2002/12/16
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pB1A011											'입력/수정용 ComProxy Dll 사용 변수 
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim adoRec
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow         
Dim RowData(5)
Dim RowDataPre

call LoadBasisGlobalInf()

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    'Multi SpreadSheet
    strSpread = Request("txtSpread")
    lgLngMaxRow       = Request("txtMaxRows")                                      
  
Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub    

Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

Sub SubBizQueryMulti()
    Dim strData
    
    On Error Resume Next  
	Err.Clear                                                                  '☜: Clear Error status
 	
	Call SubMakeSQLStatements()
	
	
	Set adoRec = Server.CreateObject("ADODB.Recordset")

	adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 	
	
	
	If Err.Number <> 0 Then										
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Call SubCloseDB(lgObjConn)                                               '☜: Close DB Connection
		Response.End
	End If

	If adoRec.EOF = True Then										'☜:해당하는 데이터가 존재하지 않습니다.
		Call DisplayMsgBox("971001", vbOKOnly, "해당하는 Major 코드", "", I_MKSCRIPT)
	End If
	
	LngRow = 0
	
	Do While Not adoRec.EOF
        LngRow = LngRow + 1
	 	RowData(0) = adoRec(0) 'major_cd
	 	RowData(1) = adoRec(1) 'major_nm
	 	RowData(2) = adoRec(2) 'minor_len
	 	RowData(3) = adoRec(3) 'type(major)
	 	RowData(4) = adoRec(4) 'minor_type
		
	 	RowDataPre = adoRec(0)
	 	adoRec.MoveNext
	 	
	 	'다음 레코드의 Major_Cd가 이전 레코드의 Major_Cd와 같을 경우 
	 	Do While RowDataPre = adoRec(0) 
	 		if RowData(4) <> adoRec(4) then
	 			if Trim(adoRec(4)) = "S" then
	 				RowData(4) = Trim(adoRec(4))
	 			End If
	 		End If			
	 		adoRec.MoveNext	
	 		if adoRec.EOF then				'중요체크.........
	 			exit Do
	 		End If	            
	     Loop
	    			
	 	strData = strData & Chr(11) & RowData(0)
	 	strData = strData & Chr(11) & RowData(1)
	 	strData = strData & Chr(11) & RowData(2)
	 	If  Trim(RowData(3)) = "S" Then
	 		strData = strData & Chr(11) & "시스템 정의"	
	 	Else
	 		strData = strData & Chr(11) & "사용자 정의"	
	 	End If
	 	If  RowData(4) = "S" Then
	 		strData = strData & Chr(11) & "Y"
	 	Else
	 		strData = strData & Chr(11) & "N"
	 	End If

	 	strData = strData & Chr(11) & LngRow

	 	strData = strData & Chr(11) & Chr(12)
	 Loop

	
%>

<Script Language=vbscript>
    Dim LngRow          
    Dim strTemp
    Dim strData

	With parent																	'☜: 화면 처리 ASP 를 지칭함 
	.frm1.txtMajorNm.value = "<%=ConvSPChars(LookUpMajor(Request("txtMajor")))%>"
	.frm1.vspdData.ReDraw = False
	 strData = "<%=ConvSPChars(strData)%>"
    .ggoSpread.Source = .frm1.vspdData 
	.ggoSpread.SSShowData strData
	.frm1.vspdData.ReDraw = True
	.DbQueryOk
	End With
</Script>	
<% 
	
	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection	
End Sub    
%>	
	
<%	
Sub SubBizSaveMulti()
				
    Dim ObjPB2G011
    
    Set ObjPB2G011 = server.CreateObject ("PB2G011.cBControlMajorCode")    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
                                                                               '☜: Clear Error status
    Call ObjPB2G011.B_CONTROL_MAJOR_CODE(gStrGlobalCollection,strSpread)
    Set ObjPB2G011 = nothing

    If CheckSYSTEMError(Err,True) = True Then                                              
		Response.End 
    End If
    on error goto 0                                                             
%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%					

End Sub

Function LookUpMajor(Byval strCode)
    Const B386_I1_major_cd = 0
    Const B386_I1_major_nm = 1
    
    Const B386_E1_major_cd = 0
    Const B386_E1_major_nm = 1
    Const B386_E1_minor_len = 2
    Const B386_E1_type = 3

	Dim ObjPB2S012		
	Dim I1_b_major
	Dim E1_b_major
	
    ReDim I1_b_major(B386_I1_major_nm)
    ReDim E1_b_major(B386_E1_type)
    
    I1_b_major(B386_I1_major_cd) = Request("txtMajor")
    
    Set ObjPB2S012 = server.CreateObject ("PB2S012.cBLookMajorCode")    
    
    On Error Resume Next
    Err.Clear                                                                            '☜: Clear Error status
    E1_b_major = ObjPB2S012.B_LOOKUP_MAJOR_CODE (gStrGlobalCollection,I1_b_major)
    Set ObjPB2S012 = nothing    

    If Err.number <> 0 and inStr(Err.Description ,"122200") > 0 then
  	LookUpMajor = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then                                              
        	Exit Function
	    End If
        on error goto 0
        LookUpMajor = E1_b_major(B386_E1_major_nm)
    End If

End Function
%>


<Script Language=vbscript RUNAT=server>
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
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
</Script>
<%

Sub SubMakeSQLStatements()
    Dim MajorCd  

	MajorCd = Replace(Request("txtMajor"), "'", "''")
	
	lgStrSQL = lgStrSQL & "select distinct a.major_cd, a.major_nm,a.minor_len, a.type, b.minor_type "  
	lgStrSQL = lgStrSQL & "from b_major a, b_minor b " 
	lgStrSQL = lgStrSQL & "where a.major_cd *= b.major_cd  and a.major_cd >=  " & FilterVar(MajorCd , "''", "S") & " order by a.major_cd"  

End Sub
%>
