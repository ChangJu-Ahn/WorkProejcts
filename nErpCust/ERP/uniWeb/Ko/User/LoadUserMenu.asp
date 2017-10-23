<!-- #Include file="../inc/IncServer.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<Script Language=vbscript src="../inc/incUni2KTV.vbs"></Script>

<%

	Err.Clear											                         	'¢Ð: Protect system from crashing

	Dim ADOConn
	Dim ADORs
	Dim strSql
	Dim UsrMenu

    Call SubOpenDB(ADOConn)                                                        '¢Ð: Make  a DB Connection
    
    strSql =		  "SELECT UPPER_MNU_ID, MNU_ID, MNU_NM, MNU_TYPE "
	strSql = strSql & "FROM Z_USR_MNU "
    strSql = strSql & "WHERE USR_ID = '" & gUsrID & "' "
    strSql = strSql & "AND LANG_CD  = '" & gLang & "' "
    strSql = strSql & "ORDER BY SYS_LVL ASC, MNU_SEQ ASC"    
		

    If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
        UsrMenu =  ""
    Else
	    While Not ADORs.EOF        
           UsrMenu = UsrMenu & ADORs("UPPER_MNU_ID") & chr(11) & ADORs("MNU_ID") & chr(11) & ADORs("MNU_NM") & chr(11) & ADORs("MNU_TYPE") & chr(11) & Chr(12)
           ADORs.MoveNext
	    WEnd        
    End If

    Call SubCloseRs(ADORs)                                                          '¢Ð: Release RecordSSet
    Call SubCloseDB(ADOConn)     

%>

<Script Language=vbscript>

	Dim strUpper
	Dim strNm
	Dim strKey
	Dim strMnuType
	Dim strImg
	Dim ndNode
    Call MakeUserMenuTreeView()
    parent.DbQueryOk

Sub MakeUserMenuTreeView()
    Dim UsrMenuCol
    Dim UsrMenuRow
    Dim iDx
	Dim ndNode
	Dim UsrMenu
	
    On Error Resume Next
    
    UsrMenu =  "<%=UsrMenu%>"
    
    
    If Trim(UsrMenu) = "" Then
       Exit Sub
    End If
    
    UsrMenuRow = Split(UsrMenu,Chr(12))

    For iDx = 0 To UBound(UsrMenuRow) - 1 
        
        UsrMenuCol = Split(UsrMenuRow(iDx),Chr(11))
            
		If UsrMenuCol(3) = "P" Then
			strImg = C_USURL
			UsrMenuCol(1) = UsrMenuCol(1) & Chr(20) & UsrMenuCol(0)			
		Else
			strImg = C_USFolder
		End If
		
		if UsrMenuCol(3) = "M" Then	
			Set ndNode = Parent.frm2.uniTree1.Nodes.Add(UsrMenuCol(0), tvwChild, UsrMenuCol(1), UsrMenuCol(2), strImg)
			ndNode.ExpandedImage = C_USOpen
			ndNode.Tag = 0
			ndNode.Parent.Tag = Cint(ndNode.Parent.Tag) + 1
		Else
			Set ndNode = Parent.frm2.uniTree1.Nodes(UsrMenuCol(0))
			ndNode.Tag = Cint(ndNode.Tag) + 1
		End If			
    Next    
    
    
    
End Sub

Sub Document_onReadyStateChange()
	Dim SelectNode
	Set SelectNode = parent.frm2.uniTree1.Nodes(1)
	SelectNode.Selected = True
	SelectNode.Expanded = True 
	parent.frm2.uniTree1.HideSelection = False
	parent.frm2.uniTree1.MousePointer = 0
End Sub

</Script>

