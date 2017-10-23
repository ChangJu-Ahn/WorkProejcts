'==========================================================================================
'   Event Name : SearchTV
'   Event Desc : Search tree view data
'==========================================================================================
Sub SearchTV(Node,strData,pFlag)
    Dim TempNode

    If Node Is Nothing Then
       Exit Sub
    End If

    Call FindTVData(Node.Child ,strData,pFlag)             'Current child
    
    If pFlag = "N" Then
       Exit Sub
    End If
    
    Call FindTVSiblingData(Node,strData,pFlag)            'Current next child

    If pFlag = "N" Then
       Exit Sub
    End If
    
    Set TempNode = Node

    Do While(1)
    
       If TempNode.Parent Is Nothing Then                 'Current sibling of parent
          Exit Sub
       End If
       
       Set TempNode = TempNode.Parent

       Call FindTVData(TempNode.Next ,strData,pFlag)       

       If pFlag = "N" Then
          Exit Sub
       End If
          
    Loop   
    
End Sub

'==========================================================================================
'   Event Name : FindTVSiblingData
'   Event Desc : Find Sibling Node
'==========================================================================================
Sub FindTVSiblingData(pNode ,strData,pFlag )

    Do While(1)
       If pNode.Next Is Nothing Then
          Exit Sub
       End If
       
       Set pNode = pNode.Next

       If UCase(pNode.Text) = UCase(strData) Then
          pNode.Selected = True 
          pFlag = "N"
          Call ExpandNodeToParent(pNode)
          Exit Sub
       End If
       
       Call FindTVData(pNode.Child ,strData,pFlag)

       If pFlag = "N" Then
          Exit Sub
       End If

    Loop   

End Sub

'==========================================================================================
'   Event Name : FindTVData
'   Event Desc : Find All Node
'==========================================================================================
Sub FindTVData(pNode ,strData,pFlag )
    Dim TempParentNode
    
    If pNode Is Nothing Or pFlag = "N" Then
       Exit Sub
    End If
    
    Do While (1)

       If pFlag = "N" Then
          Exit Sub 
       End If

       If UCase(pNode.Text) = UCase(strData) Then
       
          pNode.Selected = True 
          pFlag   = "N"
          Call ExpandNodeToParent(pNode)
          Exit Sub
          
       End If
       
       Call FindTVData(pNode.Child, strData,pFlag)
       
       If pNode.Next Is Nothing Then
          Exit Sub
       End If

       Set pNode = pNode.Next
    Loop
        
End Sub
'==========================================================================================
'   Event Name : ExpandNodeToParent
'   Event Desc : Expamd current node , parent and ancestor
'==========================================================================================
Sub ExpandNodeToParent(pNode)
    Dim TempParentNode

    Set TempParentNode = pNode
          
    Do While (1)
       If TempParentNode.Parent Is Nothing Then
          Exit Sub
       End if
       Set TempParentNode = TempParentNode.Parent
       TempParentNode.Expanded = True
    Loop     

End Sub
