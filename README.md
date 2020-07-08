# Sharepoint-Automation-with-VBA





Sub link_edit_Mode()
    Dim mySh As Worksheet
    Dim spSite As String
    
    Set mySh = Sheets("Sheet1")
    
    Dim src(0 To 1) As Variant
    
    spSite = "https://  " 'site name
    src(0) = spSite & "/_vti_bin"
    
    src(1) = "{d2078826-099e-43fa-90b2-756ad973730e}" 'GUID
    
    mySh.ListObjects.Add xlSrcExternal, src, True, xlYes, mySh.Range("A1")
    
End Sub



Sub SaveChanges()
 Dim mySh As Worksheet
   Dim lstOBJ As ListObject

   On Error GoTo errhdnler
   
   Set mySh = Sheets("Sheet1")
   Set lstOBJ = mySh.ListObjects(1)
   
   lstOBJ.UpdateChanges xlListConflictDialog
   
   Set mySh = Nothing
   Set lstOBJ = Nothing
   
Exit Sub
errhdnler:

Debug.Print Err.Description & Err.Number

End Sub


Sub refresh_Con()
 Dim mySh As Worksheet
   Dim lstOBJ As ListObject

On Error GoTo errhdnler

   Set mySh = Sheets("Sheet1")
   
   Set lstOBJ = mySh.ListObjects(1)
   
   lstOBJ.Refresh
  
   Set mySh = Nothing
   Set lstOBJ = Nothing
   
Exit Sub

errhdnler:

Debug.Print Err.Description & Err.Number

End Sub
