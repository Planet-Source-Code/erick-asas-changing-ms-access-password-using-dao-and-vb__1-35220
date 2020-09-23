Attribute VB_Name = "mod_Security"
'This sample program is designed by Erick Asas,
'SentrySoft2000 Ltd.

'The program is intended as an example of how to implement
'changing of MS password from VB.
'The program is free to use and comes as is with nowarranty
'
'clasas@mnd.philcom.com
'Initial UserName = Admin
'Initial Password = ""

Option Explicit
Public strUID As String 'var to hold username
Public strPWD As String 'var to hold user password

Public Function Open_Database(ByVal User_Name As String, ByVal User_Password As String) As Database
On Error GoTo ET


Dim dbs As Database
Dim strMdb As String
Dim strMdw As String

    ' Ensure that the Microsoft Jet workgroup information
    ' file is available.
    strMdw = App.Path & "\" & "MSSecure.mdw"
    DBEngine.SystemDB = strMdw
    ' Set the DefaultUser and DefaultPassword properties for
    ' the DBEngine object.
    DBEngine.DefaultUser = User_Name
    DBEngine.DefaultPassword = User_Password
    
    'Define Jet WorkSpace
    Dim wrkjet As Workspace
    
    Set wrkjet = CreateWorkspace("JetWorkspace", User_Name, _
      User_Password, dbUseJet)
    
    strMdb = App.Path & "\" & "MSSecure.mdb"
    'Open database for any transactions
    Set dbs = wrkjet.OpenDatabase(strMdb)
    Set Open_Database = dbs
    dbs.Close
  Exit Function

ET:
    Err.Raise Err.Number, Err.Description
    Exit Function
End Function

Public Function UpdatePassword(ByVal User_Name As String, ByVal Old_Password As String, ByVal New_Password As String) As Boolean
On Error GoTo ET

    ' Ensure that the Microsoft Jet workgroup information
    ' file is available.
    DBEngine.SystemDB = App.Path & "\" & "MSSecure.mdw"
    ' Set the DefaultUser and DefaultPassword properties for
    ' the DBEngine object.
    DBEngine.DefaultUser = User_Name
    DBEngine.DefaultPassword = Old_Password
    
    'Define Jet WorkSpace
    Dim wrkjet As Workspace
    
    Set wrkjet = CreateWorkspace("JetWorkspace", User_Name, _
      Old_Password, dbUseJet)
    ' Change Password
    With wrkjet
      .Users(User_Name).NewPassword Old_Password, New_Password
    End With
    wrkjet.Close
    UpdatePassword = True
    
Exit Function
ET:
    UpdatePassword = False
    Err.Raise Err.Number, Err.Description
    Exit Function
End Function
