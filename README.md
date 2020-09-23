<div align="center">

## Connect to Lotus Notes View using COM and Loop through records


</div>

### Description

This piece of code allows you to connect to an existing Lotus Notes View using COM and Loop through the records within that View. This is simple, but for those people out there that deal with Notes and SQL Server/Access databases, then this is a must. Very easy to manipulate.
 
### More Info
 
Be sure to reference the Lotus Domino Object!

None that I am aware of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Todd Walling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/todd-walling.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/todd-walling-connect-to-lotus-notes-view-using-com-and-loop-through-records__1-34295/archive/master.zip)

### API Declarations

```
' Declare Notes COM variables
  Dim n_Session As New Domino.NotesSession
  Dim n_Database As New Domino.NotesDatabase
  Dim n_Document As NotesDocument
  Dim n_ViewEntry As NotesViewEntry
  Dim n_View As NotesView
  Dim n_ViewNav As NotesViewNavigator
  Dim l_TestVariable As String
```


### Source Code

```
' Initialize session and set database
  n_Session.Initialize
  ' Set database
  Set n_Database = n_Session.GetDatabase("Your Server Name", "YourFileLocation\YourFile.nsf")
  ' Set to view
  Set n_View = n_Database.GetView("NameOfYourView")
  ' Set view navigator
  Set n_ViewNav = n_View.CreateViewNav
  ' Move to first record
  Set n_ViewEntry = n_ViewNav.GetFirstDocument()
  ' Loop through records within view
  Do While Not (n_ViewEntry Is Nothing)
    ' Set view to document
    Set n_Document = n_ViewEntry.Document
    ' Set local variables
    l_TestVariable = n_Document.GetItemValue("TestFieldName")(0)
    ' Get next entry
     Set n_ViewEntry = n_ViewNav.GetNextDocument(n_ViewEntry)
  Loop
  ' Clean-up
  Set n_ViewEntry = Nothing
  Set n_ViewNav = Nothing
  Set n_View = Nothing
  Set n_Document = Nothing
  Set n_Database = Nothing
  Set n_Session = Nothing
```

