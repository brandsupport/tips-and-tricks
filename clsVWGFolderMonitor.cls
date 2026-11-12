VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVWGFolderMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents WatchedItems As Outlook.Items
Public WatchedFolder As Outlook.Folder

Private Sub WatchedItems_ItemAdd(ByVal Item As Object)
    On Error Resume Next
    modVWGDispatcher.HandleMailItem Item, WatchedFolder, "ITEM_ADD"
End Sub

Private Sub WatchedItems_ItemChange(ByVal Item As Object)
    On Error Resume Next
    modVWGDispatcher.HandleMailItem Item, WatchedFolder, "ITEM_CHANGE"
End Sub
