' Module Name: Report_rptDAYEMERGENCYPHONES
' Module Type: Document Module
' Lines of Code: 28
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
'    CellPhoneTree = Null
'    If Cluster <> 1 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 1") & ", "
'    If Cluster <> 2 Then
'        CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 2")
'        If Not IsNull(CellPhoneTree) Then CellPhoneTree = CellPhoneTree & ", "
'    End If
'    If Cluster <> 3 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 3") & ", "
'    If Cluster <> 4 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 4") & ", "
'    If Cluster <> 5 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 5") & ", "
'    If Cluster <> 6 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 6") & ", "
'    If Cluster <> 7 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 7") & ", "
'    If Cluster <> 8 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 8") & ", "
'    If Cluster <> 9 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 9")
'    If Cluster <> 99 Then CellPhoneTree = CellPhoneTree & ", " & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 99")
End Sub

Private Sub Report_Current()
    If LocationName <> PReviousLocationName Then PReviousLocationName = LocationName
End Sub

Private Sub Report_Load()
    PReviousLocationName = LocationName
    Page = 0
End Sub