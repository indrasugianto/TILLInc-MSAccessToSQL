' Module Name: Report_rptRESEMERGENCYPHONES
' Module Type: Document Module
' Lines of Code: 26
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
    CellPhoneTree = Null
    If Cluster <> 1 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 1") & "  "
    If Cluster <> 2 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 2") & "  "
    If Cluster <> 3 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 3") & "  "
    If Cluster <> 4 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 4") & "  "
    If Cluster <> 5 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 5") & "  "
    If Cluster <> 6 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 6") & "  "
    If Cluster <> 7 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 7") & "  "
    If Cluster <> 8 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 8") & "  "
    If Cluster <> 9 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 9") & "  "
    If Cluster <> 10 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 10") & "  "
    If Cluster <> 11 Then CellPhoneTree = CellPhoneTree & DLookup("ClusterManagerCell", "CatClusters", "ClusterID = 11") & "  "
End Sub

Private Sub Report_Current()
    If LocationName <> PReviousLocationName Then PReviousLocationName = LocationName
End Sub

Private Sub Report_Load()
    PReviousLocationName = LocationName
    Page = 0
End Sub
