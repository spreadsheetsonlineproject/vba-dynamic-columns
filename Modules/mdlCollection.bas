Attribute VB_Name = "mdlCollection"
Option Explicit

Public Sub howCollectionWorks()

    Dim coll As New Collection
    
    coll.Add "elso", "kulcs1"
    coll.Add "masodik", "kulcs2"
    coll.Add "harmadik", "kulcs3", 2

    Call printColl(coll)
    
    Debug.Print "-----------"
    
    Debug.Print coll(2)
    
    Debug.Print "-----------"
    
    Debug.Print coll.Count
    
    coll.Add "uj", "kulcsuj", before:=3
    coll.Remove (2)
    
    Call printColl(coll)
    Debug.Print coll.item(2)
    
    Debug.Print "////////////////"

End Sub

Public Sub printColl(ByRef coll As Collection)

    Dim item As Variant
    For Each item In coll
        Debug.Print item
    Next item

End Sub
