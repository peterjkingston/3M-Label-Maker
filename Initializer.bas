Attribute VB_Name = "Initializer"
Public Sub Run(datastore As Collection)
    Deploy.All
    Dictionaries.AssembleAll datastore
End Sub
