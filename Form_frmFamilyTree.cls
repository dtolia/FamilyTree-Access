VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFamilyTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load() ' Or Button_Click

    Dim rs As DAO.Recordset
    Dim tvw As Object
    Dim nodRoot As Object
    Dim nodChild As Object
    Dim dictNodes As Object
    Dim personID As Long
    Dim parentID As Variant
    Dim personName As String
    Dim personID_Str As String
    Dim sanitizedKey As String
    Dim parentSanitizedKey As String

    Set tvw = Me.TreeView0
    Set dictNodes = CreateObject("Scripting.Dictionary")

    tvw.Nodes.Clear

    Set rs = CurrentDb.OpenRecordset("SELECT ID, NameID, ParentID FROM Family ORDER BY ID") ' Order by ID might still be helpful

    If Not rs.EOF Then
        rs.MoveFirst

        ' --- Single Pass: Create nodes and establish hierarchy directly ---
        Do While Not rs.EOF
            personID = rs!ID
            personName = rs!NameID
            parentID = rs!parentID  ' Get ParentID in each record
            personID_Str = CStr(personID)
            sanitizedKey = "NodeKey_" & Replace(personID_Str, ".", "_")
            parentSanitizedKey = "NodeKey_" & CStr(parentID) ' For potential parent lookup


            Debug.Print "Processing Record - ID:", personID, "Name:", personName, "ParentID:", parentID, "sanitizedKey:", sanitizedKey, "parentSanitizedKey:", parentSanitizedKey

            If IsNull(parentID) Or parentID = 0 Then
                ' --- Create ROOT NODE (no parent) ---
                Debug.Print "  Creating ROOT Node - Key:", sanitizedKey, "Name:", personName
                On Error Resume Next
                Set nodChild = tvw.Nodes.Add(, , sanitizedKey, personName) ' No relative, no relationship = root
                If Err.Number <> 0 Then
                    Debug.Print "  Error adding ROOT node:", Err.Number, Err.Description, "Key:", sanitizedKey
                    Err.Clear
                    GoTo SkipNode
                End If
                On Error GoTo 0
                dictNodes.Add personID, nodChild ' Store root node in dictionary

            Else
                ' --- Create CHILD NODE (under parent) ---
                If dictNodes.Exists(parentID) Then ' **Crucial: Check if parent node exists in dictionary**
                    Set nodRoot = dictNodes(parentID) ' Get parent node from dictionary

                    Debug.Print "  Creating CHILD Node - Key:", sanitizedKey, "Name:", personName, "Parent Key:", nodRoot.Key
                    On Error Resume Next
                    Set nodChild = tvw.Nodes.Add(nodRoot.Key, tvwChild, sanitizedKey, personName) ' **Add as CHILD under nodRoot**
                    If Err.Number <> 0 Then
                        Debug.Print "  Error adding CHILD node:", Err.Number, Err.Description, "Parent Key:", nodRoot.Key, "Child Key:", sanitizedKey
                        Err.Clear
                        GoTo SkipNode
                    End If
                    On Error GoTo 0
                    dictNodes.Add personID, nodChild ' Store child node in dictionary

                Else
                    Debug.Print "  WARNING: Parent node NOT FOUND for ParentID:", parentID, " while processing Child:", personName, " (ID:", personID, "). Child will be created as ROOT."
                    ' **Handle case where parent is not found (e.g., data issue) - create as root for now**
                    On Error Resume Next
                    Set nodChild = tvw.Nodes.Add(, , sanitizedKey, personName) ' Create as ROOT if parent not found
                    If Err.Number <> 0 Then
                        Debug.Print "  Error adding ROOT node (Parent Not Found Fallback):", Err.Number, Err.Description, "Key:", sanitizedKey
                        Err.Clear
                        GoTo SkipNode
                    End If
                    On Error GoTo 0
                    dictNodes.Add personID, nodChild ' Store as root (even if it should be child)

                End If
            End If ' End If - IsNull(ParentID) Or ParentID = 0


SkipNode: ' Skip to next record if node addition fails
            rs.MoveNext
        Loop    ' Loop End - Single Pass
    End If      ' End If - Recordset not EOF

    rs.Close
    Set rs = Nothing
    Set tvw = Nothing
    Set nodRoot = Nothing
    Set nodChild = Nothing
    Set dictNodes = Nothing

End Sub
