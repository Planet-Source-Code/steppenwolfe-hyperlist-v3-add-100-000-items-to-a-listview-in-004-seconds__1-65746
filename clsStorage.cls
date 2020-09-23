VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/~ Macbeth, 2nd Witch: By the prickling in my thumbs, something wicked this way comes..

Private Const NEG1                  As Long = -1
Private Const m0                    As Long = &H0
Private Const m1                    As Long = &H1
Private Const m2                    As Long = &H2
Private Const m4                    As Long = &H4
Private Const m8                    As Long = &H8
Private Const m32                   As Long = &H20

'/* array type
Private Enum eArrType
    Array_Long = 0
    Array_String = 1
    Array_Variant = 2
End Enum

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                  pSrc As Any, _
                                                                  ByVal lByteLen As Long)
                                                                  
Private m_TestType() As HLIStc

Public Event eSComplete()
Public Event eSErrCond(sRtn As String, lErr As Long)

Private m_bSorted                   As Boolean
Private m_bDimensional              As Boolean
Private m_bTags                     As Boolean
Private m_bColumns                  As Boolean
Private m_bActive                   As Boolean
Private m_lPointer                  As Long
Private m_lLowCount                 As Long
Private m_lHighCount                As Long
Private m_lColCount                 As Long
Private m_lCRef                     As Long
Private m_lPtr()                    As Long
Private m_lRetPtr                   As Long
Private m_sArrIdx()                 As String
Private m_aStorage()                As String
Private c_Key                       As Collection
Private c_Var                       As Collection
Private m_tHIStc()                  As HLIStc
Private m_RItem()                   As HLIRtn
Private m_tSubArr()                 As HLISubItm
Private c_PtrMem                    As Collection

'/* array lowbound
Public Property Get p_LowCount() As Long
    p_LowCount = m_lLowCount
End Property

Public Property Let p_LowCount(ByVal PropVal As Long)
    m_lLowCount = PropVal
End Property

'/* array highbound
Public Property Get p_HighCount() As Long
    p_HighCount = m_lHighCount
End Property

Public Property Let p_HighCount(ByVal PropVal As Long)
    m_lHighCount = PropVal
End Property

'/* sorted status
Public Property Get p_Sorted() As Boolean
    p_Sorted = m_bSorted
End Property

Private Property Let p_Sorted(ByVal PropVal As Boolean)
    m_bSorted = PropVal
End Property

'/* dimension count
Public Property Get p_ColCount() As Long
    p_ColCount = m_lColCount
End Property

Public Property Let p_ColCount(ByVal PropVal As Long)
    m_lColCount = PropVal
End Property


Private Sub Class_Initialize()
'
End Sub

'                   <<+>>
'||<<<<<<<<<<<<<<<< Build >>>>>>>>>>>>>>>>||
'                   <<+>>

Private Function Array_Check(ByRef sArray As Variant) As Boolean
'/* validity test

On Error GoTo Handler

    '/* an array
    If IsArray(sArray) Then
        On Error Resume Next
        '/* dimensioned
        If IsError(UBound(sArray)) Then
            GoTo Handler
        End If
        On Error GoTo 0
    Else
        GoTo Handler
    End If

    Array_Check = True

On Error GoTo 0
Exit Function

Handler:

End Function

Private Function Array_Counters(ByRef aStorage As Variant) As Boolean
'/* store initial array count

On Error GoTo Handler

    m_lLowCount = LBound(aStorage)
    m_lHighCount = UBound(aStorage)
    '/* actual counter
    m_lCRef = m_lHighCount
    
    '/* success
    Array_Counters = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("Array_Counters", Err.Number)

End Function

Private Function ArrayExists(ByRef sArray As Variant) As Boolean

On Error Resume Next

    If IsError(UBound(sArray)) Then
        GoTo Handler
    End If

    '/* success
    ArrayExists = True

Handler:
    On Error GoTo 0

End Function

Private Function Array_Reset() As Boolean

On Error Resume Next
    
    '/* reset tracking keys
    Key_Reset
    '/* Erase arrays
    Erase m_aStorage
    Erase m_sArrIdx
    Erase m_lPtr
    Clear_ItmPtr
    '/* reset counters
    m_lCRef = 0
    m_lLowCount = 0
    m_lHighCount = 0
    m_bTags = False
    m_bDimensional = False
    m_lColCount = 0
    m_bActive = False
    
    '/* success
    Array_Reset = True

On Error GoTo 0

End Function

Private Function Array_Type(ByRef aStorage As Variant) As eArrType
'/* get array variable type (tags)

On Error GoTo Handler

    If TypeName(aStorage) = "Long()" Then
        Array_Type = Array_Long
    ElseIf TypeName(aStorage) = "String()" Then
        Array_Type = Array_String
    ElseIf TypeName(aStorage) = "Variant()" Then
        Array_Type = Array_Variant
    Else
        Array_Type = -1
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("Array_Type", Err.Number)

End Function

Public Function Change_State() As Boolean
'/* structure dimensions have changed
'/* so adjust params

    '/* alternatively, you could reshuffle
    '/* the index, and reload the control..
    m_bSorted = False
    '/* store the dimensions
    Array_Counters m_aStorage

End Function

Private Function DeAllocate_Pointer(ByVal sKey As String, _
                                    Optional ByVal bPurge As Boolean) As Boolean

'/* resolve or purge memory pointers

Dim lPtr    As Long
Dim lC      As Long

On Error GoTo Handler

    If Not bPurge Then
        '/* get the pointer
        lPtr = c_PtrMem.Item(sKey)
        If lPtr = 0 Then GoTo Handler
        '/* release the memory
        CopyMemBr ByVal lPtr, 0&, &H4
    Else
        '/* destroy the struct last
        For lC = c_PtrMem.Count To 1 Step -1
            If Not CLng(c_PtrMem.Item(lC)) = 0 Then
                lPtr = c_PtrMem.Item(lC)
                CopyMemBr ByVal lPtr, 0&, &H4
            End If
        Next lC
    End If

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("DeAllocate_Pointer", Err.Number)

End Function

'* Name           : Build_Control
'* Purpose        : array build hub
'* Inputs         : storage pointer -long
'* Outputs        : none
'*********************************************
Public Function Build_Control(ByVal lPointer As Long) As Boolean

'~ It is possible to manipulate the same array from
'~ two seperate entities, (processes), by using
'~ array pointer duplication, however, some
'~ caution must be used to ensure secondary pointers
'~ are fully resolved upon class termination.
'~ Failure to resolve the pointers will result
'~ in the arrays remaining in memory after the
'~ application has terminated, as the runtime is
'~ unable to reclaim the resources due to intervention
'~ via copymemory.
'~ I have included a safer method, Build_CtlSafe,
'~ which destroys the passed structure immediately
'~ after removing its data members. The downside
'~ to this approach however, is that all arrays will need
'~ to be recreated each time there is a change to the
'~ passed structure, whilst with the Build_Control method,
'~ only the working array needs to be copied to sync a change.
'~ A benefit to this approach though, is that returning
'~ subitems will run faster using working arrays, rather then
'~ accessing the array through a complex structure.

Dim i       As Long
Dim lPtr    As Long
Dim lSpr    As Long
Dim lTpr    As Long

    '/* reset container
    If m_bActive Then
        Array_Reset
    End If
    
    Set c_PtrMem = New Collection
    '/* initialize local struct
    ReDim m_tHIStc(0)
    '/* copy the structure from the pointer
    CopyMemBr ByVal VarPtrArray(m_tHIStc), lPointer, &H4
    c_PtrMem.Add lPointer, "ptHIStc"
    
    With m_tHIStc(0)
        '/* create the working array
        If Array_Check(.Item) Then
            If ArrayExists(.Item) Then
                '/* high and low count
                Array_Counters .Item
                '/* create a working array for the index functions
                '/* copy the array pointer
                CopyMemBr lSpr, ByVal GetASPtr(.Item), &H4
                '/* copy the pointer to our temp array
                CopyMemBr ByVal GetASPtr(m_aStorage), lSpr, &H4
                '/* add the pointer to mem collection
                c_PtrMem.Add lSpr, "m_aStorage"
            '/* no items -destroy and abort
            Else
                DeAllocate_Pointer "ptHIStc"
                Exit Function
            End If
        End If
        
        '/* test for subitems
        If ArrayExists(.SubItem(LBound(.Item)).Text) Then
            m_lColCount = .Column
            m_bDimensional = True
        End If

        '/* we are using tags-
        '/* create a working array
        '/* for the search tags
        If Array_Check(.Tag) Then
            If ArrayExists(.Tag) Then
                m_bTags = True
                CopyMemBr lTpr, ByVal GetASPtr(.Tag), &H4
                CopyMemBr ByVal GetASPtr(m_sArrIdx), lTpr, &H4
                c_PtrMem.Add lTpr, "m_sArrIdx"
            End If
        End If
    End With

    m_bActive = True
    '/* success
    Build_Control = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("Build_Control", Err.Number)
    
End Function

Public Function Return_Item(ByVal lIndex As Long) As Long

'/* return
Dim lS      As Long
Dim lX      As Long
Dim lPtr    As Long

On Error GoTo Handler

    '/* swap sort pointer
    If m_bSorted Then lIndex = Get_Index(lIndex)
    '/* reset the pointer
    Clear_ItmPtr
    '/* redim struct
    ReDim m_RItem(0)
    With m_tHIStc(0)
        '/* copy item text
        lS = LenB(.Item(lIndex))
        If Not lS = 0 Then
            'm_RItem.RItem = Space$(lS)
            'CopyMemBr m_RItem.RItem, .Item(lIndex), lS
        End If
    End With
    
    '/* add subitems
    If m_bDimensional Then
        With m_tHIStc(0)
            '/* test subitem array
            If Array_Check(.SubItem(lIndex).Text) Then
                '/* dimension container
                'ReDim m_RItem.RSubItem(LBound(.SubItem(lIndex).Text) To UBound(.SubItem(lIndex).Text))
                '/* loop through structure
                For lX = LBound(.SubItem(lIndex).Text) To UBound(.SubItem(lIndex).Text)
                    '/* size string
                    lS = LenB(.SubItem(lIndex).Text(lX))
                    If lS = 0 Then GoTo skip
                    '/* size array item
                    'm_RItem.RSubItem(lX) = Space$(lS)
                    '/* copy
                    'CopyMemBr m_RItem.RSubItem(lX), .SubItem(lIndex).Text(lX), lS
skip:
                Next lX
            End If
        End With
    End If
    
    '/* copy struct to pointer
    'CopyMemBr m_lRetPtr, ByVal VarPtrArray(m_RItem), &H4
    Return_Item = m_lRetPtr

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("Return_Item", Err.Number)

End Function

Private Sub Clear_ItmPtr()
'/* release item return pointer

    If Not m_lRetPtr = 0 Then
        CopyMemBr ByVal m_lRetPtr, 0&, &H4
        m_lRetPtr = 0
    End If
    
End Sub


'                    <<*>>
'||<<<<<<<<<<<<<<<< Search >>>>>>>>>>>>>>>>||
'                    <<*>>

'* Name           : Item_Search
'* Purpose        : retrieve item by number
'* Inputs         : row number -long
'* Outputs        : item text -string
'*********************************************
Public Function Item_Search(ByVal lItem As Long) As String

Dim lTemp As Long

On Error Resume Next

    If m_bSorted Then
        Item_Search = m_aStorage(m_lPtr(lItem))
    Else
        Item_Search = m_aStorage(lItem)
    End If

On Error GoTo 0

End Function

'* Name           : Index_Return
'* Purpose        : retrieve sorted item index
'* Inputs         : row number -long
'* Outputs        : item index -long
'*********************************************
Public Function Index_Return(ByVal lItem As Long) As Long

Dim lTemp As Long

On Error GoTo Handler

    If Not m_bSorted Then GoTo Handler
    Index_Return = m_lPtr(lItem)

On Error GoTo 0
Exit Function

Handler:
    Index_Return = -1

End Function

'* Name           : Tag_Search
'* Purpose        : search array by row tag
'* Inputs         : search tag -string
'* Outputs        : row number -long
'*********************************************
Public Function Tag_Search(ByVal sTag As String) As Long
'/* return a row number by its tag id

On Error GoTo Handler

    If m_bSorted Then
        Tag_Search = c_Key.Item(sTag)
    Else
        Tag_Search = m_lPtr(CLng(c_Key.Item(sTag)))
    End If

On Error GoTo 0
Exit Function

Handler:
    Tag_Search = -1

End Function

'* Name           : Tag_Search
'* Purpose        : search array by row tag
'* Inputs         : search tag -string
'* Outputs        : row number -long
'*********************************************
Public Function Remove_Duplicates() As Collection
'/* fastest de-duper possible

Dim lC  As Long
Dim cT  As Collection

On Error Resume Next

    Set cT = New Collection
    lC = UBound(m_aStorage)
    '/* only unique keys will be added
    Do
        cT.Add 1, m_aStorage(lC)
        lC = lC - 1
    Loop While lC > 1
    
    '/* return unique items
    Set Remove_Duplicates = cT

On Error GoTo 0

End Function

Private Sub Tags_Init(ByRef aTagArray() As String)
'/* create our tag search collection

Dim lC As Long

    Set c_Key = New Collection
    For lC = 0 To UBound(aTagArray)
        c_Key.Add lC, aTagArray(lC)
    Next lC
    
End Sub

Private Function Key_Update(ByVal sKey As String, _
                            ByVal lIndex As Long)

'/* add a new tag key

On Error Resume Next

    c_Key.Add sKey, lIndex

On Error GoTo 0

End Function

Private Function Get_Index(ByVal lIndex As Long) As Long
'/* return sorted index pointer

    Get_Index = m_lPtr(lIndex)

End Function


'                    <<|>>
'||<<<<<<<<<<<<<<<<< Sort >>>>>>>>>>>>>>>>>||
'                    <<|>>


'* Name           : Sort_Control
'* Purpose        : sorting hub
'* Inputs         : sort type -enum
'* Outputs        : boolean
'*********************************************
Public Function Sort_Control(ByVal lSortType As Long, _
                             Optional ByVal lCol As Long) As Boolean

On Error GoTo Handler

'1GHz celeron, 100fsb, 256 mb/r
'100,000 * len- 8, semi sorted, case sensitive
'Idx TriQuickSort
'avg: 1.59
'Idx Qsort
'avg: 1.36

    '/* test for valid index array
    If Not Array_Check(m_lPtr) Then Load_Index m_lPtr
    '/* array less then min dimensions
    If Not QSI_Check Then GoTo Handler
    '/* default sort
    If lSortType = 0 Then lSortType = 1

    '/* Case - lCp
    '/* &h1 no case, &h0 case(binary)
    '/* Order - lDir
    '/* &h1 ascend, &hffff descend (+1 more, -1 less)
    Select Case lSortType
    '/* ascending case sensitive
    Case 1
        QSI_Sort m_aStorage, m_lPtr, LBound(m_aStorage), UBound(m_aStorage), &H0, &H1
    '/* reverse case sensitive
    Case 2
        QSI_Sort m_aStorage, m_lPtr, LBound(m_aStorage), UBound(m_aStorage), &H0, &HFFFF
    '/* forward case insensitive
    Case 3
        QSI_Sort m_aStorage, m_lPtr, LBound(m_aStorage), UBound(m_aStorage), &H1, &H1
    '/* reverse case insensitive
    Case 4
        QSI_Sort m_aStorage, m_lPtr, LBound(m_aStorage), UBound(m_aStorage), &H1, &HFFFF
    End Select
    
    'TestSort
    
    '/* success
    m_bSorted = True
    Sort_Control = True

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("Sort_Control", Err.Number)

End Function

Private Sub Load_Index(ByRef lPtr() As Long)

'/* preload sort index

Dim l       As Long

    ReDim lPtr(m_lHighCount)
    For l = m_lLowCount To m_lHighCount
        lPtr(l) = l
    Next l

End Sub

Private Sub TestSort()

Dim lC As Long

    For lC = LBound(m_aStorage) To UBound(m_aStorage) Step 1000
        'Debug.Print m_aStorage(lC)
        Debug.Print m_aStorage(m_lPtr(lC))
    Next lC
    
End Sub

Private Sub Key_Reset()
'/* reset tag key container

    Set c_Key = New Collection

End Sub


'                       <<->>
'||<<<<<<<<<<<<<<< IdxQuickSortV5 >>>>>>>>>>>>>>>||
'                       <<->>


Private Sub QSI_Sort(sA() As String, _
                     lIdxA() As Long, _
                     ByVal lbA As Long, _
                     ByVal ubA As Long, _
                     ByVal lCp As Long, _
                     ByVal lDr As Long)

'/* based on the awesome indexed sort by Rde (Rohan) w/ mods

Dim lo          As Long
Dim hi          As Long
Dim cnt         As Long
Dim Item        As String
Dim lpStr       As Long
Dim idxItem     As Long
Dim lpS         As Long

    '/* pre execution check
    If Not UBound(sA) > 0 Then Exit Sub
    '/* Allow for worst case senario + some
    hi = ((ubA - lbA) \ m8) + m32
    '/* Stack to hold pending lower boundries
    ReDim lbs(m1 To hi) As Long
    '/* Stack to hold pending upper boundries
    ReDim ubs(m1 To hi) As Long
    '/* Cache pointer to the string variable                                                '*** Change History ***
    lpStr = GetVarPtr(Item)                                                                 '<- Change 1. VarPtr exchanged for direct access
    '/* Cache pointer to the string array                                                   '<- pointer call in typelib, -GetVarPtr, avoiding runtime
    lpS = GetVarPtr(sA(lbA)) - (lbA * m4)                                                   '<- error checking, and saving a call per execution
                                                                                            '<- this is a considerable gain in a busy loop structure
    '/* Get pivot index position                                                            '<- warning- do not pass a null pointer though!
    Do: hi = ((ubA - lbA) \ m2) + lbA
        '/* Grab current value into item
        CopyMemBv lpStr, lpS + (lIdxA(hi) * m4), m4                                         '<- Change 2. using faster typelib copymembv,
        '/* Grab current index                                                              '<- using typelib bypasses runtime getlasterror
        idxItem = lIdxA(hi): lIdxA(hi) = lIdxA(ubA)                                         '<- check, saving a secondary call on each execution.
        '/* Set bounds                                                                      '<- function call pointer is also resolved directly to
        lo = lbA: hi = ubA                                                                  '<- routine rather then at runtime when using 'declare'
        '/* Storm right in
        Do
            If Not StrComp(Item, sA(lIdxA(lo)), lCp) = lDr Then                             '<- Change 3. simplified not structure means 2 less
                lIdxA(hi) = lIdxA(lo)                                                       '<- instructions per iteration, and 3%+ performance gain
                hi = hi - m1
                Do
                    If Not StrComp(sA(lIdxA(hi)), Item, lCp) = lDr Then
                        lIdxA(lo) = lIdxA(hi)
                        Exit Do
                    End If
                    hi = hi - m1
                Loop Until hi = lo
                '/* Found swaps or out of loop
                If hi = lo Then Exit Do
            End If
            lo = lo + m1
        Loop While hi > lo                                                                  '<- Change 4. Do While/Loop changed to faster Do/Loop While,
        '/* Re-assign current                                                               '<- assembles without additional jmp instruction
        lIdxA(hi) = idxItem
        If (lbA < lo - m1) Then
            If (ubA > lo + m1) Then cnt = cnt + m1: lbs(cnt) = lo + m1: ubs(cnt) = ubA
            ubA = lo - m1
        ElseIf (ubA > lo + m1) Then
            lbA = lo + m1
        Else
            If cnt = m0 Then Exit Do
            lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - m1
        End If
    Loop: CopyMem ByVal lpStr, 0&, m4
    
End Sub

Private Sub QSI_InitLong(ByVal lLb As Long, _
                         ByVal lUb As Long, _
                         ByRef aPtr() As Long)

'/* initialize the pointer array
Dim lC As Long

    Erase aPtr
    ReDim aPtr(lLb To lUb)
    lC = lLb
    
    Do
        aPtr(lC) = lC
        lC = lC + 1
    Loop Until lC = lUb
    
End Sub

Private Function QSI_Check() As Boolean
'/* pretest array properties

On Error GoTo Handler

    If Not UBound(m_aStorage) > 0 Then GoTo Handler
    If Not UBound(m_lPtr) > 0 Then GoTo Handler
    QSI_Check = True

Handler:
    On Error GoTo 0

End Function

Private Sub QSI_Count(Optional ByVal bDecrement As Boolean)
'/* adjust array high count

    If Not bDecrement Then
        m_lHighCount = m_lHighCount + 1
    Else
        m_lHighCount = m_lHighCount - 1
    End If

End Sub

Private Function QSI_Redimn() As Boolean
'/* redimension arrays in blocks

On Error GoTo Handler

    m_lCRef = UBound(m_sArrIdx)
    ReDim Preserve m_sArrIdx(m_lHighCount + 100)
    ReDim Preserve m_lPtr(m_lHighCount + 100)
    
    '/* success
    QSI_Redimn = True
    
Handler:
    On Error GoTo 0

End Function

Private Sub Class_Terminate()

On Error Resume Next

    '/* release all pointers
    DeAllocate_Pointer "k", True

End Sub





'> Not Implemented
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub QS_Sort(sA() As String, _
                    ByVal lbA As Long, _
                    ByVal ubA As Long, _
                    ByVal lCp As Long, _
                    ByVal lDr As Long)

'/* excellent sort by Rde(Rohan), w/ mods (up to 10% faster)

Dim lo      As Long
Dim hi      As Long
Dim cnt     As Long
Dim Item    As String
Dim lpStr   As Long
Dim lp      As Long
Dim lT      As Long
Dim lX      As Long

    '/* pre execution check
    If Not UBound(sA) > 0 Then Exit Sub
    hi = ((ubA - lbA) \ m8) + m32
    ReDim lbs(m1 To hi) As Long
    ReDim ubs(m1 To hi) As Long
    lpStr = GetVarPtr(Item)
    lp = GetVarPtr(sA(lbA)) - (lbA * m4)

    Do
        hi = ((ubA - lbA) \ m2) + lbA
        CopyMemBv lpStr, lp + (hi * m4), m4
        CopyMemBv lp + (hi * m4), lp + (ubA * m4), m4
        lo = lbA
        hi = ubA
        'lX = StrPtr(item)
        Do
            'If Not StrCompC(lX, StrPtr(sA(lo))) = 1 Then
            If Not StrComp(Item, sA(lo), lCp) = lDr Then
                CopyMemBv lp + (hi * m4), lp + (lo * m4), m4
                hi = hi - m1
                Do Until (hi = lo)
                    If Not StrComp(sA(hi), Item, lCp) = lDr Then
                    'If Not StrCompC(StrPtr(sA(hi)), lX) = 1 Then
                        CopyMemBv lp + (lo * m4), lp + (hi * m4), m4
                        Exit Do
                    End If
                    hi = hi - m1
                Loop
                If hi = lo Then Exit Do
            End If
            lo = lo + m1
        Loop While (hi > lo)
        
        CopyMemBv lp + (hi * m4), lpStr, m4
        
        If (lbA < lo - m1) Then
            If (ubA > lo + m1) Then
                cnt = cnt + m1
                lbs(cnt) = lo + m1
                ubs(cnt) = ubA
            End If
            ubA = lo - m1
        ElseIf (ubA > lo + m1) Then
            lbA = lo + m1
        Else
            If cnt = m0 Then Exit Do
            lbA = lbs(cnt)
            ubA = ubs(cnt)
            cnt = cnt - m1
        End If
    Loop

    CopyMem ByVal lpStr, 0&, m4

End Sub

Public Function Build_CtlSafe(ByVal lPointer As Long) As Boolean

'~ Here we cannabilize the structure and
'~ pass the contents to working arrays.
'~ this is much faster then access
'~ times into a complex structure.

Dim i       As Long
Dim lPtr    As Long
Dim lSpr    As Long
Dim lTpr    As Long

    '/* reset container
    If m_bActive Then
        Array_Reset
    End If
    
    '/* initialize local struct
    ReDim m_tHIStc(0)
    '/* copy the structure from the pointer
    CopyMemBr ByVal VarPtrArray(m_tHIStc), lPointer, &H4
    With m_tHIStc(0)
        '/* create the working array
        If Array_Check(.Item) Then
            If ArrayExists(.Item) Then
                '/* high and low count
                Array_Counters .Item
                '/* move the array out of the structure
                CopyMemBr lTpr, ByVal GetASPtr(.Item), &H4
                CopyMemBr ByVal GetASPtr(m_aStorage), lTpr, &H4
                '/* release the structure
                CopyMemBr ByVal GetASPtr(.Item), 0&, &H4
            '/* no items -abort
            Else
                Exit Function
            End If
        End If
        
        '/* test for subitems
        If ArrayExists(.SubItem(0).Text) Then
            ReDim m_tSubArr(LBound(.SubItem) To UBound(.SubItem))
            '/ column count
            m_lColCount = .Column
            m_bDimensional = True
            '/* copy subitems to the temporary container
            CopyMemBr lSpr, ByVal VarPtrArray(.SubItem), &H4
            CopyMemBr ByVal VarPtrArray(m_tSubArr), lSpr, &H4
            '/* release the structure
            CopyMemBr ByVal VarPtrArray(.SubItem), 0&, &H4
        End If

        '/* get search tags
        If Array_Check(.Tag) Then
            If ArrayExists(.Tag) Then
                m_bTags = True
                CopyMemBr lTpr, ByVal GetASPtr(.Tag), &H4
                CopyMemBr ByVal GetASPtr(m_sArrIdx), lTpr, &H4
                '/* release the structure
                CopyMemBr ByVal GetASPtr(.Tag), 0&, &H4
            End If
        End If
    End With
    
    '/* destroy original structure
    CopyMemBr ByVal lPointer, 0&, &H4
    m_bActive = True
    '/* success
    Build_CtlSafe = True
    Erase m_tHIStc
    
    Debug.Print UBound(m_aStorage)
    Debug.Print m_aStorage(UBound(m_aStorage))
    Debug.Print UBound(m_tSubArr)
    Debug.Print m_tSubArr(UBound(m_tSubArr)).Text(0)

On Error GoTo 0
Exit Function

Handler:
    RaiseEvent eSErrCond("Build_Control", Err.Number)
    
End Function
