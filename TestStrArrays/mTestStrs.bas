Attribute VB_Name = "mTestStrArrays"
Option Explicit

Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub ZeroMemByV Lib "kernel32" Alias "RtlZeroMemory" (ByVal lpDest As Long, ByVal lLenB As Long)

Private lAbuf() As Long

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' This is intended as a sort testing aid, and though it may be
' a handy technique to use elsewhere, care should be taken to
' ensure no memory violations occur.
'
' Distinctive usage:
'
' - This is intended to be used where the string items and
'   the number of string items in the array are not changing.
'
' - This caches the string pointers only, not the strings,
'   and is intended for use when re-ordering but not altering
'   the string array, so care must be taken to reset the cached
'   pointers whenever array items are added, removed, or modified.
'
' - When caching with CacheArrayPtrs the passed string array
'   must contain at least one item or errors will occur.
'
' - When resetting with ResetArrayPtrs the passed string array
'   size must match the cached size or errors will occur.
'
' - This uses RtlZeroMemory to nullify string pointers but only
'   when calling ResetArrayPtrs with bNullify set to True; see
'   SaveOriginal below. An un-confirmed mis-trust hangs over the
'   use of RtlZeroMemory on some OS's?
'
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub ResetArrayPtrs(sArr() As String, Optional ByVal bNullify As Boolean)

    Dim lpStr As Long, lpBuf As Long
    Dim LBd As Long, UBd As Long

    LBd = LBound(sArr)
    UBd = UBound(sArr)

    lpStr = VarPtr(sArr(LBd))  ' Cache string array pointer

    If bNullify Then
       If (UBd - LBd) Then
          ZeroMemByV lpStr, ((UBd - LBd) + 1&) * 4&
       End If
    Else
       lpBuf = VarPtr(lAbuf(LBd)) ' Cache buffer array pointer
   
       If (UBd - LBd) Then
          CopyMemByV lpStr, lpBuf, ((UBd - LBd) + 1&) * 4&
       End If
    End If
End Sub

Public Sub CacheArrayPtrs(sArr() As String)

    Dim lpStr As Long, lpBuf As Long
    Dim LBd As Long, UBd As Long

    LBd = LBound(sArr)
    UBd = UBound(sArr)

    ReDim lAbuf(LBd To UBd) As Long

    lpStr = VarPtr(sArr(LBd))  ' Cache string array pointer
    lpBuf = VarPtr(lAbuf(LBd)) ' Cache buffer array pointer

    If (UBd - LBd) Then
       CopyMemByV lpBuf, lpStr, ((UBd - LBd) + 1&) * 4&
    End If
End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Demo code only follows:

'Sub LoadArray()
'
'    ' Code to load test strings into array
'    '...
'
'    ' Whenever the array is re-loaded cache the
'    ' pointers in their original order
'    CacheArrayPtrs sTestArray
'
'End Sub

'Sub SortTest()
'
'    ' Reset the array items back to their original
'    ' positions just before each new sorting test
'    ResetArrayPtrs sTestArray
'
'    ' Do the sorting
'    strSort sTestArray
'
'    ' I leave the array sorted to access the sorted data
'    ' and don't reset until starting a new sort test
'
'End Sub

'Sub CommitChanges()
'
'    ' I cache the array pointers when one of the following occurs:
'
'    ' Load the array with new items
'    ' Add item(s) to the array
'    ' Delete item(s) from the array
'    ' Modify item(s) the strings themselves in any way
'    ' Alter the array order and wish to test or save the resulting order
'
'    CacheArrayPtrs sA
'End Sub

'Sub SaveOriginal()
'
'    ' I use the temp array when I don't wish to alter
'    ' the current sort state of the test array
'
'    Dim sTmp() As String
'
'    ' This array must be initialised to size and
'    ' must contain only null string pointers
'
'    ReDim sTmp(lb To ub) As String
'
'    ' This makes an illegal copy
'    ResetArrayPtrs sTmp
'
'    ' Code
'    '...
'
'    ' Must nullify before going out of scope
'    ResetArrayPtrs sTmp, True
'End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
