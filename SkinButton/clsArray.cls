VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' API    :   CopyMemory (defined in modDrawAPI.bas)
' Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Option Explicit
Private Type ListType
    hWnd As Long
    PrevProc As Long
    Obj As SkinnableButton
End Type
Private ObjectsList() As ListType
Private ObjectsCount As Long

Public Property Get Count() As Long
    Count = GetObjectsCount
End Property

Private Function GetObjectsCount() As Long
    GetObjectsCount = ObjectsCount
End Function
Public Sub AddObject(hWnd As Long, PrevProc As Long, Obj As SkinnableButton)
CheckBound
    Dim FirstEm As Long
    FirstEm = FindEmptyList
    If FirstEm > 0 Then
        ObjectsList(FirstEm).hWnd = hWnd
        ObjectsList(FirstEm).PrevProc = PrevProc
        Set ObjectsList(FirstEm).Obj = Obj
        ObjectsCount = ObjectsCount + 1
    End If
End Sub
Public Sub RemoveObject(hWnd As Long)
    Dim nGetObj
    nGetObj = FindObj(hWnd)
    If nGetObj > 0 Then
        ObjectsList(nGetObj).hWnd = 0
        ObjectsList(nGetObj).PrevProc = 0
        Set ObjectsList(nGetObj).Obj = Nothing
        ObjectsCount = ObjectsCount - 1
    End If
End Sub
Public Function GetProc(hWnd As Long) As Long
    Dim nGetObj As Long
    nGetObj = FindObj(hWnd)
    If nGetObj > 0 Then
        GetProc = ObjectsList(nGetObj).PrevProc
    End If
End Function
Public Function GetObj(hWnd As Long) As SkinnableButton
    Dim nGetObj As Long
    nGetObj = FindObj(hWnd)
    If nGetObj > 0 Then
        Set GetObj = ObjectsList(nGetObj).Obj
    End If
End Function


Private Sub CheckBound()

On Error GoTo ResCheck
Dim nObjs As Long
nObjs = GetObjectsCount
    If nObjs = 0 Then
        ReDim ObjectsList(1 To 1)
    Else
        If nObjs = UBound(ObjectsList) Then
            Dim nBound As Long
            nBound = nObjs + 5
            ReDim Preserve ObjectsList(1 To nBound)
        End If
    End If
    Dim EmptyLast As Long
    EmptyLast = FindEmptiesFromBottom
    If EmptyLast > 5 Then
        nObjs = UBound(ObjectsList) - EmptyLast + 5
        ReDim Preserve ObjectsList(1 To (nObjs))
    End If
ResCheck:
End Sub
Private Function FindEmptyList()
On Error Resume Next
    Dim i As Long
    For i = 1 To ObjectsCount + 1
        If ObjectsList(i).hWnd = 0 Then
            FindEmptyList = i
            Exit Function
        End If
    Next
    FindEmptyList = -1
End Function

Private Function FindObj(hWnd As Long) As Long
    Dim i As Long
    For i = 1 To UBound(ObjectsList)
        If ObjectsList(i).hWnd = hWnd Then
            FindObj = i
            Exit Function
        End If
    Next
    FindObj = -1
End Function

Private Function FindEmptiesFromBottom() As Long
    Dim i As Long, nEmpties As Long
    For i = UBound(ObjectsList) To 1 Step -1
        If ObjectsList(i).hWnd = 0 Then
            nEmpties = nEmpties + 1
        Else
           FindEmptiesFromBottom = nEmpties
           Exit Function
        End If
    Next
    FindEmptiesFromBottom = -2
End Function
Public Sub RemoveAll()
    ReDim ObjectsList(0)
    ObjectsCount = 0
End Sub
