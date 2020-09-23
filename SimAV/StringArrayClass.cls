VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringArrayClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''Not used Yet.  Will be used to read programs already
'''loaded to determine if infected

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Enum eItemPosition
   ipUseDefault = -1
End Enum

Public Enum eBinarySearchResult
   ebsItemNotFound = -1
End Enum

Private Const BUFFER_SIZE As Long = 100

Private Items() As String
Private Index As Long

Public Property Get Count() As Long

   Count = Index

End Property

Public Property Get Item(Optional Position As eItemPosition = ipUseDefault) As String

   If Index > 0 Then
      If (Position = ipUseDefault) Or (Position = Index - 1) Then
         Item = Items(Index - 1)
      ElseIf (Position < 0) Or (Position > UBound(Items)) Then
         Err.Raise 9
         Exit Property
      Else
         Item = Items(Position)
      End If
   End If
   
End Property

Public Property Let Item(Optional Position As eItemPosition = ipUseDefault, Item As String)

   If (Position = ipUseDefault) Or (Position = Index - 1) Then
      Items(Index - 1) = Item
   ElseIf (Position < 0) Or (Position > UBound(Items)) Then
      Err.Raise 9
      Exit Property
   Else
      Items(Position) = Item
   End If

End Property

Public Sub SetArray(NewItems() As String)

   Items = NewItems
   Index = UBound(Items)

End Sub

Public Sub Clear()

   ReDim Items(BUFFER_SIZE - 1)
   Index = 0

End Sub

Public Sub Add(Item As String, Optional Position As eItemPosition = ipUseDefault)

   Dim l As Long
   
   If UBound(Items) < Index Then
      ReDim Preserve Items(Index + BUFFER_SIZE)
   End If

   If (Position = ipUseDefault) Or (Position = Index) Then
      Items(Index) = Item
   ElseIf (Position < 0) Or (Position > Index) Then
      Err.Raise 9
      Exit Sub
   Else
      CopyMemory ByVal VarPtr(Items(Position + 1)), ByVal VarPtr(Items(Position)), (Index - Position) * 4&
      CopyMemory ByVal VarPtr(Items(Position)), l, 4&
      Items(Position) = Item
   End If

   Index = Index + 1

End Sub

Public Sub Remove(Optional Position As eItemPosition = ipUseDefault)

   Dim l As Long

   If Index > 0 Then
      Index = Index - 1
      
      If (Position = ipUseDefault) Or (Position = Index) Then
      ElseIf (Position < 0) Or (Position > Index) Then
         Err.Raise 9
         Exit Sub
      Else
         l = StrPtr(Items(Position))
         CopyMemory ByVal VarPtr(Items(Position)), ByVal VarPtr(Items(Position + 1)), (Index - Position + 1) * 4
         CopyMemory ByVal VarPtr(Items(Index)), l, 4&
      End If
      
      If (Index + BUFFER_SIZE) = UBound(Items) Then
         ReDim Preserve Items(UBound(Items) - BUFFER_SIZE)
      End If
   End If

End Sub

Public Function Sort(Optional RemoveMatches As Boolean) As StringArrayClass

   Dim s As StringArrayClass
   Dim l As Long
   Dim i As Long
   
   Set s = New StringArrayClass

   For l = 0 To Me.Count - 1
      If s.BinarySearch(Me.Item(l), i) = ebsItemNotFound Then
         s.Add Me.Item(l), i
      ElseIf Not RemoveMatches Then
         s.Add Me.Item(l), i
      End If
   Next

   Set Sort = s
   Set s = Nothing
   
End Function

Public Function BinarySearch(Item As String, Optional SuggestedPosition As Long, Optional PartialMatch As Boolean) As eBinarySearchResult

   Dim LeftIndex As Long
   Dim RightIndex As Long
   Dim Middle As Long
   Dim s As String
   Dim l As Long
   Dim MatchLen As Long
   
   If PartialMatch Then
      MatchLen = Len(Item)
   End If
   
   s = LCase$(Item)
       
   BinarySearch = ebsItemNotFound

   If Index > 0 Then
      If PartialMatch Then
         l = StrComp(LCase$(Left$(Items(Index - 1), MatchLen)), s)
      Else
         l = StrComp(LCase$(Items(Index - 1)), s)
      End If
      If l < 0 Then
         SuggestedPosition = Index
         Exit Function
      End If
   Else
      Exit Function
   End If
   
   LeftIndex = 0
   RightIndex = Index - 1
   
   While LeftIndex <= RightIndex
      Middle = (LeftIndex + RightIndex) \ 2
      If PartialMatch Then
         l = StrComp(LCase$(Left$(Items(Middle), MatchLen)), s)
      Else
         l = StrComp(LCase$(Items(Middle)), s)
      End If
      If l = 0 Then
         BinarySearch = Middle
         SuggestedPosition = Middle
         RightIndex = -1
      ElseIf l < 0 Then
         LeftIndex = Middle + 1
      Else
         RightIndex = Middle - 1
      End If
   Wend
   If BinarySearch = ebsItemNotFound Then
      If l < 0 Then
         SuggestedPosition = Middle + 1
      Else
         SuggestedPosition = Middle
      End If
   End If
   
End Function

Private Sub Class_Initialize()

   Clear

End Sub

Private Sub Class_Terminate()

   Erase Items

End Sub

Public Sub PrintArray()

   Dim l As Long
   
   Debug.Print "++"
   For l = 0 To Me.Count - 1
      Debug.Print "(" & l & ") " & Items(l)
   Next
   Debug.Print "--"

End Sub

