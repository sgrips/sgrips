Attribute VB_Name = "mListviewCalls"
Option Explicit

Public Enum typeFieldDefinition
    typeFieldDef_CHAR = 0
    typeFieldDef_INTEGER = 1
    typeFieldDef_DECIMAL = 2
    typeFieldDef_DATETIME = 3
    typeFieldDef_BOOLEAN = 4
    typeFieldDef_DOUBLE = 5
    
    typeFieldDef_ICON = -100
End Enum


Private Const IDC_WAIT = 32514&   ' vbHourglass
Private Const IDC_ARROW = 32512&   ' vbDefault
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long

'
' Copyright © 1997-2001 Brad Martinez, http://www.mvps.org
'
Private Const WM_DESTROY = &H2

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const OLDWNDPROC = "OldWndProc"
'

'
' Copyright © 1997-2001 Brad Martinez, http://www.mvps.org
'
Private phwndLV As Long
Private pfUndocSort As Boolean

' User-defined flag passed to and used in ListViewCompareProc
' indicating a descending sort (lvwDescending)
Private Const SORT_DESCENDING = &H80000000

' User-defined value used to mask the zero based column index values
' (below) from the lParamSort param value passed to ListViewCompareProc's .
Private Const SORT_COLUMNMASK = &HFF

Public Enum LVSortDirections
  lvsdAscending = lvwAscending  ' = 0
  lvsdDescending = lvwDescending  ' = 1
  lvsdToggle = 2
End Enum

' user-defined value returned by many listview messages indicating
' the index of no valid listview item
Private Const LVI_NOITEM = -1

' user-defined value specifying an arbitrary maximum listview item and
' subitem text length
Private Const LV_CB_ITEMTEXT = 32

Private Type POINTAPI
  x As Long
  y As Long
End Type

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

' ==================================================================
' listview defs

Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEM = (LVM_FIRST + 5)
Private Const LVM_FINDITEM = (LVM_FIRST + 13)
Private Const LVM_GETITEMTEXT = (LVM_FIRST + 45)
Private Const LVM_SORTITEMS = (LVM_FIRST + 48)

Private Type LVITEM   ' was LV_ITEM
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As String
  cchTextMax As Long
  iImage As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iIndent As Long
#End If
End Type

' LVITEM mask
Private Const LVIF_PARAM = &H4

Private Type LVFINDINFO   ' was LV_FINDINFO
  flags As LVFINDINFO_flags
  psz As Long  ' if String, must be pre-allocated
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type
 
Private Enum LVFINDINFO_flags
  LVFI_PARAM = &H1
  LVFI_STRING = &H2
  LVFI_PARTIAL = &H8
  LVFI_WRAP = &H20
  LVFI_NEARESTXY = &H40
End Enum

Private pAPISort As Boolean
Private pFixIndices As Boolean
Private pListView As ListView
'
'
'' Returns the value of the specifed zero-based listview item's lParam.
'
'Public Function GetLVItemLParam(hwndLV As Long, iItem As Long) As Long
'  Dim lvi As LVITEM
'
'  lvi.mask = LVIF_PARAM
'  lvi.iItem = iItem
'  ' Returns TRUE if successful or FALSE otherwise.
'  If ListView_GetItem(hwndLV, lvi) Then
'    GetLVItemLParam = lvi.lParam
'  End If
'
'End Function

' Returns an AddRef'd ListItem object reference from the ListItem's
' LVITEM lParam value.

' ======================================================
' For both the Mscomctl.ocx and Comctl32.ocx TreeView and ListView
' controls, the Node and ListItem's ObjPtr() values reside at the 3rd
' DWORD (@ byte offset 8) in the Node and ListItem's lParam.
'
' Is highly undocumented, for more info, see the TVItemData demo at
' http://www.mvps.org/btmtz/treeview/
' =====================================================

Private Function GetListItemFromlParam(lParam As Long) As listitem
  Dim lpli As Long
  Dim li As listitem
  
  If lParam Then
    Call MoveMemory(lpli, ByVal lParam + 8, 4)
    If lpli Then
      Call MoveMemory(li, lpli, 4)
      Set GetListItemFromlParam = li
      ' li is not AddRef'd, so we have to zero it before this proc ends,
      ' or it will be Released (causing a GPF) when it goes out of scope.
      Call MoveMemory(li, 0&, 4)
    End If
  End If
  
End Function

' Application-defined callback function, which is called by the listview during
' a sort operation each time the relative order of two listview items needs to
' be compared. (see the desciption of LVM_SORTITEMS in the SDK)

'    lParam1      - the 1st item's LVITEM lParam value
'    lParam2      - the 2nd item's LVITEM lParam value
'    lParamSort  - application-defined value that is passed to the comparison function.

' The callback function must return a negative value if the first item should
' precede the second, a positive value if the first item should follow the second,
' or zero if the two items are equivalent.

' Invoked by ListView_SortItems() below, proc is static for performance

Private Static Function ListViewCompareProc(ByVal lParam1 As Long, _
                                                                      ByVal lParam2 As Long, _
                                                                      ByVal lParamSort As Long) As Long
  Dim iColumn As Long
  Dim dwRtn As Long
  Dim lvfi As LVFINDINFO
  Dim s1 As String * LV_CB_ITEMTEXT
  Dim s2 As String * LV_CB_ITEMTEXT
  
  iColumn = lParamSort And SORT_COLUMNMASK
  
  Dim typeCol As typeFieldDefinition
  
  typeCol = toZero(getValueInTag(pListView.ColumnHeaders(iColumn + 1).Tag, "FIELDTYPE"))
  
  If pfUndocSort Then
    Select Case typeCol
      Case typeFieldDefinition.typeFieldDef_CHAR, typeFieldDef_ICON     ' this is slower than allowing the ListView to do its default text sort...
        dwRtn = StrComp(getValueItem(lParam1, iColumn, typeCol), _
                                     getValueItem(lParam2, iColumn, typeCol))
      
      Case typeFieldDefinition.typeFieldDef_DECIMAL, _
        typeFieldDefinition.typeFieldDef_DOUBLE, _
        typeFieldDefinition.typeFieldDef_INTEGER, _
        typeFieldDefinition.typeFieldDef_BOOLEAN
        dwRtn = getValueItem(lParam1, iColumn, typeCol) - _
                       getValueItem(lParam2, iColumn, typeCol) ' can possibly overflow dwRtn...
      
      Case typeFieldDefinition.typeFieldDef_DATETIME    ' see Form_Load
        dwRtn = CDate(Format$(getValueItem(lParam1, iColumn, typeCol), "yyyy-mm-dd hh.nn.ss")) - _
                       CDate(Format$(getValueItem(lParam2, iColumn, typeCol), "yyyy-mm-dd hh.nn.ss"))  ' can possibly overflow dwRtn...
                                
    End Select   ' iColumn
      
  Else
    ' use the documented (and painfully slow) way to obtain listview item or subitem
    ' text from the LVITEM.lParam, as is demonstrated in
    ' http://support.microsoft.com/support/kb/articles/Q170/8/84.asp
    lvfi.flags = LVFI_PARAM
    lvfi.lParam = lParam1
    Call ListView_GetItemText(phwndLV, ListView_FindItem(phwndLV, -1, lvfi), iColumn, s1)
    lvfi.lParam = lParam2
    Call ListView_GetItemText(phwndLV, ListView_FindItem(phwndLV, -1, lvfi), iColumn, s2)
  
    Select Case typeCol
      Case typeFieldDefinition.typeFieldDef_CHAR
        dwRtn = StrComp(s1, s2)
      Case typeFieldDefinition.typeFieldDef_DECIMAL, _
        typeFieldDefinition.typeFieldDef_DOUBLE, _
        typeFieldDefinition.typeFieldDef_INTEGER, _
        typeFieldDefinition.typeFieldDef_BOOLEAN
        dwRtn = Val(s1) - Val(s2)
      Case typeFieldDefinition.typeFieldDef_DATETIME
        dwRtn = CDate(Format$(s1, "yyyy-mm-dd hh.nn.ss")) - CDate(Format$(s2, "yyyy-mm-dd hh.nn.ss"))
    End Select   ' iColumn
  
  End If   ' g_fUndocSort

  ' if lvwDescending, reverse the sign of the return value.
  If (lParamSort And SORT_DESCENDING) Then
    ListViewCompareProc = -dwRtn
  Else
    ListViewCompareProc = dwRtn
  End If

End Function

Private Function getValueItem(liParam As Long, colID As Long, typef As typeFieldDefinition) As String
    Dim li As listitem
    
    Set li = GetListItemFromlParam(liParam)
    If colID = 0 Then
        If typef = typeFieldDef_ICON Then
            getValueItem = li.Icon
        Else
            getValueItem = li.Text
        End If
    Else
        If typef = typeFieldDef_ICON Then
            getValueItem = li.ListSubItems(colID).ReportIcon
        Else
            getValueItem = li.ListSubItems(colID)
        End If
    End If
    
    
    
    
End Function


' ==================================================================
' listview macros

' a VB string adapted version of the real macro

Private Function ListView_GetItemText(hwndLV As Long, _
                                                            i As Long, _
                                                            iSubItem As Long, _
                                                            sText As String) As Long
  Dim lvi As LVITEM
    
  ' instead of assigning lvi.pszText = sText, which allocates another string
  ' buffer for lvi.pszText and copies sText to it, simply copy VarPtr(sText)
  ' to VarPtr(lvi.pszText), so both variables contain the value StrPtr(sText)
  Call MoveMemory(ByVal VarPtr(lvi.pszText), ByVal VarPtr(sText), 4)
  lvi.cchTextMax = Len(sText)
  lvi.iSubItem = iSubItem
  
  ' though LVM_GETITEMTEXT rtns cbChars copied to lvi.pszText, the real macro rtns no value
  ListView_GetItemText = SendMessage(hwndLV, LVM_GETITEMTEXT, ByVal i, lvi)
  
  ' clear VarPtr(lvi.pszText) so VB doesn't free StrPtr(sText) when this proc ends
  Call MoveMemory(ByVal VarPtr(lvi.pszText), 0&, 4)

End Function
'
'Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
'  ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pitem)
'End Function

Private Function ListView_FindItem(hWnd As Long, iStart, plvfi As LVFINDINFO) As Long
  ListView_FindItem = SendMessage(hWnd, LVM_FINDITEM, ByVal iStart, plvfi)
End Function

Private Function ListView_SortItems(hwndLV As Long, pfnCompare As Long, lParamSort As Long) As Boolean
  ListView_SortItems = SendMessage(hwndLV, LVM_SORTITEMS, ByVal lParamSort, ByVal pfnCompare)
End Function



Private Function SubClass(hWnd As Long, lpfnNew As Long) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  
  If GetProp(hWnd, OLDWNDPROC) Then
    fSuccess = True
  Else
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)
    If lpfnOld Then
      fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
    End If
  End If
  
  If fSuccess Then
    SubClass = True
  Else
    If lpfnOld Then Call UnSubClass(hWnd)
    MsgBox "Unable to successfully subclass &H" & Hex(hWnd), vbCritical
  End If
  
End Function

Private Function UnSubClass(hWnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, OLDWNDPROC)
  If lpfnOld Then
    If RemoveProp(hWnd, OLDWNDPROC) Then
      UnSubClass = SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    End If
  End If

End Function

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg

    ' ======================================================
    ' If the specified menuitems are checked, replace the pointer to the
    ' ListView's sort callback procedure with that of our own, and set the
    ' the wParam value (the callback procedue's lParamSort value) to
    ' the values our sort procedure wants to see (the same values that
    ' are set in Form1.ListView1_ColumnClick when mnuOptsFixIndices
    ' .Checked = False).
    
    Case LVM_SORTITEMS
        If (pAPISort And pFixIndices) Then
          wParam = pListView.SortKey Or (CBool(pListView.SortOrder) And SORT_DESCENDING)
          lParam = FARPROC(AddressOf ListViewCompareProc)
        End If
    
    ' ======================================================
    ' in case it wasn't already done, Unsubclass the hWnd window...
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
      
  End Select
  
  WndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
  
End Function

' A dummy procedure that receives and returns the result
' of the AddressOf operator

Public Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function


' ======================================================================
' ListView1 sorting

' Toggles the sort order, and sorts the ListView's items or subitems under
' the respectively clicked ColumnHeader.


' iColumn is one-based ColumnHeader.Index

Public Sub SortListview(lv As ListView, iColumn As Long, dwSort As LVSortDirections, APISort As Boolean, FixIndices As Boolean, m_cHdrIcons As cLVHeaderSortIcons)
  Dim ch As ColumnHeader
  Dim dwTime As Long
  
  Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
  
  pAPISort = APISort
  pFixIndices = FixIndices
  phwndLV = lv.hWnd
  pfUndocSort = True
  
  Set pListView = lv
  
  With lv
    If (iColumn = -1) Then
      ' mnuViewArrangeAscending_Click, mnuViewArrangeDescending_Click:
      ' as if the active column's header was clicked (toggle the column's
      ' sort order, ColumnHeaders are one-based, SortKey is zero-based)
      Set ch = .ColumnHeaders(.SortKey + 1)
    Else
      ' ListView1_ColumnClick, mnuViewArrangeCol_Click:
      ' as if a column other than the active column's header was clicked
      ' (don't toggle), its sort order is determined by the currently checked
      ' mnuViewArrangeAscending/mnuViewArrangeDescending menus
      Set ch = .ColumnHeaders(iColumn)
    End If
    
    ' Toggle the clicked column's sort order only if the active column is clicked
    ' (iow, don't reverse the sort order when different columns are clicked,
    ' ColumnHeader indices are one-based, .SortKey is zero-based).
    If ((.SortKey = (iColumn - 1)) And (dwSort = lvsdToggle)) Then
      ch.Tag = setValueInTag(ch.Tag, "SORT", Abs(Not CBool(Val(getValueInTag(ch.Tag, "SORT")))))           ' have to use Val since ch.Tag starts as ""
    ElseIf (dwSort <> lvsdToggle) Then
      ch.Tag = setValueInTag(ch.Tag, "SORT", Abs(CBool(dwSort)))
    End If
    
    ' (if .Sorted = True, the ListView will be sorted when SortOrder is set)
    .Sorted = False
    
    ' Set the ListView's sort order to the ColumnHeader.Tag value determined
    ' above and the sort key to the ColumnHeader's zero-based index
    ' (ColumnHeader.Index is one-based).
    .SortOrder = Val(getValueInTag(ch.Tag, "SORT"))
    .SortKey = ch.Index - 1
    
    Select Case True
      Case APISort And FixIndices
        ' Subclass the ListView, and set it's Sorted property. When the ListView
        ' sends a LVM_SORTITEMS, intercept it and replace the ListView's
        ' internal sort procedure with our own. Once the all of the sort callbacks
        ' are done processing, the ListView will then re-order all ListItem.Index
        ' properties to that of each ListItem's new one-based in the ListView.
        Call SubClass(phwndLV, AddressOf WndProc)
        .Sorted = True
        Call UnSubClass(phwndLV)
      
      Case APISort
        ' Sort the ListView ourselves, bypassing the ListView's Sorted property.
        ' All ListItem.Index values will remain as they were before the sort. Pass
        ' the zero-based column header index and the sort order as a flag
        ' (&H80000000) to be used in ListViewCompareProc
        Call ListView_SortItems(phwndLV, AddressOf ListViewCompareProc, _
                                             .SortKey Or (CBool(.SortOrder) And SORT_DESCENDING))
      
      Case Else
        .Sorted = True
      
    End Select   ' True
    
    ' Set the header icons
    Call m_cHdrIcons.SetHeaderIcons(.SortKey, .SortOrder)
  End With   ' lv
  
  Call SetCursor(LoadCursor(0, ByVal IDC_ARROW))

End Sub


Public Sub sortListViewDestroy()
    Set pListView = Nothing
End Sub



Public Function setValueInTag(ByVal sTAG As String, ByVal key As String, ByVal Value As String) As String
    'TAG = @KEY=VALUE@KEY1=VALUE1
    Dim r As Object
    
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = "(@" & key & ")=([^@]+)"
    r.IgnoreCase = True
    
    If Not r.Test(sTAG) Then
        sTAG = sTAG & "@" & key & "=" & Value
    Else
        sTAG = r.Replace(sTAG, "$1=" & Value)
    End If
    
    setValueInTag = sTAG
End Function

Public Function getValueInTag(ByVal sTAG As String, ByVal key As String) As String
    'TAG = @KEY=VALUE@KEY1=VALUE1
    Dim r As Object
    Dim matchs As Object
    Dim matcho As Object
    
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = "(@" & key & ")=([^@]+)"
    r.IgnoreCase = True
    Set matchs = r.Execute(sTAG)
    
    For Each matcho In matchs
        getValueInTag = matcho.SubMatches(1)
    Next
End Function

