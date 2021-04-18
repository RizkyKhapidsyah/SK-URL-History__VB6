VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CURLHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'
' URLHistory
'
' URL history collection
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaildlx.com/e_morcillo
'
' Created: 05/31/1999
'
'*********************************************************************************************
Option Explicit

Dim m_History As UrlHistory
Dim m_URLs As Collection

'*********************************************************************************************
'
' AddUrl
' ------
'
' Adds a new URL to the cache
'
' Parameters
' ----------
' URL: The URL to add
' Title: The URL title (TITLE tag of web page)
'
'*********************************************************************************************
Public Sub AddURL(ByVal URL As String, ByVal Title As String, ByVal Flags As Long)

    m_History.AddURL URL, Title, Flags
    
End Sub

'*********************************************************************************************
'
' Count
' -----
'
' Returns the item count
'
'*********************************************************************************************
Public Property Get Count() As Long

    Count = m_URLs.Count
    
End Property

'*********************************************************************************************
'
' Clear
' -----
'
' Clears the URL history
'
'*********************************************************************************************
Public Sub Clear()

    m_History.ClearHistory
    
    Refresh
    
End Sub

'*********************************************************************************************
'
' Refresh
' -------
'
' Reloads the URL list.
'
' Parameters:
' -----------
' URLStartsWith: Indicates the first chars of the URL. Default is "" (All urls)
' Flags: Query flags
'
'*********************************************************************************************
Public Sub Refresh(Optional ByVal URLStartsWith As String, Optional ByVal Flags As STATURL_QUERYFLAGS)
Dim IEnm As IEnumSTATURL, STAT As STATURL
Dim C As Long, URL As URLHistoryItem

    ' Create a new enumerator object
    Set IEnm = m_History.EnumUrls
    
    ' Set the enumeration filter
    IEnm.SetFilter URLStartsWith, Flags
    
    ' Release the previous collection
    ' and create a new one
    Set m_URLs = New Collection
    
    ' Enumerate URLs
    Do Until IEnm.Next(, STAT) = 0
        
        ' Create a new URLHistoryItem
        ' object
        Set URL = New URLHistoryItem
        
        ' Initialize the URL object
        URL.Init STAT
        
        ' Add the URLHistoryItem object
        ' to the collection
        m_URLs.Add URL
        
    Loop
    
    ' Release the enumerator object
    Set IEnm = Nothing
    
End Sub

'*********************************************************************************************
'
' Item
' ----
'
' Returns an URL item
'
'*********************************************************************************************
Public Property Get Item(ByVal Index As Long) As URLHistoryItem

    Set Item = m_URLs.Item(Index)
    
End Property

'*********************************************************************************************
'
' Remove
' ------
'
' Removes an URL from the cache by its index
'
'*********************************************************************************************
Public Property Get Remove(ByVal Index As Long) As URLHistoryItem
Dim URL As URLHistoryItem

    Set URL = m_URLs.Item(Index)
    
    m_History.DeleteURL URL.URL, 0
    m_URLs.Remove Index
    
    Set URL = Nothing
    
End Property

'*********************************************************************************************
'
' DeleteURL
' ---------
'
' Removes an URL from the cache by URL
'
'*********************************************************************************************
Public Sub DeleteURL(ByVal URL As Long, ByVal Flags As Long)

    m_History.DeleteURL URL, Flags
    
    Refresh
    
End Sub

Private Sub Class_Initialize()

    Set m_History = New UrlHistory
    
    Refresh
    
End Sub

Private Sub Class_Terminate()

    Set m_History = Nothing
   
End Sub

Public Function NewEnum() As Variant
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = m_URLs.[_NewEnum]
    
End Function

