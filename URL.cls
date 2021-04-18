VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "URLHistoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'
' URLHistory
'
' URL history item
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

Dim m_STAT As STATURL
'*********************************************************************************************
'
' Ptr2StrW
'
' Returns a string from a unicode string pointer.
'
'*********************************************************************************************
Private Function Ptr2StrW(Ptr As Long) As String
Dim sRtn() As Byte

    ' Check if the pointer is valid
    If Ptr <> 0 Then
    
        ReDim sRtn(lstrlenW(ByVal Ptr) * 2)
        
        ' Copy the string to the byte array
        MoveMemory sRtn(0), ByVal Ptr, UBound(sRtn)
        Ptr2StrW = sRtn()
        
    End If
    
End Function

'*********************************************************************************************
'
' Init
'
' Initializes the object data
'
'*********************************************************************************************
Friend Sub Init(STAT As STATURL)

    m_STAT = STAT
    
End Sub

'*********************************************************************************************
'
' LastVisited
'
' Returns the last visited date
'
'*********************************************************************************************
Public Property Get LastVisited() As Date
Dim ST As SYSTEMTIME, SerialDate As Double

    FileTimeToLocalFileTime m_STAT.ftLastVisited, m_STAT.ftLastVisited
    FileTimeToSystemTime m_STAT.ftLastVisited, ST
    SystemTimeToVariantTime ST, SerialDate
    
    LastVisited = SerialDate
    
End Property

'*********************************************************************************************
'
' Expires
'
'*********************************************************************************************
Public Property Get Expires() As Date
Dim ST As SYSTEMTIME, SerialDate As Double

    FileTimeToLocalFileTime m_STAT.ftExpires, m_STAT.ftExpires
    FileTimeToSystemTime m_STAT.ftExpires, ST
    SystemTimeToVariantTime ST, SerialDate
    
    Expires = SerialDate
    
End Property

'*********************************************************************************************
'
' LastUpdated
'
' Returns the last update date of the URL
'
'*********************************************************************************************
Public Property Get LastUpdated() As Date
Dim ST As SYSTEMTIME, SerialDate As Double

    FileTimeToLocalFileTime m_STAT.ftLastUpdated, m_STAT.ftLastUpdated
    FileTimeToSystemTime m_STAT.ftLastUpdated, ST
    SystemTimeToVariantTime ST, SerialDate
    
    LastUpdated = SerialDate
    
End Property

'*********************************************************************************************
'
' URL
'
' Returns the URL
'
'*********************************************************************************************
Public Property Get URL() As String

    URL = Ptr2StrW(m_STAT.pwcsUrl)
    
End Property

'*********************************************************************************************
'
' Flags
'
' Returns the URL flags
'
'*********************************************************************************************
Public Property Get Flags() As STATURLFLAGS

    Flags = m_STAT.dwFlags
    
End Property

'*********************************************************************************************
'
' Title
'
' Returns the URL title
'
'*********************************************************************************************
Public Property Get Title() As String

    Title = Ptr2StrW(m_STAT.pwcsTitle)
    
End Property

Private Sub Class_Terminate()

    ' Release the strings from
    ' m_STAT variable
    CoTaskMemFree m_STAT.pwcsTitle
    CoTaskMemFree m_STAT.pwcsUrl
    
End Sub
