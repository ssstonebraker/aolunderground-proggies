Attribute VB_Name = "mVariables"
Public iRandomCount As Integer
Private m_sComputerName As String
Private m_byCommandLight As Byte
Private m_lTrackLength As Long
Private m_bTrackIsPlaying As Boolean
Private m_bAudioAllOff As Boolean
Private m_bRandomSet As Boolean
Private m_iVolumeSetting As Integer


Property Get sComputerName() As String
    sComputerName = m_sComputerName
End Property
Property Let sComputerName(newValue As String)
    m_sComputerName = newValue
End Property

Property Get byCommandLight() As Byte
    byCommandLight = m_byCommandLight
End Property
Property Let byCommandLight(newValue As Byte)
    m_byCommandLight = newValue
End Property

Property Get lTrackLength() As Long
    lTrackLength = m_lTrackLength
End Property
Property Let lTrackLength(newValue As Long)
    m_lTrackLength = newValue
End Property

Property Get bTrackIsPlaying() As Boolean
    bTrackIsPlaying = m_bTrackIsPlaying
End Property
Property Let bTrackIsPlaying(newValue As Boolean)
    m_bTrackIsPlaying = newValue
End Property

Property Get bAudioAllOff() As Boolean
    bAudioAllOff = m_bAudioAllOff
End Property
Property Let bAudioAllOff(newValue As Boolean)
    m_bAudioAllOff = newValue
End Property

Property Get bRandomSet() As Boolean
    bRandomSet = m_bRandomSet
End Property
Property Let bRandomSet(newValue As Boolean)
    m_bRandomSet = newValue
End Property

Property Get iVolumeSetting() As Integer
    iVolumeSetting = m_iVolumeSetting
End Property
Property Let iVolumeSetting(newValue As Integer)
    m_iVolumeSetting = newValue
End Property
