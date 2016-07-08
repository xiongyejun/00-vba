VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReedSolomon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Const m_nMaxSize As Long = 1024
Private m_objField As clsFiniteField

Private m_nDataCount As Long
Private m_nChecksumCount As Long
Private m_nPrimitiveRoot As Long
Private m_nStartingPower As Long

Private m_nGeneratorPolynomial(m_nMaxSize - 1) As Long

Friend Property Get Field() As clsFiniteField
    Set Field = m_objField
End Property

Friend Property Set Field(ByVal obj As clsFiniteField)
    Set m_objField = obj
End Property

Friend Property Get PrimitiveRoot() As Long
    PrimitiveRoot = m_nPrimitiveRoot
End Property

Friend Property Let PrimitiveRoot(ByVal n As Long)
    m_nPrimitiveRoot = n
End Property

Friend Property Get DataCount() As Long
    DataCount = m_nDataCount
End Property

Friend Property Get ChecksumCount() As Long
    ChecksumCount = m_nChecksumCount
End Property

Friend Property Get StartingPower() As Long
    StartingPower = m_nStartingPower
End Property

Friend Property Get GeneratorPolynomialCoefficient(ByVal nDegree As Long) As Long
    GeneratorPolynomialCoefficient = m_nGeneratorPolynomial(nDegree)
End Property

Friend Function Init(ByVal nDataCount As Long, ByVal nChecksumCount As Long, Optional ByVal objField As clsFiniteField, Optional ByVal nPrimitiveRoot As Long, Optional ByVal nStartingPower As Long) As Boolean
    Dim i As Long, j As Long
    Dim t As Long
    If nDataCount <= 0 Or nChecksumCount <= 0 Or nDataCount + nChecksumCount > m_nMaxSize Then Exit Function
    If nStartingPower < 0 Then Exit Function
    If objField Is Nothing Then Set objField = m_objField _
       Else Set m_objField = objField
    If objField Is Nothing Then Exit Function
    If nPrimitiveRoot = 0 Then nPrimitiveRoot = m_nPrimitiveRoot
    If nPrimitiveRoot = 0 Then nPrimitiveRoot = m_objField.PrimitiveRoot
    If nPrimitiveRoot = 0 Then Exit Function
    m_nDataCount = nDataCount
    m_nChecksumCount = nChecksumCount
    m_nPrimitiveRoot = nPrimitiveRoot
    m_nStartingPower = nStartingPower
    If nStartingPower And 1 Then
        t = nPrimitiveRoot
        nStartingPower = nStartingPower Xor 1
    Else
        t = 1
    End If
    i = 1
    j = nPrimitiveRoot
    Do Until nStartingPower = 0
        i = i + i
        j = objField.MultiplyLUT(j, j)
        If nStartingPower And i Then
            t = objField.MultiplyLUT(t, j)
            nStartingPower = nStartingPower Xor i
        End If
    Loop
    t = objField.Negative(t)
    Erase m_nGeneratorPolynomial
    m_nGeneratorPolynomial(0) = t
    m_nGeneratorPolynomial(1) = 1
    For i = 2 To nChecksumCount
        t = objField.MultiplyLUT(t, nPrimitiveRoot)
        m_nGeneratorPolynomial(i) = 1
        For j = i - 1 To 1 Step -1
            m_nGeneratorPolynomial(j) = objField.Add(m_nGeneratorPolynomial(j - 1), _
                                   objField.MultiplyLUT(m_nGeneratorPolynomial(j), t))
        Next j
        m_nGeneratorPolynomial(0) = objField.MultiplyLUT(m_nGeneratorPolynomial(0), t)
    Next i
    Init = True
End Function

Friend Sub Encode(ByRef nData() As Long)
    Dim t(m_nMaxSize - 1) As Long
    Dim i As Long, j As Long, k As Long
    CopyMemory t(m_nChecksumCount), nData(m_nChecksumCount), m_nDataCount * 4&
    For i = m_nDataCount - 1 To 0 Step -1
        k = t(i + m_nChecksumCount)
        For j = 0 To m_nChecksumCount - 1
            t(i + j) = m_objField.Subtract(t(i + j), m_objField.MultiplyLUT(k, m_nGeneratorPolynomial(j)))
        Next j
    Next i
    For i = 0 To m_nChecksumCount - 1
        nData(i) = m_objField.Negative(t(i))
    Next i
End Sub

