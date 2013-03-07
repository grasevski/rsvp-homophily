VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Result generator module
' Just create and init a new ResultGenerator and call generateResults
'
' N.B. The following dependencies must be added for this class to
' work (Tools->References...):
'   Microsoft Scripting Runtime
'   Microsoft ActiveX Data Objects 2.8 Library
Option Explicit

' Genders can be chosen
Enum Gender
    Male
    Female
End Enum

' Attribute (id,name) pair
Private Type AttributeValue
    ix As Long
    name As String
End Type

' Sender Interest LHS data fields
Private Enum LhsField
    NLhsIRhsS
    ULhsI
    ULhsS
End Enum

' Sender Interest RHS data fields
Private Enum RhsField
    Rhs
    URhsS
End Enum

' Sheets and tables used in sql queries to generate results
Private Const PROPERTY_SHEET = "PJT2_TR_PROPERTY"
Private Const PROPERTY_TABLE = "[" & PROPERTY_SHEET & "$]"
Private Const METRIC_SHEET = "PJT2_TR_METRIC"
Private Const METRIC_TABLE = "[" & METRIC_SHEET & "$]"
Private Const USERS_SHEET = "RHSUSERS_BY_SENDER"
Private Const USERS_TABLE = "[" & USERS_SHEET & "$]"

' Various offsets and positionings for the results tables
Private Const PAGE_SPLIT_OFFSET = 5
Private Const SENDER_INTEREST_OFFSET = 10
Private Const EXPECTED_SENDER_INTEREST_OFFSET = 4
Private Const RELATIVE_SENDER_INTEREST_OFFSET = 5
Private Const RECEIVER_INTEREST_OFFSET = 6
Private Const EXPECTED_RECEIVER_INTEREST_OFFSET = 4
Private Const RELATIVE_RECEIVER_INTEREST_OFFSET = 5
Private Const HARMONIC_MEAN_OFFSET = 5
Private Const GEOMETRIC_MEAN_OFFSET = 6
Private Const INFO_OFFSET = 2
Private Const RECEIVER_DATA_SPACE = 3
Private Const HOMOPHILY_DATA_SPACE = 5
Private Const HOMOPHILY_RESULTS_OFFSET = 5
Private Const Y0 = 5
Private Const X0 = 4

' Gender constants
Private Const GENDER_OFFSET = 134
Private Const NUM_GENDERS = 2

' Data structures used for querying results
Private cn As New ADODB.Connection
Private rs As New ADODB.Recordset
Private attributeValues() As AttributeValue
Private attributeIndices As New Dictionary
Private genders(1 To NUM_GENDERS) As String
Private gender1 As Gender, gender2 As Gender

' Initialize the SQL connection and properties query
Private Sub Class_Initialize()
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
        & ActiveWorkbook.FullName _
        & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
    rs.Open "select distinct propertyid, fieldname from " _
        & PROPERTY_TABLE, cn
    genders(1) = "Male"
    genders(2) = "Female"
End Sub

' Close the properties query and the SQL connection
Private Sub Class_Terminate()
    rs.Close
    cn.Close
End Sub

' Main program - removes previous results sheets and creates new ones
Sub generateResults(Optional g1 As Gender = Male, Optional g2 As Gender = Female)
    Const NOTIMPORTANT_IX = 0
    Const NOTIMPORTANT_NAME = "Not Important"
    Const NULL_IX = 1000000
    Const NULL_NAME = "Null"
    Dim s As Worksheet
    Dim rsAttribute As ADODB.Recordset
    Dim cmdAttribute As New ADODB.Command
    Dim i As Integer
    Dim zeroIncluded As Boolean
    For Each s In Sheets
        If s.name <> PROPERTY_SHEET _
        And s.name <> METRIC_SHEET _
        And s.name <> USERS_SHEET _
        Then s.Delete
    Next
    cmdAttribute.ActiveConnection = cn
    gender1 = g1
    gender2 = g2
    While Not rs.EOF
        cmdAttribute.CommandText = "select count(*) from " & _
            PROPERTY_TABLE & " where propertyid=? and valueid=0"
        cmdAttribute.Parameters(0).Type = adInteger
        cmdAttribute.Parameters(0) = rs.Fields(0)
        Set rsAttribute = cmdAttribute.Execute
        zeroIncluded = rsAttribute.Fields(0) > 0
        rsAttribute.Close
        
        cmdAttribute.CommandText = _
            "select count(*) from (select distinct valueid from " _
            & PROPERTY_TABLE & " where propertyid=? and valueid<>0)"
        cmdAttribute.Parameters(0).Type = adInteger
        cmdAttribute.Parameters(0) = rs.Fields(0)
        Set rsAttribute = cmdAttribute.Execute
        ReDim attributeValues(1 To rsAttribute.Fields(0) + 2)
        attributeIndices.RemoveAll
        rsAttribute.Close
        
        cmdAttribute.CommandText = _
            "select distinct valueid, valuename from " _
            & PROPERTY_TABLE & _
            " where propertyid=? order by valueid"
        cmdAttribute.Parameters(0).Type = adInteger
        cmdAttribute.Parameters(0) = rs.Fields(0)
        Set rsAttribute = cmdAttribute.Execute
        If zeroIncluded Then
            attributeValues(1).name = rsAttribute.Fields(1)
            rsAttribute.MoveNext
        Else
            attributeValues(1).name = NOTIMPORTANT_NAME
        End If
        attributeValues(1).ix = 0
        attributeIndices.Add "ix0", 1
        For i = 2 To UBound(attributeValues) - 1
            attributeValues(i).ix = rsAttribute.Fields(0)
            attributeValues(i).name = rsAttribute.Fields(1)
            attributeIndices.Add "ix" & attributeValues(i).ix, i
            rsAttribute.MoveNext
        Next
        rsAttribute.Close
        
        attributeValues(UBound(attributeValues)).ix = NULL_IX
        attributeValues(UBound(attributeValues)).name = NULL_NAME
        attributeIndices.Add "ix" & attributeValues(UBound(attributeValues)).ix, UBound(attributeValues)
        Sheets.Add after:=Sheets(Sheets.Count)
        ActiveSheet.name = rs.Fields(1)
        senderInterest
        receiverInterest
        combineResults
        siSri
        homophilyMetric
        rs.MoveNext
    Wend
End Sub

' Writes sender interest tables to the sheet
Private Sub senderInterest()
    Dim anchorRange As Range
    Set anchorRange = Range("A1")
    anchorRange = "Sender Interest"
    anchorRange.Font.Bold = True
    Set anchorRange = anchorRange.Offset(Y0, X0)
    senderInterestData anchorRange
    senderInterestResults anchorRange
    addColumnRowInfo _
        anchorRange.Offset(4 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + INFO_OFFSET)
End Sub

' Writes sender interest data tables to the sheet
Private Sub senderInterestData(anchorRange As Range)
    Dim rsMetric As ADODB.Recordset
    Dim cmdMetric As New ADODB.Command
    
    cmdMetric.ActiveConnection = cn
    
    anchorRange.Offset(-1, -2) = "Data (LHS2RHS)"
    
    anchorRange.Offset(-1) = "n(LHSi,RHSj)"
    anchorRange.Offset(NLhsIRhsS + UBound(attributeValues) + 1, 1).Formula _
        = "=sum(" & _
        columnRange(anchorRange).Address(False, False) & ")"
    rowRange(anchorRange.Offset(NLhsIRhsS + UBound(attributeValues))).FillRight
    outlineTable anchorRange
    
    anchorRange.Offset(-1, UBound(attributeValues) + PAGE_SPLIT_OFFSET) _
        = "n(LHSs,RHSj)"
    anchorRange.Offset(NLhsIRhsS + UBound(attributeValues) + 1, UBound(attributeValues) + PAGE_SPLIT_OFFSET + 1).Formula _
        = "=sum(" & _
        columnRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False) _
        & ")"
    rowRange(anchorRange.Offset(NLhsIRhsS + UBound(attributeValues), UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FillRight
    rowRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = "=sum(" & rowRange(anchorRange).Address(False, False) _
        & ")"
    tableRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FillDown
    outlineTable anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)
    anchorRange.Offset(, Rhs + UBound(attributeValues) + 1) _
        = "RHS"
    anchorRange.Offset(, URhsS + UBound(attributeValues) + 1) _
        = "u(RHSs)"
    anchorRange.Offset(1, URhsS + UBound(attributeValues) + 1).Formula _
        = "=sum(" & _
        columnRange(anchorRange.Offset(, Rhs + UBound(attributeValues))).Address(False, False) _
        & ")"
    anchorRange.Offset(NLhsIRhsS + UBound(attributeValues) + 1) _
        = "n(LHSi,RHSs)"
    anchorRange.Offset(ULhsI + UBound(attributeValues) + 1) _
        = "u(LHSi)"
    anchorRange.Offset(ULhsS + UBound(attributeValues) + 1) _
        = "u(LHSs)"
    anchorRange.Offset(ULhsS + UBound(attributeValues) + 1, 1).Formula _
        = "=sum(" & _
        rowRange(anchorRange.Offset(ULhsI + UBound(attributeValues))).Address(False, False) _
        & ")"
    anchorRange.Offset(UBound(attributeValues) + 1, UBound(attributeValues) + PAGE_SPLIT_OFFSET) _
        = "n(LHSs,RHSs)"
    
    cmdMetric.CommandText = _
        "select lhs_value_id, rhs_value_id, lhs2rhs from " _
        & METRIC_TABLE & _
        " where attribute_id=? and lhs_gender=? and rhs_gender=?"
    cmdMetric.Parameters(0).Type = adInteger
    cmdMetric.Parameters(0) = rs.Fields(0)
    cmdMetric.Parameters(1).Type = adInteger
    cmdMetric.Parameters(1) = gender1 + GENDER_OFFSET
    cmdMetric.Parameters(2).Type = adInteger
    cmdMetric.Parameters(2) = gender2 + GENDER_OFFSET
    Set rsMetric = cmdMetric.Execute
    While Not rsMetric.EOF
        If attributeIndices.Exists("ix" & rsMetric.Fields(0)) _
        And attributeIndices.Exists("ix" & rsMetric.Fields(1)) _
        Then _
            anchorRange.Offset(attributeIndices("ix" & rsMetric.Fields(1)), attributeIndices("ix" & rsMetric.Fields(0))) _
            = rsMetric.Fields(2)
        rsMetric.MoveNext
    Wend
    rsMetric.Close
    
    cmdMetric.CommandText = _
        "select distinct lhs_value_id, lhs_users from " & _
        METRIC_TABLE & " where attribute_id=? and lhs_gender=?"
    cmdMetric.Parameters(0).Type = adInteger
    cmdMetric.Parameters(0) = rs.Fields(0)
    cmdMetric.Parameters(1).Type = adInteger
    cmdMetric.Parameters(1) = gender1 + GENDER_OFFSET
    Set rsMetric = cmdMetric.Execute
    While Not rsMetric.EOF
        If attributeIndices.Exists("ix" & rsMetric.Fields(0)) Then _
            anchorRange.Offset(ULhsI + UBound(attributeValues) + 1, attributeIndices("ix" & rsMetric.Fields(0))) _
            = rsMetric.Fields(1)
        rsMetric.MoveNext
    Wend
    rsMetric.Close
    
    cmdMetric.CommandText = _
        "select distinct rhs_value_id, rhs_users from " & _
        METRIC_TABLE & " where attribute_id=? and rhs_gender=?"
    cmdMetric.Parameters(0).Type = adInteger
    cmdMetric.Parameters(0) = rs.Fields(0)
    cmdMetric.Parameters(1).Type = adInteger
    cmdMetric.Parameters(1) = gender2 + GENDER_OFFSET
    Set rsMetric = cmdMetric.Execute
    While Not rsMetric.EOF
        If attributeIndices.Exists("ix" & rsMetric.Fields(0)) Then _
            anchorRange.Offset(attributeIndices("ix" & rsMetric.Fields(0)), Rhs + UBound(attributeValues) + 1) _
            = rsMetric.Fields(1)
        rsMetric.MoveNext
    Wend
    rsMetric.Close
    
    'Formatting
    Range(anchorRange.Cells(1), anchorRange.Offset(ULhsI + UBound(attributeValues) + 1, Rhs + UBound(attributeValues) + 1).Cells(1)).BorderAround , xlThick, xlColorIndexAutomatic
    Range(anchorRange.Offset(ULhsS + UBound(attributeValues) + 1).Cells(1), anchorRange.Offset(ULhsS + UBound(attributeValues) + 1, 1).Cells(1)).BorderAround , xlThick, xlColorIndexAutomatic
    Range(anchorRange.Offset(, URhsS + UBound(attributeValues) + 1).Cells(1), anchorRange.Offset(1, URhsS + UBound(attributeValues) + 1).Cells(1)).BorderAround , xlThick, xlColorIndexAutomatic
    Range(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET).Cells(1), anchorRange.Offset(NLhsIRhsS + UBound(attributeValues) + 1, 2 * UBound(attributeValues) + PAGE_SPLIT_OFFSET).Cells(1)).BorderAround , xlThick, xlColorIndexAutomatic
End Sub

' Writes sender interest results tables to the sheet
Private Sub senderInterestResults(anchorRange As Range)
    Dim resultsRange As Range
    
    Set resultsRange = _
        anchorRange.Offset(UBound(attributeValues) + SENDER_INTEREST_OFFSET)
    
    columnRange(resultsRange).FormulaArray = _
        divisionFormula(columnRange(anchorRange).Address(False, False), anchorRange.Offset(NLhsIRhsS + UBound(attributeValues) + 1, 1).Address(False, False))
    tableRange(resultsRange).FillRight
    outlineInterestTable resultsRange, _
        "Absolute Sender Interest I", _
        "n(LHSi,RHSj)/n(LHSi,RHSs)"
    
    outlineInterestTable _
        resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET), _
        "Absolute Sender Interest II", _
        "{n(LHSi,RHSj)/n(LHSs,RHSj)}"
    tableRange(resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = _
        divisionFormula(tableRange(anchorRange).Address(False, False), tableRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False))
    
    Set resultsRange = _
        resultsRange.Offset(UBound(attributeValues) + EXPECTED_SENDER_INTEREST_OFFSET)
    
    rowRange(resultsRange).FormulaArray = _
        divisionFormula(anchorRange.Offset(1, Rhs + UBound(attributeValues) + 1).Address(False, False), anchorRange.Offset(1, URhsS + UBound(attributeValues) + 1).Address(, False))
    tableRange(resultsRange).FillDown
    outlineInterestTable resultsRange, _
        "Expected Sender Interest I", _
        "u(RHSj)/u(RHSs)"
    
    columnRange(resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = _
        divisionFormula(anchorRange.Offset(ULhsI + UBound(attributeValues) + 1, 1).Address(False, False), anchorRange.Offset(ULhsS + UBound(attributeValues) + 1, 1).Address(False))
    tableRange(resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FillRight
    outlineInterestTable _
        resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET), _
        "Expected Sender Interest II", _
        "{u(LHSi)/u(LHSs)}"
    
    Set resultsRange = _
        resultsRange.Offset(UBound(attributeValues) + RELATIVE_SENDER_INTEREST_OFFSET)
    
    outlineInterestTable resultsRange, _
        "Relative Sender Interest", _
        "{n(LHSi,RHSj)/n(LHSi,RHSs)}/{u(RHSj)/u(RHSs)}"
    tableRange(resultsRange).FormulaArray = _
        divisionFormula(tableRange(anchorRange.Offset(UBound(attributeValues) + SENDER_INTEREST_OFFSET)).Address(False, False), tableRange(anchorRange.Offset(2 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET)).Address(False, False))
    
    outlineInterestTable _
        resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET), _
        "Comparative Sender Interest", _
        "{n(LHSi,RHSj)/n(LHSs,RHSj)}/{u(LHSi)/u(LHSs)}"
    tableRange(resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = _
        divisionFormula(tableRange(anchorRange.Offset(UBound(attributeValues) + SENDER_INTEREST_OFFSET, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False), tableRange(anchorRange.Offset(2 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False))
End Sub

' Writes receiver interest tables to the sheet
Private Sub receiverInterest()
    Dim anchorRange As Range
    Set anchorRange = _
        Range("A1").Offset(4 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + INFO_OFFSET + 2 * Y0)
    anchorRange = "Receiver Interest"
    anchorRange.Font.Bold = True
    Set anchorRange = anchorRange.Offset(Y0, X0)
    receiverInterestData anchorRange
    receiverInterestResults anchorRange
    addColumnRowInfo _
        anchorRange.Offset(7 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + RELATIVE_RECEIVER_INTEREST_OFFSET + INFO_OFFSET + 1)
End Sub

' Writes receiver interest data tables to the sheet
Private Sub receiverInterestData(anchorRange As Range)
    Dim rsMetric As ADODB.Recordset
    Dim cmdMetric As New ADODB.Command
    Dim dataRange As Range
    
    cmdMetric.ActiveConnection = cn
    
    anchorRange.Offset(-1, -2) = "Data (LHS2RHS)"
    anchorRange.Offset(-1) = "n(LHSi,RHSj)"
    outlineTable anchorRange
    
    cmdMetric.CommandText = _
        "select lhs_value_id, rhs_value_id, lhs2rhs from " _
        & METRIC_TABLE & _
        " where attribute_id=? and lhs_gender=? and rhs_gender=?"
    cmdMetric.Parameters(0).Type = adInteger
    cmdMetric.Parameters(0) = rs.Fields(0)
    cmdMetric.Parameters(1).Type = adInteger
    cmdMetric.Parameters(1) = gender1 + GENDER_OFFSET
    cmdMetric.Parameters(2).Type = adInteger
    cmdMetric.Parameters(2) = gender2 + GENDER_OFFSET
    Set rsMetric = cmdMetric.Execute
    While Not rsMetric.EOF
        If attributeIndices.Exists("ix" & rsMetric.Fields(0)) _
        And attributeIndices.Exists("ix" & rsMetric.Fields(1)) _
        Then _
            anchorRange.Offset(attributeIndices("ix" & rsMetric.Fields(1)), attributeIndices("ix" & rsMetric.Fields(0))) _
            = rsMetric.Fields(2)
        rsMetric.MoveNext
    Wend
    rsMetric.Close
    
    Set dataRange = _
        anchorRange.Offset(UBound(attributeValues) + RECEIVER_DATA_SPACE)
    
    dataRange.Offset(-1) = "n(LHSi,RHSj,+)"
    outlineTable dataRange
    
    cmdMetric.CommandText = _
        "select lhs_value_id, rhs_value_id, lhs2rhs_t from " _
        & METRIC_TABLE & _
        " where attribute_id=? and lhs_gender=? and rhs_gender=?"
    cmdMetric.Parameters(0).Type = adInteger
    cmdMetric.Parameters(0) = rs.Fields(0)
    cmdMetric.Parameters(1).Type = adInteger
    cmdMetric.Parameters(1) = gender1 + GENDER_OFFSET
    cmdMetric.Parameters(2).Type = adInteger
    cmdMetric.Parameters(2) = gender2 + GENDER_OFFSET
    Set rsMetric = cmdMetric.Execute
    While Not rsMetric.EOF
        If attributeIndices.Exists("ix" & rsMetric.Fields(0)) _
        And attributeIndices.Exists("ix" & rsMetric.Fields(1)) _
        Then _
            dataRange.Offset(attributeIndices("ix" & rsMetric.Fields(1)), attributeIndices("ix" & rsMetric.Fields(0))) _
            = rsMetric.Fields(2)
        rsMetric.MoveNext
    Wend
    rsMetric.Close
    
    Set dataRange = _
        dataRange.Offset(UBound(attributeValues) + RECEIVER_DATA_SPACE + 1)
    
    dataRange.Offset(-1) = "n(LHSi,RHSs)"
    columnRange(dataRange).FormulaArray = "=sum(" & _
        columnRange(anchorRange).Address(False, False) & ")"
    tableRange(dataRange).FillRight
    outlineTable dataRange
    
    dataRange.Offset(-1, UBound(attributeValues) + PAGE_SPLIT_OFFSET) _
        = "n(LHSs,RHSj)"
    rowRange(dataRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = "=sum(" & rowRange(anchorRange).Address(False, False) _
        & ")"
    tableRange(dataRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FillDown
    outlineTable _
        dataRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)
    
    Set dataRange = _
        dataRange.Offset(UBound(attributeValues) + RECEIVER_DATA_SPACE)
    dataRange.Offset(-1) = "n(LHSi,RHSs,+)"
    columnRange(dataRange).FormulaArray = "=sum(" & _
        columnRange(anchorRange.Offset(UBound(attributeValues) + RECEIVER_DATA_SPACE)).Address(False, False) _
        & ")"
    tableRange(dataRange).FillRight
    outlineTable dataRange
    
    dataRange.Offset(-1, UBound(attributeValues) + PAGE_SPLIT_OFFSET) = "n(LHSs,RHSj,+)"
    rowRange(dataRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = "=sum(" & _
        rowRange(anchorRange.Offset(UBound(attributeValues) + RECEIVER_DATA_SPACE)).Address(False, False) _
        & ")"
    tableRange(dataRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FillDown
    outlineTable _
        dataRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)
End Sub

' Writes receiver interest results tables to the sheet
Private Sub receiverInterestResults(anchorRange As Range)
    Dim resultsRange As Range
    
    Set resultsRange = _
        anchorRange.Offset(4 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + 1)
    
    outlineInterestTable resultsRange, _
        "Absolute Success Rate", _
        "n(LHSi,RHSj,+)/n(LHSi,RHSj)"
    tableRange(resultsRange).FormulaArray = _
        divisionFormula(tableRange(anchorRange.Offset(UBound(attributeValues) + RECEIVER_DATA_SPACE)).Address(False, False), tableRange(anchorRange).Address(False, False))
    
    Set resultsRange = _
        resultsRange.Offset(UBound(attributeValues) + EXPECTED_RECEIVER_INTEREST_OFFSET)
    
    outlineInterestTable resultsRange, _
        "Expected Success Rate", _
        "n(LHSi,RHSs,+)/n(LHSi,RHSs)"
    tableRange(resultsRange).FormulaArray = _
        divisionFormula(tableRange(anchorRange.Offset(3 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + 1)).Address(False, False), tableRange(anchorRange.Offset(2 * UBound(attributeValues) + 2 * RECEIVER_DATA_SPACE + 1)).Address(False, False))
    
    outlineInterestTable _
        resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET), _
        "Expected Success Rate", _
        "n(LHSs,RHSj,+)/n(LHSs,RHSj)"
    tableRange(resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = _
        divisionFormula(tableRange(anchorRange.Offset(3 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + 1, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False), tableRange(anchorRange.Offset(2 * UBound(attributeValues) + 2 * RECEIVER_DATA_SPACE + 1, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False))
    
    Set resultsRange = _
        resultsRange.Offset(UBound(attributeValues) + RELATIVE_RECEIVER_INTEREST_OFFSET)
    
    outlineInterestTable resultsRange, _
        "Relative Success Rate (=SRI)", _
        "{n(LHSi,RHSj,+)/n(LHSi,RHSj)}/{n(LHSi,RHSs,+)/n(LHSi,RHSs)}"
    tableRange(resultsRange).FormulaArray = _
        divisionFormula(tableRange(anchorRange.Offset(4 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + 1)).Address(False, False), tableRange(anchorRange.Offset(5 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + 1)).Address(False, False))
    
    outlineInterestTable _
        resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET), _
        "Relative Success Rate", _
        "{n(LHSi,RHSj,+)/n(LHSi,RHSj)}/{n(LHSs,RHSj,+)/n(LHSs,RHSj)}"
    tableRange(resultsRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = _
        divisionFormula(tableRange(anchorRange.Offset(4 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + 1)).Address(False, False), tableRange(anchorRange.Offset(5 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + 1, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False))
End Sub

' Writes combined results tables to the sheet
Private Sub combineResults()
    Dim anchorRange As Range, resultsRange As Range
    Set anchorRange = Range("A1").Offset(Y0, X0)
    Set resultsRange = _
        Range("A1").Offset(11 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + RELATIVE_RECEIVER_INTEREST_OFFSET + 2 * INFO_OFFSET + 4 * Y0 + 1)
    resultsRange = "Combine Measure"
    resultsRange.Font.Bold = True
    Set resultsRange = resultsRange.Offset(Y0, X0)
    fillResultsTable anchorRange, _
        resultsRange, _
        "Arithmetic Mean", _
        "average", _
        "or"
    Set resultsRange = _
        resultsRange.Offset(UBound(attributeValues) + HARMONIC_MEAN_OFFSET)
    fillResultsTable anchorRange, _
        resultsRange, _
        "Harmonic Mean", _
        "harmean"
    Set resultsRange = _
        resultsRange.Offset(UBound(attributeValues) + GEOMETRIC_MEAN_OFFSET)
    fillResultsTable anchorRange, _
        resultsRange, _
        "Geometric Mean", _
        "geomean"
End Sub

' Combine results helper function
Private Sub fillResultsTable(anchorRange As Range, r As Range, tableName As String, measure As String, Optional cond As String = "and")
    Dim rColumn As Range
    Dim s1 As String, s2 As String, r1 As String, r2 As String, xs As String
    Set rColumn = _
        anchorRange.Offset(3 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + 1, 1)
    s1 = rColumn.Address(False, False)
    s2 = _
        rColumn.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET).Address(False, False)
    Set rColumn = _
        rColumn.Offset(7 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + RELATIVE_RECEIVER_INTEREST_OFFSET + INFO_OFFSET + 2 * Y0 + 1)
    r1 = rColumn.Address(False, False)
    r2 = _
        rColumn.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET).Address(False, False)
    Set rColumn = columnRange(r)
    xs = "(" & s1 & "," & s2 & "," & r1 & "," & r2 & ")"
    r.Offset(1, 1).Formula = _
        "=if(" & cond & xs & "," & measure & xs & ")"
    rColumn.FillDown
    tableRange(r).FillRight
    outlineInterestTable r, tableName
End Sub

' Writes SiSri table to the sheet
Private Sub siSri()
    Dim siRange As Range, sriRange As Range, anchorRange As Range
    Set siRange = _
        Range("A1").Offset(3 * UBound(attributeValues) + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET)
    Set sriRange = _
        siRange.Offset(7 * UBound(attributeValues) + INFO_OFFSET + 2 * Y0 + 3 * RECEIVER_DATA_SPACE + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + RELATIVE_RECEIVER_INTEREST_OFFSET + 1)
    Set anchorRange = _
        sriRange.Offset(4 * UBound(attributeValues) + 4 * Y0 + INFO_OFFSET + HARMONIC_MEAN_OFFSET + GEOMETRIC_MEAN_OFFSET)
    anchorRange = "SI * SRI"
    anchorRange.Font.Bold = True
    Set siRange = siRange.Offset(Y0, X0)
    Set sriRange = sriRange.Offset(Y0, X0)
    Set anchorRange = anchorRange.Offset(Y0, X0)
    outlineInterestTable anchorRange
    tableRange(anchorRange).FormulaArray = "=" & _
        tableRange(siRange).Address(False, False) & "*" & _
        tableRange(sriRange).Address(False, False)
End Sub

' Writes homophily metric tables to the sheet
Private Sub homophilyMetric()
    Dim anchorRange As Range
    Set anchorRange = _
        Range("A1").Offset(15 * UBound(attributeValues) + 8 * Y0 + 3 * RECEIVER_DATA_SPACE + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + INFO_OFFSET + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + RELATIVE_RECEIVER_INTEREST_OFFSET + HARMONIC_MEAN_OFFSET + GEOMETRIC_MEAN_OFFSET + 1)
    anchorRange = "Homophily Metric"
    anchorRange.Font.Bold = True
    Set anchorRange = anchorRange.Offset(Y0, X0)
    homophilyMetricData anchorRange
    homophilyMetricResults anchorRange
End Sub

' Writes homophily metric data tables to the sheet
Private Sub homophilyMetricData(anchorRange As Range)
    Dim rsMetric As ADODB.Recordset
    Dim cmdMetric As New ADODB.Command
    Dim dataRange As Range, r1 As Range, r2 As Range
    
    cmdMetric.ActiveConnection = cn
    
    anchorRange.Offset(-1) = "RHSu(LHSi,RHSj)"
    anchorRange.Offset(UBound(attributeValues) + 1) _
        = "RHSu(LHSi,RHSs)"
    cmdMetric.CommandText = _
        "select lhs_value_id, rhs_value_id, rhs_users from " _
        & USERS_TABLE & _
        " where attribute_id=? and lhs_gender=? and rhs_gender=?"
    cmdMetric.Parameters(0).Type = adInteger
    cmdMetric.Parameters(0) = rs.Fields(0)
    cmdMetric.Parameters(1).Type = adInteger
    cmdMetric.Parameters(1) = gender1 + GENDER_OFFSET
    cmdMetric.Parameters(2).Type = adInteger
    cmdMetric.Parameters(2) = gender2 + GENDER_OFFSET
    Set rsMetric = cmdMetric.Execute
    While Not rsMetric.EOF
        If attributeIndices.Exists("ix" & rsMetric.Fields(0)) _
        And attributeIndices.Exists("ix" & rsMetric.Fields(1)) _
        Then _
            anchorRange.Offset(attributeIndices("ix" & rsMetric.Fields(1)), attributeIndices("ix" & rsMetric.Fields(0))) _
            = rsMetric.Fields(2)
        rsMetric.MoveNext
    Wend
    rsMetric.Close
    anchorRange.Offset(UBound(attributeValues) + 1, 1).Formula _
        = "=sum(" & columnRange(anchorRange).Address(False, False) _
        & ")"
    rowRange(anchorRange.Offset(UBound(attributeValues))).FillRight
    outlineTable anchorRange
    
    anchorRange.Offset(-1, UBound(attributeValues) + PAGE_SPLIT_OFFSET) _
        = "RHSu(LHSi,RHSj)/RHSu(LHSi,RHSs)"
    columnRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FormulaArray _
        = _
        divisionFormula(columnRange(anchorRange).Address(False, False), anchorRange.Offset(UBound(attributeValues) + 1, 1).Address(, False))
    tableRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).FillRight
    outlineTable anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)
    
    Set dataRange = _
        anchorRange.Offset(UBound(attributeValues) + HOMOPHILY_DATA_SPACE)
    dataRange.Offset(-1) = "n(LHSi,RHSj,+) expected"
    Set r1 = Range("A1").Offset(Y0, X0)
    Set r2 = _
        r1.Offset(9 * UBound(attributeValues) + 3 * RECEIVER_DATA_SPACE + 2 * Y0 + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + INFO_OFFSET + RECEIVER_INTEREST_OFFSET + EXPECTED_RECEIVER_INTEREST_OFFSET + 1, UBound(attributeValues) + PAGE_SPLIT_OFFSET)
    Set r1 = r1.Offset(NLhsIRhsS + UBound(attributeValues) + 1, 1)
    columnRange(dataRange).FormulaArray = "=" & _
        r1.Address(, False) & "*" & _
        columnRange(anchorRange.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET)).Address(False, False) _
        & "*" & columnRange(r2).Address(False, False)
    tableRange(dataRange).FillRight
    outlineTable dataRange
    
    ' Formatting
    Range(anchorRange.Offset(UBound(attributeValues) + 1).Cells(1), anchorRange.Offset(UBound(attributeValues) + 1, UBound(attributeValues)).Cells(1)).BorderAround , xlThick, xlColorIndexAutomatic
End Sub

' Writes homophily metric results tables to the sheet
Private Sub homophilyMetricResults(anchorRange As Range)
    Dim resultsRange As Range, actualRange As Range, expectedRange As Range
    Set actualRange = _
        Range("A1").Offset(5 * UBound(attributeValues) + 3 * Y0 + RECEIVER_DATA_SPACE + SENDER_INTEREST_OFFSET + EXPECTED_SENDER_INTEREST_OFFSET + RELATIVE_SENDER_INTEREST_OFFSET + INFO_OFFSET, X0)
    Set expectedRange = _
        anchorRange.Offset(UBound(attributeValues) + HOMOPHILY_DATA_SPACE)
    Set resultsRange = _
        expectedRange.Offset(UBound(attributeValues) + HOMOPHILY_RESULTS_OFFSET)
    tableRange(resultsRange).FormulaArray = _
        divisionFormula(tableRange(actualRange).Address(False, False), tableRange(expectedRange).Address(False, False))
    outlineInterestTable resultsRange, , "H(LHSi,RHSj)"
End Sub

' Draws up a results table, with heat map etc
Private Sub outlineInterestTable(r As Range, Optional tableName As String = "", Optional tableDescription As String = "")
    Dim i As Integer
    r.Offset(-2, -2) = tableName
    r.Offset(-2, -2).Font.Bold = True
    r.Offset(-2) = tableDescription
    r.Offset(-1, 1) = genders(gender1 + 1) & " Sender"
    r.Offset(1, -2) = genders(gender2 + 1) & " Receiver"
    For i = LBound(attributeValues) To UBound(attributeValues)
        r.Offset(i, -1) = attributeValues(i).name
    Next
    outlineTable r
    With tableRange(r)
        .NumberFormat = "0.00"
        .FormatConditions.AddColorScale 3
    End With
End Sub

' Draws up a table with attributes on each axis
Private Sub outlineTable(r As Range)
    Dim i As Integer
    For i = LBound(attributeValues) To UBound(attributeValues)
        r.Offset(i) = attributeValues(i).ix
        r.Offset(, i) = attributeValues(i).ix
    Next
    tableRange(r).BorderAround , xlThick, xlColorIndexAutomatic
    Range(r.Cells(1), r.Offset(UBound(attributeValues), UBound(attributeValues)).Cells(1)).BorderAround , xlThick, xlColorIndexAutomatic
End Sub

' Adds info on interpreting the data
Private Sub addColumnRowInfo(r As Range, Optional columnInfo As String = "", Optional rowInfo As String = "")
    r = "Interpreted by Column"
    r.Font.Bold = True
    r.Offset(1) = columnInfo
    r.Offset(1).Font.Bold = True
    r.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET) = "Interpreted by Row"
    r.Offset(, UBound(attributeValues) + PAGE_SPLIT_OFFSET).Font.Bold = True
    r.Offset(1, UBound(attributeValues) + PAGE_SPLIT_OFFSET) = rowInfo
    r.Offset(1, UBound(attributeValues) + PAGE_SPLIT_OFFSET).Font.Bold = True
End Sub

' Returns a square range of an area of values indexed by (attribute,attribute)
Private Function tableRange(r As Range) As Range
    Set tableRange = _
        Range(r.Offset(1, 1).Cells(1), r.Offset(UBound(attributeValues), UBound(attributeValues)).Cells(1))
End Function

' Returns the first row of an area of values indexed by (attribute,attribute)
Private Function rowRange(r As Range) As Range
    Set rowRange = _
        Range(r.Offset(1, 1).Cells(1), r.Offset(1, UBound(attributeValues)).Cells(1))
End Function

' Returns the first column of an area of values indexed by (attribute,attribute)
Private Function columnRange(r As Range) As Range
    Set columnRange = _
        Range(r.Offset(1, 1).Cells(1), r.Offset(UBound(attributeValues), 1).Cells(1))
End Function

' Wrapper around division formulae, to avoid errors such as division by zero
Private Function divisionFormula(numerator As String, denominator As String) As String
    divisionFormula = "=if(" & denominator & "," & numerator & "/" _
        & denominator & ")"
End Function
