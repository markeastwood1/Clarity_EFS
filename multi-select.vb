'
' Description here, Mark Eastwood
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'
Option Explicit
' this essentially makes string compare case-insensitive ... use binary for case-sensitive
Option Compare Text

Sub DropdownMultiSelect(Target As Range)
    On Error Resume Next

    Dim oldValue As String
    Dim newValue As String

    Application.EnableEvents = False

    With Target
        newValue = .Value2
        Application.Undo
        oldValue = .Value2
        .Value2 = newValue
    End With

    Application.EnableEvents = True

    'ThisWorkbook.UnProtect_This_Sheet
    Dim DelimiterType As String
    DelimiterType = ", "

    ' there are only certain cells that I want to allow multi-select...
    ' we've already filtered cells with validation (in general) and cells (more specifically) that have the "list validation" metadata
    ' but now we want to allow this only for specific cells... which is where the below is used
    Dim MULTI_SELECT_FRAUD_FIRST As Range
    Dim MULTI_SELECT_FRAUD_THIRD As Range
    Dim MULTI_SELECT_LIFE_ACCT As Range
    Dim MULTI_SELECT_LIFE_COLLECT As Range
    Dim MULTI_SELECT_LIFE_ORIG As Range
    Dim MULTI_SELECT_LIFE_APP_FRAUD As Range

    Dim MULTI_SELECT_CELLS As Range

    ' besides hoping to make maintenance simpler
    ' there are limits to the number of unions and the number of "line continuation"
    Set MULTI_SELECT_FRAUD_FIRST = Union(Range(FRAUD_FIRST_PORTFOLIOS), Range(FRAUD_FIRST_CARD_TYPES), Range(FRAUD_FIRST_DATA_SOURCES), Range(FRAUD_FIRST_TAGS), Range(RETAIL_FRAUD_PORTFOLIOS))
    Set MULTI_SELECT_FRAUD_THIRD = Union(Range(FRAUD_THIRD_PORTFOLIOS), Range(FRAUD_THIRD_CARD_TYPES), Range(FRAUD_THIRD_DATA_SOURCES), Range(FRAUD_THIRD_TAGS))
    Set MULTI_SELECT_LIFE_ACCT = Union(Range(LIFECYCLE_ACCT_PORTFOLIOS), Range(LIFECYCLE_ACCT_INDUSTRY_TYPE), Range(LIFECYCLE_ACCT_RISK_CATEGORY), Range(LIFECYCLE_ACCT_PRODUCT_TYPE), Range(LIFECYCLE_ACCT_PRODUCT_SUBTYPE), Range(LIFECYCLE_ACCT_COLLATERALIZE), Range(LIFECYCLE_ACCT_TYPES_FEATURE), Range(LIFECYCLE_ACCT_DATA_SRC_FEATURE))
    Set MULTI_SELECT_LIFE_COLLECT = Union(Range(LIFECYCLE_COLLECT_PORTFOLIOS), Range(LIFECYCLE_COLLECT_INDUSTRY_TYPE), Range(LIFECYCLE_COLLECT_RISK_CATEGORY), Range(LIFECYCLE_COLLECT_PRODUCT_TYPE), Range(LIFECYCLE_COLLECT_PRODUCT_SUBTYPE), Range(LIFECYCLE_COLLECT_COLLATERALIZE), Range(LIFECYCLE_COLLECT_TYPES_FEATURE), Range(LIFECYCLE_COLLECT_DATA_SRC_FEATURE))
    Set MULTI_SELECT_LIFE_ORIG = Union(Range(LIFECYCLE_ORIG_PORTFOLIOS), Range(LIFECYCLE_ORIG_INDUSTRY_TYPE), Range(LIFECYCLE_ORIG_RISK_CATEGORY), Range(LIFECYCLE_ORIG_PRODUCT_TYPE), Range(LIFECYCLE_ORIG_PRODUCT_SUBTYPE), Range(LIFECYCLE_ORIG_COLLATERALIZE), Range(LIFECYCLE_ORIG_PROD_STRUCT), Range(LIFECYCLE_ORIG_ACCT_TYPES_FEATURE), Range(LIFECYCLE_ORIG_DATA_SRC_FEATURE), Range(LIFECYCLE_ORIG_TAGS))
    Set MULTI_SELECT_LIFE_APP_FRAUD = Union(Range(LIFECYCLE_APP_FRAUD_PORTFOLIOS), Range(LIFECYCLE_APP_FRAUD_INDUSTRY_TYPE), Range(LIFECYCLE_APP_FRAUD_RISK_CATEGORY), Range(LIFECYCLE_APP_FRAUD_PRODUCT_TYPE), Range(LIFECYCLE_APP_FRAUD_PRODUCT_SUBTYPE), Range(LIFECYCLE_APP_FRAUD_COLLATERALIZE), Range(LIFECYCLE_APP_FRAUD_DATA_SRC_FEATURE), Range(LIFECYCLE_APP_FRAUD_TAGS))

    Set MULTI_SELECT_CELLS = Union(MULTI_SELECT_FRAUD_FIRST, MULTI_SELECT_FRAUD_THIRD, MULTI_SELECT_LIFE_ACCT, MULTI_SELECT_LIFE_COLLECT, MULTI_SELECT_LIFE_ORIG, MULTI_SELECT_LIFE_APP_FRAUD)

    On Error Resume Next

   'did the change happen where we are interested?
    Dim inMultiSelectCell As Range
    Set inMultiSelectCell = Intersect(Target, MULTI_SELECT_CELLS)

    If inMultiSelectCell Is Nothing Then
        Exit Sub
    End If

    If oldValue = "" Then
         Exit Sub
    Else
        If newValue = "" Then
            Exit Sub
        Else
            Application.EnableEvents = False
            Target.Value2 = oldValue & DelimiterType & newValue
            Application.EnableEvents = True
         End If
    End If

    'ThisWorkbook.Protect_This_Sheet
End Sub
