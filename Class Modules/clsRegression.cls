VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsRegression Class Module

' Properties to store data and model information
Private pDataX As Range          ' Range for independent variable data (x-values)
Private pDataY As Range          ' Range for dependent variable data (y-values)
Private pCoefficients As Variant ' Array to store coefficients of the fitted curve
Private pDegree As Integer       ' Degree of the polynomial for polynomial regression
Private pModelType As String     ' String to store the type of regression model used ("Linear", "Polynomial", etc.)

' Property to Set the x-values (independent variable)
Public Property Let DataX(Value As Range)
    Set pDataX = Value
End Property

' Property to Get the x-values (independent variable)
Public Property Get DataX() As Range
    Set DataX = pDataX
End Property

' Property to Set the y-values (dependent variable)
Public Property Let DataY(Value As Range)
    Set pDataY = Value
End Property

' Property to Get the y-values (dependent variable)
Public Property Get DataY() As Range
    Set DataY = pDataY
End Property

' Method to perform Linear Regression
' This method calculates the slope (m) and intercept (b) of the best-fit line
' The coefficients are stored in the pCoefficients array
Public Sub FitLinear()
    Dim n As Long
    Dim sumX As Double, sumY As Double
    Dim sumXY As Double, sumX2 As Double
    Dim m As Double, B As Double
    Dim i As Long

    n = pDataX.Count
    For i = 1 To n
        sumX = sumX + pDataX.Cells(i, 1).Value
        sumY = sumY + pDataY.Cells(i, 1).Value
        sumXY = sumXY + pDataX.Cells(i, 1).Value * pDataY.Cells(i, 1).Value
        sumX2 = sumX2 + pDataX.Cells(i, 1).Value ^ 2
    Next i

    ' Calculate slope (m)
    m = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX ^ 2)

    ' Calculate y-intercept (b)
    B = (sumY - m * sumX) / n

    ' Store coefficients in an array
    ReDim pCoefficients(1 To 2)
    pCoefficients(1) = m
    pCoefficients(2) = B

    ' Set the model type to "Linear"
    pModelType = "Linear"
End Sub

' Method to perform Polynomial Regression
' This method calculates the coefficients of a polynomial fit of the specified degree
' The coefficients are stored in the pCoefficients array
Public Sub FitPolynomial(degree As Integer)
    Dim n As Long
    Dim i As Long, j As Long, k As Long
    Dim X() As Double, Y() As Double, A() As Double, B() As Double
    Dim coefficients() As Double

    pDegree = degree
    n = pDataX.Count

    ' Resize arrays for input data (X) and dependent data (Y)
    ReDim X(n, degree) As Double
    ReDim Y(n) As Double
    ReDim A(degree, degree) As Double
    ReDim B(degree) As Double

    ' Fill X and Y arrays with data
    For i = 1 To n
        Y(i) = pDataY.Cells(i, 1).Value
        For j = 0 To degree
            X(i, j) = pDataX.Cells(i, 1).Value ^ j
        Next j
    Next i

    ' Calculate A matrix (X^T * X)
    For i = 0 To degree
        For j = 0 To degree
            A(i, j) = 0
            For k = 1 To n
                A(i, j) = A(i, j) + X(k, i) * X(k, j)
            Next k
        Next j
    Next i

    ' Calculate B vector (X^T * Y)
    For i = 0 To degree
        B(i) = 0
        For k = 1 To n
            B(i) = B(i) + X(k, i) * Y(k)
        Next k
    Next i

    ' Solve the system of linear equations A * coefficients = B
    coefficients = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(A), B)

    ' Store the coefficients in the pCoefficients array
    ReDim pCoefficients(0 To degree) As Double
    For i = 0 To degree
        pCoefficients(i) = coefficients(i)
    Next i

    ' Set the model type to "Polynomial"
    pModelType = "Polynomial"
End Sub

' Method to predict the y-value for a given x-value based on the fitted model
' This method returns the predicted y-value using the stored coefficients
Public Function Predict(xValue As Double) As Double
    Dim yValue As Double
    Dim i As Integer

    Select Case pModelType
        Case "Linear"
            ' For linear regression: y = mx + b
            yValue = pCoefficients(1) * xValue + pCoefficients(2)
        Case "Polynomial"
            ' For polynomial regression: y = a0 + a1*x + a2*x^2 + ... + an*x^n
            yValue = 0
            For i = 0 To pDegree
                yValue = yValue + pCoefficients(i) * xValue ^ i
            Next i
        ' Additional models can be added here
    End Select

    Predict = yValue
End Function

' Method to get a summary of the fitted model
' This method returns a string describing the fitted model and its coefficients
Public Function GetSummary() As String
    Dim summary As String
    Dim i As Integer

    Select Case pModelType
        Case "Linear"
            summary = "Linear Regression: y = " & pCoefficients(1) & " * x + " & pCoefficients(2)
        Case "Polynomial"
            summary = "Polynomial Regression (Degree " & pDegree & "): y = "
            For i = 0 To pDegree
                If i > 0 Then
                    summary = summary & " + "
                End If
                summary = summary & pCoefficients(i) & " * x^" & i
            Next i
        ' Additional models can be summarized here
    End Select

    GetSummary = summary
End Function

