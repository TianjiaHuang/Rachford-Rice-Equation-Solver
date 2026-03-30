Attribute VB_Name = "Module1"
Option Explicit

Public Sub RR_Huang_RunFromSheet()
    Dim z() As Double, k() As Double, beta0() As Double
    Dim beta() As Double
    Dim tol As Double, maxIter As Long
    Dim nc As Long, NPm1 As Long
    Dim betaStartRow As Long
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Calculator")

    nc = CLng(ws.Range("B4").Value)
    NPm1 = CLng(ws.Range("B5").Value)
    tol = CDbl(ws.Range("B6").Value)
    maxIter = CLng(ws.Range("B7").Value)

    betaStartRow = 12 + NPm1 + 2

    z = RangeToRowVector(ws.Range("B10").Resize(1, nc))
    k = RangeToMatrix(ws.Range("B12").Resize(NPm1, nc))
    beta0 = RangeToColumnVector(ws.Range("B" & betaStartRow).Resize(NPm1, 1))

    beta = RR_Huang_Core(z, k, beta0, tol, maxIter)

    ws.Range("C" & betaStartRow & ":C" & (betaStartRow + Application.Max(NPm1, 20))).ClearContents
    ws.Range("C" & betaStartRow - 1).Value = "beta"

    For i = 1 To UBound(beta)
        ws.Cells(betaStartRow + i - 1, 3).Value = beta(i)
    Next i
End Sub

Private Function RR_Huang_Core(z() As Double, k() As Double, beta0() As Double, tol As Double, maxIter As Long) As Double()
    Dim NPm1 As Long, nc As Long
    Dim beta() As Double, betaNew() As Double, betaTrial() As Double
    Dim a() As Double, theta() As Double, b() As Double
    Dim t() As Double, alpha() As Double, grad() As Double, Hess() As Double, d() As Double
    Dim tTrial() As Double, alphaTrial() As Double, gradTrial() As Double, HessTrial() As Double
    Dim denom() As Double, numer() As Double
    Dim gradNorm As Double, lambdaMax As Double, lambda As Double
    Dim dg As Double, ddg As Double, s As Double
    Dim maxLineSearchIter As Long, tolLineSearch As Double
    Dim i As Long, j As Long, iter As Long, n As Long
    Dim Kmax As Double, Kmin As Double, tempMin As Double

    NPm1 = UBound(k, 1)
    nc = UBound(k, 2)

    ReDim beta(1 To NPm1)
    ReDim betaNew(1 To NPm1)
    ReDim a(1 To NPm1, 1 To nc)
    ReDim theta(1 To nc)
    ReDim b(1 To nc)

    For j = 1 To NPm1
        beta(j) = beta0(j)
    Next j

    For j = 1 To NPm1
        For i = 1 To nc
            a(j, i) = 1# - k(j, i)
        Next i
    Next j

    For i = 1 To nc
        theta(i) = 1#
        For j = 1 To NPm1
            Kmax = RowMax(k, j, nc)
            Kmin = RowMin(k, j, nc)

            If k(j, i) > 1# Then
                theta(i) = Min2(theta(i), (1# - Kmin) / (k(j, i) - Kmin))
            Else
                theta(i) = Min2(theta(i), (Kmax - 1#) / (Kmax - k(j, i)))
            End If
        Next j
    Next i

    For i = 1 To nc
        tempMin = 1# - k(1, i) * z(i)
        For j = 2 To NPm1
            tempMin = Min2(tempMin, 1# - k(j, i) * z(i))
        Next j
        b(i) = Min2(1# - z(i) / theta(i), tempMin)
    Next i

    maxLineSearchIter = 10
    tolLineSearch = 0.001

    For iter = 1 To maxIter
        t = ComputeT(a, beta, NPm1, nc)
        alpha = ComputeAlpha(a, t, NPm1, nc)
        grad = ComputeGradient(alpha, z, NPm1, nc)

        gradNorm = VecInfNorm(grad)
        If gradNorm < tol Then
            RR_Huang_Core = beta
            Exit Function
        End If

        Hess = ComputeHessian(alpha, z, NPm1, nc)
        d = SolveLinearSystem(Hess, NegateVector(grad), NPm1)

        denom = TransposeMatVec(a, d, NPm1, nc)
        numer = SubtractVectors(b, TransposeMatVec(a, beta, NPm1, nc), nc)

        lambdaMax = 1#
        For i = 1 To nc
            If denom(i) > 0# Then
                lambda = numer(i) / denom(i)
                lambdaMax = Max2(0#, Min2(lambdaMax, lambda))
            End If
        Next i

        s = 1#
        For j = 1 To NPm1
            betaNew(j) = beta(j)
        Next j

        For n = 1 To maxLineSearchIter
            ReDim betaTrial(1 To NPm1)
            For i = 1 To NPm1
                betaTrial(i) = beta(i) + s * lambdaMax * d(i)
            Next i

            tTrial = ComputeT(a, betaTrial, NPm1, nc)
            alphaTrial = ComputeAlpha(a, tTrial, NPm1, nc)
            gradTrial = ComputeGradient(alphaTrial, z, NPm1, nc)

            dg = lambdaMax * DotProduct(gradTrial, d, NPm1)

            If dg < tolLineSearch Then
                For i = 1 To NPm1
                    betaNew(i) = betaTrial(i)
                Next i
                Exit For
            End If

            HessTrial = ComputeHessian(alphaTrial, z, NPm1, nc)
            ddg = lambdaMax * lambdaMax * QuadForm(d, HessTrial, NPm1)

            If Abs(ddg) < 0.00000000000001 Then Exit For

            s = s - dg / ddg
            If s < 0# Then s = 0#
            If s > 1# Then s = 1#

            For i = 1 To NPm1
                betaNew(i) = betaTrial(i)
            Next i
        Next n

        For j = 1 To NPm1
            beta(j) = betaNew(j)
        Next j
    Next iter

    RR_Huang_Core = beta
End Function

Private Function RangeToRowVector(rng As Range) As Double()
    Dim data As Variant, arr() As Double, i As Long, n As Long
    data = rng.Value2
    n = rng.Columns.Count
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CDbl(data(1, i))
    Next i
    RangeToRowVector = arr
End Function

Private Function RangeToColumnVector(rng As Range) As Double()
    Dim data As Variant, arr() As Double, i As Long, n As Long
    data = rng.Value2
    n = rng.Rows.Count
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CDbl(data(i, 1))
    Next i
    RangeToColumnVector = arr
End Function

Private Function RangeToMatrix(rng As Range) As Double()
    Dim data As Variant, arr() As Double, i As Long, j As Long, nr As Long, nc As Long
    data = rng.Value2
    nr = rng.Rows.Count
    nc = rng.Columns.Count
    ReDim arr(1 To nr, 1 To nc)
    For i = 1 To nr
        For j = 1 To nc
            arr(i, j) = CDbl(data(i, j))
        Next j
    Next i
    RangeToMatrix = arr
End Function

Private Function ComputeT(a() As Double, beta() As Double, NPm1 As Long, nc As Long) As Double()
    Dim t() As Double, i As Long, j As Long
    ReDim t(1 To nc)
    For i = 1 To nc
        t(i) = 1#
        For j = 1 To NPm1
            t(i) = t(i) - a(j, i) * beta(j)
        Next j
    Next i
    ComputeT = t
End Function

Private Function ComputeAlpha(a() As Double, t() As Double, NPm1 As Long, nc As Long) As Double()
    Dim alpha() As Double, i As Long, j As Long
    ReDim alpha(1 To NPm1, 1 To nc)
    For j = 1 To NPm1
        For i = 1 To nc
            alpha(j, i) = a(j, i) / t(i)
        Next i
    Next j
    ComputeAlpha = alpha
End Function

Private Function ComputeGradient(alpha() As Double, z() As Double, NPm1 As Long, nc As Long) As Double()
    Dim grad() As Double, i As Long, j As Long
    ReDim grad(1 To NPm1)
    For j = 1 To NPm1
        grad(j) = 0#
        For i = 1 To nc
            grad(j) = grad(j) + alpha(j, i) * z(i)
        Next i
    Next j
    ComputeGradient = grad
End Function

Private Function ComputeHessian(alpha() As Double, z() As Double, NPm1 As Long, nc As Long) As Double()
    Dim H() As Double, i As Long, j As Long, k As Long
    ReDim H(1 To NPm1, 1 To NPm1)
    For j = 1 To NPm1
        For k = 1 To NPm1
            H(j, k) = 0#
            For i = 1 To nc
                H(j, k) = H(j, k) + alpha(j, i) * alpha(k, i) * z(i)
            Next i
        Next k
    Next j
    ComputeHessian = H
End Function

Private Function SolveLinearSystem(a() As Double, b() As Double, n As Long) As Double()
    Dim m() As Double, x() As Double
    Dim i As Long, j As Long, k As Long, pivot As Long
    Dim factor As Double, temp As Double
    ReDim m(1 To n, 1 To n + 1)
    ReDim x(1 To n)
    For i = 1 To n
        For j = 1 To n
            m(i, j) = a(i, j)
        Next j
        m(i, n + 1) = b(i)
    Next i
    For k = 1 To n
        pivot = k
        For i = k + 1 To n
            If Abs(m(i, k)) > Abs(m(pivot, k)) Then pivot = i
        Next i
        If pivot <> k Then
            For j = k To n + 1
                temp = m(k, j)
                m(k, j) = m(pivot, j)
                m(pivot, j) = temp
            Next j
        End If
        If Abs(m(k, k)) < 0.00000000000001 Then Err.Raise vbObjectError + 513, , "Singular Hessian matrix."
        For i = k + 1 To n
            factor = m(i, k) / m(k, k)
            For j = k To n + 1
                m(i, j) = m(i, j) - factor * m(k, j)
            Next j
        Next i
    Next k
    For i = n To 1 Step -1
        x(i) = m(i, n + 1)
        For j = i + 1 To n
            x(i) = x(i) - m(i, j) * x(j)
        Next j
        x(i) = x(i) / m(i, i)
    Next i
    SolveLinearSystem = x
End Function

Private Function TransposeMatVec(a() As Double, x() As Double, nr As Long, nc As Long) As Double()
    Dim y() As Double, i As Long, j As Long
    ReDim y(1 To nc)
    For i = 1 To nc
        y(i) = 0#
        For j = 1 To nr
            y(i) = y(i) + a(j, i) * x(j)
        Next j
    Next i
    TransposeMatVec = y
End Function

Private Function NegateVector(x() As Double) As Double()
    Dim y() As Double, i As Long
    ReDim y(1 To UBound(x))
    For i = 1 To UBound(x)
        y(i) = -x(i)
    Next i
    NegateVector = y
End Function

Private Function SubtractVectors(x() As Double, y() As Double, n As Long) As Double()
    Dim z() As Double, i As Long
    ReDim z(1 To n)
    For i = 1 To n
        z(i) = x(i) - y(i)
    Next i
    SubtractVectors = z
End Function

Private Function DotProduct(x() As Double, y() As Double, n As Long) As Double
    Dim i As Long, s As Double
    s = 0#
    For i = 1 To n
        s = s + x(i) * y(i)
    Next i
    DotProduct = s
End Function

Private Function QuadForm(x() As Double, a() As Double, n As Long) As Double
    Dim i As Long, j As Long, s As Double
    s = 0#
    For i = 1 To n
        For j = 1 To n
            s = s + x(i) * a(i, j) * x(j)
        Next j
    Next i
    QuadForm = s
End Function

Private Function VecInfNorm(x() As Double) As Double
    Dim i As Long, m As Double
    m = Abs(x(1))
    For i = 2 To UBound(x)
        If Abs(x(i)) > m Then m = Abs(x(i))
    Next i
    VecInfNorm = m
End Function

Private Function RowMax(a() As Double, rowNum As Long, nc As Long) As Double
    Dim i As Long, m As Double
    m = a(rowNum, 1)
    For i = 2 To nc
        If a(rowNum, i) > m Then m = a(rowNum, i)
    Next i
    RowMax = m
End Function

Private Function RowMin(a() As Double, rowNum As Long, nc As Long) As Double
    Dim i As Long, m As Double
    m = a(rowNum, 1)
    For i = 2 To nc
        If a(rowNum, i) < m Then m = a(rowNum, i)
    Next i
    RowMin = m
End Function

Private Function Min2(a As Double, b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

Private Function Max2(a As Double, b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function
