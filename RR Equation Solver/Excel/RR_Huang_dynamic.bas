Attribute VB_Name = "RR_Huang_Excel"
Option Explicit

Private Const INPUT_SHEET As String = "Calculator"
Private Const RESULTS_SHEET As String = "Results"
Private Const EPS_PIVOT As Double = 0.00000000000001#

Public Sub RR_Huang_RunFromSheet()
    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    Dim z() As Double, k() As Double, beta0() As Double
    Dim beta() As Double, residual() As Double, point() As Double
    Dim tol As Double, maxIter As Long
    Dim nc As Long, NPm1 As Long
    Dim betaStartRow As Long
    Dim iterCount As Long, ierr As Long
    Dim messageText As String

    Set wsIn = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsOut = GetOrCreateResultsSheet()

    If Not ReadAndValidateInputs(wsIn, z, k, beta0, tol, maxIter, nc, NPm1, betaStartRow, messageText) Then
        WriteRunStatus wsOut, "VBA", 0, -1, messageText
        MsgBox messageText, vbExclamation, "RR Huang"
        Exit Sub
    End If

    beta = RR_Huang_Core(z, k, beta0, tol, maxIter, residual, point, iterCount, ierr)

    WriteBetaToCalculator wsIn, beta, betaStartRow, "beta"
    WriteRunStatus wsOut, "VBA", iterCount, ierr, SolverNote(ierr, iterCount, maxIter)
    WriteResidualHistory wsOut, residual, iterCount
    WritePointHistory wsOut, point, iterCount, NPm1
End Sub

Public Sub RR_Huang_RunFromSheet_Auto()
    RR_Huang_RunFromSheet
End Sub

Public Sub RR_Huang_ResetResults()
    Dim wsOut As Worksheet

    Set wsOut = GetOrCreateResultsSheet()
    wsOut.Cells.ClearContents
End Sub

Private Function ReadAndValidateInputs(ws As Worksheet, z() As Double, k() As Double, beta0() As Double, tol As Double, _
    maxIter As Long, nc As Long, NPm1 As Long, betaStartRow As Long, messageText As String) As Boolean
    Dim i As Long
    Dim zSum As Double

    On Error GoTo InvalidInput

    nc = CLng(ws.Range("B4").Value)
    NPm1 = CLng(ws.Range("B5").Value)
    tol = CDbl(ws.Range("B6").Value)
    maxIter = CLng(ws.Range("B7").Value)

    If nc <= 0 Then
        messageText = "NC in cell B4 must be a positive integer."
        Exit Function
    End If

    If NPm1 <= 0 Then
        messageText = "NP-1 in cell B5 must be a positive integer."
        Exit Function
    End If

    If tol <= 0# Then
        messageText = "Tolerance in cell B6 must be positive."
        Exit Function
    End If

    If maxIter <= 0 Then
        messageText = "Maximum iterations in cell B7 must be a positive integer."
        Exit Function
    End If

    betaStartRow = 12 + NPm1 + 2

    z = RangeToRowVector(ws.Range("B10").Resize(1, nc))
    k = RangeToMatrix(ws.Range("B12").Resize(NPm1, nc))
    beta0 = RangeToColumnVector(ws.Range("B" & betaStartRow).Resize(NPm1, 1))

    zSum = 0#
    For i = 1 To nc
        If z(i) < 0# Then
            messageText = "All z values must be nonnegative."
            Exit Function
        End If
        zSum = zSum + z(i)
    Next i

    If Abs(zSum - 1#) > 0.000001 Then
        messageText = "The z values must sum to 1 within 1E-6."
        Exit Function
    End If

    ReadAndValidateInputs = True
    Exit Function

InvalidInput:
    messageText = "Unable to read inputs from the Calculator sheet. Check that all required cells contain numeric values."
End Function

Private Function ReadAndValidateEstimatorInputs(ws As Worksheet, z() As Double, k() As Double, nc As Long, NPm1 As Long, _
    betaStartRow As Long, messageText As String) As Boolean
    Dim i As Long
    Dim zSum As Double

    On Error GoTo InvalidInput

    nc = CLng(ws.Range("B4").Value)
    NPm1 = CLng(ws.Range("B5").Value)

    If nc <= 0 Then
        messageText = "NC in cell B4 must be a positive integer."
        Exit Function
    End If

    If NPm1 <= 0 Then
        messageText = "NP-1 in cell B5 must be a positive integer."
        Exit Function
    End If

    betaStartRow = 12 + NPm1 + 2
    z = RangeToRowVector(ws.Range("B10").Resize(1, nc))
    k = RangeToMatrix(ws.Range("B12").Resize(NPm1, nc))

    zSum = 0#
    For i = 1 To nc
        If z(i) < 0# Then
            messageText = "All z values must be nonnegative."
            Exit Function
        End If
        zSum = zSum + z(i)
    Next i

    If Abs(zSum - 1#) > 0.000001 Then
        messageText = "The z values must sum to 1 within 1E-6."
        Exit Function
    End If

    ReadAndValidateEstimatorInputs = True
    Exit Function

InvalidInput:
    messageText = "Unable to read estimator inputs from the Calculator sheet."
End Function

Private Function RR_Huang_Core(z() As Double, k() As Double, beta0() As Double, tol As Double, maxIter As Long, _
    residual() As Double, point() As Double, iterCount As Long, ierr As Long) As Double()
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
    ReDim residual(1 To maxIter)
    ReDim point(1 To maxIter, 1 To NPm1)
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
    ierr = 1

    For iter = 1 To maxIter
        For j = 1 To NPm1
            point(iter, j) = beta(j)
        Next j

        t = ComputeT(a, beta, NPm1, nc)
        alpha = ComputeAlpha(a, t, NPm1, nc)
        grad = ComputeGradient(alpha, z, NPm1, nc)

        gradNorm = VecInfNorm(grad)
        residual(iter) = gradNorm

        If gradNorm < tol Then
            ierr = 0
            iterCount = iter
            RR_Huang_Core = beta
            Exit Function
        End If

        Hess = ComputeHessian(alpha, z, NPm1, nc)
        d = SolveLinearSystem(Hess, NegateVector(grad), NPm1, ierr)
        If ierr <> 0 Then
            iterCount = iter
            RR_Huang_Core = beta
            Exit Function
        End If

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
            If Abs(ddg) < EPS_PIVOT Then Exit For

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

        iterCount = iter
    Next iter

    RR_Huang_Core = beta
End Function

Private Function Initial_Estimation_Gradient_Huang_Core(z() As Double, k() As Double, vertices() As Double, _
    gradVals() As Double, weights() As Double, vertexCount As Long, noteText As String) As Double()
    Dim NPm1 As Long, nc As Long, numConstraints As Long
    Dim A() As Double, b() As Double, a_i() As Double
    Dim centroid() As Double
    Dim combo() As Long, idx() As Long
    Dim M() As Double, rhs() As Double, point() As Double
    Dim i As Long, j As Long, r As Long
    Dim theta As Double, K1 As Double, Kn As Double, constraintBound As Double
    Dim gradNorm As Double, invGradSum As Double

    NPm1 = UBound(k, 1)
    nc = UBound(k, 2)
    numConstraints = nc + NPm1 + 1

    ReDim centroid(1 To NPm1)
    ReDim A(1 To numConstraints, 1 To NPm1)
    ReDim b(1 To numConstraints)
    ReDim a_i(1 To NPm1, 1 To nc)

    For j = 1 To NPm1
        For i = 1 To nc
            a_i(j, i) = 1# - k(j, i)
        Next i
    Next j

    For i = 1 To nc
        theta = 1#
        For j = 1 To NPm1
            K1 = RowMax(k, j, nc)
            Kn = RowMin(k, j, nc)
            If k(j, i) > 1# Then
                theta = Min2(theta, (1# - Kn) / (k(j, i) - Kn))
            Else
                theta = Min2(theta, (K1 - 1#) / (K1 - k(j, i)))
            End If
            A(i, j) = 1# - k(j, i)
        Next j

        constraintBound = 1# - z(i) / theta
        For j = 1 To NPm1
            constraintBound = Min2(constraintBound, 1# - k(j, i) * z(i))
        Next j
        b(i) = constraintBound
    Next i

    For j = 1 To NPm1
        A(nc + j, j) = -1#
        b(nc + j) = 0#
    Next j

    For j = 1 To NPm1
        A(numConstraints, j) = 1#
    Next j
    b(numConstraints) = 1#

    ReDim idx(1 To NPm1)
    For i = 1 To NPm1
        idx(i) = i
    Next i

    Do
        ReDim M(1 To NPm1, 1 To NPm1)
        ReDim rhs(1 To NPm1)

        For r = 1 To NPm1
            For j = 1 To NPm1
                M(r, j) = A(idx(r), j)
            Next j
            rhs(r) = b(idx(r))
        Next r

        If MatrixRank(M, NPm1) = NPm1 Then
            point = SolveLinearSystem(M, rhs, NPm1, i)
            If i = 0 Then
                If PointSatisfiesConstraints(A, b, point, numConstraints, NPm1) Then
                    AppendUniqueVertex vertices, vertexCount, point, NPm1
                End If
            End If
        End If
    Loop While NextCombination(idx, numConstraints, NPm1)

    If vertexCount = 0 Then
        noteText = "No feasible region found."
        Initial_Estimation_Gradient_Huang_Core = centroid
        Exit Function
    End If

    ReDim gradVals(1 To vertexCount)
    ReDim weights(1 To vertexCount)

    invGradSum = 0#
    For i = 1 To vertexCount
        point = GetVertex(vertices, i, NPm1)
        gradNorm = GradientNormAtPoint(a_i, z, point, NPm1, nc)
        gradVals(i) = gradNorm
        If gradNorm > 0# Then
            weights(i) = 1# / gradNorm
            invGradSum = invGradSum + weights(i)
        End If
    Next i

    If invGradSum <= 0# Then
        noteText = "Gradient weights could not be computed."
        Initial_Estimation_Gradient_Huang_Core = centroid
        Exit Function
    End If

    For i = 1 To vertexCount
        weights(i) = weights(i) / invGradSum
        For j = 1 To NPm1
            centroid(j) = centroid(j) + weights(i) * vertices(i, j)
        Next j
    Next i

    noteText = "Centroid computed successfully."
    Initial_Estimation_Gradient_Huang_Core = centroid
End Function

Private Sub WriteBetaToCalculator(ws As Worksheet, beta() As Double, betaStartRow As Long, headerText As String)
    Dim i As Long

    ws.Range("C" & betaStartRow & ":C" & (betaStartRow + Application.Max(UBound(beta), 20))).ClearContents
    ws.Range("C" & betaStartRow - 1).Value = headerText

    For i = 1 To UBound(beta)
        ws.Cells(betaStartRow + i - 1, 3).Value = beta(i)
    Next i
End Sub

Private Sub WriteRunStatus(ws As Worksheet, solverName As String, iterCount As Long, ierr As Long, noteText As String)
    ws.Range("A1:B6").ClearContents
    ws.Range("A1").Value = "Solver"
    ws.Range("B1").Value = solverName
    ws.Range("A2").Value = "iterCount"
    ws.Range("B2").Value = iterCount
    ws.Range("A3").Value = "ierr"
    ws.Range("B3").Value = ierr
    ws.Range("A4").Value = "Note"
    ws.Range("B4").Value = noteText
End Sub

Private Sub WriteResidualHistory(ws As Worksheet, residual() As Double, iterCount As Long)
    Dim i As Long

    ws.Range("A11:B100000").ClearContents
    ws.Range("A11").Value = "Iteration"
    ws.Range("B11").Value = "Residual"

    For i = 1 To iterCount
        ws.Cells(11 + i, 1).Value = i
        ws.Cells(11 + i, 2).Value = residual(i)
    Next i
End Sub

Private Sub WritePointHistory(ws As Worksheet, point() As Double, iterCount As Long, NPm1 As Long)
    Dim i As Long, j As Long

    ws.Range("D11:ZZ100000").ClearContents
    ws.Cells(11, 4).Value = "Iteration"
    For j = 1 To NPm1
        ws.Cells(11, 4 + j).Value = "beta_" & j
    Next j

    For i = 1 To iterCount
        ws.Cells(11 + i, 4).Value = i
        For j = 1 To NPm1
            ws.Cells(11 + i, 4 + j).Value = point(i, j)
        Next j
    Next i
End Sub

Private Function SolverNote(ierr As Long, iterCount As Long, maxIter As Long) As String
    Select Case ierr
        Case 0
            SolverNote = "Converged."
        Case 2
            SolverNote = "Singular Hessian matrix."
        Case Else
            If iterCount >= maxIter Then
                SolverNote = "Maximum iterations reached without convergence."
            Else
                SolverNote = "Solver stopped before convergence."
            End If
    End Select
End Function

Private Function GetOrCreateResultsSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateResultsSheet = ThisWorkbook.Worksheets(RESULTS_SHEET)
    On Error GoTo 0

    If GetOrCreateResultsSheet Is Nothing Then
        Set GetOrCreateResultsSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateResultsSheet.Name = RESULTS_SHEET
    End If
End Function

Private Function GetOrCreateGradientSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateGradientSheet = ThisWorkbook.Worksheets("InitialGradient")
    On Error GoTo 0

    If GetOrCreateGradientSheet Is Nothing Then
        Set GetOrCreateGradientSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateGradientSheet.Name = "InitialGradient"
    End If
End Function

Private Sub WriteInitialGradientSheet(ws As Worksheet, centroid() As Double, vertices() As Double, gradVals() As Double, _
    weights() As Double, vertexCount As Long, NPm1 As Long, noteText As String)
    Dim i As Long, j As Long

    ws.Cells.ClearContents
    ws.Range("A1").Value = "Status"
    ws.Range("B1").Value = noteText
    ws.Range("A3").Value = "Centroid"
    For j = 1 To NPm1
        ws.Cells(3, 1 + j).Value = "beta_" & j
        ws.Cells(4, 1 + j).Value = centroid(j)
    Next j

    ws.Range("A6").Value = "Vertices"
    ws.Cells(7, 1).Value = "Vertex"
    For j = 1 To NPm1
        ws.Cells(7, 1 + j).Value = "beta_" & j
    Next j
    ws.Cells(7, NPm1 + 2).Value = "grad_norm"
    ws.Cells(7, NPm1 + 3).Value = "weight"

    For i = 1 To vertexCount
        ws.Cells(7 + i, 1).Value = i
        For j = 1 To NPm1
            ws.Cells(7 + i, 1 + j).Value = vertices(i, j)
        Next j
        ws.Cells(7 + i, NPm1 + 2).Value = gradVals(i)
        ws.Cells(7 + i, NPm1 + 3).Value = weights(i)
    Next i
End Sub

Private Function NextCombination(idx() As Long, n As Long, k As Long) As Boolean
    Dim i As Long, j As Long

    For i = k To 1 Step -1
        If idx(i) < n - k + i Then
            idx(i) = idx(i) + 1
            For j = i + 1 To k
                idx(j) = idx(j - 1) + 1
            Next j
            NextCombination = True
            Exit Function
        End If
    Next i
End Function

Private Function MatrixRank(a() As Double, n As Long) As Long
    Dim m() As Double
    Dim i As Long, j As Long, k As Long, pivot As Long
    Dim factor As Double, temp As Double
    Dim rankVal As Long

    ReDim m(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            m(i, j) = a(i, j)
        Next j
    Next i

    rankVal = 0
    For k = 1 To n
        pivot = 0
        For i = k To n
            If Abs(m(i, k)) > EPS_PIVOT Then
                pivot = i
                Exit For
            End If
        Next i
        If pivot = 0 Then Exit For

        If pivot <> k Then
            For j = k To n
                temp = m(k, j)
                m(k, j) = m(pivot, j)
                m(pivot, j) = temp
            Next j
        End If

        rankVal = rankVal + 1
        For i = k + 1 To n
            factor = m(i, k) / m(k, k)
            For j = k To n
                m(i, j) = m(i, j) - factor * m(k, j)
            Next j
        Next i
    Next k

    MatrixRank = rankVal
End Function

Private Function PointSatisfiesConstraints(A() As Double, b() As Double, point() As Double, numConstraints As Long, NPm1 As Long) As Boolean
    Dim i As Long, j As Long
    Dim lhs As Double

    PointSatisfiesConstraints = True
    For i = 1 To numConstraints
        lhs = 0#
        For j = 1 To NPm1
            lhs = lhs + A(i, j) * point(j)
        Next j
        If lhs > b(i) + 0.00000001 Then
            PointSatisfiesConstraints = False
            Exit Function
        End If
    Next i
End Function

Private Sub AppendUniqueVertex(vertices() As Double, vertexCount As Long, point() As Double, NPm1 As Long)
    Dim i As Long, j As Long
    Dim isDuplicate As Boolean
    Dim resized() As Double

    For i = 1 To vertexCount
        isDuplicate = True
        For j = 1 To NPm1
            If Abs(vertices(i, j) - point(j)) > 0.00000001 Then
                isDuplicate = False
                Exit For
            End If
        Next j
        If isDuplicate Then Exit Sub
    Next i

    vertexCount = vertexCount + 1
    If vertexCount = 1 Then
        ReDim vertices(1 To 1, 1 To NPm1)
    Else
        ReDim resized(1 To vertexCount, 1 To NPm1)
        For i = 1 To vertexCount - 1
            For j = 1 To NPm1
                resized(i, j) = vertices(i, j)
            Next j
        Next i
        vertices = resized
    End If

    For j = 1 To NPm1
        vertices(vertexCount, j) = Round(point(j), 8)
    Next j
End Sub

Private Function GetVertex(vertices() As Double, vertexIndex As Long, NPm1 As Long) As Double()
    Dim point() As Double
    Dim j As Long

    ReDim point(1 To NPm1)
    For j = 1 To NPm1
        point(j) = vertices(vertexIndex, j)
    Next j
    GetVertex = point
End Function

Private Function GradientNormAtPoint(a_i() As Double, z() As Double, point() As Double, NPm1 As Long, nc As Long) As Double
    Dim t() As Double, alpha() As Double, grad() As Double

    t = ComputeT(a_i, point, NPm1, nc)
    alpha = ComputeAlpha(a_i, t, NPm1, nc)
    grad = ComputeGradient(alpha, z, NPm1, nc)
    GradientNormAtPoint = EuclideanNorm(grad)
End Function

Private Function EuclideanNorm(x() As Double) As Double
    Dim i As Long
    Dim s As Double

    s = 0#
    For i = 1 To UBound(x)
        s = s + x(i) * x(i)
    Next i
    EuclideanNorm = Sqr(s)
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

Private Function SolveLinearSystem(a() As Double, b() As Double, n As Long, ierr As Long) As Double()
    Dim m() As Double, x() As Double
    Dim i As Long, j As Long, k As Long, pivot As Long
    Dim factor As Double, temp As Double
    ReDim m(1 To n, 1 To n + 1)
    ReDim x(1 To n)
    ierr = 0

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

        If Abs(m(k, k)) < EPS_PIVOT Then
            ierr = 2
            SolveLinearSystem = x
            Exit Function
        End If

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
