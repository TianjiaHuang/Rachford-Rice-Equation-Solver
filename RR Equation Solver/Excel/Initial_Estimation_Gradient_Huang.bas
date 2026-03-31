Attribute VB_Name = "Initial_Estimation_Gradient_Huang"
Option Explicit

Private Const INPUT_SHEET As String = "Calculator"
Private Const OUTPUT_SHEET As String = "InitialGradient"
Private Const EPS_PIVOT As Double = 0.00000000000001#

Public Sub InitialGradient_CreateSheet()
    Dim wsIn As Worksheet
    Dim wsOut As Worksheet
    Dim z() As Double, k() As Double
    Dim nc As Long, NPm1 As Long, betaStartRow As Long
    Dim centroid() As Double, vertices() As Double, gradVals() As Double, weights() As Double
    Dim vertexCount As Long
    Dim messageText As String

    Set wsIn = ThisWorkbook.Worksheets(INPUT_SHEET)
    Set wsOut = GetOrCreateGradientSheet()

    If Not ReadAndValidateEstimatorInputs(wsIn, z, k, nc, NPm1, betaStartRow, messageText) Then
        wsOut.Cells.ClearContents
        wsOut.Range("A1").Value = "Status"
        wsOut.Range("B1").Value = messageText
        MsgBox messageText, vbExclamation, "Initial Gradient Huang"
        Exit Sub
    End If

    centroid = Initial_Estimation_Gradient_Huang_Core(z, k, vertices, gradVals, weights, vertexCount, messageText)
    WriteInitialGuessToCalculator wsIn, centroid, betaStartRow
    WriteInitialGradientSheet wsOut, centroid, vertices, gradVals, weights, vertexCount, NPm1, messageText
End Sub

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

Private Sub WriteInitialGuessToCalculator(ws As Worksheet, centroid() As Double, betaStartRow As Long)
    Dim i As Long

    ws.Range("B" & betaStartRow & ":B" & (betaStartRow + Application.Max(UBound(centroid), 20))).ClearContents
    For i = 1 To UBound(centroid)
        ws.Cells(betaStartRow + i - 1, 2).Value = centroid(i)
    Next i
End Sub

Private Function Initial_Estimation_Gradient_Huang_Core(z() As Double, k() As Double, vertices() As Double, _
    gradVals() As Double, weights() As Double, vertexCount As Long, noteText As String) As Double()
    Dim NPm1 As Long, nc As Long, numConstraints As Long
    Dim A() As Double, b() As Double, a_i() As Double
    Dim centroid() As Double
    Dim idx() As Long
    Dim M() As Double, rhs() As Double, point() As Double
    Dim i As Long, j As Long, r As Long, ierr As Long
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
            point = SolveLinearSystem(M, rhs, NPm1, ierr)
            If ierr = 0 Then
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

Private Function GetOrCreateGradientSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateGradientSheet = ThisWorkbook.Worksheets(OUTPUT_SHEET)
    On Error GoTo 0

    If GetOrCreateGradientSheet Is Nothing Then
        Set GetOrCreateGradientSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateGradientSheet.Name = OUTPUT_SHEET
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

Private Function EuclideanNorm(x() As Double) As Double
    Dim i As Long
    Dim s As Double

    s = 0#
    For i = 1 To UBound(x)
        s = s + x(i) * x(i)
    Next i
    EuclideanNorm = Sqr(s)
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
