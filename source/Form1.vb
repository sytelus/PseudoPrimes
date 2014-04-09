Public Class Form1

    Dim knownPrimes As New Collections.Generic.List(Of Long)
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'WriteNumbersToFile(GeneratePrimeList2(100000), "c:\temp\primeList.txt")
        knownPrimes.AddRange(ReadNumbersFromFile("c:\temp\primeList.txt"))
    End Sub

    Private Sub WriteNumbersToFile(ByVal numbers As Integer(), ByVal fileName As String)
        Dim numsInString As New System.Text.StringBuilder
        For i As Integer = 0 To numbers.Length - 1
            numsInString.Append(numbers(i).ToString + ",")
        Next
        System.IO.File.WriteAllText(fileName, numsInString.ToString)
    End Sub
    Private Function ReadNumbersFromFile(ByVal fileName As String) As Long()
        Dim numsAsStrings As String() = System.IO.File.ReadAllText(fileName).Split(","c)
        Dim nums(numsAsStrings.Length - 1) As Long
        For i As Integer = 0 To numsAsStrings.Length - 2
            nums(i) = CType(numsAsStrings(i), Long)
        Next
        Return nums
    End Function

    Private Function GeneratePrimeList2(ByVal numOfPrimes As Int32) As Int32()
        Dim primes(numOfPrimes - 1) As Int32
        Dim cntr As Int32 = 5
        Dim pcntr As Int32 = 2
        Dim i As Int32

        Dim maxPrimeSqrRoot As Double


        primes(0) = 2
        primes(1) = 3

        Do Until pcntr = numOfPrimes



            'get the highest factor to check against if under the 
            'predetermined maximum
            maxPrimeSqrRoot = Math.Sqrt(cntr)
            'for each prime other than two
            For i = 1 To pcntr - 1
                'check for clean division
                If cntr Mod primes(i) = 0 Then
                    'not prime
                    Exit For
                Else
                    'if all primes up to maxfactor used or all primes used then it is prime
                    If primes(i + 1) > maxPrimeSqrRoot OrElse i = pcntr - 1 Then
                        primes(pcntr) = cntr
                        pcntr += 1
                        Exit For
                    End If
                End If
            Next
            cntr += 2
        Loop
        Return primes

    End Function

    Private Function GetPseudoPrimeSeriesMultiplier(ByVal pseudoLevel As Integer) As Long
        Dim primes As Integer() = GeneratePrimeList2(pseudoLevel)
        Dim multiplier As Long = 1
        For primeNumIndex As Integer = 0 To primes.Length - 1
            multiplier *= primes(primeNumIndex)
        Next
        Return multiplier
    End Function

    Private Function GetPseudoPrimeBaseSeries(ByVal pseudoLevel As Integer) As Long()
        Dim primes As Integer() = GeneratePrimeList2(pseudoLevel)
        Dim multiplier As Long = GetPseudoPrimeSeriesMultiplier(pseudoLevel)

        Dim psdPrimesList As New ArrayList

        Dim cnt As Integer = 0
        Dim testNumCount As Long = multiplier + GeneratePrimeList2(pseudoLevel + 1)(pseudoLevel) - 1
        For testNum As Long = 3 To testNumCount ' CType(testNumCount / 1000, Long) Step 2
            Dim isDivisible As Boolean = False
            For primeNumIndex As Integer = 0 To primes.Length - 1
                If testNum Mod primes(primeNumIndex) = 0 Then
                    isDivisible = True
                    Exit For
                End If
            Next

            If Not isDivisible Then
                psdPrimesList.Add(testNum)
                cnt += 1
            End If
        Next

        Dim psdPrimes As Long() = DirectCast(psdPrimesList.ToArray(GetType(Long)), Long())
        Return psdPrimes
    End Function


    Private Function GetCountOfBinaryPrimes(ByVal psdPrimes As Long()) As Long
        Dim lastPsdPrime As Long = 0
        Dim binPsdPrimeCount As Long = 0
        For psdPrimeIndex As Integer = 0 To psdPrimes.Length - 1
            If psdPrimes(psdPrimeIndex) - lastPsdPrime = 2 Then
                binPsdPrimeCount += 1
            End If
            lastPsdPrime = psdPrimes(psdPrimeIndex)
        Next
        Return binPsdPrimeCount
    End Function

    Private Sub WritePseudoPrimeSeriesesToExcel(ByVal pseudoLevel As Integer)
        Dim psdPrimes As Long() = GetPseudoPrimeBaseSeries(pseudoLevel)
        Dim multiplier As Long = GetPseudoPrimeSeriesMultiplier(pseudoLevel)

        Dim excelSheetContent As New System.Text.StringBuilder
        excelSheetContent.Append("<Table border=0>")
        For n As Long = 0 To 15 ' testNumCount
            excelSheetContent.Append("<TR>")
            For seriesIndex As Integer = 0 To psdPrimes.Length - 1

                Dim psdPrime As Long = psdPrimes(seriesIndex) + multiplier * n
                If n = 0 AndAlso seriesIndex > 0 AndAlso seriesIndex < (psdPrimes.Length - 1) AndAlso ((psdPrimes(seriesIndex + 1) - psdPrimes(seriesIndex) = 2) OrElse (psdPrimes(seriesIndex) - psdPrimes(seriesIndex - 1) = 2)) Then
                    excelSheetContent.Append("<TD bgcolor=yellow>")
                Else
                    excelSheetContent.Append("<TD>")
                End If

                If knownPrimes.Contains(psdPrime) Then
                    excelSheetContent.Append("<font  name=arial size=-1 color=green>")
                    excelSheetContent.Append(psdPrime)
                Else
                    excelSheetContent.Append("<font  name=arial size=-1 >")
                    excelSheetContent.Append(GetFactors(psdPrime))
                End If

                excelSheetContent.Append("</font></TD>")
            Next
            excelSheetContent.Append("</TR>")
        Next
        excelSheetContent.Append("</Table>")
        IO.File.WriteAllText("c:\temp\prm-" + pseudoLevel.ToString + ".html", excelSheetContent.ToString)
        'Process.Start("c:\temp\prm.html")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim pseudoLevel As Integer = CType(NumericUpDown1.Value, Integer)
        'WritePseudoPrimeSeriesesToExcel(pseudoLevel)

        Dim psdPrimes As Long() = GetPseudoPrimeBaseSeries(pseudoLevel)
        Dim multiplier As Long = GetPseudoPrimeSeriesMultiplier(pseudoLevel)
        Dim series1 As Long() = GetSeriesMembers(psdPrimes(0), multiplier, CType(psdPrimes(psdPrimes.Length - 1), Integer))
        Dim series2 As Long() = GetSeriesMembers(psdPrimes(1), multiplier, CType(psdPrimes(psdPrimes.Length - 1), Integer))

        Dim htmlContent As New System.Text.StringBuilder
        htmlContent.Append("<Table border=1>")
        For psdPrimeIndex As Integer = 0 To psdPrimes.Length - 1
            htmlContent.Append("<TR>")

            htmlContent.Append("<TD>")
            htmlContent.Append(psdPrimeIndex)
            htmlContent.Append("</TD>")

            htmlContent.Append("<TD>")
            Dim rawCoeff As Double = psdPrimes(psdPrimes.Length - 1)
            For thisUpperLimitIndex As Integer = 0 To psdPrimeIndex
                rawCoeff *= ((psdPrimes(thisUpperLimitIndex) - 2.0) / (psdPrimes(thisUpperLimitIndex)))
            Next
            htmlContent.Append(rawCoeff)
            htmlContent.Append("</TD>")

            htmlContent.Append("<TD>")
            Dim undivisiblePairCount As Integer = GetUndivisiblePairCount(series1, series2, psdPrimes, psdPrimeIndex)
            htmlContent.Append(undivisiblePairCount)
            htmlContent.Append("</TD>")

            htmlContent.Append("</TR>")
        Next
        htmlContent.Append("</Table>")

        IO.File.WriteAllText("c:\temp\PRM.html", htmlContent.ToString)
        Process.Start("c:\temp\prm.html")
    End Sub

    Private Function GetUndivisiblePairCount(ByVal series1 As Long(), ByVal series2 As Long(), ByVal psdPrimes As Long(), ByVal psdPrimesUpperLimit As Integer) As Integer
        Dim undivisiblePairCount As Integer = 0
        For seriesIndex As Integer = 0 To series1.Length - 1
            Dim isDivisible As Boolean = False
            For psdPrimeIndex As Integer = 0 To psdPrimesUpperLimit
                If ((series1(seriesIndex) Mod psdPrimes(psdPrimeIndex)) = 0) Or ((series2(seriesIndex) Mod psdPrimes(psdPrimeIndex)) = 0) Then
                    isDivisible = True
                    Exit For
                End If
            Next
            If isDivisible = False Then
                undivisiblePairCount += 1
            End If
        Next
        Return undivisiblePairCount
    End Function

    Private Function GetSeriesMembers(ByVal start As Long, ByVal multiplier As Long, ByVal howMany As Integer) As Long()
        Dim series(howMany) As Long
        For i As Integer = 1 To howMany
            series(i) = start + (multiplier * i)
        Next
        Return series
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim num As Long = CType(NumericUpDown2.Value, Long)
        Dim knownPrimeIndex As Integer = 0
        Dim factors As String = ""
        Do
            If num Mod knownPrimes(knownPrimeIndex) = 0 Then
                factors += knownPrimes(knownPrimeIndex).ToString + " * " + (num / knownPrimes(knownPrimeIndex)).ToString + vbCrLf
            End If
            knownPrimeIndex += 1
        Loop While num >= knownPrimes(knownPrimeIndex) / 2
        TextBox1.Text = factors
    End Sub

    Private Function GetFactors(ByVal num As Long) As String
        Dim knownPrimeIndex As Integer = 0
        Dim factors As String = ""
        Do
            If num Mod knownPrimes(knownPrimeIndex) = 0 Then
                factors += knownPrimes(knownPrimeIndex).ToString + "."
            End If
            knownPrimeIndex += 1
        Loop While num >= knownPrimes(knownPrimeIndex) / 2

        Return factors
    End Function
End Class
