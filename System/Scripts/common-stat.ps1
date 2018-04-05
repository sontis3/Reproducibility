# найти медиану
function Get-Median {
    <# 
    .Synopsis 
        Gets a median 
    .Description 
        Gets the median of a series of numbers 
    .Example 
        Get-Median 2,4,6,8 
    .Link 
        Get-Average 
    .Link 
        Get-StandardDeviation 
    #>
    param(
        # The numbers to average
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [Double[]]$Number
    )
    
    begin {
        $numberSeries = @()
    }
    
    process {
        $numberSeries += $number
    }
    
    end {
        $sortedNumbers = @($numberSeries | Sort-Object)
        if ($numberSeries.Count % 2) {
            # Odd, pick the middle
            $sortedNumbers[($sortedNumbers.Count / 2) - 1]
        } else {
            # Even, average the middle two
            ($sortedNumbers[($sortedNumbers.Count / 2)] + $sortedNumbers[($sortedNumbers.Count / 2) - 1]) / 2
        }                        
    }
}

# найти стандартное отклонение
function Get-StandardDeviation {            
    [CmdletBinding()]            
    param (            
        [double[]]$numbers            
    )            

    $avg = $numbers | Measure-Object -Average | Select-Object Count, Average            

    $popdev = 0            

    foreach ($number in $numbers){            
        $popdev +=  [math]::pow(($number - $avg.Average), 2)            
    }            

    $sd = [math]::sqrt($popdev / ($avg.Count-1))            
    $sd
}

# найти ковариацию
function Get-Covariance {            
    [CmdletBinding()]            
    param (            
        [double[]]$numbers1,            
        [double[]]$numbers2
    )

    $avg1 = $numbers1 | Measure-Object -Average | Select-Object Average            
    $avg2 = $numbers2 | Measure-Object -Average | Select-Object Average            

    $cov = 0                        
    for ($i=0; $i -le ($numbers1.length -1); $i++) {                        
        $cov += ($numbers1[$i]-$avg1.Average) * ($numbers2[$i]-$avg2.Average)              
    }

    $cov /= ($numbers1.length - 1)
    $cov
}

# коэффициент корреляции
function Get-Correlation {            
    [CmdletBinding()]            
    param (            
        [double[]]$numbers1,            
        [double[]]$numbers2
    )

    $sd1 = Get-StandardDeviation -numbers $numbers1
    $sd2 = Get-StandardDeviation -numbers $numbers2

    $cov = Get-Covariance -numbers1 $numbers1 -numbers2 $numbers2

    $correlation = $cov / ($sd1 * $sd2)                        
    $correlation
}

# наклон линейной регрессии
function Get-Slope {            
    [CmdletBinding()]            
    param (            
        [double[]]$x,            
        [double[]]$y
    )

    $sumXi = ($x | Measure-Object -Sum).Sum
    $sumYi = ($y | Measure-Object -Sum).Sum

    $xy = @()
    $x2 = @()
    for ($i = 0; $i -lt $x.Count; $i++) {
        $xy += $x[$i] * $y[$i]
        $x2 += $x[$i] * $x[$i]
    }
    $sumXYi = ($xy | Measure-Object -Sum).Sum
    $sumXi2 = ($x2 | Measure-Object -Sum).Sum

    $slope = ($sumXi * $sumYi - $x.Count * $sumXYi) / ($sumXi * $sumXi - $x.Count * $sumXi2)

    $slope
}

# посчитать общие статистики
function Get-CommonStats {            
    [CmdletBinding()]            
    param (            
        [PSCustomObject[]]$samples            
    )            

    $stats = [PSCustomObject]@{
        means = @();
        gmeans = @();
        median = @();
        minimum = @();
        maximum = @();
        SD = @();
        RSD = @();
    }
    foreach ($item in $samples) {
        if ($item.values[0].GetType().Name -eq "String") {
            $stats.means += @{name = $item.name; value = ""}
            $stats.gmeans += @{name = $item.name; value = ""}
            $stats.median += @{name = $item.name; value = ""}
            $stats.minimum += @{name = $item.name; value = ""}
            $stats.maximum += @{name = $item.name; value = ""}
            $stats.SD += @{name = $item.name; value = ""}
            $stats.RSD += @{name = $item.name; value = ""}
        } else {
            # $compcol = [Double]1.0
            # $item.values | ForEach-Object {
            #     $compcol *= $_
            # }
            $compcol = 0.0
            $valCount = 0
            $isNegativeExist = $false
            $item.values | ForEach-Object {
                if ($_ -gt 0.0) {
                    $compcol += [math]::Log($_)
                    $valCount++
                } elseif ($_ -lt 0.0) {
                    $isNegativeExist = $true
                }
            }
            $mo = $item.values | Measure-Object -Average -Minimum -Maximum
            $stats.means += @{name = $item.name; value = $mo.Average}
            if ($isNegativeExist) {
                $stats.gmeans += @{name = $item.name; value = "-"}
            } elseif ($valCount -gt 0) {
                $stats.gmeans += @{name = $item.name; value = [math]::Exp($compcol / $valCount)}
            } else {
                $stats.gmeans += @{name = $item.name; value = 0.0}
            }
            # $stats.gmeans += @{name = $item.name; value = [math]::Exp($compcol / $item.values.Count)}
            # $stats.gmeans += @{name = $item.name; value = [math]::pow($compcol, 1 / $item.values.Count)}
            $stats.median += @{name = $item.name; value = Get-Median $item.values}
            $stats.minimum += @{name = $item.name; value = $mo.Minimum}
            $stats.maximum += @{name = $item.name; value = $mo.Maximum}
            $sd = Get-StandardDeviation $item.values
            $stats.SD += @{name = $item.name; value = $sd}
            if ($mo.Average -ne 0) {
                $stats.RSD += @{name = $item.name; value = $sd / $mo.Average * 100.0}
            } else {
                $stats.RSD += @{name = $item.name; value = 0.00}
            }
        }
        # try {
        # }
        # catch {
        # }
    }

    $stats
}

#############################################################################################
#   input = z-value (-inf to +inf)
#   output = p under Standard Normal curve from -inf to z
#   e.g., if z = 0.0, function returns 0.5000
#   ACM Algorithm #209
function Gauss ($z)
{
#   double y; // 209 scratch variable
#   double p; // result. called 'z' in 209
#   double w; // 209 scratch variable
#   if (z == 0.0)
#     p = 0.0;
#   else
#   {
#     y = Math.Abs(z) / 2;
#     if (y >= 3.0)
#     {
#       p = 1.0;
#     }
#     else if (y < 1.0)
#     {
#       w = y * y;
#       p = ((((((((0.000124818987 * w
#         - 0.001075204047) * w + 0.005198775019) * w
#         - 0.019198292004) * w + 0.059054035642) * w
#         - 0.151968751364) * w + 0.319152932694) * w
#         - 0.531923007300) * w + 0.797884560593) * y * 2.0;
#     }
#     else
#     {
#       y = y - 2.0;
#       p = (((((((((((((-0.000045255659 * y
#         + 0.000152529290) * y - 0.000019538132) * y
#         - 0.000676904986) * y + 0.001390604284) * y
#         - 0.000794620820) * y - 0.002034254874) * y
#         + 0.006549791214) * y - 0.010557625006) * y
#         + 0.011630447319) * y - 0.009279453341) * y
#         + 0.005353579108) * y - 0.002141268741) * y
#         + 0.000535310849) * y + 0.999936657524;
#     }
#   }
#   if (z > 0.0)
#     return (p + 1.0) / 2;
#   else
#     return (1.0 - p) / 2;
    if ($z -eq 0.0) {
        $p = 0.0
    }
    else {
        $y = [math]::Abs($z) / 2
        if ($y -ge 3.0) {
            $p = 1.0
        }
        elseif ($y -lt 1.0) {
            $w = $y * $y
            $p = ((((((((0.000124818987 * $w `
                - 0.001075204047) * $w + 0.005198775019) * $w `
                - 0.019198292004) * $w + 0.059054035642) * $w `
                - 0.151968751364) * $w + 0.319152932694) * $w `
                - 0.531923007300) * $w + 0.797884560593) * $y * 2.0
        }
        else {
            $y = $y - 2.0
            $p = (((((((((((((-0.000045255659 * $y `
                + 0.000152529290) * $y - 0.000019538132) * $y `
                - 0.000676904986) * $y + 0.001390604284) * $y `
                - 0.000794620820) * $y - 0.002034254874) * $y `
                + 0.006549791214) * $y - 0.010557625006) * $y `
                + 0.011630447319) * $y - 0.009279453341) * $y `
                + 0.005353579108) * $y - 0.002141268741) * $y `
                + 0.000535310849) * $y + 0.999936657524
        }
    }
    if ($z -gt 0.0) {
        $p = ($p + 1.0) / 2
    }
    else {
        $p = (1.0 - $p) / 2
    }

    $p
}

# Calculating the Area Under the T-Distribution
# ACM algorithm #395
function Student ($t, $df) {
#   double n = df; // to sync with ACM parameter name
#   double a, b, y;
#   t = t * t;
#   y = t / n;
#   b = y + 1.0;
#   if (y > 1.0E-6) y = Math.Log(b);
#   a = n - 0.5;
#   b = 48.0 * a * a;
#   y = a * y;
#   y = (((((-0.4 * y - 3.3) * y - 24.0) * y - 85.5) /
#     (0.8 * y * y + 100.0 + b) + y + 3.0) / b + 1.0) * Math.Sqrt(y);
#   return 2.0 * Gauss(-y); // ACM algorithm 209
    $y = $t * $t / $df
    if ($y -gt 1.0E-6) {
        $y = [math]::Log($y + 1.0)
    }
    $a = $df - 0.5
    $b = 48.0 * $a * $a
    $y *= $a
    $y = (((((-0.4 * $y - 3.3) * $y - 24.0) * $y - 85.5) / (0.8 * $y * $y + 100.0 + $b) + $y + 3.0) / $b + 1.0) * [math]::Sqrt($y)

    $y = Gauss $(-$y)

    $y * 2.0
}

# найти p-value из t-value
function TtoP ($t, $df) {
    $absT = [math]::Abs($t)
    $t2 = $t * $t
    switch ($df) {
        1 { $p = 1 - 2 * [math]::atan($absT) / [math]::PI }
        2 { $p = 1 - $absT / [math]::Sqrt($t2 + 2)}
        3 { $p = 1 - 2 * ([math]::atan($absT / [math]::Sqrt(3)) + $absT * [math]::Sqrt(3) / ($t2 + 3)) / [math]::PI }
        4 { $p = 1 - $absT * (1 + 2 / ($t2 + 4)) / [math]::Sqrt($t2 + 4) }
        Default {
            # https://msdn.microsoft.com/en-us/magazine/mt620016.aspx
            $p = Student $t $df
        }
    }

    $p
}

function LJspin ($q, $i, $j, $b) {
  $zz = 1
  $z = $zz
  $k = $i
  while ($k -le $j) {
    $zz *= $q * $k / ($k - $b)
    $z += $zz 
    $k += 2
  }
  $z
}

# f cumulative distribution function
# http://www.jstor.org/pss/2683414   exact P-Value for an F-Test by hand
# Algorithm 346: F-test probabilities
# https://gist.github.com/Robsteranium/2662186
# https://www.easycalculation.com/statistics/f-test-p-value.php


function Fspin ($f, $df1, $df2) {
    $x = $df2 / ($df1 * $f + $df2)
    if ($df1 % 2 -eq 0) {
        $ret = (LJspin (1 - $x) $df2 ($df1 + $df2 - 4) ($df2 - 2)) * [math]::pow($x, $df2 / 2)
    }
    elseif ($df2 % 2 -eq 0) {
        $ret = 1 - ((LJspin $x $df1 ($df1 + $df2 - 4) ($df1 - 2)) * [math]::pow(1 - $x, $df1 / 2))
    }
    else {
        $tan = [math]::Atan([math]::Sqrt($df1 * $f / $df2))
        $a = 2 * $tan / [math]::PI
        $sat = [math]::Sin($tan)
        $cot = [math]::Cos($tan)
        if ($df2 -ge 1) {
            $a += $sat * $cot * (LJspin ($cot * $cot) 2 ($df2 - 3) -1) * 2 / [math]::PI
        }
        $c = 4 * (LJspin ($sat * $sat) ($df2 + 1) ($df1 + $df2 - 4) ($df2 - 2)) * $sat * [math]::pow($cot, $df2) / [math]::PI
        if ($df1 -eq 1) {
            $ret = 1 - $a
        }
        elseif ($df2 -eq 1) {
            $ret = 1 - $a + $c / 2
        }
        else {
            $k = 2
            while ($k -le ($df2 - 1) / 2) {
                $c *= $k / ($k - .5); 
                $k++
            }
            $ret = 1 - $a + $c
        }
    }

    $ret
}

# двухвыборочный t-критерий для независимых выборок с одинаковой дисперсией http://libguides.library.kent.edu/SPSS/IndependentTTest
function Get-T {            
    [CmdletBinding()]            
    param (            
        [int]$nT,            
        [double]$meanT,            
        [double]$sdT,            
        [int]$nR,
        [double]$meanR,
        [double]$sdR
    )

    $dF = $nT + $nR - 2     # число степеней свободы (degrees of freedom)
    $sP = [math]::sqrt((($nT - 1) * [math]::pow($sdT, 2) + ($nR - 1) * [math]::pow($sdR, 2)) / $dF)       # несмещенная оценка дисперсии (pooled standard deviation)
    $t = ($meanT - $meanR) / ($sP * [math]::sqrt((1 / $nT) + (1 / $nR)))
    $p = TtoP $t $dF

    if ($sdT -ge $sdR) {
        $sdMax = $sdT
        $sdMin = $sdR
        $ndf = $nT - 1
        $ddf = $nR - 1
    }
    else {
        $sdMax = $sdR
        $sdMin = $sdT
        $ndf = $nR - 1
        $ddf = $nT - 1
    }

    $fva = ($sdMax * $sdMax) / ($sdMin * $sdMin)
    $pvar = (Fspin $fva $ndf $ddf) * 2

    $result = [PSCustomObject]@{
        # meanT = $meanT;
        # meanR = $meanR;
        tValue = $t;
        dF = $dF;
        p = $p;
        fva = $fva;
        pVar = $pvar; 
    }

    $result
}
#############################################################################################


# вычисление квантилей распределения Стьюдента методом Хилла
# http://www.boost.org/doc/libs/1_63_0/boost/math/special_functions/detail/t_distribution_inv.hpp
# http://serostanov.blogspot.ru/2010/05/blog-post.html
# https://fossies.org/linux/gama/lib/gnu_gama/statan.cpp
# http://book2s.com/java/src/package/edu/cmu/tetrad/util/probutils.html

# функция ошибок
function Erfc () {
    [CmdletBinding()]            
    param (
        [double]$x
    )
    $a = 8 * ([math]::PI - 3) / (3 * [math]::PI * (4 - [math]::PI))
    $Result = $x / [math]::Abs($x) * [math]::Sqrt(1 - [math]::Exp(-$x * $x * (4 / [math]::PI + $a * $x * $x) / (1 + $a * $x * $x)))

    (1 - $Result)
}

# Asymptotic inverse expansion about normal
# обратная функция распределения стандартного нормального распределения
# returns a negative normal deviate at the lower tail probability level p 
function InverseNormalCDF () {
    [CmdletBinding()]            
    param (
        [double]$p
    )
    $a1 = -3.969683028665376e+01
    $a2 = 2.209460984245205e+02
    $a3 = -2.759285104469687e+02
    $a4 = 1.383577518672690e+02
    $a5 = -3.066479806614716e+01
    $a6 = 2.506628277459239e+00
    $b1 = -5.447609879822406e+01
    $b2 = 1.615858368580409e+02
    $b3 = -1.556989798598866e+02
    $b4 = 6.680131188771972e+01
    $b5 = -1.328068155288572e+01
    $c1 = -7.784894002430293e-03
    $c2 = -3.223964580411365e-01
    $c3 = -2.400758277161838e+00
    $c4 = -2.549732539343734e+00
    $c5 = 4.374664141464968e+00
    $c6 = 2.938163982698783e+00
    $d1 = 7.784695709041462e-03
    $d2 = 3.224671290700398e-01
    $d3 = 2.445134137142996e+00
    $d4 = 3.754408661907416e+00
    # break-points for tails
    $p_low = 0.02425
    $p_high = 1.0 - $p_low

    $q = $x = 0

    if (($p -gt 0.0) -and ($p -lt $p_low))              # Rational approximation for lower region
    {
        $q = [math]::sqrt(-2.0 * [math]::Log($p))
        $x = ((((($c1 * $q + $c2) * $q + $c3) * $q + $c4) * $q + $c5) * $q + $c6) / (((($d1 * $q + $d2) * $q + $d3) * $q + $d4) * $q + 1)
    } elseif (($p -ge $p_low) -and ($p -le $p_high))    # Rational approximation for central region
    {
        $q = $p - 0.5;
        $r = $q * $q;
        $x = ((((($a1 * $r + $a2) * $r + $a3) * $r + $a4) * $r + $a5) * $r + $a6) * $q / ((((($b1 * $r + $b2) * $r + $b3) * $r + $b4) * $r + $b5) * $r + 1.0)
    } elseif (($p -gt $p_high) -and ($p -lt 1.0))         # Rational approximation for upper region
    {
        $q = [math]::Sqrt(-2.0 * [math]::Log(1.0 - $p))
        $x = -((((($c1 * $q + $c2) * $q + $c3) * $q + $c4) * $q + $c5) * $q + $c6) / (((($d1 * $q + $d2) * $q + $d3) * $q + $d4) * $q + 1.0)
    } else {
        Write-Error "Параметр в ф-ции InverseNormalCDF () $p > 1.0"
    }
    # непонятный код. без него ф-ция дает правильный результат
    # if (($p -gt 0.0) -and ($p -lt 1.0))
    # {
    #     $e = 0.5 * (Erfc (-$x / [math]::Sqrt(2.0))) - $p
    #     $u = $e * [math]::Sqrt(2.0 * [math]::PI) * [math]::Exp($x * $x / 2.0)
    #     $x = $x - $u / (1.0 + $x * $u / 2.0)
    # }

    $x
}

# Student's t inverse cumulative distribution function
# n > 2
function Get-Inverse-Students-T-Hill () {
    [CmdletBinding()]            
    param (
        [double]$p,
        [int]$n
    )            
    $a = 1.0 / ($n - 0.5)
    $b = 48.0 / ($a * $a)
    $c = ((20700.0 * $a / $b - 98.0) * $a - 16.0) * $a + 96.36
    $d = ((94.5 / ($b + $c) - 3.0) / $b + 1.0) * [math]::sqrt($a * [math]::PI / 2.0) * $n
    $x = $d * $p
    $y = [math]::pow($x, (2.0 / $n))
    if ($y -gt (0.05 + $a)) {
        $x = InverseNormalCDF ($p * 0.5)
        $y = $x * $x
        if ($n -lt 5) {
            $c += 0.3 * ($n - 4.5) * ($x + 0.6)
        }
        $c += (((0.05 * $d * $x - 5.0) * $x - 7.0) * $x - 2.0) * $x + $b
        $y = (((((0.4 * $y + 6.3) * $y + 36.0) * $y + 94.5) / $c - $y - 3.0) / $b + 1.0) * $x
        $y = $a * $y * $y
        if ($y -gt 0.002) {
            $y = [math]::Exp($y) - 1.0
        } else {
            $y += 0.5 * $y * $y
        }
    }
    else {
        $y = ((1.0 / ((($n + 6.0) / ($n * $y) - 0.089 * $d - 0.822) * ($n + 2.0) * 3.0) + 0.5 / ($n + 4.0)) * $y - 1.0) * ($n + 1.0) / ($n + 2.0) + 1.0 / $y
    }
    [math]::Sqrt($n * $y)
}

# Get-Inverse-Students-T-Hill -p 0.05 -n 23
# Get-Inverse-Students-T-Hill -p 0.2 -n 23
# Get-Inverse-Students-T-Hill -p 0.1 -n 16
# Get-Inverse-Students-T-Hill -p 0.2 -n 16