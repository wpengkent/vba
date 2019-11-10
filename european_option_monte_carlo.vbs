'---------------------------
' Set lower bound for array
'--------------------------
Option Base 1

'--------------------------------------------------------
' Inputs
'  s = spot price (share price)
'  k = strike price
'  t = time to maturity
'  signma = volatility
'  r = risk-free rate
'  time_step = number of time steps
'  n_sim = number of random iteration per time step
'--------------------------------------------------------

Function EuropeanOption(
                        s As Double,
                        k As Double,
                        t As Double,
                        sigma As Double,
                        r As Double,
                        time_step As Double,
                        n_sim As Double
                       ) As Variant

    '-----------------------------------------------------------------
    ' dt = fraction of time
    ' z = random number (shock)
    ' dlns = log normal return
    ' option_price = option price for each time step
    ' stock_price_array = two dimentional array to store stock price
    ' option_price_array = vector to store option price for each sim
    '-----------------------------------------------------------------
    Dim dt, z, dlns, option_price, stock_price_array(), option_price_array() As Double
    ReDim stock_price_array(n_sim, time_step + 1)
    ReDim option_price_array(n_sim)
    dt = t / time_step

    For i = 1 To n_sim
        'Initialze array, store spot price
        stock_price_array(i, 1) = s

        ' Initialze random-number genrator
        Randomize

        For j = 1 To time_step
            z = WorksheetFunction.NormSInv(Rnd())
            dlns = (r - sigma^2 / 2) * dt + z * sigma * dt^0.5
            stock_price_array(i, j + 1) = stock_price_array(i, j) * Exp(dlns)
        Next j
        option_price_array(i) = WorksheetFunction.Max(stock_price_array(i, j) - k, 0) * Exp(-r * t)
    Next i
    option_price = 0
    For i = 1 To n_sim
        option_price = option_price + option_price_array(i)
    Next i
    option_price = option_price / n_sim
    EuropeanOption= option_price
End Function

testing