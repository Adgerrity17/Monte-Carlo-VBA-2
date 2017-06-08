Attribute VB_Name = "Module1"
Function MC_Call(S0, Exercise, Mean, Sigma, Interest, Time, Divisions, Runs)

deltat = Time / Divisions
interestdelta = Exp(Interest * deltat)

Up = Exp(Mean * deltat + Sigma * Sqr(deltat))
down = Exp(Mean * deltat - Sigma * Sqr(deltat))

pathlength = Int(Time / deltat)

'Risk Neutral Probabilities
piup = (interestdelta - down) / (Up - down)
pidown = 1 - piup

Temp = 0

For Index = 1 To Runs
    Upcounter = 0
    'generate terminal price
    For j = 1 To pathlength
    If Rnd > pidown Then Upcounter = Upcounter + 1
        If S0 * Up ^ (Upcounter + pathlenght - j) * down ^ (j - Upcounter) < X Then GoTo Compute
        Next j
Compute:
        Callvalue = Application.Max(S0 * (Up ^ Upcounter) * (down ^ (pathlenght - Upcounter)) - Exercise, 0) / (interestdelta ^ pathlenght)
        Temp = Temp + Callvalue
Next Index

MC_Call = Temp / Runs
        
        
End Function
