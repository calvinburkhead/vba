Function ConvertBytesToHumanReadable(bytes)
    Dim KiB, MiB, GiB, TiB, PiB, EiB, ZiB, YiB, B, suffixDivisor, suffixDisplay
    
    B = Application.WorksheetFunction.Power(2, 0)
    KiB = Application.WorksheetFunction.Power(2, 10)
    MiB = Application.WorksheetFunction.Power(2, 20)
    GiB = Application.WorksheetFunction.Power(2, 30)
    TiB = Application.WorksheetFunction.Power(2, 40)
    PiB = Application.WorksheetFunction.Power(2, 50)
    EiB = Application.WorksheetFunction.Power(2, 60)
    ZiB = Application.WorksheetFunction.Power(2, 70)
    YiB = Application.WorksheetFunction.Power(2, 80)
    
    If Not (IsNumeric(bytes)) Then bytes = 0
    Select Case Int(Log(bytes) / Log(KiB))
    Case 1
        suffixDivisor = KiB
        suffixDisplay = "KiB"
    Case 2
        suffixDivisor = MiB
        suffixDisplay = "MiB"
    Case 3
        suffixDivisor = GiB
        suffixDisplay = "GiB"
    Case 4
        suffixDivisor = TiB
        suffixDisplay = "TiB"
    Case 5
        suffixDivisor = PiB
        suffixDisplay = "PiB"
    Case 6
        suffixDivisor = EiB
        suffixDisplay = "EiB"
    Case 7
        suffixDivisor = ZiB
        suffixDisplay = "ZiB"
    Case 8
        suffixDivisor = YiB
        suffixDisplay = "YiB"
    Case Else
        suffixDivisor = B
        suffixDisplay = "B"
    End Select
        
        ConvertBytesToHumanReadable = Round(bytes / suffixDivisor, 2) & " " & suffixDisplay
End Function

Function ConvertBytesToGibibytes(bytes)
    Dim GiB, suffixDivisor
    
    GiB = Application.WorksheetFunction.Power(2, 30)
    
    If Not (IsNumeric(bytes)) Then bytes = 0
    ConvertBytesToGibibytes = Round(bytes / GiB, 2)
End Function

