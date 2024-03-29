Option Explicit
Sub Main()
    'Clear information
    Debug.Clear

'Define OpenStaad Reference
Dim Func As Object
Set Func = GetObject( , "StaadPro.OpenSTAAD")

'Define general parameter
 Dim i As Integer
 Dim j As Integer
 Dim Result As Variant

'Deleting the exsiting group name before create the new groups
    Dim GroupName() As String
    ReDim GroupName(6) As String
    GroupName(0) = "_B"
    GroupName(1) = "_FB"
    GroupName(2) = "_NB"
    GroupName(3) = "_FNB"
    GroupName(4) = "_FFNB"
    GroupName(5) = "_BB"
    GroupName(6) = "_Max"
    'Deleting Group Operation
        'For i = 0 To 6
            'Func.Geometry.DeleteGroup(GroupName(i))
        'Next

'Silent Analysis the staad file before the operation
        'Func.AnalyzeEx(1, 0, 0)

'Check total member in the model
    Dim TotalMem As Long
    TotalMem = Func.Geometry.GetMemberCount()
    'Debug.Print "Total Member in model: " ; TotalMem

    'Get Member Data then the Member Data after check beam or column will be used to be the parameter to collect all data by using to be 2 dimonsion array
        'Define MemData() to be 2 Dimensional Array (8,TotalMem-1)
            '(MemNo, MaxMo, MoLoadNo, MoGroup, MaxShear, ShearNo, ShearGroup, BeamOrCol,Group )
    Dim MemData() As Variant
    ReDim MemData(9,TotalMem-1) As Variant
    Dim MemList() As Long
    ReDim MemList(TotalMem-1) As Long
    'Debug.Print "MemList Array size: " ; UBound(MemList)
    Func.Geometry.GetBeamList(MemList())
        'Copy MemList to MemData and Printing all member list
        'For i = 0 To TotalMem-1
            'MemData(1,i) = MemList(i)
            ''Debug.Print "Mem No : " ; MemData(1,i)
        'Next

'Check Load Case & Load Combination
    'Check Load Case then get the name
        Dim TotalLoadCase As Long
        TotalLoadCase = Func.Load.GetPrimaryLoadCaseCount()
        'Debug.Print "Total Load Case : " ; TotalLoadCase
            'Get the load case number
            Dim LoadCaseNum() As Long
            ReDim LoadCaseNum(TotalLoadCase-1) As Long
            Func.Load.GetPrimaryLoadCaseNumbers(LoadCaseNum())
                'Check Load Case number by printing
                'For i = 0 To UBound(LoadCaseNum)
                    'Debug.Print "Load Case Number: " ; LoadCaseNum(i)
                'Next
    'Check Load Combination
        Dim TotalLoadComb As Long
        TotalLoadComb = Func.Load.GetLoadCombinationCaseCount()
        'Debug.Print "Total Load Combination: "; TotalLoadComb
            'Get Load Combination number
            Dim LoadCombNum() As Long
            ReDim LoadCombNum(TotalLoadComb-1) As Long
            Func.Load.GetLoadCombinationCaseNumbers(LoadCombNum())
                'Check Load Comb Number by printing
                'For i = 0 To UBound(LoadCombNum)
                    'Debug.Print "Load Combination Number: " ; LoadCombNum(i)
                'Next
    'Assembly Load Case & Load Combination
    Dim TotalLoads() As Long
    ReDim TotalLoads(UBound(LoadCaseNum) + UBound(LoadCombNum)+1)
    'Debug.Print "UBound(TotalLoads): " ; UBound(TotalLoads)
    For i = 0 To UBound(LoadCaseNum)
        TotalLoads(i) = LoadCaseNum(i)
        'Debug.Print "A"
    Next
    i = 0
    For j = UBound(LoadCaseNum)+1 To UBound(TotalLoads)
        TotalLoads(j) = LoadCombNum(i)
        i = i +1
        'Debug.Print "B"
    Next
        'Check Total Loads by printing
        'For i = 0 To UBound(TotalLoads)
            'Debug.Print "All Load : " ; TotalLoads(i)
        'Next

'Moment and Shear Section to set the category of member
    'Finding Max Moment Value by comparing every load case in model
    Dim MoDir As String
    MoDir = "MZ"
    Dim MinMo As Double
    Dim MinMoPos As Double
    Dim MaxMo As Double
    Dim MaxMoPos As Double
    Dim AbsoMaxMo As Double
    For i = 0 To TotalMem-1
            MemData(1,i) = MemList(i)
        For j = 0 To UBound(TotalLoads())
            'Member No have to be "Long" only to complete the operation
            Func.Output.GetMinMaxBendingMoment(MemList(i), MoDir, TotalLoads(j), MinMo, MinMoPos, MaxMo, MaxMoPos)
                If Abs(MaxMo) > Abs(MinMo) Then
                    AbsoMaxMo = Abs(MaxMo)
                    Else
                    AbsoMaxMo = Abs(MinMo)
                End If
                If MemData(2,i) < AbsoMaxMo Then
                    MemData(2,i) = AbsoMaxMo
                    MemData(3,i) = TotalLoads(j)
                    Else
                    'Nothing
                End If
            'Debug.Print "Information : " ; MemList(i) ; TotalLoads(j) ; MaxMo
        Next
                'Check Information after the Moment Operation
                'Debug.Print "Mem No : " ; MemData(1,i) ; " Max Moment : " ; MemData(2,i) ; " Load No : " ; MemData(3,i)
    Next

    'Finding Max Shear Value by comparing every load case in model
    Dim ShearDir As String
    ShearDir = "FY"
    Dim MinShear As Double
    Dim MinShearPos As Double
    Dim MaxShear As Double
    Dim MaxShearPos As Double
    Dim AbsoMaxShear As Double
    For i = 0 To TotalMem-1
        For j = 0 To UBound(TotalLoads())
            'Member No have to be "Long" only to complete the operation
            Func.Output.GetMinMaxShearForce(MemList(i), ShearDir, TotalLoads(j), MinShear, MinShearPos, MaxShear, MaxShearPos)
                If Abs(MaxShear) > Abs(MinShear) Then
                    AbsoMaxShear = Abs(MaxShear)
                    Else
                    AbsoMaxShear = Abs(MinShear)
                End If
                If MemData(5,i) < AbsoMaxShear Then
                    MemData(5,i) = AbsoMaxShear
                    MemData(6,i) = TotalLoads(j)
                    Else
                    'Nothing
                End If
        Next
                'Check Information after the Shear Operation
                'Debug.Print "Mem No : " ; MemData(1,i) ; " Max Shear : " ; MemData(5,i) ; " Load No : " ; MemData(6,i)
    Next

    'Category the member by Moment First then Shear. Finally compare between two type of then and select the maximum group value
        'Moment Group check
            For i = 0 To TotalMem-1
                If MemData(2,i)*101.9644 <= 2870 Then
                    MemData(4,i) = 1
                    Else
                    If MemData(2,i)*101.9644 <= 2875 Then
                        MemData(4,i) = 2
                        Else
                        If MemData(2,i)*101.9644 <= 5230 Then
                            MemData(4,i) = 3
                            Else
                            If MemData(2,i)*101.9644 <= 5235 Then
                                MemData(4,i) = 4
                                Else
                                If MemData(2,i)*101.9644 <= 5240 Then
                                    MemData(4,i) = 5
                                    Else
                                    If MemData(2,i)*101.9644 <= 12230 Then
                                        MemData(4,i) = 6
                                        Else
                                        MemData(4,i) = 7
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next

        'Shear Group check
            For i = 0 To TotalMem-1
                If MemData(5,i)*101.9644 <= 3195 Then
                    MemData(7,i) = 1
                    Else
                    If MemData(5,i)*101.9644 <= 3610 Then
                        MemData(7,i) = 2
                        Else
                        If MemData(5,i)*101.9644 <= 4035 Then
                            MemData(7,i) = 3
                            Else
                            If MemData(5,i)*101.9644 <= 4350 Then
                                MemData(7,i) = 4
                                Else
                                If MemData(5,i)*101.9644 <= 4840 Then
                                    MemData(7,i) = 5
                                    Else
                                    If MemData(5,i)*101.9644 <= 8130 Then
                                        MemData(7,i) = 6
                                        Else
                                        MemData(7,i) = 7
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        'Select the maximum
            For i = 0 To TotalMem-1
                If MemData(4,i) < MemData(7,i) Then
                    MemData(9,i) = MemData(7,i)
                    Else
                    MemData(9,i) = MemData(4,i)
                End If
            Next

    'Creating Group Operation but Delete the existing group and check "Is it a Beam" First
        For j = 0 To 6
            Func.Geometry.DeleteGroup(GroupName(j))
            For i = 0 To TotalMem-1
                'Have to use MemList(i) becasue MemData(1,i) is Variant that did not work for Staad API
                If Func.Geometry.IsBeam(MemList(i), 10) = True Then
                    If MemData(9,i) = j+1 Then
                        Func.Geometry.SelectBeam(MemList(i))
                        Else
                        'Nothing
                    End If
                    Else
                    'Nothing
                End If
            Next
            Func.Geometry.CreateGroup(2, GroupName(j))
            Func.Geometry.ClearMemberSelection()
        Next

'Silent Analysis the staad file after the operation
        Func.AnalyzeEx(1, 0, 0)
End Sub
