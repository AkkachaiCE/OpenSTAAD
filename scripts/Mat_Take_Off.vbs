Option Explicit

Sub Main()
 Debug.Clear
    Dim StaadFunc As Object
    Set StaadFunc = GetObject( , "StaadPro.OpenSTAAD")

Dim i As Integer
Dim j As Integer

'Get Total Member
    Dim MemCount As Long
    MemCount = StaadFunc.Geometry.GetMemberCount()
    'Debug.Print "Total Member : " ; MemCount
        Dim MemData() As Variant
        ReDim MemData(MemCount-1,8) As Variant
'Get Member List
    Dim MemList() As Long
    ReDim MemList(MemCount-1) As Long
    StaadFunc.Geometry.GetBeamList(MemList())

For i = 0 To MemCount-1
    MemData(i , 1) = MemList(i)
    'Debug.Print "Mem No : " ; MemList(i)
Next

'Get Member length
Dim MemLength() As Double
ReDim MemLength(MemCount-1) As Double
For i = 0 To MemCount-1
    MemLength(i) = StaadFunc.Geometry.GetBeamLength(MemList(i))
    'Debug.Print "Mem No : " ; MemList(i) ; " Length : " ; Round(MemLength(i),3) ; " m."
    MemData(i, 3) = MemLength(i)
Next

'Get Member Property & Section name
Dim MemProps As Variant
Dim WidMem As Double
Dim DepMem As Double
Dim AxMem As Double
Dim AyMem As Double
Dim AzMem As Double
Dim IxMem As Double
Dim IyMem As Double
Dim IzMem As Double

Dim SecMem As String
For i = 0 To MemCount-1
    SecMem = StaadFunc.Property.GetBeamSectionName(MemList(i))
    MemProps = StaadFunc.Property.GetBeamProperty(MemList(i), WidMem, DepMem, AxMem, AyMem, AzMem, IxMem, IyMem, IzMem)
    'Debug.Print "Mem No : " ; MemList(i) ; " Name : " ; SecMem ; " CrossSection : " ; AxMem
    MemData(i, 2) = SecMem
    MemData(i, 4) = AxMem
Next

'Get Material Name & Property
Dim MemMat() As String
ReDim MemMat(MemCount-1) As String
For i = 0 To MemCount-1
    MemMat(i) = StaadFunc.Property.GetBeamMaterialName(MemList(i))
    'Debug.Print " No : " ; MemList(i) ; " Mat Name : " ; MemMat(i)
    MemData(i, 5) = MemMat(i)
Next

    Dim MemElas As Double
    Dim MemPoi As Double
    Dim MemDen() As Double
    ReDim MemDen(MemCount-1) As Double
    Dim MemAlp As Double
    Dim MemDam As Double
    For i = 0 To MemCount-1
        StaadFunc.Property.GetMaterialProperty(MemMat(i), MemElas, MemPoi, MemDen(i), MemAlp, MemDam)
        'Debug.Print "Mat Name : " ; MemMat(i) ; " Density : " ; MemDen(i)
        MemData(i, 6) = MemDen(i)
'Get Weight
        MemData(i, 7) = MemData(i, 4) * MemData(i, 3) * MemData(i, 6)
        'Debug.Print "No :" ; MemData(i, 1) ; " Section : " ; MemData(i, 2) ; " Length : " ; MemData(i, 3) ; " Area : " ; MemData(i, 4) ; " Material : " ; MemData(i, 5) ; " Density : " ; MemData(i, 6) ; " Weight : " ; MemData(i, 7)
    Next

'Select Case And Sumation of values
    Dim SecCount As Long
    SecCount = StaadFunc.Property.GetSectionPropertyCount()
    'Debug.Print "Number of Different Section : " ; SecCount
    For i = 0 To MemCount-1
        MemData(i, 8) = StaadFunc.Property.GetBeamSectionPropertyRefNo(MemList(i))
        'Debug.Print "Ref No : "; MemData(i, 8)
        'Debug.Print MemData(i, 8)
    Next

'Finding Max & Min of Ref Member
'MaxRef Finding
Dim MaxRef As Long
MaxRef = 0
Dim MinRef As Long
MinRef = 1
For i = 0 To MemCount-1
    'If i = MemCount-1 Then
        'Nothing

        'Else
            If MaxRef < MemData(i, 8) Then
        MaxRef = MemData(i, 8)
                Else
                
        'Nothing

            End If
    'End If
Next
    'Debug.Print "Max Ref No : " ; MaxRef
Dim RefList() As Variant
ReDim RefList(SecCount-1,5) As Variant
'MinRef Finding
'For i = 0 To MemCount-1
    ''If i = MemCount-1 Then
        ''Nothing
        ''Else
            'If MinRef > MemData(i, 8) Then
        'MinRef = MemData(i, 8)
                'Else
                ''MinRef = MinRef
        ''Nothing
            'End If
    ''End If
'Next
    'Debug.Print "Min Ref No : " ; MinRef
'Check Member Ref no
Dim k As Integer
k = 0
For j = MinRef To MaxRef
    'Debug.Print "Ref Mem : "; j
    For i = 0 To MemCount-1
        If j = MemData(i, 8) Then
            RefList(k,1) = MemData(i, 8)
            k = k+1
            Exit For
            Else
            'Nothing
        End If

    Next
    'Debug.Print "Ref Array no : " ; RefList(k,1)
    'k = k+1
Next

'For i = 0 To SecCount-1
    'Debug.Print " Ref No : "; RefList(i,1)
'Next

'Quantity Section
For j = 0 To UBound(RefList())
    For i = 0 To MemCount-1
        If MemData(i, 8) = RefList(j,1) Then
            RefList(j, 2) = MemData(i, 2)
            RefList(j, 3) = RefList(j, 3) + MemData(i, 3)
            RefList(j, 4) = RefList(j, 4) + MemData(i, 7)
            RefList(j, 5) = MemData(i, 5)
            Else
            'Nothing
        End If

    Next
    'Debug.Print "Section Name : " ; RefList(j, 2) ; " Sum Length : " ; Round(RefList(j, 3),3) ; " Sum Weight : " ; Round(RefList(j, 4),3) ; " Material : " ; RefList(j, 5)
Next

'Convert KN to Kg >> 101.9644 and round up the decimal
'Summary Length and Weight
Dim SumLen As Double
Dim SumWei As Double
For i = 0 To UBound(RefList)
    SumLen = SumLen + RefList(i, 3)
    SumWei = SumWei + RefList(i, 4)
Next
    'Debug.Print "Sum Length : " ; Round(SumLen,2)
    'Debug.Print "Sum Weight : " ; Round(SumWei,3)

'Report the result to the table (Write Table Section)
'Dim SheetNo As Long
Dim TableNo1 As Long
Dim SheetNo1 As Long
SheetNo1 = StaadFunc.Table.CreateReport("Material Take Off")

TableNo1 = StaadFunc.Table.AddTable(SheetNo1, "Quantity Of Material", SecCount+1, 4)

StaadFunc.Table.SetColumnHeader(SheetNo1, TableNo1, 1, "Section Name")
StaadFunc.Table.SetColumnUnitString(SheetNo1, TableNo1, 1, "")

StaadFunc.Table.SetColumnHeader(SheetNo1, TableNo1, 2, "Sum of Length")
StaadFunc.Table.SetColumnUnitString(SheetNo1, TableNo1, 2, "(m)")

StaadFunc.Table.SetColumnHeader(SheetNo1, TableNo1, 3, "Sum of Weight")
StaadFunc.Table.SetColumnUnitString(SheetNo1, TableNo1, 3, "(Kg)")

StaadFunc.Table.SetColumnHeader(SheetNo1, TableNo1, 4, "Material")
StaadFunc.Table.SetColumnUnitString(SheetNo1, TableNo1, 4, "")

'Fill Data to Table
'Convert KN to Kg and roundup to 3 Decimal
For i = 0 To SecCount-1
    RefList(i, 3) = Round(RefList(i, 3), 2)
    RefList(i, 4) = Round((RefList(i, 4))*101.9644, 3)
Next

Dim ResTable As Variant
For i = 1 To SecCount
    For j = 1 To 4
        'Debug.Print "RefList : "; RefList(i, j)
        ResTable = StaadFunc.Table.SetCellValue(SheetNo1, TableNo1, i, j, CStr(RefList(i-1, j+1)))
    Next
Next
        StaadFunc.Table.SetCellValue(SheetNo1, TableNo1, SecCount+1, 1, "Sum")
        'StaadFunc.Table.SetCellValue(SheetNo1, TableNo1, SecCount+1, 2, CStr(Round(SumLen,2)))
        StaadFunc.Table.SetCellValue(SheetNo1, TableNo1, SecCount+1, 3, CStr(Round((SumWei)*101.9644,3)))

End Sub
