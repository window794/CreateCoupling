Option Explicit
Option Base 1

Sub CreateCP()
    
    Dim clsCs As CastSpell: Set clsCs = New CastSpell 'VBA高速化おまじないを格納したクラスのインスタンス化
    
    Dim seme As String
    Dim uke As String
    
    Dim luck As Long
    Dim lastrow As Long
    
    Dim arrMember As Variant
    
    Worksheets(1).Activate
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    arrMember = Range(Cells(1, 1), Cells(lastrow, 1)) 'メンバを配列に格納

    '攻めを出す
    Randomize '乱数初期化
    luck = Int(lastrow * Rnd + 1)
    seme = arrMember(luck, 1)
    
    '受けを出す
    Randomize '乱数初期化
    luck = Int(lastrow * Rnd + 1)
    uke = arrMember(luck, 1)
    
    Do Until seme <> uke '攻めと受けが同じ値だったら、受けを出し直す
        
        Randomize
        luck = Int(lastrow * Rnd + 1)
        uke = arrMember(luck, 1)
        
    Loop
    
    
    '結果を出す
    Worksheets(2).Activate
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    lastrow = lastrow + 1
    
    Cells(lastrow, 1) = Now
    Cells(lastrow, 2) = seme
    Cells(lastrow, 3).Value = "×"
    Cells(lastrow, 4).Value = uke
    
    Range("A:D").EntireColumn.AutoFit
    Cells(lastrow, 1).Select
    
    Set clsCs = Nothing
    
    MsgBox "Done"

End Sub
