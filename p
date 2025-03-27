value = 0x000000  # 分割したい値（例: 24ビット値）

# 上位8ビットを取得
upper = (value >> 16) & 0xFF

# 中位8ビットを取得
middle = (value >> 8) & 0xFF

# 下位8ビットを取得
lower = value & 0xFF

# 結果を出力
print(f"Upper: {upper:#04x}")
print(f"Middle: {middle:#04x}")
print(f"Lower: {lower:#04x}")

Dim readSize As Long
Dim bytesProcessed As Long
Dim readBuffer(255) As Byte

readSize = 1 ' 読み取るバイト数
bytesProcessed = 0

ftstatus = FT260_I2CMaster_Read(ftHandle, i2cAddr, 0, readBuffer(0), readSize, bytesProcessed)
Dim readSize As Long
Dim bytesProcessed As Long
Dim readBuffer(4) As Byte ' 5バイト分のバッファ (0~4の5つの要素)

readSize = 5 ' 読み取るバイト数
bytesProcessed = 0

ftstatus = FT260_I2CMaster_Read(ftHandle, i2cAddr, 0, readBuffer(0), readSize, bytesProcessed)

Sub Execute_I2C_Command()
    Dim ftHandle As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ftstatus As Long
    Dim i2cAddr As Byte
    Dim regAddr As Integer
    Dim readSize As Long
    Dim bytesProcessed As Long
    Dim readBuffer(4) As Byte ' 5バイトのバッファを確保

    ' ワークシート設定
    Set ws = ThisWorkbook.Sheets(2)
    
    ' 最終行を取得（B列の最終データ行）
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' FT260を開く
    ftHandle = FT260OpenHandle(ws)
    If ftHandle <= 0 Then
        MsgBox "FT260の接続に失敗しました。", vbCritical
        Exit Sub
    End If

    ' データ処理ループ
    For i = 4 To lastRow
        Dim deviceAddr As Byte
        Dim bank As Byte
        Dim addr As Byte
        Dim readValue As String

        ' セルの値を取得（I2Cアドレス）
        bank = ws.Cells(i, 2).Value
        deviceAddr = ws.Cells(i, 3).Value
        addr = ws.Cells(i, 4).Value

        ' I2Cアドレスの設定（16進変換）
        i2cAddr = Val("&H" & deviceAddr)

        ' 読み取り処理
        If ws.Cells(i, 5).Value <> "" Then
            readSize = 5 ' 5バイト読み取る
            bytesProcessed = 0

            ftstatus = FT260_I2CMaster_Read(ftHandle, i2cAddr, 0, readBuffer(0), readSize, bytesProcessed)

            If ftstatus = 0 Then
                ' 読み取ったデータをExcelに書き込み
                readValue = ""
                For j = 0 To readSize - 1
                    readValue = readValue & Right("0" & Hex(readBuffer(j)), 2) & " "
                Next j
                ws.Cells(i, 7).Value = Trim(readValue) ' RESULT列にデータを表示
            Else
                ws.Cells(i, 7).Value = "ERR" ' エラー表示
            End If
        End If
    Next i

    ' FT260を閉じる
    FT260_Close ftHandle
    MsgBox "I2C通信が完了しました。", vbInformation
End Sub

