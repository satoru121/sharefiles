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

VBA
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Sub Execute_I2C_Command()
    Dim ftHandle As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ftstatus As Long
    Dim i2cAddr As Byte
    Dim regAddr As Integer
    Dim WaitTimeMs As Long
    Dim I2CSpeed As Long

    ' ワークシート設定
    Set ws = ThisWorkbook.Sheets(2)
    
    ' B列（I2CADR）の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' FT260接続
    ftHandle = FT260OpenHandle(ws)
    If ftstatus <> 0 Then
        MsgBox "FT260の接続に失敗しました。エラーコード: " & ftstatus, vbCritical
        Exit Sub
    End If

    ' I2C速度の設定
    I2CSpeed = ws.Cells(1, 8).Value ' 例: セル H1 に 100000, 400000, 1000000 のいずれかを設定
    ftstatus = FT260_I2CMaster_SetClockSpeed(ftHandle, I2CSpeed)
    If ftstatus <> 0 Then
        MsgBox "I2C速度の設定に失敗しました。エラーコード: " & ftstatus, vbCritical
        Exit Sub
    End If

    ' 各行のデータを処理
    For i = 4 To lastRow
        Dim deviceAddr As Byte
        Dim bank As Byte
        Dim addr As Byte
        Dim readSize As Byte
        Dim writeData As Byte
        Dim bytesProcessed As Long

        ' I2CADR取得
        i2cAddr = Val("&H" & ws.Cells(i, 2).Value)
        bank = Val(ws.Cells(i, 3).Value)
        addr = Val("&H" & ws.Cells(i, 4).Value)
        regAddr = Val("&H" & ws.Cells(i, 4).Value)

        ' Wait時間取得
        WaitTimeMs = ws.Cells(i, 9).Value ' 例: セル I列 に 10, 100 などの値を入力

        ' Read処理
        If ws.Cells(i, 5).Value <> "" Then
            readSize = Val(ws.Cells(i, 5).Value)
            Dim readBuffer(0) As Byte
            
            ftstatus = FT260_I2CMaster_Read(ftHandle, i2cAddr, 0, readBuffer(0), readSize, bytesProcessed)
            If ftstatus = 0 Then
                ws.Cells(i, 7).Value = Right("0" & Hex(readBuffer(0)), 2) ' 読み出しデータ
            Else
                ws.Cells(i, 7).Value = "ERR"
            End If
        End If
        
        ' 指定した時間待機
        Sleep (WaitTimeMs)

        ' Write処理
        If ws.Cells(i, 6).Value <> "" Then
            Dim writeBuffer(1) As Byte
            writeBuffer(0) = regAddr
            writeBuffer(1) = Val("&H" & ws.Cells(i, 6).Value)
            
            ftstatus = FT260_I2CMaster_Write(ftHandle, i2cAddr, 0, writeBuffer(0), 2, bytesProcessed)
            If ftstatus <> 0 Then
                ws.Cells(i, 7).Value = "ERR"
            End If
        End If
        
        ' 指定した時間待機
        Sleep (WaitTimeMs)
    Next i

    ' FT260接続解除
    FT260_Close (ftHandle)
    MsgBox "I2C通信が完了しました。", vbInformation
End Sub
使用方法
I2C速度の設定

H1 セルに 100000、400000、1000000 のいずれかを入力

Wait時間の設定

I列 (9列目) に 10 や 100 などのミリ秒単位の値を設定

各行の処理後にその値だけ待機

コードの実行

Execute_I2C_Command を実行すれば、設定された速度でI2C通信が行われ、適切なWait時間が適用される

これでI2C通信がより柔軟に制御できるようになりました！






