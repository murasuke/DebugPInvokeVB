Module DebugPInvoke
    ' API定義
    Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
        (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Object) As Integer
    Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
    Declare Function DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As Integer) As Integer

    Sub Main()
        Const PRINTER_NAME = "Microsoft Print to PDF"

        ' DeviceContext取得
        Dim hDC = CreateDC("WINSPOOL", PRINTER_NAME, 0, 0)

        ' プリンターの描画サイズを取得
        Dim width_pix = GetDeviceCaps(hDC, 8) ' 8:HORZRES, ピクセル単位の画面の幅
        Dim height_pix = GetDeviceCaps(hDC, 10) '10:VERTRES, ピクセル単位（ラスタ行数）の画面の高さ
        Debug.WriteLine(PRINTER_NAME & " : " & width_pix & " × " & height_pix)

        DeleteDC(hDC)
    End Sub
End Module
