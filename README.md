' ==================================================================================
' HƯỚNG DẪN CÀI ĐẶT ĐẦY ĐỦ - VBA VỚI USERFORM
' ==================================================================================
'
' BƯỚC 1: TẠO USERFORM
' --------------------
' 1. Mở file Excel nhập liệu
' 2. Nhấn Alt + F11 (mở VBA Editor)
' 3. Click Insert > UserForm (tạo form mới)
' 4. Trong cửa sổ Toolbox (nếu không thấy: View > Toolbox), kéo thả:
'    a. Label (từ Toolbox vào form)
'    b. ListBox (từ Toolbox vào form)
'    c. CommandButton (kéo 2 cái - một cho OK, một cho Cancel)
'
' 5. THIẾT LẬP PROPERTIES (cửa sổ Properties bên phải, nếu không thấy: View > Properties Window):
'
'    Chọn UserForm (click vào vùng trống của form):
'    - Tìm dòng (Name): gõ vào: frmSelectProduct
'    - Tìm dòng Caption: gõ vào: Chọn Sản Phẩm
'    - Tìm dòng Width: gõ: 420
'    - Tìm dòng Height: gõ: 380
'
'    Chọn Label (click vào Label):
'    - (Name): lblTitle
'    - Caption: Tìm thấy nhiều sản phẩm phù hợp
'    - Left: 10
'    - Top: 10
'    - Width: 380
'    - Height: 30
'    - Font: Click vào nút [...] bên cạnh Font, chọn Arial, 11, Bold
'    - ForeColor: Click vào [...], chọn màu xanh dương đậm
'
'    Chọn ListBox (click vào ListBox):
'    - (Name): lstProducts
'    - Left: 10
'    - Top: 50
'    - Width: 380
'    - Height: 220
'    - Font: Arial, 10
'
'    Chọn Button thứ nhất (click vào button):
'    - (Name): btnOK
'    - Caption: ✓ Chọn (hoặc OK)
'    - Left: 120
'    - Top: 290
'    - Width: 90
'    - Height: 35
'    - Font: Arial, 10, Bold
'
'    Chọn Button thứ hai (click vào button còn lại):
'    - (Name): btnCancel
'    - Caption: ✗ Hủy (hoặc Cancel)
'    - Left: 220
'    - Top: 290
'    - Width: 90
'    - Height: 35
'    - Font: Arial, 10, Bold
'
' 6. QUAN TRỌNG: Kiểm tra lại (Name) của UserForm PHẢI là: frmSelectProduct
'
' 7. Double-click vào UserForm (vào vùng trống) để mở code editor
' 8. Copy và dán CODE PHẦN A bên dưới vào
'
' ==================================================================================


' ==================================================================================
' PHẦN A: CODE CHO USERFORM (frmSelectProduct)
' ==================================================================================
' Copy code này vào UserForm (double-click UserForm để mở code editor)
' ==================================================================================

Option Explicit

' Biến lưu kết quả
Private selectedValue As Variant
Private selectedText As String
Private isCancelled As Boolean
Private itemsData() As String
Private pricesData() As Variant

' Khởi tạo form khi load
Private Sub UserForm_Initialize()
    ' Thiết lập kích thước và vị trí
    Me.Width = 420
    Me.Height = 380
    
    ' Căn giữa màn hình
    Me.StartUpPosition = 0 ' Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    ' Thiết lập màu nền
    Me.BackColor = RGB(240, 248, 255) ' Alice Blue
    
    ' Thiết lập ListBox
    lstProducts.BackColor = RGB(255, 255, 255)
    lstProducts.Font.Name = "Segoe UI"
    lstProducts.Font.Size = 10
End Sub

' Hàm hiển thị danh sách sản phẩm
Public Sub ShowSelection(items() As String, prices() As Variant, searchTerm As String)
    Dim i As Integer
    
    isCancelled = True
    
    ' Lưu data
    itemsData = items
    pricesData = prices
    
    ' Cập nhật tiêu đề
    lblTitle.Caption = "🔍 Tìm thấy " & UBound(items) & " sản phẩm cho: """ & searchTerm & """"
    
    ' Xóa danh sách cũ
    lstProducts.Clear
    
    ' Thêm sản phẩm vào ListBox
    For i = 1 To UBound(items)
        lstProducts.AddItem (i & ". " & items(i) & " │ " & Format(prices(i), "#,##0") & " VNĐ")
    Next i
    
    ' Chọn item đầu tiên mặc định
    If lstProducts.ListCount > 0 Then
        lstProducts.ListIndex = 0
    End If
    
    ' Focus vào ListBox
    lstProducts.SetFocus
    
    ' Hiển thị form
    Me.Show
End Sub

' Khi nhấn nút OK
Private Sub btnOK_Click()
    If lstProducts.ListIndex >= 0 Then
        isCancelled = False
        
        ' Lấy index thực tế (vì thêm số thứ tự ở đầu)
        Dim idx As Integer
        idx = lstProducts.ListIndex + 1
        
        ' Lấy tên và giá từ data gốc
        selectedText = itemsData(idx)
        selectedValue = pricesData(idx)
        
        Me.Hide
    Else
        MsgBox "⚠️ Vui lòng chọn một sản phẩm!", vbExclamation, "Chưa Chọn"
    End If
End Sub

' Khi nhấn nút Cancel
Private Sub btnCancel_Click()
    isCancelled = True
    Me.Hide
End Sub

' Khi double-click vào ListBox (tương đương nhấn OK)
Private Sub lstProducts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOK_Click
End Sub

' Khi nhấn Enter trong ListBox
Private Sub lstProducts_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then ' Enter key
        btnOK_Click
    ElseIf KeyAscii = 27 Then ' Esc key
        btnCancel_Click
    End If
End Sub

' Properties để lấy kết quả
Public Property Get SelectedPrice() As Variant
    SelectedPrice = selectedValue
End Property

Public Property Get SelectedName() As String
    SelectedName = selectedText
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = isCancelled
End Property


' ==================================================================================
' BƯỚC 2: CODE CHO SHEET NHẬP LIỆU
' ==================================================================================
' 1. Trong VBA Editor, tìm sheet "NhapLieu" (hoặc sheet bạn dùng) ở cửa sổ Project bên trái
' 2. Double-click vào sheet đó
' 3. Copy và dán CODE PHẦN B bên dưới vào
' ==================================================================================


' ==================================================================================
' PHẦN B: CODE CHO SHEET NHẬP LIỆU
' ==================================================================================
' Copy code này vào Sheet "NhapLieu" (hoặc sheet bạn sử dụng)
' ==================================================================================

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim filePath As String
    Dim wbDonGia As Workbook
    Dim wsDonGia As Worksheet
    Dim tenSP As String
    Dim lastRow As Long
    Dim i As Long
    Dim matchCount As Integer
    Dim matchItems() As String
    Dim matchPrices() As Variant
    Dim wb As Workbook
    Dim frm As frmSelectProduct
    
    ' ===== KIỂM TRA ĐIỀU KIỆN =====
    ' Chỉ xử lý khi sửa cột B (Tên Vật Tư) từ dòng 2 trở đi
    If Target.Column <> 2 Or Target.Row < 2 Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    
    ' ===== CẤU HÌNH ĐƯỜNG DẪN =====
    ' ⚠️ QUAN TRỌNG: SỬA ĐƯỜNG DẪN FILE ĐƠN GIÁ Ở ĐÂY
    filePath = "C:\DuLieu\DonGia.xlsx"
    
    ' HOẶC dùng Desktop:
    ' filePath = "C:\Users\" & Environ("USERNAME") & "\Desktop\DonGia.xlsx"
    
    ' HOẶC cùng thư mục với file nhập liệu:
    ' filePath = ThisWorkbook.Path & "\DonGia.xlsx"
    
    ' ===== KIỂM TRA FILE TỒN TẠI =====
    If Dir(filePath) = "" Then
        MsgBox "⚠️ KHÔNG TÌM THẤY FILE ĐƠN GIÁ!" & vbCrLf & vbCrLf & _
               "Đường dẫn: " & filePath & vbCrLf & vbCrLf & _
               "Vui lòng:" & vbCrLf & _
               "1. Kiểm tra file có tồn tại" & vbCrLf & _
               "2. Sửa đường dẫn trong VBA (Alt+F11)", _
               vbExclamation, "Lỗi File"
        Exit Sub
    End If
    
    ' ===== TẮT CẬP NHẬT =====
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' ===== MỞ FILE ĐƠN GIÁ =====
    Set wbDonGia = Nothing
    
    ' Kiểm tra file đã mở chưa
    For Each wb In Workbooks
        If UCase(wb.FullName) = UCase(filePath) Then
            Set wbDonGia = wb
            Exit For
        End If
    Next wb
    
    ' Nếu chưa mở thì mở file
    If wbDonGia Is Nothing Then
        Set wbDonGia = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True)
    End If
    
    ' Tìm sheet đơn giá
    On Error Resume Next
    Set wsDonGia = wbDonGia.Sheets("BangGia")
    If wsDonGia Is Nothing Then
        Set wsDonGia = wbDonGia.Sheets(1)
    End If
    On Error GoTo ErrorHandler
    
    ' ===== TÌM KIẾM SẢN PHẨM =====
    tenSP = Trim(Target.Value)
    
    If tenSP <> "" Then
        lastRow = wsDonGia.Cells(wsDonGia.Rows.Count, "A").End(xlUp).Row
        matchCount = 0
        
        ' Tìm tất cả sản phẩm khớp
        For i = 2 To lastRow
            If InStr(1, UCase(Trim(wsDonGia.Cells(i, "A").Value)), UCase(tenSP), vbTextCompare) > 0 Then
                matchCount = matchCount + 1
                ReDim Preserve matchItems(1 To matchCount)
                ReDim Preserve matchPrices(1 To matchCount)
                matchItems(matchCount) = Trim(wsDonGia.Cells(i, "A").Value)
                matchPrices(matchCount) = wsDonGia.Cells(i, "B").Value
            End If
        Next i
        
        ' ===== XỬ LÝ KẾT QUẢ =====
        If matchCount = 0 Then
            ' ===== KHÔNG TÌM THẤY =====
            Target.Offset(0, 1).Value = "❌ Không tìm thấy"
            
            wbDonGia.Close SaveChanges:=False
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            
            MsgBox "❌ Không tìm thấy sản phẩm: """ & tenSP & """" & vbCrLf & vbCrLf & _
                   "Gợi ý:" & vbCrLf & _
                   "• Kiểm tra chính tả" & vbCrLf & _
                   "• Thử từ khóa ngắn hơn" & vbCrLf & _
                   "• Xem danh sách trong file đơn giá", _
                   vbInformation, "Không Tìm Thấy"
            Exit Sub
            
        ElseIf matchCount = 1 Then
            ' ===== CHỈ 1 KẾT QUẢ - TỰ ĐỘNG ĐIỀN =====
            Target.Value = matchItems(1)
            Target.Offset(0, 1).Value = matchPrices(1)
            
        Else
            ' ===== NHIỀU KẾT QUẢ - DÙNG USERFORM ĐỂ CHỌN =====
            wbDonGia.Close SaveChanges:=False
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            
            ' Tạo và hiển thị UserForm
            Set frm = New frmSelectProduct
            frm.ShowSelection matchItems, matchPrices, tenSP
            
            ' Xử lý kết quả từ UserForm
            If Not frm.Cancelled Then
                Application.EnableEvents = False
                Target.Value = frm.SelectedName
                Target.Offset(0, 1).Value = frm.SelectedPrice
                Application.EnableEvents = True
            Else
                Target.Offset(0, 1).Value = "❌ Đã hủy"
            End If
            
            Unload frm
            Set frm = Nothing
            Exit Sub
        End If
    Else
        ' Xóa giá nếu xóa tên
        Target.Offset(0, 1).Value = ""
    End If
    
    ' ===== ĐÓNG FILE VÀ BẬT LẠI CẬP NHẬT =====
    wbDonGia.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

' ===== XỬ LÝ LỖI =====
ErrorHandler:
    On Error Resume Next
    If Not wbDonGia Is Nothing Then
        wbDonGia.Close SaveChanges:=False
    End If
    On Error GoTo 0
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "⚠️ LỖI: " & Err.Number & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & _
               "File: " & filePath, _
               vbCritical, "Lỗi VBA"
    End If
End Sub


' ==================================================================================
' TÓM TẮT CÀI ĐẶT
' ==================================================================================
'
' ✅ BƯỚC 1: TẠO USERFORM
'    - Insert > UserForm
'    - Đặt tên: frmSelectProduct
'    - Thêm: 1 Label, 1 ListBox, 2 Buttons
'    - Đặt tên controls theo hướng dẫn
'    - Dán PHẦN A vào UserForm
'
' ✅ BƯỚC 2: CODE SHEET
'    - Double-click sheet "NhapLieu"
'    - Dán PHẦN B vào
'    - Sửa đường dẫn file đơn giá
'
' ✅ BƯỚC 3: LƯU FILE
'    - File > Save As
'    - Chọn "Excel Macro-Enabled Workbook (.xlsm)"
'    - Lưu file
'
' ✅ BƯỚC 4: TEST
'    - Đóng VBA Editor (Alt + Q)
'    - Enable Macro khi mở file
'    - Nhập tên vật tư vào cột B
'    - Nếu trùng → UserForm xuất hiện!
'
' ==================================================================================
'
' CẤU TRÚC FILE:
' - File nhập liệu: Cột B = Tên vật tư, Cột C = Giá
' - File đơn giá: Cột A = Tên vật tư, Cột B = Giá (không cần VBA code)
'
' TÍNH NĂNG:
' ✅ Tìm kiếm thông minh (chứa từ khóa)
' ✅ UserForm đẹp khi có nhiều kết quả trùng
' ✅ Double-click hoặc Enter để chọn nhanh
' ✅ Esc để hủy
' ✅ Tự động điền khi chỉ 1 kết quả
'
' ==================================================================================
