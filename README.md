# JoinCellsSavedFormated_ExcelWorksheet
 Join Cells Saved Formated - Excel Worksheet
HÀM UDF TỰ ĐỘNG NỐI CHUỖI GIỮ ĐỊNH DẠNG

với Hàm S_JoinF

=S_joinF(toCell,sentenceSpace,Cells,...)

Hướng dẫn sử dụng hàm:

![image](https://user-images.githubusercontent.com/58664571/157865372-b3872a6c-28a6-40c4-8dbd-277f79d1ed8e.png)


Cách viết hàm nhanh, gõ vào ô chuỗi =S_JoinF và ấn tổ hợp phím Ctrl+Shift+A

Ví dụ: gộp các chuỗi từ các ô C1 đến C4, phân cách là dấu cách, trả vào ô B1
### Cách 1: =S_JoinF(B1, " ",C1,C2,C3,C4)
### Cách 2: =S_JoinF(B1, " ",C1:C4)
### Cách 3 (gộp tại ô giá trị): =S_JoinF(C1:C4, " ",C1:C4)
Để tự động Gộp ô ừ B1 đến B5 hãy gõ thêm B1:B5: =S_JoinF(B1:B5, " ",C1:C4)

Ở đây ô C1 là ô đầu tiên nhập vào nên được chọn làm ô để đặt chiều rộng cột ô đã gộp

				

![Join_Fonts_Formating](https://user-images.githubusercontent.com/58664571/157867247-2b802a15-b20f-4cce-89ad-efc67d157146.jpg)
