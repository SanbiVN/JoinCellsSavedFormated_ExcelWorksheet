# JoinCellsSavedFormated_ExcelWorksheet
 Join Cells Save Formatted - Excel Worksheet
HÀM UDF TỰ ĐỘNG NỐI CHUỖI GIỮ ĐỊNH DẠNG

với Hàm joinCells

=joinCells(toCell,sentenceSpace,Cells,...)

Hướng dẫn sử dụng hàm:

<p align="center"><img title="vbe" src="https://user-images.githubusercontent.com/58664571/157865372-b3872a6c-28a6-40c4-8dbd-277f79d1ed8e.png" width="760"></p>

Cách viết hàm nhanh, gõ vào ô chuỗi =joinCells và ấn tổ hợp phím Ctrl+Shift+A

Ví dụ: gộp các chuỗi từ các ô C1 đến C4, phân cách là dấu cách, trả vào ô B1
### Cách 1: =joinCells(B1, " ",C1,C2,C3,C4)
### Cách 2: =joinCells(B1, " ",C1:C4)
### Cách 3 (gộp tại ô giá trị): =joinCells(C1:C4, " ",C1:C4)
Để tự động Gộp ô ừ B1 đến B5 hãy gõ thêm B1:B5: =joinCells(B1:B5, " ",C1:C4)

Ở đây ô C1 là ô đầu tiên nhập vào nên được chọn làm ô để đặt chiều rộng cột ô đã gộp

<p align="center"><img title="udf_joinCellsFormated" src="https://github.com/SanbiVN/JoinCellsSavedFormated_ExcelWorksheet/assets/58664571/818e6a20-6e4d-42f3-8733-b04a3f9464cd" width="760"></p>

<p align="center"><img title="Join_Fonts_Formating" src="https://user-images.githubusercontent.com/58664571/157867247-2b802a15-b20f-4cce-89ad-efc67d157146.jpg" width="760"></p>
			

