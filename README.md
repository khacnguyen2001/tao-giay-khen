# Website tạo giấy khen hàng loạt

Tạo giấy khen PNG cho toàn bộ học sinh từ file Excel, đặt tên file theo **mã học sinh** (`HS001.png`, `HS002.png`…) và đóng gói thành `giay_khen_hoc_sinh.zip`.

- Chạy hoàn toàn trong trình duyệt (Chrome hoặc Safari trên macOS), **không cần cài đặt, không cần Internet, không có backend**.
- Dữ liệu học sinh và ảnh **không được gửi đi đâu cả**: mọi xử lý diễn ra trên máy của bạn.

## Cấu trúc thư mục

```
website_tao_giay_khen/
├── index.html                  Giao diện chính
├── style.css                   Thiết kế giao diện
├── script.js                   Đọc Excel, xem trước, tạo ảnh, đóng gói ZIP
├── chay_website.command        (Tuỳ chọn) Mở website qua máy chủ cục bộ
├── fonts/
│   ├── PlayfairDisplay-BoldItalic.ttf          Font gốc (Google Fonts 1.203, có đủ dấu tiếng Việt)
│   ├── PlayfairDisplay-BoldItalic-NoLiga.ttf   Cùng font đó, đã tắt chữ ghép (Th, fi, fl…) như Canva
│   ├── playfair-bold-italic.js                 Bản NoLiga, nhúng sẵn dạng base64 (website dùng file này)
│   └── OFL.txt                          Giấy phép font (SIL Open Font License)
├── libs/
│   ├── xlsx.full.min.js        SheetJS 0.18.5 — đọc file Excel
│   └── jszip.min.js            JSZip 3.10.1 — đóng gói ZIP
└── mau_thu/                    File mẫu để chạy thử
    ├── danh_sach_hoc_sinh_mau.xlsx
    ├── khung_giay_khen_goc.png
    └── giay_khen_mau_co_ten.png
```

Mọi thư viện và font đều nằm sẵn trong thư mục, bạn không cần tải thêm gì.

## Cách chạy

**Cách 1 — đơn giản nhất:** nhấp đúp vào `index.html` (hoặc kéo file vào cửa sổ Chrome/Safari).
Website được thiết kế để chạy trực tiếp như vậy, kể cả với `file://`.

**Cách 2 — qua máy chủ cục bộ** (chỉ dùng nếu trình duyệt báo lỗi ở cách 1):

- Nhấp đúp vào `chay_website.command`. Lần đầu, nếu macOS chặn, hãy nhấp chuột phải → **Mở** → **Mở**.
- Hoặc mở Terminal tại thư mục dự án và chạy:

  ```bash
  python3 -m http.server 8000
  ```

  rồi mở `http://localhost:8000` trong trình duyệt. Nhấn `Ctrl + C` trong Terminal để dừng.

## Quy trình sử dụng

1. **Khu vực A** — tải 3 file (bấm vào ô hoặc kéo thả):
   - File Excel danh sách học sinh (`.xlsx`, `.xls`);
   - Ảnh khung giấy khen gốc, **chưa có tên** (`.png`, `.jpg`, `.webp`);
   - Ảnh mẫu **đã có tên**. Ảnh này chỉ để tham chiếu vị trí, không bao giờ được dùng làm nền ảnh xuất.
2. Kiểm tra **trang tính**, **cột mã** và **cột tên** (website tự đoán theo tiêu đề cột, bạn có thể đổi lại), rồi bấm **Đọc danh sách học sinh**.
3. **Khu vực B, C** — tên được đặt ở vị trí mặc định X = 1019, Y = 688. Website cũng so sánh ảnh mẫu với ảnh khung và hiện vị trí gợi ý, nhưng chỉ áp dụng khi bạn bấm **Dùng vị trí dò được**. Sau đó:
   - Chọn chế độ xem **“Chồng ảnh mẫu để đối chiếu”** để kiểm tra tên mới có trùng vị trí tên trong ảnh mẫu không;
   - Kéo tên trên ảnh để chỉnh vị trí, nhấp đúp để đặt tên tại một điểm, hoặc dùng phím mũi tên (1 px; giữ Shift: 10 px);
   - Kiểm tra font **Playfair Display Bold Italic**, cỡ chữ **114 px** và màu **#5e1a36**.
4. **Khu vực D** — bấm **Tạo thử 3 giấy khen**. Ảnh thử gồm học sinh đầu tiên, học sinh có tên dài nhất và học sinh cuối cùng. Mở từng ảnh để kiểm tra, rồi đánh dấu **“Tôi đã kiểm tra ảnh thử”**.
5. Bấm **Tạo toàn bộ giấy khen**, chờ thanh tiến trình chạy xong, rồi bấm **Tải ZIP**.

## Các quy tắc đã được áp dụng

| Nội dung | Cách xử lý |
|---|---|
| Font | Playfair Display, Bold (700), Italic. Font được tải và xác minh trước khi cho phép xuất ảnh. Chữ ghép đã được tắt (ví dụ “Th” trong “Thị” không bị nối thành một ký tự) để chữ giống hệt Canva. |
| Cỡ chữ | Mặc định 114 **pixel** trên ảnh gốc (Canvas tính theo px). Chỉ thay đổi khi bạn tự chỉnh. |
| Vị trí mặc định | X = 1019, Y = 688 (tính cho khung 2000 × 1414 px). Nếu ảnh khung nhỏ hơn thì dùng tâm ảnh. Kết quả dò từ ảnh mẫu chỉ là gợi ý: bấm **Dùng vị trí dò được** nếu muốn áp dụng. |
| Màu chữ | `#5e1a36`. Màu đen `#000000` bị chặn. Nếu bạn chọn màu khác, website hiện cảnh báo. |
| Căn chỉnh | X, Y là điểm neo. “Giữa chữ” đặt Y ở giữa chiều cao chữ hoa, nên mọi tên có cùng đường chân chữ. |
| Tên quá dài | Website cảnh báo trong bảng và trên khung xem trước, rồi hỏi xác nhận trước khi tạo. Chỉ tự thu nhỏ khi bạn bật tuỳ chọn, và ghi rõ từng trường hợp. Tên tràn ra ngoài ảnh thì không cho tạo. |
| Mã học sinh | Đọc dưới dạng chuỗi, giữ số 0 ở đầu (`00012`, hoặc ô số định dạng `0000`). |
| Thứ tự | Giữ đúng thứ tự các dòng trong Excel. |
| Dòng lỗi | Thiếu mã, thiếu tên, mã trùng (không phân biệt hoa/thường, vì macOS coi `HS001` và `hs001` là cùng một file), mã có ký tự cấm trong tên file. Các dòng này được liệt kê và bỏ qua. |
| Ảnh xuất | PNG, đúng kích thước pixel của ảnh khung gốc. Mỗi ảnh được vẽ lại từ đầu trên nền ảnh khung, nên không bị chồng tên từ ảnh trước. |
| ZIP | `giay_khen_hoc_sinh.zip`. Các PNG được lưu nguyên (không nén lại) và website kiểm tra không có tên file trùng. Nút tải chỉ mở khi đã tạo xong. |
| An toàn kết quả | Khi bạn đổi cấu hình, ảnh khung hoặc danh sách, kết quả cũ bị huỷ và bạn phải tạo thử lại. Bấm Huỷ thì loại bỏ toàn bộ ảnh đang tạo dở. |

## Định dạng file Excel

Tối thiểu 2 cột, tên cột tuỳ ý. Ví dụ:

| Mã HS | Họ và tên |
|---|---|
| HS001 | Nguyễn Văn An |
| HS002 | Trần Thị Bích Ngọc |

Mẹo: để giữ số 0 ở đầu mã, hãy định dạng cột mã là **Text** trong Excel, hoặc dùng định dạng số kiểu `00000`.

## Kiểm thử đã thực hiện

Đã kiểm thử tự động trên Google Chrome (mở trực tiếp bằng `file://`) với các file trong `mau_thu/` và một danh sách 300 học sinh:

1. Đọc được Excel, tự nhận đúng cột “Mã HS” và “Họ và tên”, giữ số 0 ở đầu (`0012`, `00001`).
2. Tên có dấu hiển thị đúng. Font có đủ toàn bộ ký tự tiếng Việt.
3. Phát hiện đủ 7 dòng lỗi: 2 cặp mã trùng (trong đó có `HS001`/`hs001`), thiếu tên, thiếu mã, dòng trống.
4. Vị trí dò từ ảnh mẫu lệch dưới 1 px so với vị trí thật (tâm X 859/860, chân chữ Y 560/560).
5. Ảnh xuất có kích thước 1600 × 1131 px, đúng bằng ảnh gốc. Ngoài vùng tên, nền không thay đổi pixel nào. Màu chữ đúng `#5e1a36`.
6. Tên file đúng theo mã. ZIP đủ số ảnh, không trùng tên.
7. Kéo chuột, huỷ giữa chừng và tải một ảnh riêng lẻ đều hoạt động. 300 giấy khen được tạo trong khoảng 13 giây mà trang không bị treo.

## Thay font

Website dùng `fonts/playfair-bold-italic.js`. Đây là font nhúng dạng base64, nhờ vậy website chạy được cả khi mở trực tiếp file.
Font này được tạo từ file TTF gốc, sau khi **tắt chữ ghép** (`liga`: Th, fi, fl…) để giống cách Canva hiển thị.

Nếu thay `PlayfairDisplay-BoldItalic.ttf` bằng phiên bản khác, hãy tạo lại 2 file kia bằng Terminal (cần Python 3):

```bash
cd fonts
python3 -m pip install fonttools
python3 -c "
from fontTools.ttLib import TTFont; from fontTools import subset
f=TTFont('PlayfairDisplay-BoldItalic.ttf')
tags={r.FeatureTag for t in ('GSUB','GPOS') for r in f[t].table.FeatureList.FeatureRecord}
o=subset.Options(); o.layout_features=sorted(tags-{'liga'}); o.name_IDs=['*']; o.name_languages=['*']
o.glyph_names=True; o.notdef_outline=True; o.legacy_kern=True; o.hinting=True; o.layout_scripts=['*']
s=subset.Subsetter(o); s.populate(glyphs=f.getGlyphOrder(), unicodes=f.getBestCmap().keys()); s.subset(f)
f.save('PlayfairDisplay-BoldItalic-NoLiga.ttf')"
{ printf 'window.PLAYFAIR_BOLD_ITALIC_BASE64 = "'; base64 -i PlayfairDisplay-BoldItalic-NoLiga.ttf | tr -d '\n'; printf '";\n'; } > playfair-bold-italic.js
```

Nếu muốn **giữ** chữ ghép, ở dòng cuối hãy dùng `PlayfairDisplay-BoldItalic.ttf` thay cho file `-NoLiga.ttf`.
Nếu xoá file JS, website sẽ tải trực tiếp file `-NoLiga.ttf`. Khi đó nên chạy theo **Cách 2** (máy chủ cục bộ).

## Xử lý sự cố

- **Danh sách đọc ra không đúng file mong muốn**: website chỉ đọc file bạn chọn trên trang và không bao giờ tự đọc thư mục `mau_thu/`. Mỗi lần bấm “Đọc danh sách học sinh”, website đọc lại file đó và ghi rõ nguồn (tên file, dung lượng, giờ sửa) ngay trên bảng danh sách. Nếu bạn sửa file Excel sau khi chọn, trình duyệt sẽ báo cần chọn lại: hãy bấm **Đổi file** và chọn lại file đó.

- **“Không tải được font”**: kiểm tra thư mục `fonts/` có đủ file không, hoặc chạy theo Cách 2.
- **Vị trí dò không chính xác**: ảnh mẫu cần cùng khung với ảnh gốc, chỉ khác phần tên. Nếu khác, hãy kéo tên bằng chuột rồi đối chiếu ở chế độ “Chồng ảnh mẫu”.
- **Ảnh rất lớn (trên 4096 × 4096 px)**: Safari có thể giới hạn kích thước Canvas, nên dùng Chrome.
- **Danh sách rất dài (hàng nghìn học sinh, ảnh lớn)**: ZIP được tạo trong bộ nhớ trình duyệt. Nếu máy yếu, hãy chia danh sách thành nhiều file Excel.

## Đổi giá trị mặc định

Các giá trị mặc định nằm ở đầu file `script.js`:

```js
const DEFAULT_SIZE = 114;  // cỡ chữ (px)
const DEFAULT_X = 1019;    // vị trí X (px)
const DEFAULT_Y = 688;     // vị trí Y (px)
const DEFAULT_COLOR = '#5e1a36';
```
