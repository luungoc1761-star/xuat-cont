# LPVN Smart Warehouse 3D & Container Shipment Reconciliation Hub

Hệ thống Web Dashboard 3D động & Đối soát Kế hoạch Xuất hàng Container theo Tháng / Tuần với Tồn kho **Kho 01** và **Kho 60**, tích hợp **Multi-Agent Verification Engine** và **OneDrive Realtime Auto-Sync**.

## 1. Kết Quả Đối Soát Chuẩn Xác Từ 2 File Nguồn Trên Repo
- **File Tồn Kho (`Stock list.xlsx` — Sheet `Stock Balance With Batch`)**:
  - Tổng số dòng lô (Batch) toàn bộ các kho: `24,296` dòng.
  - Lọc nghiêm ngặt **Kho 01** và **Kho 60**:
    - **Kho 01**: `463` lô batch | `140` mã hàng duy nhất | Tổng tồn: **`210,016` pcs**
    - **Kho 60**: `2,055` lô batch | `100` mã hàng duy nhất | Tổng tồn: **`127,651` pcs**
    - **Tổng Gộp (Kho 01 + Kho 60)**: `2,518` lô batch | **`220` mã hàng duy nhất** (`20` mã nằm ở cả 2 kho) | Tổng tồn: **`337,667` pcs**
  - **Tồn Kho Thành Phẩm Theo Mã Trái (LHF) / Phải (RHF)**:
    - Thành phẩm đầu `8` (`8MMMMDVV04`, ký tự thứ 6: `1` = Trái/LHF, `2` = Phải/RHF): **`101` mã** thuộc **`51` cặp Model/Feature**, tổng tồn **`120,754` pcs** (`60,379` Trái / `60,375` Phải).
    - **`46` cặp Model cân bằng 100%** (bao gồm Model `#8333` lệch trong từng kho đơn lẻ `01: 1,179L/1,400R` và `60: 1,509L/1,288R` nhưng cân bằng tuyệt đối `2,688L = 2,688R` khi gộp Kho 01 + Kho 60).
    - **`5` cặp Model lệch Trái/Phải**: `#3565` (`0L / 180R`), `#3611` (`1,344L / 1,216R`), `#5025` (`864L / 810R`), `#5171` (`2,480L / 2,560R`), `#8695` (`3,068L / 3,016R`).
- **File Kế Hoạch Xuất Hàng (`NOV. SHIPMENT -OCT22.xlsx`)**:
  - **Kế hoạch tháng (`NOV-full`)**: `236` dòng đơn hàng thuộc `27` PO (`26` cụm Container):
    - *Phân bổ lũy kế FIFO (Kho 60 -> Kho 01)*: **`152` dòng đủ hàng xuất** (`12` PO đủ 100%) | **`84` dòng thiếu hàng** (`15` PO thiếu hàng, gồm `24` dòng thiếu một phần và `60` dòng hết tồn kho = 0).
    - *Đối chiếu trực tiếp từng dòng với tổng tồn (Kho 01 + 60)*: **`206` dòng đủ hàng xuất** (`19` PO đủ 100%) | **`30` dòng thiếu hàng** (`8` PO thiếu hàng).
  - **Kế hoạch tuần (`OCT-WK 42`)**: `52` dòng đơn hàng thuộc `4` PO (`PKL# 245` đến `248`):
    - **`46` dòng đủ hàng xuất** (`2` PO đủ 100%: `436862`, `436864`) | **`6` dòng thiếu hàng** (`2` PO thiếu: `0201-0003002760`, `436866`).

## 2. Khởi Chạy Ứng Dụng & Đồng Bộ OneDrive Realtime

```bash
# Cài đặt thư viện
npm install

# Chạy ứng dụng Web (tự động theo dõi file Stock list.xlsx & NOV. SHIPMENT -OCT22.xlsx)
npm run dev

# Hoặc chạy tiến trình đồng bộ OneDrive -> Repo -> Git main tự động
ONE_DRIVE_DIR="/path/to/OneDrive/stock list" AUTO_GIT_PUSH=true npm run watch-onedrive
```
