import { Fragment, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

const STORAGE_KEY = "nhatKyThiCong";
const DAILY_LOG_STORAGE_KEY = "dailyLogs";
const CONTRACT_STORAGE_KEY = "hopDongThiCong";
const BOQ_STORAGE_KEY = "boqItems";
const WBS_STORAGE_KEY = "wbsPlans";
const MATERIAL_IN_STORAGE_KEY = "materialIns";
const LABOR_TEAM_STORAGE_KEY = "laborTeams";
const SUBCONTRACTOR_STORAGE_KEY = "subcontractors";
const SUBCONTRACT_STORAGE_KEY = "subContracts";
const EQUIPMENT_LIST_STORAGE_KEY = "equipmentList";
const EQUIPMENT_TRANSFER_STORAGE_KEY = "equipmentTransfers";
const ACCEPTANCE_BATCH_STORAGE_KEY = "acceptanceBatches";
const ACCEPTANCE_STORAGE_KEY = "acceptances";
const INVOICE_STORAGE_KEY = "invoices";
const COLLECTION_STORAGE_KEY = "collections";
const OTHER_COST_STORAGE_KEY = "otherCosts";
const CASH_FLOW_PLAN_STORAGE_KEY = "cashFlowPlans";
const DEDUCTION_PLAN_STORAGE_KEY = "deductionPlans";
const LABOR_PAYMENT_PLAN_STORAGE_KEY = "laborPaymentPlans";
const PROJECT_DOCUMENT_STORAGE_KEY = "projectDocuments";
const DOCUMENT_TEMPLATE_STORAGE_KEY = "documentTemplates";
const SYSTEM_CATALOG_STORAGE_KEY = "systemCatalogs";
const SYSTEM_SETTINGS_STORAGE_KEY = "systemSettings";
const USERS_STORAGE_KEY = "users";
const CURRENT_USER_STORAGE_KEY = "currentUser";
const ALL_PROJECTS = "Tất cả";
const ALL_CONTRACTS = "all";
const WBS_STATUSES = ["Chưa bắt đầu", "Đang thi công", "Hoàn thành", "Chậm tiến độ"];
const DOCUMENT_TABS = ["Pháp lý", "Thiết kế", "BOQ", "Vật tư", "Thi công", "Nghiệm thu", "Thanh toán", "Quyết toán"];
const DOCUMENT_TEMPLATE_GROUPS = [
  "Biểu mẫu dự toán",
  "Biểu mẫu tiến độ WBS",
  "Biểu mẫu nhật ký thi công",
  "Biểu mẫu nghiệm thu",
  "Biểu mẫu vật tư",
  "Biểu mẫu nhân công",
  "Biểu mẫu thanh toán",
  "Quy trình / Hướng dẫn thi công",
  "TDS / MSDS sản phẩm",
  "Hồ sơ mẫu",
  "Khác",
];
const DOCUMENT_TEMPLATE_FILE_TYPES = ["Excel", "Word", "PDF", "Image", "Khác"];
const DOCUMENT_TEMPLATE_STATUSES = ["Đang hiệu lực", "Hết hiệu lực", "Dự thảo"];
const DAILY_IMAGE_TYPES = [
  "Ảnh trước thi công",
  "Ảnh trong quá trình thi công",
  "Ảnh sau thi công / hoàn thiện",
  "Ảnh phát sinh / lỗi hiện trường",
];

const projects = [
  { ma: "CT001", ten: "Biệt thự Thảo Điền" },
  { ma: "CT002", ten: "Nhà phố Bình Dương" },
  { ma: "CT003", ten: "Villa Đồng Nai" },
];

const sampleWbsPlans = [
  {
    maWbs: "WBS-01.01",
    hopDongCongTrinh: projects[0].ten,
    hangMucCongViec: "Sơn lót tường trong",
    khoiLuongKeHoach: 100,
  },
  {
    maWbs: "WBS-01.02",
    hopDongCongTrinh: projects[1].ten,
    hangMucCongViec: "Sơn phủ hoàn thiện",
    khoiLuongKeHoach: 120,
  },
];

const constructionRoleDefinitions = [
  { id: "ROLE-001", maVaiTro: "ADMIN", tenVaiTro: "Admin hệ thống", nhomVaiTro: "Quản trị", moTa: "Toàn quyền tất cả module và cấu hình hệ thống.", trangThai: "Hoạt động" },
  { id: "ROLE-002", maVaiTro: "TGD", tenVaiTro: "TGĐ", nhomVaiTro: "Điều hành", moTa: "Xem toàn bộ dashboard, báo cáo, hợp đồng, công nợ và lãi/lỗ.", trangThai: "Hoạt động" },
  { id: "ROLE-003", maVaiTro: "GDTC", tenVaiTro: "Giám đốc Thi công", nhomVaiTro: "Điều hành thi công", moTa: "Xem toàn bộ dự án thi công, duyệt chi phí khác, duyệt nghiệm thu nội bộ và xem báo cáo tổng hợp.", trangThai: "Hoạt động" },
  { id: "ROLE-004", maVaiTro: "PM", tenVaiTro: "Giám đốc Dự án / PM", nhomVaiTro: "Quản lý dự án", moTa: "Quản lý công trình được giao, thêm/sửa WBS, kiểm tra nhật ký, vật tư, nhân công, thiết bị, nghiệm thu và báo cáo.", trangThai: "Hoạt động" },
  { id: "ROLE-005", maVaiTro: "CHT", tenVaiTro: "Chỉ huy trưởng", nhomVaiTro: "Hiện trường", moTa: "Nhập nhật ký thi công, cập nhật khối lượng, vật tư, nhân công, thiết bị và xem WBS/BOQ liên quan.", trangThai: "Hoạt động" },
  { id: "ROLE-006", maVaiTro: "QS", tenVaiTro: "QS", nhomVaiTro: "Khối lượng", moTa: "Theo dõi khối lượng, lập nghiệm thu và đối chiếu khối lượng thực hiện - nghiệm thu.", trangThai: "Hoạt động" },
  { id: "ROLE-007", maVaiTro: "QC", tenVaiTro: "QC", nhomVaiTro: "Chất lượng", moTa: "Xem nhật ký, cập nhật kiểm tra chất lượng, tham gia nghiệm thu và upload chứng cứ chất lượng.", trangThai: "Hoạt động" },
  { id: "ROLE-008", maVaiTro: "TKDA", tenVaiTro: "Thư ký dự án", nhomVaiTro: "Hồ sơ", moTa: "Quản lý hồ sơ dự án, biên bản nghiệm thu, hồ sơ thanh toán, hồ sơ hoàn công, nhắc việc hồ sơ và hỗ trợ PM trong công tác chứng từ dự án.", trangThai: "Hoạt động" },
  { id: "ROLE-009", maVaiTro: "KTCT", tenVaiTro: "Kế toán công trình", nhomVaiTro: "Tài chính công trình", moTa: "Quản lý hóa đơn, công nợ, thu tiền, theo dõi nghiệm thu và báo cáo tài chính công trình.", trangThai: "Hoạt động" },
  { id: "ROLE-010", maVaiTro: "KHOCT", tenVaiTro: "Kho công trình", nhomVaiTro: "Kho công trình", moTa: "Nhập vật tư về công trình, theo dõi xuất dùng vật tư từ nhật ký và tồn kho.", trangThai: "Hoạt động" },
];

const constructionModulePermissions = [
  "Dashboard",
  "Hợp đồng",
  "BOQ - Dự toán",
  "Kế hoạch thi công WBS",
  "Nhật ký thi công",
  "Vật tư",
  "Nhân công",
  "Nhà thầu phụ / Tổ đội",
  "Máy móc thiết bị",
  "Nghiệm thu",
  "Hóa đơn - Công nợ",
  "Chi phí khác",
  "Báo cáo",
  "Hồ sơ dự án",
  "Danh mục hệ thống",
  "Cài đặt hệ thống",
];

const rolePermissionDefaults = {
  ADMIN: { all: ["xem", "them", "sua", "xoa", "duyet"] },
  TGD: {
    "Dashboard": ["xem"],
    "Hợp đồng": ["xem"],
    "Hóa đơn - Công nợ": ["xem"],
    "Báo cáo": ["xem"],
  },
  GDTC: {
    "Dashboard": ["xem"],
    "Kế hoạch thi công WBS": ["xem"],
    "Nhật ký thi công": ["xem"],
    "Vật tư": ["xem"],
    "Nhân công": ["xem"],
    "Nhà thầu phụ / Tổ đội": ["xem"],
    "Máy móc thiết bị": ["xem"],
    "Nghiệm thu": ["xem", "duyet"],
    "Chi phí khác": ["xem", "duyet"],
    "Báo cáo": ["xem"],
  },
  PM: {
    "Dashboard": ["xem"],
    "Hợp đồng": ["xem"],
    "BOQ - Dự toán": ["xem"],
    "Kế hoạch thi công WBS": ["xem", "them", "sua"],
    "Nhật ký thi công": ["xem", "sua"],
    "Vật tư": ["xem"],
    "Nhân công": ["xem"],
    "Nhà thầu phụ / Tổ đội": ["xem"],
    "Máy móc thiết bị": ["xem"],
    "Nghiệm thu": ["xem", "them", "sua"],
    "Chi phí khác": ["xem", "them", "sua"],
    "Báo cáo": ["xem"],
    "Hồ sơ dự án": ["xem", "them", "sua"],
  },
  CHT: {
    "BOQ - Dự toán": ["xem"],
    "Kế hoạch thi công WBS": ["xem"],
    "Nhật ký thi công": ["xem", "them", "sua"],
    "Vật tư": ["xem", "them", "sua"],
    "Nhân công": ["xem", "them", "sua"],
    "Máy móc thiết bị": ["xem", "them", "sua"],
  },
  QS: {
    "BOQ - Dự toán": ["xem", "sua"],
    "Kế hoạch thi công WBS": ["xem"],
    "Nhật ký thi công": ["xem"],
    "Nghiệm thu": ["xem", "them", "sua"],
    "Báo cáo": ["xem"],
  },
  QC: {
    "Nhật ký thi công": ["xem", "them", "sua"],
    "Nghiệm thu": ["xem", "them", "sua"],
    "Hồ sơ dự án": ["xem", "them", "sua"],
  },
  TKDA: {
    "Hợp đồng": ["xem"],
    "BOQ - Dự toán": ["xem"],
    "Kế hoạch thi công WBS": ["xem"],
    "Nhật ký thi công": ["xem"],
    "Nghiệm thu": ["xem", "them", "sua"],
    "Hóa đơn - Công nợ": ["xem"],
    "Báo cáo": ["xem"],
    "Hồ sơ dự án": ["xem", "them", "sua"],
  },
  KTCT: {
    "Nghiệm thu": ["xem"],
    "Hóa đơn - Công nợ": ["xem", "them", "sua"],
    "Báo cáo": ["xem"],
  },
  KHOCT: {
    "Vật tư": ["xem", "them", "sua"],
    "Nhật ký thi công": ["xem"],
    "Báo cáo": ["xem"],
  },
};

const dataCatalogTabs = [
  {
    key: "customers",
    label: "Khách hàng",
    storageKey: "customers",
    columns: [
      { key: "maKhachHang", label: "Mã KH", required: true },
      { key: "tenKhachHang", label: "Tên khách hàng", required: true },
      { key: "nguoiDaiDien", label: "Người đại diện" },
      { key: "soDienThoai", label: "Số điện thoại" },
      { key: "email", label: "Email", required: true },
      { key: "matKhau", label: "Mật khẩu", type: "password", required: true },
      { key: "diaChi", label: "Địa chỉ", wide: true },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Đang hợp tác", "Tiềm năng", "Tạm dừng"] },
    ],
    samples: [
      { id: "CAT-CUS-001", maKhachHang: "KH-001", tenKhachHang: "Công ty TNHH Minh Phú", nguoiDaiDien: "Nguyễn Minh Phú", soDienThoai: "0909000001", email: "minhphu@example.com", diaChi: "Thảo Điền, TP.HCM", trangThai: "Đang hợp tác" },
      { id: "CAT-CUS-002", maKhachHang: "KH-002", tenKhachHang: "Gia đình Anh Khoa", nguoiDaiDien: "Trần Anh Khoa", soDienThoai: "0909000002", email: "anhkhoa@example.com", diaChi: "Dĩ An, Bình Dương", trangThai: "Tiềm năng" },
    ],
  },
  {
    key: "projects",
    label: "Công trình",
    storageKey: "projectsCatalog",
    columns: [
      { key: "maCongTrinh", label: "Mã CT", required: true },
      { key: "tenCongTrinh", label: "Tên công trình", required: true },
      { key: "khachHang", label: "Khách hàng" },
      { key: "diaChi", label: "Địa chỉ", wide: true },
      { key: "quanLyDuAn", label: "Quản lý dự án" },
      { key: "ngayKhoiCong", label: "Ngày khởi công", type: "date" },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Đang chuẩn bị", "Đang thi công", "Tạm dừng", "Hoàn thành"] },
    ],
    samples: [
      { id: "CAT-PRO-001", maCongTrinh: "CT001", tenCongTrinh: "Biệt thự Thảo Điền", khachHang: "Công ty TNHH Minh Phú", diaChi: "Quận 2, TP.HCM", quanLyDuAn: "Lê Hoàng Nam", ngayKhoiCong: "2026-06-01", trangThai: "Đang thi công" },
      { id: "CAT-PRO-002", maCongTrinh: "CT002", tenCongTrinh: "Nhà phố Bình Dương", khachHang: "Gia đình Anh Khoa", diaChi: "Dĩ An, Bình Dương", quanLyDuAn: "Phạm Quốc Huy", ngayKhoiCong: "2026-06-10", trangThai: "Đang chuẩn bị" },
    ],
  },
  {
    key: "wbsStandards",
    label: "WBS chuẩn",
    storageKey: "wbsStandards",
    columns: [
      { key: "maWbs", label: "Mã WBS", required: true },
      { key: "hangMucCongViec", label: "Hạng mục", required: true },
      { key: "nhomCongViec", label: "Nhóm công việc" },
      { key: "dvt", label: "ĐVT" },
      { key: "dinhMucVatTu", label: "ĐM vật tư", type: "number" },
      { key: "dinhMucNhanCong", label: "ĐM nhân công", type: "number" },
      { key: "dinhMucThietBi", label: "ĐM thiết bị", type: "number" },
    ],
    samples: [
      { id: "CAT-WBS-001", maWbs: "WBS-01.01", hangMucCongViec: "Sơn lót tường trong", nhomCongViec: "Sơn nội thất", dvt: "m2", dinhMucVatTu: "0.12", dinhMucNhanCong: "0.03", dinhMucThietBi: "0.01" },
      { id: "CAT-WBS-002", maWbs: "WBS-01.02", hangMucCongViec: "Sơn phủ hoàn thiện", nhomCongViec: "Sơn nội thất", dvt: "m2", dinhMucVatTu: "0.18", dinhMucNhanCong: "0.04", dinhMucThietBi: "0.01" },
    ],
  },
  {
    key: "materials",
    label: "Vật tư",
    storageKey: "materialsCatalog",
    columns: [
      { key: "maVatTu", label: "Mã vật tư", required: true },
      { key: "tenVatTu", label: "Tên vật tư", required: true },
      { key: "nhomVatTu", label: "Nhóm" },
      { key: "dvt", label: "ĐVT" },
      { key: "donGiaDuKien", label: "Đơn giá DK", type: "number" },
      { key: "nhaCungCapMacDinh", label: "NCC mặc định" },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Đang dùng", "Ngưng dùng"] },
    ],
    samples: [
      { id: "CAT-MAT-001", maVatTu: "VT-001", tenVatTu: "Sơn lót kháng kiềm", nhomVatTu: "Sơn lót", dvt: "kg", donGiaDuKien: "85000", nhaCungCapMacDinh: "NCC Sơn Hòa Bình", trangThai: "Đang dùng" },
      { id: "CAT-MAT-002", maVatTu: "VT-002", tenVatTu: "Sơn phủ hoàn thiện", nhomVatTu: "Sơn phủ", dvt: "kg", donGiaDuKien: "125000", nhaCungCapMacDinh: "NCC Sơn Hòa Bình", trangThai: "Đang dùng" },
    ],
  },
  {
    key: "personnel",
    label: "Nhân sự",
    storageKey: "personnelCatalog",
    columns: [
      { key: "maNhanSu", label: "Mã NS", required: true },
      { key: "hoTen", label: "Họ tên", required: true },
      { key: "chucVu", label: "Chức vụ" },
      { key: "phongBan", label: "Phòng ban" },
      { key: "soDienThoai", label: "Số điện thoại" },
      { key: "email", label: "Email" },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Đang làm", "Nghỉ phép", "Nghỉ việc"] },
    ],
    samples: [
      { id: "CAT-PER-001", maNhanSu: "NS-001", hoTen: "Lê Hoàng Nam", chucVu: "Quản lý dự án", phongBan: "Thi công", soDienThoai: "0909111222", email: "nam@example.com", trangThai: "Đang làm" },
      { id: "CAT-PER-002", maNhanSu: "NS-002", hoTen: "Nguyễn Bảo Anh", chucVu: "Giám sát", phongBan: "Thi công", soDienThoai: "0909333444", email: "baoanh@example.com", trangThai: "Đang làm" },
    ],
  },
  {
    key: "teams",
    label: "Tổ đội",
    storageKey: "teamsCatalog",
    columns: [
      { key: "maToDoi", label: "Mã tổ", required: true },
      { key: "tenToDoi", label: "Tên tổ đội", required: true },
      { key: "nguoiDaiDien", label: "Người đại diện" },
      { key: "soDienThoai", label: "Số điện thoại" },
      { key: "loaiToDoi", label: "Loại tổ đội" },
      { key: "donGiaCongNgay", label: "Đơn giá công/ngày", type: "number" },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Sẵn sàng", "Đang thi công", "Tạm ngưng"] },
    ],
    samples: [
      { id: "CAT-TEAM-001", maToDoi: "TD-001", tenToDoi: "Tổ sơn đá A", nguoiDaiDien: "Anh Tuấn", soDienThoai: "0909555666", loaiToDoi: "Sơn đá", donGiaCongNgay: "550000", trangThai: "Sẵn sàng" },
      { id: "CAT-TEAM-002", maToDoi: "TD-002", tenToDoi: "Tổ bột trét B", nguoiDaiDien: "Chị Hương", soDienThoai: "0909777888", loaiToDoi: "Bột trét", donGiaCongNgay: "500000", trangThai: "Đang thi công" },
    ],
  },
  {
    key: "equipment",
    label: "Máy móc thiết bị",
    storageKey: "equipmentCatalog",
    columns: [
      { key: "maThietBi", label: "Mã TB", required: true },
      { key: "tenThietBi", label: "Tên thiết bị", required: true },
      { key: "loaiThietBi", label: "Loại thiết bị" },
      { key: "nguonSoHuu", label: "Nguồn sở hữu", type: "select", options: ["Công ty", "Thuê ngoài", "Nhà cung cấp"] },
      { key: "donGiaThueNgay", label: "Đơn giá/ngày", type: "number" },
      { key: "donGiaThueGio", label: "Đơn giá/giờ", type: "number" },
      { key: "tinhTrangHienTai", label: "Tình trạng", type: "select", options: ["Sẵn sàng", "Đang sử dụng", "Bảo trì"] },
    ],
    samples: [
      { id: "CAT-EQP-001", maThietBi: "TB-001", tenThietBi: "Máy phun sơn Graco", loaiThietBi: "Máy phun", nguonSoHuu: "Công ty", donGiaThueNgay: "0", donGiaThueGio: "0", tinhTrangHienTai: "Sẵn sàng" },
      { id: "CAT-EQP-002", maThietBi: "TB-002", tenThietBi: "Giàn giáo di động", loaiThietBi: "Giàn giáo", nguonSoHuu: "Thuê ngoài", donGiaThueNgay: "300000", donGiaThueGio: "0", tinhTrangHienTai: "Đang sử dụng" },
    ],
  },
  {
    key: "suppliers",
    label: "Nhà cung cấp",
    storageKey: "suppliersCatalog",
    columns: [
      { key: "maNhaCungCap", label: "Mã NCC", required: true },
      { key: "tenNhaCungCap", label: "Tên NCC", required: true },
      { key: "nhomCungCap", label: "Nhóm cung cấp" },
      { key: "nguoiLienHe", label: "Người liên hệ" },
      { key: "soDienThoai", label: "Số điện thoại" },
      { key: "email", label: "Email" },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Đang hợp tác", "Dự phòng", "Ngưng hợp tác"] },
    ],
    samples: [
      { id: "CAT-SUP-001", maNhaCungCap: "NCC-001", tenNhaCungCap: "NCC Sơn Hòa Bình", nhomCungCap: "Sơn", nguoiLienHe: "Chị Mai", soDienThoai: "0911000001", email: "mai@example.com", trangThai: "Đang hợp tác" },
      { id: "CAT-SUP-002", maNhaCungCap: "NCC-002", tenNhaCungCap: "Thiết bị Nam Phát", nhomCungCap: "Máy móc", nguoiLienHe: "Anh Phát", soDienThoai: "0911000002", email: "phat@example.com", trangThai: "Dự phòng" },
    ],
  },
  {
    key: "costTypes",
    label: "Loại chi phí",
    storageKey: "costTypesCatalog",
    columns: [
      { key: "maChiPhi", label: "Mã chi phí", required: true },
      { key: "tenChiPhi", label: "Tên chi phí", required: true },
      { key: "nhomChiPhi", label: "Nhóm chi phí" },
      { key: "hinhThucThanhToan", label: "Thanh toán", type: "select", options: ["Tiền mặt", "Chuyển khoản", "Công nợ"] },
      { key: "canPheDuyet", label: "Cần duyệt", type: "select", options: ["Có", "Không"] },
      { key: "ghiChu", label: "Ghi chú", wide: true },
    ],
    samples: [
      { id: "CAT-COST-001", maChiPhi: "CP-001", tenChiPhi: "Vận chuyển vật tư", nhomChiPhi: "Vận chuyển", hinhThucThanhToan: "Chuyển khoản", canPheDuyet: "Có", ghiChu: "Dùng cho chi phí giao nhận vật tư" },
      { id: "CAT-COST-002", maChiPhi: "CP-002", tenChiPhi: "Chi phí phát sinh hiện trường", nhomChiPhi: "Phát sinh", hinhThucThanhToan: "Tiền mặt", canPheDuyet: "Có", ghiChu: "Cần ghi rõ lý do phát sinh" },
    ],
  },
  {
    key: "projectStatuses",
    label: "Trạng thái dự án",
    storageKey: "projectStatusesCatalog",
    columns: [
      { key: "maTrangThai", label: "Mã TT", required: true },
      { key: "tenTrangThai", label: "Tên trạng thái", required: true },
      { key: "nhomTrangThai", label: "Nhóm" },
      { key: "mauHienThi", label: "Màu hiển thị" },
      { key: "thuTu", label: "Thứ tự", type: "number" },
      { key: "ghiChu", label: "Ghi chú", wide: true },
    ],
    samples: [
      { id: "CAT-STS-001", maTrangThai: "TT-001", tenTrangThai: "Đang chuẩn bị", nhomTrangThai: "Trước thi công", mauHienThi: "amber", thuTu: "1", ghiChu: "Chưa bắt đầu hiện trường" },
      { id: "CAT-STS-002", maTrangThai: "TT-002", tenTrangThai: "Đang thi công", nhomTrangThai: "Thi công", mauHienThi: "blue", thuTu: "2", ghiChu: "Đang triển khai công việc" },
      { id: "CAT-STS-003", maTrangThai: "TT-003", tenTrangThai: "Hoàn thành", nhomTrangThai: "Kết thúc", mauHienThi: "green", thuTu: "3", ghiChu: "Đã nghiệm thu/kết thúc" },
    ],
  },
  {
    key: "users",
    label: "Người dùng",
    storageKey: "usersCatalog",
    columns: [
      { key: "maNguoiDung", label: "Mã user", required: true },
      { key: "hoTen", label: "Họ tên", required: true },
      { key: "email", label: "Email" },
      { key: "sdt", label: "Số điện thoại" },
      { key: "chucDanh", label: "Chức danh" },
      { key: "congTrinhPhanCong", label: "Công trình được phân công", wide: true },
      { key: "roles", label: "Vai trò", type: "roles", required: true },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Hoạt động", "Khóa"] },
    ],
    samples: [
      { id: "CAT-USR-001", maNguoiDung: "ADMIN", hoTen: "Admin hệ thống", email: "admin@sonhoabinh.vn", matKhau: "123456", sdt: "0909000000", chucDanh: "Quản trị hệ thống", congTrinhPhanCong: "Tất cả", roles: ["ADMIN"], trangThai: "Hoạt động" },
      { id: "CAT-USR-002", maNguoiDung: "U-002", hoTen: "Nguyễn Văn A", email: "vana@example.com", matKhau: "123456", sdt: "0909111222", chucDanh: "PM kiêm Chỉ huy trưởng", congTrinhPhanCong: "CT001 - Biệt thự Thảo Điền", roles: ["PM", "CHT", "QS"], trangThai: "Hoạt động" },
      { id: "CAT-USR-003", maNguoiDung: "U-003", hoTen: "Trần Văn B", email: "vanb@example.com", matKhau: "123456", sdt: "0909333444", chucDanh: "QS/QC công trình", congTrinhPhanCong: "CT002 - Nhà phố Bình Dương", roles: ["QS", "QC"], trangThai: "Hoạt động" },
      { id: "CAT-USR-004", maNguoiDung: "U-004", hoTen: "Lê Văn C", email: "vanc@example.com", matKhau: "123456", sdt: "0909555666", chucDanh: "Thư ký dự án", congTrinhPhanCong: "CT001, CT002", roles: ["TKDA", "KTCT"], trangThai: "Hoạt động" },
    ],
  },
];

const systemCatalogTabs = dataCatalogTabs.filter((tab) => tab.key !== "users");

const systemSettingsTabs = [
  dataCatalogTabs.find((tab) => tab.key === "users"),
  {
    key: "roles",
    label: "Vai trò",
    storageKey: "rolesCatalog",
    columns: [
      { key: "maVaiTro", label: "Mã vai trò", required: true },
      { key: "tenVaiTro", label: "Tên vai trò", required: true },
      { key: "nhomVaiTro", label: "Nhóm vai trò" },
      { key: "moTa", label: "Mô tả", wide: true },
      { key: "trangThai", label: "Trạng thái", type: "select", options: ["Hoạt động", "Tạm khóa"] },
    ],
    samples: constructionRoleDefinitions,
  },
  { key: "permissions", label: "Phân quyền", storageKey: "permissionMatrix" },
  { key: "approvalWorkflows", label: "Quy trình phê duyệt", storageKey: "approvalWorkflows" },
  { key: "systemLogs", label: "Nhật ký hệ thống", storageKey: "systemLogs" },
  { key: "backup", label: "Sao lưu dữ liệu", storageKey: "backupTools" },
].filter(Boolean);

const defaultPermissionMatrix = buildDefaultPermissionMatrix();

const defaultApprovalWorkflows = [
  { id: "WF-001", name: "Chi phí khác", steps: ["Chỉ huy trưởng / PM tạo đề nghị", "Giám đốc Thi công duyệt", "TGĐ duyệt nếu vượt hạn mức"] },
  { id: "WF-002", name: "Nghiệm thu", steps: ["Chỉ huy trưởng / QS lập", "PM kiểm tra", "QC xác nhận chất lượng", "Thư ký dự án hoàn thiện hồ sơ", "Kế toán công trình theo dõi thanh toán"] },
  { id: "WF-003", name: "Hồ sơ thanh toán", steps: ["Thư ký dự án lập hồ sơ", "PM kiểm tra", "Kế toán công trình kiểm tra số liệu", "TGĐ / Giám đốc Thi công xem báo cáo"] },
];

const defaultSystemLogs = [
  { id: "LOG-001", nguoiDung: "Khoa", thoiGian: "10/06/2026", chucNang: "BOQ", hanhDong: "Chỉnh sửa" },
  { id: "LOG-002", nguoiDung: "Admin", thoiGian: "10/06/2026", chucNang: "Người dùng", hanhDong: "Thêm mới" },
];

const defaultSystemSettings = {
  roles: constructionRoleDefinitions,
  permissions: defaultPermissionMatrix,
  approvalWorkflows: defaultApprovalWorkflows,
  systemLogs: defaultSystemLogs,
};

const defaultSystemCatalogs = dataCatalogTabs.reduce((catalogs, tab) => {
  catalogs[tab.key] = tab.samples;
  return catalogs;
}, {});

const editableSettingTabs = systemSettingsTabs.filter((tab) => tab.columns);

const defaultSystemCatalogForm = [...systemCatalogTabs, ...editableSettingTabs].reduce((forms, tab) => {
  forms[tab.key] = tab.columns.reduce((form, column) => {
    form[column.key] = column.type === "roles" ? [] : column.type === "select" ? column.options?.[0] ?? "" : "";
    return form;
  }, {});
  return forms;
}, {});

const mainMenus = [
  { key: "dashboard", label: "Dashboard" },
  { key: "contracts", label: "Hợp đồng" },
  { key: "boq", label: "BOQ - Dự toán" },
  { key: "wbs", label: "Kế hoạch thi công WBS" },
  { key: "daily-log", label: "Nhật ký thi công" },
  { key: "materials", label: "Vật tư" },
  { key: "labor", label: "Nhân công" },
  { key: "subcontractors", label: "Nhà thầu phụ / Tổ đội" },
  { key: "equipment", label: "Máy móc thiết bị" },
  { key: "acceptance", label: "Nghiệm thu" },
  { key: "invoices", label: "Hóa đơn - Công nợ" },
  { key: "other-costs", label: "Chi phí khác" },
  { key: "reports", label: "Báo cáo" },
  { key: "documents", label: "Hồ sơ dự án" },
  { key: "document-templates", label: "📚 Biểu mẫu & Tài liệu thi công" },
  { key: "system-catalog", label: "Danh mục hệ thống" },
  { key: "system-settings", label: "⚙️ Cài đặt hệ thống" },
];

const MENU_KEY_ALIASES = {
  "daily-log": "dailyLogs",
  "other-costs": "otherCosts",
  documents: "projectDocuments",
  "document-templates": "documentTemplates",
  "system-catalog": "masterData",
  "system-settings": "settings",
  dailyLogs: "dailyLogs",
  otherCosts: "otherCosts",
  projectDocuments: "projectDocuments",
  documentTemplates: "documentTemplates",
  masterData: "masterData",
  settings: "settings",
};

const ROLE_MENU_PERMISSIONS = {
  ADMIN: ["*"],
  TGD: [
    "dashboard",
    "contracts",
    "boq",
    "wbs",
    "dailyLogs",
    "materials",
    "labor",
    "subcontractors",
    "equipment",
    "acceptance",
    "invoices",
    "otherCosts",
    "reports",
    "projectDocuments",
    "documentTemplates",
  ],
  GDTC: [
    "dashboard",
    "contracts",
    "boq",
    "wbs",
    "dailyLogs",
    "materials",
    "labor",
    "subcontractors",
    "equipment",
    "acceptance",
    "invoices",
    "otherCosts",
    "reports",
    "projectDocuments",
    "documentTemplates",
  ],
  PM: [
    "dashboard",
    "contracts",
    "boq",
    "wbs",
    "dailyLogs",
    "materials",
    "labor",
    "subcontractors",
    "equipment",
    "acceptance",
    "invoices",
    "otherCosts",
    "reports",
    "projectDocuments",
    "documentTemplates",
  ],
  CHT: [
    "dashboard",
    "wbs",
    "dailyLogs",
    "materials",
    "labor",
    "subcontractors",
    "equipment",
    "acceptance",
    "projectDocuments",
    "documentTemplates",
  ],
  QS: [
    "dashboard",
    "contracts",
    "boq",
    "wbs",
    "dailyLogs",
    "acceptance",
    "reports",
    "projectDocuments",
    "documentTemplates",
  ],
  QC: [
    "dashboard",
    "dailyLogs",
    "acceptance",
    "projectDocuments",
    "documentTemplates",
  ],
  TKDA: [
    "dashboard",
    "contracts",
    "wbs",
    "dailyLogs",
    "acceptance",
    "invoices",
    "reports",
    "projectDocuments",
    "documentTemplates",
  ],
  KTCT: [
    "dashboard",
    "contracts",
    "acceptance",
    "invoices",
    "otherCosts",
    "reports",
    "projectDocuments",
  ],
  KHOCT: [
    "dashboard",
    "dailyLogs",
    "materials",
    "projectDocuments",
    "documentTemplates",
  ],
};

const sidebarGroups = [
  {
    key: "business",
    items: mainMenus.filter((item) => !["system-catalog", "system-settings"].includes(item.key)),
  },
  {
    key: "catalog",
    title: "DANH MỤC HỆ THỐNG",
    items: [mainMenus.find((item) => item.key === "system-catalog")],
  },
  {
    key: "settings",
    title: "CÀI ĐẶT HỆ THỐNG",
    items: [mainMenus.find((item) => item.key === "system-settings")],
  },
];

const defaultLogForm = {
  ngay: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  maWbs: "",
  hangMucCongViec: "",
  giamSat: "",
  thoiTiet: "Nắng",
  noiDung: "",
  dvtKhoiLuong: "m2",
  khoiLuongKeHoach: "",
  khoiLuongThucHienNgay: "",
  maVatTu: "",
  tenVatTu: "",
  dvtVatTu: "kg",
  soLuongVatTu: "",
  toDoi: "",
  soNguoi: "",
  soGio: "",
  ghiChuNhanCong: "",
  maThietBi: "",
  tenThietBi: "",
  soGioHoatDong: "",
  nguonThietBi: "Công ty",
  vanDe: "",
  ghiChu: "",
};

const defaultDailyImageMeta = {
  loaiAnh: DAILY_IMAGE_TYPES[0],
  nguoiTaiLen: "",
  ghiChuAnh: "",
};

const defaultContractForm = {
  maHopDong: "",
  tenCongTrinh: "",
  khachHang: "",
  diaChiCongTrinh: "",
  giaTriHopDong: "",
  ngayKy: "",
  ngayKhoiCong: "",
  ngayHoanThanhKeHoach: "",
  pmPhuTrach: "",
  giamSatPhuTrach: "",
  trangThai: "Đang chuẩn bị",
};

const defaultBoqForm = {
  hopDongCongTrinh: "",
  maWbs: "",
  hangMucCongViec: "",
  dvt: "m2",
  khoiLuongHopDong: "",
  donGiaHopDong: "",
  dinhMucVatTu: "",
  dinhMucNhanCong: "",
  dinhMucMmtb: "",
  chiPhiKeHoach: "",
  ghiChu: "",
};

const defaultWbsForm = {
  hopDongCongTrinh: "",
  maWbs: "",
  hangMucCongViec: "",
  ngayBatDauKeHoach: "",
  ngayKetThucKeHoach: "",
  khoiLuongKeHoach: "",
  nguoiPhuTrach: "",
  trangThai: WBS_STATUSES[0],
  ghiChu: "",
};

const defaultMaterialInForm = {
  ngayNhap: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  maVatTu: "",
  tenVatTu: "",
  dvt: "kg",
  soLuongNhap: "",
  nguonCap: "Nhà máy",
  nguoiGiao: "",
  nguoiNhan: "",
  ghiChu: "",
};

const defaultLaborTeamForm = {
  maToDoi: "",
  tenToDoi: "",
  nguoiDaiDien: "",
  soDienThoai: "",
  loaiToDoi: "Sơn đá",
  donGiaCongNgay: "",
  ghiChu: "",
};

const defaultSubcontractorForm = {
  maDoiTac: "",
  tenDoiTac: "",
  nguoiDaiDien: "",
  soDienThoai: "",
  loaiDoiTac: "Tổ đội thi công",
  hangMucChuyenMon: "Sơn đá",
  ghiChu: "",
};

const defaultSubContractForm = {
  hopDongCongTrinh: "",
  maWbs: "",
  hangMucCongViec: "",
  doiTac: "",
  dvt: "m2",
  khoiLuongGiaoKhoan: "",
  donGiaKhoan: "",
  ngayBatDau: "",
  ngayKetThuc: "",
  trangThai: "Đang giao khoán",
  ghiChu: "",
};

const defaultEquipmentForm = {
  maThietBi: "",
  tenThietBi: "",
  loaiThietBi: "Máy phun",
  nguonSoHuu: "Công ty",
  donViSoHuu: "",
  donGiaThueNgay: "",
  donGiaThueGio: "",
  tinhTrangHienTai: "Sẵn sàng",
  ghiChu: "",
};

const defaultEquipmentTransferForm = {
  ngayNhan: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  maThietBi: "",
  tenThietBi: "",
  nguon: "Công ty",
  nguoiGiao: "",
  nguoiNhan: "",
  tinhTrangKhiNhan: "",
  ngayTraDuKien: "",
  ghiChu: "",
};

const defaultAcceptanceBatchForm = {
  soBienBan: "",
  ngayNghiemThu: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  khachHangChuDauTu: "",
  daiDienBenA: "",
  daiDienBenB: "",
  ghiChu: "",
};

const defaultAcceptanceForm = {
  soBienBan: "",
  hopDongCongTrinh: "",
  maWbs: "",
  hangMucCongViec: "",
  dvt: "m2",
  khoiLuongHopDong: "",
  khoiLuongDaThucHienLuyKe: "",
  khoiLuongNghiemThuKyNay: "",
  donGiaHopDong: "",
  ghiChu: "",
};

const defaultInvoiceForm = {
  soHoaDon: "",
  ngayXuatHoaDon: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  soBienBan: "",
  giaTriNghiemThu: "",
  tyLeXuatHoaDon: "100",
  vatPercent: "10",
  ghiChu: "",
};

const defaultCollectionForm = {
  soPhieuThu: "",
  ngayThuTien: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  soHoaDon: "",
  soTienThu: "",
  hinhThucThu: "Chuyển khoản",
  nguoiNop: "",
  ghiChu: "",
};

const defaultOtherCostForm = {
  ngayPhatSinh: new Date().toISOString().slice(0, 10),
  hopDongCongTrinh: "",
  nhomChiPhi: "Vận chuyển",
  noiDungChiPhi: "",
  soTien: "",
  nguoiDeNghi: "",
  nguoiDuyet: "",
  trangThai: "Chờ duyệt",
  hinhThucThanhToan: "Tiền mặt",
  ghiChu: "",
};

const defaultProjectDocumentForm = {
  hopDongCongTrinh: "",
  nhomHoSo: DOCUMENT_TABS[0],
  maHoSo: "",
  tenHoSo: "",
  loaiHoSo: "",
  ngayTaiLen: new Date().toISOString().slice(0, 10),
  nguoiTai: "",
  trangThai: "Đã tải lên",
  fileName: "",
  fileType: "",
  fileSize: 0,
  fileData: "",
};

const defaultDocumentTemplateForm = {
  maTaiLieu: "",
  tenTaiLieu: "",
  nhomTaiLieu: DOCUMENT_TEMPLATE_GROUPS[0],
  loaiFile: DOCUMENT_TEMPLATE_FILE_TYPES[0],
  phienBan: "1.0",
  ngayHieuLuc: new Date().toISOString().slice(0, 10),
  ngayHetHieuLuc: "",
  nguoiBanHanh: "",
  phongBanPhuTrach: "Phòng Thi công",
  moTaHuongDan: "",
  trangThai: DOCUMENT_TEMPLATE_STATUSES[0],
  fileName: "",
  fileType: "",
  fileSize: 0,
  fileData: "",
  downloadCount: 0,
};

function App() {
  const [currentUser, setCurrentUser] = useState(readCurrentUser);
  const [loginForm, setLoginForm] = useState({ email: "", matKhau: "" });
  const [loginError, setLoginError] = useState("");
  const [activeMenu, setActiveMenu] = useState(mainMenus[0].key);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [selectedContractId, setSelectedContractId] = useState(ALL_CONTRACTS);
  const [logForm, setLogForm] = useState(defaultLogForm);
  const [constructionLogs, setConstructionLogs] = useState(readSavedLogs);
  const [isLogFormOpen, setIsLogFormOpen] = useState(false);
  const [dailyImageMeta, setDailyImageMeta] = useState(defaultDailyImageMeta);
  const [dailyLogImages, setDailyLogImages] = useState([]);
  const [visibleImageLogId, setVisibleImageLogId] = useState(null);
  const [contracts, setContracts] = useState(readSavedContracts);
  const [contractForm, setContractForm] = useState(defaultContractForm);
  const [isContractFormOpen, setIsContractFormOpen] = useState(false);
  const [boqItems, setBoqItems] = useState(readSavedBoqItems);
  const [boqForm, setBoqForm] = useState(defaultBoqForm);
  const [isBoqFormOpen, setIsBoqFormOpen] = useState(false);
  const [wbsPlans, setWbsPlans] = useState(readSavedWbsPlans);
  const [wbsForm, setWbsForm] = useState(defaultWbsForm);
  const [isWbsFormOpen, setIsWbsFormOpen] = useState(false);
  const [materialIns, setMaterialIns] = useState(readSavedMaterialIns);
  const [materialInForm, setMaterialInForm] = useState(defaultMaterialInForm);
  const [isMaterialInFormOpen, setIsMaterialInFormOpen] = useState(false);
  const [laborTeams, setLaborTeams] = useState(readSavedLaborTeams);
  const [laborTeamForm, setLaborTeamForm] = useState(defaultLaborTeamForm);
  const [isLaborTeamFormOpen, setIsLaborTeamFormOpen] = useState(false);
  const [subcontractors, setSubcontractors] = useState(readSavedSubcontractors);
  const [subcontractorForm, setSubcontractorForm] = useState(defaultSubcontractorForm);
  const [isSubcontractorFormOpen, setIsSubcontractorFormOpen] = useState(false);
  const [subContracts, setSubContracts] = useState(readSavedSubContracts);
  const [subContractForm, setSubContractForm] = useState(defaultSubContractForm);
  const [isSubContractFormOpen, setIsSubContractFormOpen] = useState(false);
  const [equipmentList, setEquipmentList] = useState(readSavedEquipmentList);
  const [equipmentForm, setEquipmentForm] = useState(defaultEquipmentForm);
  const [isEquipmentFormOpen, setIsEquipmentFormOpen] = useState(false);
  const [equipmentTransfers, setEquipmentTransfers] = useState(readSavedEquipmentTransfers);
  const [equipmentTransferForm, setEquipmentTransferForm] = useState(defaultEquipmentTransferForm);
  const [isEquipmentTransferFormOpen, setIsEquipmentTransferFormOpen] = useState(false);
  const [acceptanceBatches, setAcceptanceBatches] = useState(readSavedAcceptanceBatches);
  const [acceptances, setAcceptances] = useState(readSavedAcceptances);
  const [acceptanceBatchForm, setAcceptanceBatchForm] = useState(defaultAcceptanceBatchForm);
  const [acceptanceForm, setAcceptanceForm] = useState(defaultAcceptanceForm);
  const [isAcceptanceBatchFormOpen, setIsAcceptanceBatchFormOpen] = useState(false);
  const [isAcceptanceFormOpen, setIsAcceptanceFormOpen] = useState(false);
  const [invoices, setInvoices] = useState(readSavedInvoices);
  const [collections, setCollections] = useState(readSavedCollections);
  const [invoiceForm, setInvoiceForm] = useState(defaultInvoiceForm);
  const [collectionForm, setCollectionForm] = useState(defaultCollectionForm);
  const [isInvoiceFormOpen, setIsInvoiceFormOpen] = useState(false);
  const [isCollectionFormOpen, setIsCollectionFormOpen] = useState(false);
  const [otherCosts, setOtherCosts] = useState(readSavedOtherCosts);
  const [otherCostForm, setOtherCostForm] = useState(defaultOtherCostForm);
  const [isOtherCostFormOpen, setIsOtherCostFormOpen] = useState(false);
  const [reportFilters, setReportFilters] = useState({
    project: ALL_PROJECTS,
    fromDate: "",
    toDate: "",
  });
  const [projectDocuments, setProjectDocuments] = useState(readSavedProjectDocuments);
  const [projectDocumentForm, setProjectDocumentForm] = useState(defaultProjectDocumentForm);
  const [activeDocumentTab, setActiveDocumentTab] = useState(DOCUMENT_TABS[0]);
  const [documentSearch, setDocumentSearch] = useState("");
  const [documentProjectFilter, setDocumentProjectFilter] = useState(ALL_PROJECTS);
  const [documentTemplates, setDocumentTemplates] = useState(readSavedDocumentTemplates);
  const [documentTemplateForm, setDocumentTemplateForm] = useState(defaultDocumentTemplateForm);
  const [documentTemplateSearch, setDocumentTemplateSearch] = useState("");
  const [documentTemplateFilters, setDocumentTemplateFilters] = useState({
    nhomTaiLieu: ALL_PROJECTS,
    loaiFile: ALL_PROJECTS,
    trangThai: ALL_PROJECTS,
  });
  const [editingDocumentTemplateId, setEditingDocumentTemplateId] = useState(null);
  const [systemCatalogs, setSystemCatalogs] = useState(readSavedSystemCatalogs);
  const [activeSystemCatalogTab, setActiveSystemCatalogTab] = useState(systemCatalogTabs[0].key);
  const [systemCatalogForm, setSystemCatalogForm] = useState(defaultSystemCatalogForm[systemCatalogTabs[0].key]);
  const [systemCatalogSearch, setSystemCatalogSearch] = useState("");
  const [editingSystemCatalogItemId, setEditingSystemCatalogItemId] = useState(null);
  const [isSystemCatalogFormOpen, setIsSystemCatalogFormOpen] = useState(false);
  const [systemSettings, setSystemSettings] = useState(readSavedSystemSettings);
  const [activeSystemSettingsTab, setActiveSystemSettingsTab] = useState(systemSettingsTabs[0].key);
  const [systemSettingsForm, setSystemSettingsForm] = useState(defaultSystemCatalogForm[systemSettingsTabs[0].key]);
  const [systemSettingsSearch, setSystemSettingsSearch] = useState("");
  const [editingSystemSettingsItemId, setEditingSystemSettingsItemId] = useState(null);
  const [isSystemSettingsFormOpen, setIsSystemSettingsFormOpen] = useState(false);
  const [excelImportErrors, setExcelImportErrors] = useState([]);

  const activeMenuKey = canViewMenu(activeMenu) ? activeMenu : getPreferredMenuKey();
  const canAccessActiveMenu = canViewMenu(activeMenuKey);
  const activeMenuInfo = mainMenus.find((item) => item.key === activeMenuKey) ?? mainMenus[0];
  const visibleSidebarGroups = sidebarGroups
    .map((group) => ({
      ...group,
      items: group.items.filter((item) => item && canViewMenu(item.key)),
    }))
    .filter((group) => group.items.length > 0);
  const contractFilterOptions = useMemo(() => getContractFilterOptions(contracts), [contracts]);
  const selectedContractOption = contractFilterOptions.find((option) => option.id === selectedContractId);
  const selectedContractName = selectedContractOption?.name || "";
  const selectedContractLabel = selectedContractOption?.label || "";

  function isMatchContract(item) {
    return isMatchContractFilter(item, selectedContractId, selectedContractName, selectedContractLabel);
  }

  const filteredContracts = contracts.filter(isMatchContract);
  const filteredBoqItems = boqItems.filter(isMatchContract);
  const filteredWbsPlans = wbsPlans.filter(isMatchContract);
  const filteredDailyLogs = constructionLogs.filter(isMatchContract);
  const filteredMaterialIns = materialIns.filter(isMatchContract);
  const filteredEquipmentTransfers = equipmentTransfers.filter(isMatchContract);
  const filteredAcceptances = acceptances.filter(isMatchContract);
  const filteredInvoices = invoices.filter(isMatchContract);
  const filteredCollections = collections.filter(isMatchContract);
  const filteredOtherCosts = otherCosts.filter(isMatchContract);
  const filteredSubContracts = subContracts.filter(isMatchContract);

  function handleLogChange(event) {
    const { name, value } = event.target;
    const wbsSource = wbsPlans.length ? wbsPlans : getCatalogSeed("wbsStandards").length ? getCatalogSeed("wbsStandards") : sampleWbsPlans;
    const linkedWbsPlan = name === "maWbs" ? wbsSource.find((plan) => plan.maWbs === value) : null;

    setLogForm({
      ...logForm,
      [name]: value,
      ...(linkedWbsPlan
        ? {
            hopDongCongTrinh: linkedWbsPlan.hopDongCongTrinh,
            hangMucCongViec: linkedWbsPlan.hangMucCongViec,
            khoiLuongKeHoach: linkedWbsPlan.khoiLuongKeHoach,
          }
        : {}),
    });
  }

  function handleDailyImageMetaChange(event) {
    setDailyImageMeta({ ...dailyImageMeta, [event.target.name]: event.target.value });
  }

  async function handleDailyImageUpload(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    const compressedImages = await Promise.all(
      files.map(async (file) => ({
        id: crypto.randomUUID(),
        loaiAnh: dailyImageMeta.loaiAnh,
        tenFile: file.name,
        fileType: file.type,
        fileSize: file.size,
        ngayTaiLen: new Date().toISOString().slice(0, 10),
        nguoiTaiLen: dailyImageMeta.nguoiTaiLen,
        ghiChuAnh: dailyImageMeta.ghiChuAnh,
        fileAnh: await compressImageFile(file),
      })),
    );

    setDailyLogImages([...dailyLogImages, ...compressedImages]);
    event.target.value = "";
  }

  function handleRemoveDailyImage(imageId) {
    setDailyLogImages(dailyLogImages.filter((image) => image.id !== imageId));
  }

  function saveLogs(nextLogs) {
    setConstructionLogs(nextLogs);
    localStorage.setItem(DAILY_LOG_STORAGE_KEY, JSON.stringify(nextLogs));
  }

  function handleSaveLog(event) {
    event.preventDefault();

    if (!logForm.ngay || !logForm.hopDongCongTrinh || !logForm.maWbs || !logForm.noiDung.trim()) {
      alert("Vui lòng nhập ngày, hợp đồng/công trình, mã WBS và nội dung công việc.");
      return;
    }

    const previousCumulative = constructionLogs
      .filter((log) => log.hopDongCongTrinh === logForm.hopDongCongTrinh && log.maWbs === logForm.maWbs)
      .reduce((sum, log) => sum + toNumber(log.khoiLuongThucHienNgay), 0);
    const khoiLuongLuyKe = previousCumulative + toNumber(logForm.khoiLuongThucHienNgay);
    const tyLeHoanThanh = toNumber(logForm.khoiLuongKeHoach)
      ? (khoiLuongLuyKe / toNumber(logForm.khoiLuongKeHoach)) * 100
      : 0;
    const nextLog = {
      id: crypto.randomUUID(),
      ...logForm,
      hinhAnhHienTruong: dailyLogImages,
      khoiLuongLuyKe,
      tyLeHoanThanh,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveLogs([nextLog, ...constructionLogs]);
    setLogForm({
      ...defaultLogForm,
      ngay: logForm.ngay,
      hopDongCongTrinh: logForm.hopDongCongTrinh,
      giamSat: logForm.giamSat,
    });
    setDailyImageMeta(defaultDailyImageMeta);
    setDailyLogImages([]);
    setIsLogFormOpen(false);
  }

  function handleDeleteLog(logId) {
    if (!confirm("Xóa nhật ký này?")) return;
    saveLogs(constructionLogs.filter((log) => log.id !== logId));
  }

  function handleCancelLogForm() {
    setLogForm(defaultLogForm);
    setDailyImageMeta(defaultDailyImageMeta);
    setDailyLogImages([]);
    setIsLogFormOpen(false);
  }

  function handleContractChange(event) {
    setContractForm({ ...contractForm, [event.target.name]: event.target.value });
  }

  function saveContracts(nextContracts) {
    setContracts(nextContracts);
    localStorage.setItem(CONTRACT_STORAGE_KEY, JSON.stringify(nextContracts));
    localStorage.setItem("contracts", JSON.stringify(nextContracts));
  }

  function handleSaveContract(event) {
    event.preventDefault();

    if (!contractForm.maHopDong.trim() || !contractForm.tenCongTrinh.trim() || !contractForm.khachHang.trim()) {
      alert("Vui lòng nhập mã hợp đồng, tên công trình và khách hàng.");
      return;
    }

    const nextContract = {
      id: crypto.randomUUID(),
      ...contractForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveContracts([nextContract, ...contracts]);
    setContractForm(defaultContractForm);
    setIsContractFormOpen(false);
  }

  function handleCancelContractForm() {
    setContractForm(defaultContractForm);
    setIsContractFormOpen(false);
  }

  function handleBoqChange(event) {
    setBoqForm({ ...boqForm, [event.target.name]: event.target.value });
  }

  function saveBoqItems(nextItems) {
    setBoqItems(nextItems);
    localStorage.setItem(BOQ_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveBoqItem(event) {
    event.preventDefault();

    if (!boqForm.hopDongCongTrinh || !boqForm.maWbs.trim() || !boqForm.hangMucCongViec.trim()) {
      alert("Vui lòng nhập hợp đồng/công trình, mã WBS và hạng mục công việc.");
      return;
    }

    const thanhTienHopDong = toNumber(boqForm.khoiLuongHopDong) * toNumber(boqForm.donGiaHopDong);
    const nextItem = {
      id: crypto.randomUUID(),
      ...boqForm,
      thanhTienHopDong,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveBoqItems([nextItem, ...boqItems]);
    setBoqForm(defaultBoqForm);
    setIsBoqFormOpen(false);
  }

  function handleCancelBoqForm() {
    setBoqForm(defaultBoqForm);
    setIsBoqFormOpen(false);
  }

  function handleWbsChange(event) {
    const { name, value } = event.target;
    const linkedBoqItem = name === "maWbs" ? boqItems.find((item) => item.maWbs === value) : null;

    setWbsForm({
      ...wbsForm,
      [name]: value,
      ...(linkedBoqItem
        ? {
            hopDongCongTrinh: linkedBoqItem.hopDongCongTrinh,
            hangMucCongViec: linkedBoqItem.hangMucCongViec,
            khoiLuongKeHoach: linkedBoqItem.khoiLuongHopDong,
          }
        : {}),
    });
  }

  function saveWbsPlans(nextPlans) {
    setWbsPlans(nextPlans);
    localStorage.setItem(WBS_STORAGE_KEY, JSON.stringify(nextPlans));
  }

  function handleSaveWbsPlan(event) {
    event.preventDefault();

    if (!wbsForm.hopDongCongTrinh || !wbsForm.maWbs.trim() || !wbsForm.hangMucCongViec.trim()) {
      alert("Vui lòng nhập hợp đồng/công trình, mã WBS và hạng mục công việc.");
      return;
    }

    const nextPlan = {
      id: crypto.randomUUID(),
      ...wbsForm,
      soNgayKeHoach: calculatePlanDays(wbsForm.ngayBatDauKeHoach, wbsForm.ngayKetThucKeHoach),
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveWbsPlans([nextPlan, ...wbsPlans]);
    setWbsForm(defaultWbsForm);
    setIsWbsFormOpen(false);
  }

  function handleCancelWbsForm() {
    setWbsForm(defaultWbsForm);
    setIsWbsFormOpen(false);
  }

  function handleMaterialInChange(event) {
    setMaterialInForm({ ...materialInForm, [event.target.name]: event.target.value });
  }

  function saveMaterialIns(nextItems) {
    setMaterialIns(nextItems);
    localStorage.setItem(MATERIAL_IN_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveMaterialIn(event) {
    event.preventDefault();

    if (
      !materialInForm.ngayNhap ||
      !materialInForm.hopDongCongTrinh ||
      !materialInForm.maVatTu.trim() ||
      !materialInForm.tenVatTu.trim()
    ) {
      alert("Vui lòng nhập ngày nhập, hợp đồng/công trình, mã vật tư và tên vật tư.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...materialInForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveMaterialIns([nextItem, ...materialIns]);
    setMaterialInForm(defaultMaterialInForm);
    setIsMaterialInFormOpen(false);
  }

  function handleCancelMaterialInForm() {
    setMaterialInForm(defaultMaterialInForm);
    setIsMaterialInFormOpen(false);
  }

  function handleLaborTeamChange(event) {
    setLaborTeamForm({ ...laborTeamForm, [event.target.name]: event.target.value });
  }

  function saveLaborTeams(nextTeams) {
    setLaborTeams(nextTeams);
    localStorage.setItem(LABOR_TEAM_STORAGE_KEY, JSON.stringify(nextTeams));
  }

  function handleSaveLaborTeam(event) {
    event.preventDefault();

    if (!laborTeamForm.maToDoi.trim() || !laborTeamForm.tenToDoi.trim()) {
      alert("Vui lòng nhập mã tổ đội và tên tổ đội.");
      return;
    }

    const nextTeam = {
      id: crypto.randomUUID(),
      ...laborTeamForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveLaborTeams([nextTeam, ...laborTeams]);
    setLaborTeamForm(defaultLaborTeamForm);
    setIsLaborTeamFormOpen(false);
  }

  function handleCancelLaborTeamForm() {
    setLaborTeamForm(defaultLaborTeamForm);
    setIsLaborTeamFormOpen(false);
  }

  function handleSubcontractorChange(event) {
    setSubcontractorForm({ ...subcontractorForm, [event.target.name]: event.target.value });
  }

  function saveSubcontractors(nextItems) {
    setSubcontractors(nextItems);
    localStorage.setItem(SUBCONTRACTOR_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveSubcontractor(event) {
    event.preventDefault();

    if (!subcontractorForm.maDoiTac.trim() || !subcontractorForm.tenDoiTac.trim()) {
      alert("Vui lòng nhập mã đối tác và tên nhà thầu phụ / tổ đội.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...subcontractorForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveSubcontractors([nextItem, ...subcontractors]);
    setSubcontractorForm(defaultSubcontractorForm);
    setIsSubcontractorFormOpen(false);
  }

  function handleCancelSubcontractorForm() {
    setSubcontractorForm(defaultSubcontractorForm);
    setIsSubcontractorFormOpen(false);
  }

  function handleSubContractChange(event) {
    const { name, value } = event.target;
    const linkedWbsPlan = name === "maWbs" ? wbsPlans.find((plan) => plan.maWbs === value) : null;

    setSubContractForm({
      ...subContractForm,
      [name]: value,
      ...(linkedWbsPlan
        ? {
            hopDongCongTrinh: linkedWbsPlan.hopDongCongTrinh,
            hangMucCongViec: linkedWbsPlan.hangMucCongViec,
            khoiLuongGiaoKhoan: linkedWbsPlan.khoiLuongKeHoach,
          }
        : {}),
    });
  }

  function saveSubContracts(nextItems) {
    setSubContracts(nextItems);
    localStorage.setItem(SUBCONTRACT_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveSubContract(event) {
    event.preventDefault();

    if (
      !subContractForm.hopDongCongTrinh ||
      !subContractForm.maWbs.trim() ||
      !subContractForm.hangMucCongViec.trim() ||
      !subContractForm.doiTac
    ) {
      alert("Vui lòng nhập hợp đồng/công trình, mã WBS, hạng mục và nhà thầu phụ / tổ đội.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...subContractForm,
      thanhTienKhoan: toNumber(subContractForm.khoiLuongGiaoKhoan) * toNumber(subContractForm.donGiaKhoan),
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveSubContracts([nextItem, ...subContracts]);
    setSubContractForm(defaultSubContractForm);
    setIsSubContractFormOpen(false);
  }

  function handleCancelSubContractForm() {
    setSubContractForm(defaultSubContractForm);
    setIsSubContractFormOpen(false);
  }

  function handleEquipmentChange(event) {
    setEquipmentForm({ ...equipmentForm, [event.target.name]: event.target.value });
  }

  function saveEquipmentList(nextItems) {
    setEquipmentList(nextItems);
    localStorage.setItem(EQUIPMENT_LIST_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveEquipment(event) {
    event.preventDefault();

    if (!equipmentForm.maThietBi.trim() || !equipmentForm.tenThietBi.trim()) {
      alert("Vui lòng nhập mã thiết bị và tên thiết bị.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...equipmentForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveEquipmentList([nextItem, ...equipmentList]);
    setEquipmentForm(defaultEquipmentForm);
    setIsEquipmentFormOpen(false);
  }

  function handleCancelEquipmentForm() {
    setEquipmentForm(defaultEquipmentForm);
    setIsEquipmentFormOpen(false);
  }

  function handleEquipmentTransferChange(event) {
    const { name, value } = event.target;
    const linkedEquipment = name === "maThietBi" ? equipmentList.find((item) => item.maThietBi === value) : null;

    setEquipmentTransferForm({
      ...equipmentTransferForm,
      [name]: value,
      ...(linkedEquipment
        ? {
            tenThietBi: linkedEquipment.tenThietBi,
            nguon: linkedEquipment.nguonSoHuu,
          }
        : {}),
    });
  }

  function saveEquipmentTransfers(nextItems) {
    setEquipmentTransfers(nextItems);
    localStorage.setItem(EQUIPMENT_TRANSFER_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveEquipmentTransfer(event) {
    event.preventDefault();

    if (
      !equipmentTransferForm.ngayNhan ||
      !equipmentTransferForm.hopDongCongTrinh ||
      !equipmentTransferForm.maThietBi.trim() ||
      !equipmentTransferForm.tenThietBi.trim()
    ) {
      alert("Vui lòng nhập ngày nhận, công trình, mã thiết bị và tên thiết bị.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...equipmentTransferForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveEquipmentTransfers([nextItem, ...equipmentTransfers]);
    setEquipmentTransferForm(defaultEquipmentTransferForm);
    setIsEquipmentTransferFormOpen(false);
  }

  function handleCancelEquipmentTransferForm() {
    setEquipmentTransferForm(defaultEquipmentTransferForm);
    setIsEquipmentTransferFormOpen(false);
  }

  function handleAcceptanceBatchChange(event) {
    setAcceptanceBatchForm({ ...acceptanceBatchForm, [event.target.name]: event.target.value });
  }

  function saveAcceptanceBatches(nextItems) {
    setAcceptanceBatches(nextItems);
    localStorage.setItem(ACCEPTANCE_BATCH_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveAcceptanceBatch(event) {
    event.preventDefault();

    if (!acceptanceBatchForm.soBienBan.trim() || !acceptanceBatchForm.hopDongCongTrinh) {
      alert("Vui lòng nhập số biên bản nghiệm thu và hợp đồng/công trình.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...acceptanceBatchForm,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveAcceptanceBatches([nextItem, ...acceptanceBatches]);
    setAcceptanceBatchForm(defaultAcceptanceBatchForm);
    setIsAcceptanceBatchFormOpen(false);
  }

  function handleCancelAcceptanceBatchForm() {
    setAcceptanceBatchForm(defaultAcceptanceBatchForm);
    setIsAcceptanceBatchFormOpen(false);
  }

  function handleAcceptanceChange(event) {
    const { name, value } = event.target;
    const linkedBoq = name === "maWbs" ? boqItems.find((item) => item.maWbs === value) : null;

    setAcceptanceForm({
      ...acceptanceForm,
      [name]: value,
      ...(linkedBoq
        ? {
            hopDongCongTrinh: linkedBoq.hopDongCongTrinh,
            hangMucCongViec: linkedBoq.hangMucCongViec,
            dvt: linkedBoq.dvt,
            khoiLuongHopDong: linkedBoq.khoiLuongHopDong,
            donGiaHopDong: linkedBoq.donGiaHopDong,
            khoiLuongDaThucHienLuyKe: getExecutedQuantity(constructionLogs, linkedBoq.hopDongCongTrinh, linkedBoq.maWbs),
          }
        : {}),
    });
  }

  function saveAcceptances(nextItems) {
    setAcceptances(nextItems);
    localStorage.setItem(ACCEPTANCE_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveAcceptance(event) {
    event.preventDefault();

    if (!acceptanceForm.soBienBan || !acceptanceForm.hopDongCongTrinh || !acceptanceForm.maWbs) {
      alert("Vui lòng nhập số biên bản, hợp đồng/công trình và mã WBS.");
      return;
    }

    const previousAccepted = getAcceptedQuantity(acceptances, acceptanceForm.hopDongCongTrinh, acceptanceForm.maWbs);
    const acceptedThisPeriod = toNumber(acceptanceForm.khoiLuongNghiemThuKyNay);
    const acceptedCumulative = previousAccepted + acceptedThisPeriod;
    const remaining = toNumber(acceptanceForm.khoiLuongHopDong) - acceptedCumulative;
    const executedNotAccepted = toNumber(acceptanceForm.khoiLuongDaThucHienLuyKe) - previousAccepted;
    const warning =
      acceptedThisPeriod > executedNotAccepted
        ? "Nghiệm thu vượt khối lượng thực hiện"
        : acceptedCumulative > toNumber(acceptanceForm.khoiLuongHopDong)
          ? "Nghiệm thu vượt hợp đồng"
          : "";

    const nextItem = {
      id: crypto.randomUUID(),
      ...acceptanceForm,
      khoiLuongNghiemThuLuyKe: acceptedCumulative,
      khoiLuongConLai: remaining,
      thanhTienNghiemThuKyNay: acceptedThisPeriod * toNumber(acceptanceForm.donGiaHopDong),
      canhBao: warning,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveAcceptances([nextItem, ...acceptances]);
    setAcceptanceForm(defaultAcceptanceForm);
    setIsAcceptanceFormOpen(false);
  }

  function handleCancelAcceptanceForm() {
    setAcceptanceForm(defaultAcceptanceForm);
    setIsAcceptanceFormOpen(false);
  }

  function handleInvoiceChange(event) {
    const { name, value } = event.target;
    const acceptedValue = name === "soBienBan" ? getAcceptanceValueByBatch(acceptances, value) : null;
    const batch = name === "soBienBan" ? acceptanceBatches.find((item) => item.soBienBan === value) : null;

    setInvoiceForm({
      ...invoiceForm,
      [name]: value,
      ...(acceptedValue
        ? {
            hopDongCongTrinh: acceptedValue.hopDongCongTrinh,
            giaTriNghiemThu: acceptedValue.giaTriNghiemThu,
          }
        : {}),
      ...(batch && !acceptedValue ? { hopDongCongTrinh: batch.hopDongCongTrinh } : {}),
    });
  }

  function saveInvoices(nextItems) {
    setInvoices(nextItems);
    localStorage.setItem(INVOICE_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveInvoice(event) {
    event.preventDefault();

    if (!invoiceForm.soHoaDon.trim() || !invoiceForm.hopDongCongTrinh || !invoiceForm.soBienBan) {
      alert("Vui lòng nhập số hóa đơn, hợp đồng/công trình và số biên bản nghiệm thu.");
      return;
    }

    const giaTriTruocVat = toNumber(invoiceForm.giaTriNghiemThu) * (toNumber(invoiceForm.tyLeXuatHoaDon) / 100);
    const tienVat = giaTriTruocVat * (toNumber(invoiceForm.vatPercent) / 100);
    const tongGiaTriHoaDon = giaTriTruocVat + tienVat;

    const nextItem = {
      id: crypto.randomUUID(),
      ...invoiceForm,
      giaTriTruocVat,
      tienVat,
      tongGiaTriHoaDon,
      canhBao: getInvoiceWarning({
        ...invoiceForm,
        giaTriTruocVat,
      }),
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveInvoices([nextItem, ...invoices]);
    setInvoiceForm(defaultInvoiceForm);
    setIsInvoiceFormOpen(false);
  }

  function handleCancelInvoiceForm() {
    setInvoiceForm(defaultInvoiceForm);
    setIsInvoiceFormOpen(false);
  }

  function handleCollectionChange(event) {
    const { name, value } = event.target;
    const linkedInvoice = name === "soHoaDon" ? invoices.find((item) => item.soHoaDon === value) : null;

    setCollectionForm({
      ...collectionForm,
      [name]: value,
      ...(linkedInvoice ? { hopDongCongTrinh: linkedInvoice.hopDongCongTrinh } : {}),
    });
  }

  function saveCollections(nextItems) {
    setCollections(nextItems);
    localStorage.setItem(COLLECTION_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveCollection(event) {
    event.preventDefault();

    if (!collectionForm.soPhieuThu.trim() || !collectionForm.hopDongCongTrinh || !collectionForm.soHoaDon) {
      alert("Vui lòng nhập số phiếu thu, hợp đồng/công trình và số hóa đơn.");
      return;
    }

    const invoice = invoices.find((item) => item.soHoaDon === collectionForm.soHoaDon);
    const totalCollectedForInvoice = collections
      .filter((item) => item.soHoaDon === collectionForm.soHoaDon)
      .reduce((sum, item) => sum + toNumber(item.soTienThu), 0);
    const canhBao =
      invoice && totalCollectedForInvoice + toNumber(collectionForm.soTienThu) > toNumber(invoice.tongGiaTriHoaDon)
        ? "Thu vượt hóa đơn"
        : "";
    const nextItem = {
      id: crypto.randomUUID(),
      ...collectionForm,
      canhBao,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveCollections([nextItem, ...collections]);
    setCollectionForm(defaultCollectionForm);
    setIsCollectionFormOpen(false);
  }

  function handleCancelCollectionForm() {
    setCollectionForm(defaultCollectionForm);
    setIsCollectionFormOpen(false);
  }

  function handleOtherCostChange(event) {
    setOtherCostForm({ ...otherCostForm, [event.target.name]: event.target.value });
  }

  function saveOtherCosts(nextItems) {
    setOtherCosts(nextItems);
    localStorage.setItem(OTHER_COST_STORAGE_KEY, JSON.stringify(nextItems));
  }

  function handleSaveOtherCost(event) {
    event.preventDefault();

    if (!otherCostForm.ngayPhatSinh || !otherCostForm.hopDongCongTrinh || !otherCostForm.noiDungChiPhi.trim()) {
      alert("Vui lòng nhập ngày phát sinh, hợp đồng/công trình và nội dung chi phí.");
      return;
    }

    const nextItem = {
      id: crypto.randomUUID(),
      ...otherCostForm,
      canhBao: getOtherCostWarning(otherCostForm),
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveOtherCosts([nextItem, ...otherCosts]);
    setOtherCostForm(defaultOtherCostForm);
    setIsOtherCostFormOpen(false);
  }

  function handleCancelOtherCostForm() {
    setOtherCostForm(defaultOtherCostForm);
    setIsOtherCostFormOpen(false);
  }

  async function handleExcelUpload(event, templateKey) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const workbook = await readExcelWorkbook(file);
      const result = parseExcelImport(workbook, templateKey);

      if (result.errors.length) {
        setExcelImportErrors(result.errors);
        alert(`File Excel có ${result.errors.length} lỗi dữ liệu. Vui lòng xem bảng lỗi để chỉnh sửa.`);
        return;
      }

      setExcelImportErrors([]);
      const action = chooseExcelImportAction();
      if (action === "cancel") return;

      applyExcelImportResult(templateKey, result, action);
    } catch (error) {
      setExcelImportErrors([{ row: "-", column: "File", message: error.message || "Không đọc được file Excel", suggestion: "Kiểm tra lại định dạng .xlsx" }]);
    } finally {
      event.target.value = "";
    }
  }

  function applyExcelImportResult(templateKey, result, action) {
    if (templateKey === "contracts") {
      const nextItems = mergeImportedRows(contracts, result.contracts, action, (item) => item.maHopDong);
      saveContracts(nextItems);
      alert(`Đã import thành công ${result.contracts.length} dòng dữ liệu`);
      return;
    }

    if (templateKey === "boq") {
      const nextBoqItems = mergeImportedRows(boqItems, result.boqItems, action, (item) => item.maHopDong || getContractCodeFromText(item.hopDongCongTrinh));
      const nextOtherCosts = mergeImportedRows(otherCosts, result.otherCosts, action, (item) => item.maHopDong || getContractCodeFromText(item.hopDongCongTrinh));
      const cashFlowPlans = mergeImportedRows(readStorageArray(CASH_FLOW_PLAN_STORAGE_KEY), result.cashFlowPlans, action, (item) => item.maHopDong);
      const deductionPlans = mergeImportedRows(readStorageArray(DEDUCTION_PLAN_STORAGE_KEY), result.deductionPlans || [], action, (item) => item.maHopDong);
      const laborPaymentPlans = mergeImportedRows(readStorageArray(LABOR_PAYMENT_PLAN_STORAGE_KEY), result.laborPaymentPlans, action, (item) => item.maHopDong);

      saveBoqItems(nextBoqItems);
      saveOtherCosts(nextOtherCosts);
      localStorage.setItem(CASH_FLOW_PLAN_STORAGE_KEY, JSON.stringify(cashFlowPlans));
      localStorage.setItem(DEDUCTION_PLAN_STORAGE_KEY, JSON.stringify(deductionPlans));
      localStorage.setItem(LABOR_PAYMENT_PLAN_STORAGE_KEY, JSON.stringify(laborPaymentPlans));
      alert(`Đã import thành công ${result.boqItems.length} dòng dữ liệu BOQ`);
      return;
    }

    if (templateKey === "wbs") {
      const nextItems = mergeImportedRows(wbsPlans, result.wbsPlans, action, (item) => item.maHopDong || getContractCodeFromText(item.hopDongCongTrinh));
      saveWbsPlans(nextItems);
      alert(`Đã import thành công ${result.wbsPlans.length} dòng dữ liệu`);
      return;
    }

    if (templateKey === "materials") {
      const nextItems = mergeImportedRows(materialIns, result.materialIns, action, (item) => item.maHopDong || getContractCodeFromText(item.hopDongCongTrinh));
      saveMaterialIns(nextItems);
      alert(`Đã import thành công ${result.materialIns.length} dòng dữ liệu`);
      return;
    }

    if (templateKey === "labor") {
      const nextItems = mergeImportedRows(laborTeams, result.laborTeams, action, (item) => item.maToDoi);
      saveLaborTeams(nextItems);
      alert(`Đã import thành công ${result.laborTeams.length} dòng dữ liệu`);
      return;
    }

    if (templateKey === "equipment") {
      const nextItems = mergeImportedRows(equipmentList, result.equipmentList, action, (item) => item.maThietBi);
      saveEquipmentList(nextItems);
      alert(`Đã import thành công ${result.equipmentList.length} dòng dữ liệu`);
      return;
    }

    if (templateKey === "acceptance") {
      const nextItems = mergeImportedRows(acceptances, result.acceptances, action, (item) => item.maHopDong || getContractCodeFromText(item.hopDongCongTrinh));
      saveAcceptances(nextItems);
      alert(`Đã import thành công ${result.acceptances.length} dòng dữ liệu`);
    }
  }

  function handleReportFilterChange(event) {
    setReportFilters({ ...reportFilters, [event.target.name]: event.target.value });
  }

  function handleProjectDocumentChange(event) {
    setProjectDocumentForm({ ...projectDocumentForm, [event.target.name]: event.target.value });
  }

  function handleProjectDocumentFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      setProjectDocumentForm((currentForm) => ({
        ...currentForm,
        fileName: file.name,
        fileType: file.type || "application/octet-stream",
        fileSize: file.size,
        fileData: reader.result,
      }));
    };
    reader.readAsDataURL(file);
  }

  function saveProjectDocuments(nextDocuments) {
    setProjectDocuments(nextDocuments);
    localStorage.setItem(PROJECT_DOCUMENT_STORAGE_KEY, JSON.stringify(nextDocuments));
  }

  function handleSaveProjectDocument(event) {
    event.preventDefault();

    if (!projectDocumentForm.hopDongCongTrinh || !projectDocumentForm.maHoSo.trim() || !projectDocumentForm.tenHoSo.trim()) {
      alert("Vui lòng nhập công trình, mã hồ sơ và tên hồ sơ.");
      return;
    }

    const nextDocument = {
      id: crypto.randomUUID(),
      ...projectDocumentForm,
      nhomHoSo: activeDocumentTab,
      createdAt: new Date().toLocaleString("vi-VN"),
    };

    saveProjectDocuments([nextDocument, ...projectDocuments]);
    setProjectDocumentForm({ ...defaultProjectDocumentForm, nhomHoSo: activeDocumentTab });
  }

  function handleCancelProjectDocument() {
    setProjectDocumentForm({ ...defaultProjectDocumentForm, nhomHoSo: activeDocumentTab });
  }

  function handleChangeDocumentTab(tab) {
    setActiveDocumentTab(tab);
    setProjectDocumentForm((currentForm) => ({ ...currentForm, nhomHoSo: tab }));
  }

  function handleDocumentTemplateChange(event) {
    setDocumentTemplateForm({ ...documentTemplateForm, [event.target.name]: event.target.value });
  }

  function handleDocumentTemplateFilterChange(event) {
    setDocumentTemplateFilters({ ...documentTemplateFilters, [event.target.name]: event.target.value });
  }

  function handleDocumentTemplateFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      setDocumentTemplateForm((currentForm) => ({
        ...currentForm,
        fileName: file.name,
        fileType: file.type || "application/octet-stream",
        fileSize: file.size,
        fileData: reader.result,
        loaiFile: detectDocumentTemplateFileType(file),
      }));
    };
    reader.readAsDataURL(file);
  }

  function saveDocumentTemplates(nextTemplates) {
    setDocumentTemplates(nextTemplates);
    localStorage.setItem(DOCUMENT_TEMPLATE_STORAGE_KEY, JSON.stringify(nextTemplates));
  }

  function handleSaveDocumentTemplate(event) {
    event.preventDefault();

    if (!documentTemplateForm.maTaiLieu.trim() || !documentTemplateForm.tenTaiLieu.trim()) {
      alert("Vui lòng nhập mã tài liệu và tên tài liệu.");
      return;
    }

    const nextTemplate = {
      id: editingDocumentTemplateId || crypto.randomUUID(),
      ...documentTemplateForm,
      downloadCount: toNumber(documentTemplateForm.downloadCount),
      updatedAt: new Date().toLocaleString("vi-VN"),
    };
    const nextTemplates = editingDocumentTemplateId
      ? documentTemplates.map((item) => (item.id === editingDocumentTemplateId ? nextTemplate : item))
      : [nextTemplate, ...documentTemplates];

    saveDocumentTemplates(nextTemplates);
    setDocumentTemplateForm(defaultDocumentTemplateForm);
    setEditingDocumentTemplateId(null);
  }

  function handleCancelDocumentTemplate() {
    setDocumentTemplateForm(defaultDocumentTemplateForm);
    setEditingDocumentTemplateId(null);
  }

  function handleEditDocumentTemplate(template) {
    setDocumentTemplateForm({ ...defaultDocumentTemplateForm, ...template });
    setEditingDocumentTemplateId(template.id);
  }

  function handleViewDocumentTemplateDescription(template) {
    alert(template.moTaHuongDan || "Tài liệu chưa có mô tả / hướng dẫn sử dụng.");
  }

  function handleChangeDocumentTemplateStatus(templateId, nextStatus) {
    saveDocumentTemplates(documentTemplates.map((item) => (item.id === templateId ? { ...item, trangThai: nextStatus } : item)));
  }

  function handleDownloadDocumentTemplate(template) {
    if (!template.fileData) {
      alert("Tài liệu chưa có file đính kèm.");
      return;
    }

    const link = document.createElement("a");
    link.href = template.fileData;
    link.download = template.fileName || `${template.maTaiLieu || "tai-lieu"}`;
    link.click();
    saveDocumentTemplates(
      documentTemplates.map((item) =>
        item.id === template.id ? { ...item, downloadCount: toNumber(item.downloadCount) + 1 } : item,
      ),
    );
  }

  function getActiveSystemCatalogDefinition(tabKey = activeSystemCatalogTab) {
    return systemCatalogTabs.find((tab) => tab.key === tabKey) ?? systemCatalogTabs[0];
  }

  function handleChangeSystemCatalogTab(tabKey) {
    setActiveSystemCatalogTab(tabKey);
    setSystemCatalogForm(defaultSystemCatalogForm[tabKey]);
    setSystemCatalogSearch("");
    setEditingSystemCatalogItemId(null);
    setIsSystemCatalogFormOpen(false);
  }

  function handleSystemCatalogChange(event) {
    setSystemCatalogForm({ ...systemCatalogForm, [event.target.name]: event.target.value });
  }

  function saveSystemCatalogs(nextCatalogs) {
    setSystemCatalogs(nextCatalogs);
    localStorage.setItem(SYSTEM_CATALOG_STORAGE_KEY, JSON.stringify(nextCatalogs));
    publishSystemCatalogs(nextCatalogs);
  }

  function handleOpenSystemCatalogForm() {
    setSystemCatalogForm(defaultSystemCatalogForm[activeSystemCatalogTab]);
    setEditingSystemCatalogItemId(null);
    setIsSystemCatalogFormOpen(true);
  }

  function handleEditSystemCatalogItem(item) {
    setSystemCatalogForm({ ...defaultSystemCatalogForm[activeSystemCatalogTab], ...item });
    setEditingSystemCatalogItemId(item.id);
    setIsSystemCatalogFormOpen(true);
  }

  function handleCancelSystemCatalogForm() {
    setSystemCatalogForm(defaultSystemCatalogForm[activeSystemCatalogTab]);
    setEditingSystemCatalogItemId(null);
    setIsSystemCatalogFormOpen(false);
  }

  function handleSaveSystemCatalogItem(event) {
    event.preventDefault();

    const activeDefinition = getActiveSystemCatalogDefinition();
    const missingColumn = activeDefinition.columns.find((column) => column.required && !String(systemCatalogForm[column.key] || "").trim());
    if (missingColumn) {
      alert(`Vui lòng nhập ${missingColumn.label}.`);
      return;
    }

    const nextItem = {
      id: editingSystemCatalogItemId || crypto.randomUUID(),
      ...systemCatalogForm,
      updatedAt: new Date().toLocaleString("vi-VN"),
    };
    const currentItems = systemCatalogs[activeSystemCatalogTab] || [];
    const nextItems = editingSystemCatalogItemId
      ? currentItems.map((item) => (item.id === editingSystemCatalogItemId ? nextItem : item))
      : [nextItem, ...currentItems];

    saveSystemCatalogs({ ...systemCatalogs, [activeSystemCatalogTab]: nextItems });
    setSystemCatalogForm(defaultSystemCatalogForm[activeSystemCatalogTab]);
    setEditingSystemCatalogItemId(null);
    setIsSystemCatalogFormOpen(false);
  }

  function getActiveSystemSettingsDefinition(tabKey = activeSystemSettingsTab) {
    return systemSettingsTabs.find((tab) => tab.key === tabKey) ?? systemSettingsTabs[0];
  }

  function getSystemSettingsItems(tabKey = activeSystemSettingsTab) {
    if (tabKey === "users") return systemCatalogs.users || [];
    return systemSettings[tabKey] || [];
  }

  function handleChangeSystemSettingsTab(tabKey) {
    setActiveSystemSettingsTab(tabKey);
    setSystemSettingsForm(defaultSystemCatalogForm[tabKey] || {});
    setSystemSettingsSearch("");
    setEditingSystemSettingsItemId(null);
    setIsSystemSettingsFormOpen(false);
  }

  function handleSystemSettingsChange(event) {
    setSystemSettingsForm({ ...systemSettingsForm, [event.target.name]: event.target.value });
  }

  function saveSystemSettings(nextSettings) {
    setSystemSettings(nextSettings);
    localStorage.setItem(SYSTEM_SETTINGS_STORAGE_KEY, JSON.stringify(nextSettings));
    localStorage.setItem("permissionMatrix", JSON.stringify(nextSettings.permissions || []));
    localStorage.setItem("approvalWorkflows", JSON.stringify(nextSettings.approvalWorkflows || []));
    localStorage.setItem("systemLogs", JSON.stringify(nextSettings.systemLogs || []));
    localStorage.setItem("rolesCatalog", JSON.stringify(nextSettings.roles || []));
  }

  function saveSystemSettingsTab(tabKey, nextItems) {
    if (tabKey === "users") {
      const nextCatalogs = { ...systemCatalogs, users: nextItems };
      saveSystemCatalogs(nextCatalogs);
      return;
    }

    saveSystemSettings({ ...systemSettings, [tabKey]: nextItems });
  }

  function handleOpenSystemSettingsForm() {
    setSystemSettingsForm(defaultSystemCatalogForm[activeSystemSettingsTab] || {});
    setEditingSystemSettingsItemId(null);
    setIsSystemSettingsFormOpen(true);
  }

  function handleEditSystemSettingsItem(item) {
    setSystemSettingsForm({ ...(defaultSystemCatalogForm[activeSystemSettingsTab] || {}), ...item });
    setEditingSystemSettingsItemId(item.id);
    setIsSystemSettingsFormOpen(true);
  }

  function handleCancelSystemSettingsForm() {
    setSystemSettingsForm(defaultSystemCatalogForm[activeSystemSettingsTab] || {});
    setEditingSystemSettingsItemId(null);
    setIsSystemSettingsFormOpen(false);
  }

  function handleSaveSystemSettingsItem(event) {
    event.preventDefault();

    const activeDefinition = getActiveSystemSettingsDefinition();
    const missingColumn = activeDefinition.columns?.find((column) => column.required && !String(systemSettingsForm[column.key] || "").trim());
    if (missingColumn) {
      alert(`Vui lòng nhập ${missingColumn.label}.`);
      return;
    }

    const currentItems = getSystemSettingsItems();
    const nextItem = {
      id: editingSystemSettingsItemId || crypto.randomUUID(),
      ...systemSettingsForm,
      updatedAt: new Date().toLocaleString("vi-VN"),
    };
    const normalizedItem = activeSystemSettingsTab === "users" ? normalizeUserItem(nextItem) : nextItem;
    const nextItems = editingSystemSettingsItemId
      ? currentItems.map((item) => (item.id === editingSystemSettingsItemId ? normalizedItem : item))
      : [normalizedItem, ...currentItems];

    saveSystemSettingsTab(activeSystemSettingsTab, nextItems);
    setSystemSettingsForm(defaultSystemCatalogForm[activeSystemSettingsTab] || {});
    setEditingSystemSettingsItemId(null);
    setIsSystemSettingsFormOpen(false);
  }

  function handleTogglePermission(rowId, permissionKey) {
    const nextPermissions = (systemSettings.permissions || []).map((row) =>
      row.id === rowId ? { ...row, [permissionKey]: !row[permissionKey] } : row,
    );
    saveSystemSettings({ ...systemSettings, permissions: nextPermissions });
  }

  function handleExportLocalStorage() {
    const data = {};
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      data[key] = localStorage.getItem(key);
    }

    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `thicong-backup-${new Date().toISOString().slice(0, 10)}.json`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function handleImportLocalStorage(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const importedData = JSON.parse(reader.result);
        Object.entries(importedData).forEach(([key, value]) => {
          localStorage.setItem(key, value);
        });
        alert("Nhập dữ liệu thành công. Vui lòng tải lại trang để áp dụng toàn bộ dữ liệu.");
      } catch {
        alert("File JSON không hợp lệ.");
      }
      event.target.value = "";
    };
    reader.readAsText(file);
  }

  function handleRestoreLocalStorage() {
    if (!confirm("Khôi phục dữ liệu mẫu hệ thống? Dữ liệu danh mục và cài đặt hiện tại sẽ được ghi về mặc định.")) return;
    saveSystemCatalogs(defaultSystemCatalogs);
    saveSystemSettings(defaultSystemSettings);
    setSystemSettings(defaultSystemSettings);
    setSystemSettingsForm(defaultSystemCatalogForm[activeSystemSettingsTab] || {});
  }

  function handleLoginChange(event) {
    setLoginForm({ ...loginForm, [event.target.name]: event.target.value });
    setLoginError("");
  }

  function handleLoginSubmit(event) {
    event.preventDefault();
    const users = readStoredUsers(systemCatalogs.users);
    const loginEmail = normalizeText(loginForm.email);
    const matchedUser = users.find((user) => normalizeText(user.email) === loginEmail);

    if (!matchedUser) {
      setLoginError("Email không tồn tại.");
      return;
    }

    if (isLockedUser(matchedUser)) {
      setLoginError("Tài khoản đang bị khóa.");
      return;
    }

    if ((matchedUser.matKhau || "") !== loginForm.matKhau) {
      setLoginError("Mật khẩu không đúng.");
      return;
    }

    const sessionUser = stripSensitiveUser(matchedUser);
    localStorage.setItem(CURRENT_USER_STORAGE_KEY, JSON.stringify(sessionUser));
    setCurrentUser(sessionUser);
    setActiveMenu(getPreferredMenuKey(sessionUser, activeMenu));
    setLoginForm({ email: "", matKhau: "" });
    setLoginError("");
  }

  function handleLogout() {
    localStorage.removeItem(CURRENT_USER_STORAGE_KEY);
    setCurrentUser(null);
    setIsMobileMenuOpen(false);
  }

  function handleChangeMenu(menuKey) {
    setActiveMenu(menuKey);
    setIsMobileMenuOpen(false);
  }

  if (!currentUser) {
    return (
      <LoginPage
        error={loginError}
        form={loginForm}
        onChange={handleLoginChange}
        onSubmit={handleLoginSubmit}
      />
    );
  }

  return (
    <div className={isMobileMenuOpen ? "app-shell mobile-menu-open" : "app-shell"}>
      <div
        className="mobile-drawer-backdrop"
        onClick={() => setIsMobileMenuOpen(false)}
        aria-hidden="true"
      />

      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">TC</span>
          <div>
            <strong>Quản Lý Thi Công</strong>
            <small>Công ty CP Sơn và Chất phủ Hòa Bình</small>
          </div>
        </div>

        <button className="drawer-close" type="button" onClick={() => setIsMobileMenuOpen(false)} aria-label="Đóng menu">
          X
        </button>

        <div className="drawer-user">
          <strong>{currentUser.hoTen || currentUser.email}</strong>
          <span>{formatUserRoles(currentUser.roles)}</span>
          <button className="secondary" type="button" onClick={handleLogout}>Đăng xuất</button>
        </div>

        <nav className="nav-list" aria-label="Menu chính">
          {visibleSidebarGroups.map((group) => (
            <Fragment key={group.key}>
              {group.title && <div className="nav-section-title">{group.title}</div>}
              {group.items.filter(Boolean).map((item) => (
                <button
                  key={item.key}
                  className={activeMenuKey === item.key ? "nav-item active" : "nav-item"}
                  onClick={() => handleChangeMenu(item.key)}
                  type="button"
                >
                  {item.label}
                </button>
              ))}
            </Fragment>
          ))}
        </nav>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div className="topbar-title">
            <button
              className="mobile-menu-button"
              type="button"
              onClick={() => setIsMobileMenuOpen(true)}
              aria-label="Mở menu"
              aria-expanded={isMobileMenuOpen}
            >
              ☰
            </button>
            <p className="eyebrow">Công ty CP Sơn và Chất phủ Hòa Bình</p>
            <h1>
              <span className="desktop-page-title">{activeMenuInfo.label}</span>
              <span className="mobile-app-title">Quản Lý Thi Công</span>
            </h1>
          </div>

          <div className="topbar-actions">
          <label className="project-filter">
            <span>Lọc Hợp đồng / Công trình</span>
            <select value={selectedContractId} onChange={(event) => setSelectedContractId(event.target.value)}>
              <option value={ALL_CONTRACTS}>{ALL_PROJECTS}</option>
              {contractFilterOptions.map((contract) => (
                <option key={contract.id} value={contract.id}>{contract.label}</option>
              ))}
            </select>
          </label>

          <div className="user-session">
            <div>
              <strong>{currentUser.hoTen || currentUser.email}</strong>
              <span>{formatUserRoles(currentUser.roles)}</span>
            </div>
            <button className="secondary" type="button" onClick={handleLogout}>Đăng xuất</button>
          </div>
          </div>
        </header>

        <ImportErrorPanel errors={excelImportErrors} onClear={() => setExcelImportErrors([])} />

        {!canAccessActiveMenu && (
          <Panel title="Không có quyền">
            <p className="placeholder-text">Bạn không có quyền truy cập chức năng này.</p>
          </Panel>
        )}

        {canAccessActiveMenu && activeMenuKey === "dashboard" && (
          <Dashboard
            acceptances={filteredAcceptances}
            acceptanceBatches={acceptanceBatches}
            boqItems={filteredBoqItems}
            collections={filteredCollections}
            contracts={filteredContracts}
            dailyLogs={filteredDailyLogs}
            equipmentList={equipmentList}
            equipmentTransfers={filteredEquipmentTransfers}
            invoices={filteredInvoices}
            laborTeams={laborTeams}
            materialIns={filteredMaterialIns}
            otherCosts={filteredOtherCosts}
            subContracts={filteredSubContracts}
            wbsPlans={filteredWbsPlans}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "contracts" && (
          <ContractsPage
            contracts={contracts}
            form={contractForm}
            isFormOpen={isContractFormOpen}
            onCancel={handleCancelContractForm}
            onChange={handleContractChange}
            onDownloadTemplate={() => downloadExcelTemplate("contracts")}
            onOpenForm={() => setIsContractFormOpen(true)}
            onSubmit={handleSaveContract}
            onUploadExcel={(event) => handleExcelUpload(event, "contracts")}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "boq" && (
          <BoqPage
            contracts={contracts}
            form={boqForm}
            isFormOpen={isBoqFormOpen}
            items={boqItems}
            onCancel={handleCancelBoqForm}
            onChange={handleBoqChange}
            onDownloadTemplate={() => downloadExcelTemplate("boq")}
            onOpenForm={() => setIsBoqFormOpen(true)}
            onSubmit={handleSaveBoqItem}
            onUploadExcel={(event) => handleExcelUpload(event, "boq")}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "wbs" && (
          <WbsPlanPage
            boqItems={boqItems}
            contracts={contracts}
            form={wbsForm}
            isFormOpen={isWbsFormOpen}
            plans={wbsPlans}
            onCancel={handleCancelWbsForm}
            onChange={handleWbsChange}
            onDownloadTemplate={() => downloadExcelTemplate("wbs")}
            onOpenForm={() => setIsWbsFormOpen(true)}
            onSubmit={handleSaveWbsPlan}
            onUploadExcel={(event) => handleExcelUpload(event, "wbs")}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "daily-log" && (
          <DailyLogPage
            allLogs={constructionLogs}
            contracts={contracts}
            form={logForm}
            imageMeta={dailyImageMeta}
            images={dailyLogImages}
            isFormOpen={isLogFormOpen}
            logs={filteredDailyLogs}
            onChange={handleLogChange}
            onChangeImageMeta={handleDailyImageMetaChange}
            onCancel={handleCancelLogForm}
            onDelete={handleDeleteLog}
            onRemoveImage={handleRemoveDailyImage}
            onOpenForm={() => setIsLogFormOpen(true)}
            onSubmit={handleSaveLog}
            onUploadImages={handleDailyImageUpload}
            onToggleImages={setVisibleImageLogId}
            visibleImageLogId={visibleImageLogId}
            wbsPlans={wbsPlans}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "materials" && (
          <MaterialsPage
            contracts={contracts}
            dailyLogs={constructionLogs}
            form={materialInForm}
            ins={materialIns}
            isFormOpen={isMaterialInFormOpen}
            onCancel={handleCancelMaterialInForm}
            onChange={handleMaterialInChange}
            onDownloadTemplate={() => downloadExcelTemplate("materials")}
            onOpenForm={() => setIsMaterialInFormOpen(true)}
            onSubmit={handleSaveMaterialIn}
            onUploadExcel={(event) => handleExcelUpload(event, "materials")}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "labor" && (
          <LaborPage
            dailyLogs={constructionLogs}
            form={laborTeamForm}
            isFormOpen={isLaborTeamFormOpen}
            onCancel={handleCancelLaborTeamForm}
            onChange={handleLaborTeamChange}
            onDownloadTemplate={() => downloadExcelTemplate("labor")}
            onOpenForm={() => setIsLaborTeamFormOpen(true)}
            onSubmit={handleSaveLaborTeam}
            onUploadExcel={(event) => handleExcelUpload(event, "labor")}
            teams={laborTeams}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "subcontractors" && (
          <SubcontractorsPage
            contracts={contracts}
            dailyLogs={constructionLogs}
            subContractForm={subContractForm}
            subContracts={subContracts}
            subcontractorForm={subcontractorForm}
            subcontractors={subcontractors}
            wbsPlans={wbsPlans}
            isSubContractFormOpen={isSubContractFormOpen}
            isSubcontractorFormOpen={isSubcontractorFormOpen}
            onCancelSubContract={handleCancelSubContractForm}
            onCancelSubcontractor={handleCancelSubcontractorForm}
            onChangeSubContract={handleSubContractChange}
            onChangeSubcontractor={handleSubcontractorChange}
            onOpenSubContractForm={() => setIsSubContractFormOpen(true)}
            onOpenSubcontractorForm={() => setIsSubcontractorFormOpen(true)}
            onSubmitSubContract={handleSaveSubContract}
            onSubmitSubcontractor={handleSaveSubcontractor}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "equipment" && (
          <EquipmentPage
            contracts={contracts}
            dailyLogs={constructionLogs}
            equipmentForm={equipmentForm}
            equipmentList={equipmentList}
            equipmentTransferForm={equipmentTransferForm}
            equipmentTransfers={equipmentTransfers}
            isEquipmentFormOpen={isEquipmentFormOpen}
            isEquipmentTransferFormOpen={isEquipmentTransferFormOpen}
            onCancelEquipment={handleCancelEquipmentForm}
            onCancelEquipmentTransfer={handleCancelEquipmentTransferForm}
            onChangeEquipment={handleEquipmentChange}
            onChangeEquipmentTransfer={handleEquipmentTransferChange}
            onDownloadTemplate={() => downloadExcelTemplate("equipment")}
            onOpenEquipmentForm={() => setIsEquipmentFormOpen(true)}
            onOpenEquipmentTransferForm={() => setIsEquipmentTransferFormOpen(true)}
            onSubmitEquipment={handleSaveEquipment}
            onSubmitEquipmentTransfer={handleSaveEquipmentTransfer}
            onUploadExcel={(event) => handleExcelUpload(event, "equipment")}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "acceptance" && (
          <AcceptancePage
            acceptanceBatchForm={acceptanceBatchForm}
            acceptanceBatches={acceptanceBatches}
            acceptanceForm={acceptanceForm}
            acceptances={acceptances}
            boqItems={boqItems}
            contracts={contracts}
            dailyLogs={constructionLogs}
            isAcceptanceBatchFormOpen={isAcceptanceBatchFormOpen}
            isAcceptanceFormOpen={isAcceptanceFormOpen}
            onCancelAcceptance={handleCancelAcceptanceForm}
            onCancelAcceptanceBatch={handleCancelAcceptanceBatchForm}
            onChangeAcceptance={handleAcceptanceChange}
            onChangeAcceptanceBatch={handleAcceptanceBatchChange}
            onDownloadTemplate={() => downloadExcelTemplate("acceptance")}
            onOpenAcceptanceBatchForm={() => setIsAcceptanceBatchFormOpen(true)}
            onOpenAcceptanceForm={() => setIsAcceptanceFormOpen(true)}
            onSubmitAcceptance={handleSaveAcceptance}
            onSubmitAcceptanceBatch={handleSaveAcceptanceBatch}
            onUploadExcel={(event) => handleExcelUpload(event, "acceptance")}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "invoices" && (
          <InvoicesPage
            acceptanceBatches={acceptanceBatches}
            acceptances={acceptances}
            collectionForm={collectionForm}
            collections={collections}
            contracts={contracts}
            invoiceForm={invoiceForm}
            invoices={invoices}
            isCollectionFormOpen={isCollectionFormOpen}
            isInvoiceFormOpen={isInvoiceFormOpen}
            onCancelCollection={handleCancelCollectionForm}
            onCancelInvoice={handleCancelInvoiceForm}
            onChangeCollection={handleCollectionChange}
            onChangeInvoice={handleInvoiceChange}
            onOpenCollectionForm={() => setIsCollectionFormOpen(true)}
            onOpenInvoiceForm={() => setIsInvoiceFormOpen(true)}
            onSubmitCollection={handleSaveCollection}
            onSubmitInvoice={handleSaveInvoice}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "other-costs" && (
          <OtherCostsPage
            contracts={contracts}
            form={otherCostForm}
            isFormOpen={isOtherCostFormOpen}
            onCancel={handleCancelOtherCostForm}
            onChange={handleOtherCostChange}
            onOpenForm={() => setIsOtherCostFormOpen(true)}
            onSubmit={handleSaveOtherCost}
            otherCosts={otherCosts}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "reports" && (
          <ReportsPage
            acceptances={acceptances}
            acceptanceBatches={acceptanceBatches}
            boqItems={boqItems}
            collections={collections}
            contracts={contracts}
            dailyLogs={constructionLogs}
            equipmentList={equipmentList}
            equipmentTransfers={equipmentTransfers}
            filters={reportFilters}
            invoices={invoices}
            laborTeams={laborTeams}
            materialIns={materialIns}
            onFilterChange={handleReportFilterChange}
            otherCosts={otherCosts}
            subContracts={subContracts}
            wbsPlans={wbsPlans}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "documents" && (
          <ProjectDocumentsPage
            activeTab={activeDocumentTab}
            contracts={contracts}
            documents={projectDocuments}
            form={projectDocumentForm}
            projectFilter={documentProjectFilter}
            search={documentSearch}
            onCancel={handleCancelProjectDocument}
            onChange={handleProjectDocumentChange}
            onChangeFile={handleProjectDocumentFileChange}
            onChangeProjectFilter={setDocumentProjectFilter}
            onChangeSearch={setDocumentSearch}
            onChangeTab={handleChangeDocumentTab}
            onSubmit={handleSaveProjectDocument}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "document-templates" && (
          <DocumentTemplatesPage
            editingTemplateId={editingDocumentTemplateId}
            filters={documentTemplateFilters}
            form={documentTemplateForm}
            search={documentTemplateSearch}
            templates={documentTemplates}
            onCancel={handleCancelDocumentTemplate}
            onChange={handleDocumentTemplateChange}
            onChangeFile={handleDocumentTemplateFileChange}
            onChangeFilter={handleDocumentTemplateFilterChange}
            onChangeSearch={setDocumentTemplateSearch}
            onChangeStatus={handleChangeDocumentTemplateStatus}
            onDownload={handleDownloadDocumentTemplate}
            onEdit={handleEditDocumentTemplate}
            onSubmit={handleSaveDocumentTemplate}
            onViewDescription={handleViewDocumentTemplateDescription}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "system-catalog" && (
          <SystemCatalogPage
            activeTab={activeSystemCatalogTab}
            catalogs={systemCatalogs}
            editingItemId={editingSystemCatalogItemId}
            form={systemCatalogForm}
            isFormOpen={isSystemCatalogFormOpen}
            search={systemCatalogSearch}
            tabs={systemCatalogTabs}
            onCancel={handleCancelSystemCatalogForm}
            onChange={handleSystemCatalogChange}
            onChangeSearch={setSystemCatalogSearch}
            onChangeTab={handleChangeSystemCatalogTab}
            onEdit={handleEditSystemCatalogItem}
            onOpenForm={handleOpenSystemCatalogForm}
            onSubmit={handleSaveSystemCatalogItem}
          />
        )}

        {canAccessActiveMenu && activeMenuKey === "system-settings" && (
          <SystemSettingsPage
            activeTab={activeSystemSettingsTab}
            catalogs={systemCatalogs}
            editingItemId={editingSystemSettingsItemId}
            form={systemSettingsForm}
            isFormOpen={isSystemSettingsFormOpen}
            search={systemSettingsSearch}
            settings={systemSettings}
            tabs={systemSettingsTabs}
            onCancel={handleCancelSystemSettingsForm}
            onChange={handleSystemSettingsChange}
            onChangeSearch={setSystemSettingsSearch}
            onChangeTab={handleChangeSystemSettingsTab}
            onEdit={handleEditSystemSettingsItem}
            onExport={handleExportLocalStorage}
            onImport={handleImportLocalStorage}
            onOpenForm={handleOpenSystemSettingsForm}
            onRestore={handleRestoreLocalStorage}
            onSubmit={handleSaveSystemSettingsItem}
            onTogglePermission={handleTogglePermission}
          />
        )}

        {canAccessActiveMenu &&
          activeMenuKey !== "dashboard" &&
          activeMenuKey !== "contracts" &&
          activeMenuKey !== "boq" &&
          activeMenuKey !== "wbs" &&
          activeMenuKey !== "materials" &&
          activeMenuKey !== "labor" &&
          activeMenuKey !== "subcontractors" &&
          activeMenuKey !== "equipment" &&
          activeMenuKey !== "acceptance" &&
          activeMenuKey !== "invoices" &&
          activeMenuKey !== "other-costs" &&
          activeMenuKey !== "reports" &&
          activeMenuKey !== "documents" &&
          activeMenuKey !== "document-templates" &&
          activeMenuKey !== "system-catalog" &&
          activeMenuKey !== "system-settings" &&
          activeMenuKey !== "daily-log" && (
          <PlaceholderPage title={activeMenuInfo.label} />
        )}
      </main>
    </div>
  );
}

function Dashboard({
  acceptanceBatches,
  acceptances,
  boqItems,
  collections,
  contracts,
  dailyLogs,
  equipmentList,
  equipmentTransfers,
  invoices,
  laborTeams,
  materialIns,
  otherCosts,
  subContracts,
  wbsPlans,
}) {
  const reportData = getConstructionReportsData({
    acceptanceBatches,
    acceptances,
    boqItems,
    collections,
    contracts,
    dailyLogs,
    equipmentList,
    equipmentTransfers,
    filters: { project: ALL_PROJECTS, fromDate: "", toDate: "" },
    invoices,
    laborTeams,
    materialIns,
    otherCosts,
    subContracts,
    wbsPlans,
    useProvidedData: true,
  });
  const dashboard = getDashboardSummaryData(reportData, { boqItems, contracts, wbsPlans, dailyLogs });

  return (
    <section className="dashboard-layout">
      <div className="dashboard-groups">
        <DashboardMetricGroup title="Quy mô dự án">
          <DashboardMetric label="Tổng số công trình" value={dashboard.projectCount} />
          <DashboardMetric label="Tổng giá trị hợp đồng" value={formatCurrency(dashboard.totalContractValue)} />
          <DashboardMetric label="Tổng giá trị BOQ" value={formatCurrency(dashboard.totalBoqValue)} />
          <DashboardMetric label="Tổng giá trị kế hoạch" value={formatCurrency(dashboard.totalPlanValue)} />
        </DashboardMetricGroup>

        <DashboardMetricGroup title="Tiến độ thực hiện">
          <DashboardMetric label="Giá trị thực hiện" value={formatCurrency(dashboard.totalExecutedValue)} />
          <DashboardMetric label="Giá trị nghiệm thu" value={formatCurrency(dashboard.totalAcceptanceValue)} />
          <DashboardMetric label="Doanh thu trước VAT" value={formatCurrency(dashboard.totalInvoiceBeforeVat)} />
          <DashboardMetric label="Tiền đã thu" value={formatCurrency(dashboard.totalCollections)} />
        </DashboardMetricGroup>

        <DashboardMetricGroup title="Dòng tiền">
          <DashboardMetric label="Công nợ phải thu" value={formatCurrency(dashboard.totalDebt)} />
          <DashboardMetric label="Nghiệm thu chưa xuất hóa đơn" value={formatCurrency(dashboard.acceptanceNotInvoicedValue)} />
          <DashboardMetric label="Đã xuất hóa đơn chưa thu" value={formatCurrency(dashboard.invoiceNotCollectedValue)} />
        </DashboardMetricGroup>

        <DashboardMetricGroup title="Hiệu quả">
          <DashboardMetric label="Chi phí thực tế" value={formatCurrency(dashboard.totalActualCost)} />
          <DashboardMetric label="Lợi nhuận gộp" value={formatCurrency(dashboard.totalProfit)} />
          <DashboardMetric label="Biên lợi nhuận %" value={formatPercent(dashboard.profitMargin)} />
        </DashboardMetricGroup>

        <DashboardMetricGroup title="Thi công">
          <DashboardMetric label="Khối lượng kế hoạch" value={formatNumber(dashboard.totalPlannedQuantity)} />
          <DashboardMetric label="Khối lượng thực hiện" value={formatNumber(dashboard.totalExecutedQuantity)} />
          <DashboardMetric label="Khối lượng nghiệm thu" value={formatNumber(dashboard.totalAcceptedQuantity)} />
          <DashboardMetric label="Tỷ lệ hoàn thành %" value={formatPercent(dashboard.completionRate)} />
        </DashboardMetricGroup>

        <DashboardMetricGroup title="Nguồn lực">
          <DashboardMetric label="Tổng nhân công" value={formatNumber(dashboard.totalWorkers)} />
          <DashboardMetric label="Tổng giờ công" value={formatNumber(dashboard.totalLaborHours)} />
          <DashboardMetric label="Tổng giờ thiết bị" value={formatNumber(dashboard.totalEquipmentHours)} />
          <DashboardMetric label="Tổng vật tư sử dụng" value={formatNumber(dashboard.totalMaterialUsage)} />
        </DashboardMetricGroup>
      </div>

      <Panel title="Cảnh báo điều hành">
        <div className="warning-list">
          {reportData.alerts.map((alert) => (
            <span className="status" key={alert}>{alert}</span>
          ))}
        </div>
        <EmptyState show={!reportData.alerts.length} text="Chưa có cảnh báo điều hành theo dữ liệu hiện tại." />
      </Panel>

      <div className="dashboard-lists">
        <Panel title="Tiến độ theo công trình">
          <ResponsiveTable>
            <thead>
              <tr>
                <th>Hợp đồng / Công trình</th>
                <th>Khối lượng kế hoạch</th>
                <th>Khối lượng thực hiện</th>
                <th>Tỷ lệ hoàn thành</th>
              </tr>
            </thead>
            <tbody>
              {dashboard.projectProgressRows.map((item) => (
                <tr key={item.hopDongCongTrinh}>
                  <td>{item.hopDongCongTrinh}</td>
                  <td>{formatNumber(item.khoiLuongKeHoach)}</td>
                  <td>{formatNumber(item.khoiLuongThucHien)}</td>
                  <td>{formatPercent(item.tyLeHoanThanh)}</td>
                </tr>
              ))}
            </tbody>
          </ResponsiveTable>
        </Panel>

        <Panel title="Top công trình theo doanh thu">
          <ResponsiveTable>
            <thead>
              <tr>
                <th>Hợp đồng / Công trình</th>
                <th>Doanh thu</th>
                <th>Đã thu tiền</th>
              </tr>
            </thead>
            <tbody>
              {dashboard.topRevenueProjects.map((item) => (
                <tr key={item.hopDongCongTrinh}>
                  <td>{item.hopDongCongTrinh}</td>
                  <td>{formatCurrency(item.doanhThu)}</td>
                  <td>{formatCurrency(item.daThuTien)}</td>
                </tr>
              ))}
            </tbody>
          </ResponsiveTable>
        </Panel>

        <Panel title="Top công trình theo lợi nhuận">
          <ResponsiveTable>
            <thead>
              <tr>
                <th>Hợp đồng / Công trình</th>
                <th>Lợi nhuận</th>
                <th>Biên lợi nhuận</th>
              </tr>
            </thead>
            <tbody>
              {dashboard.topProfitProjects.map((item) => (
                <tr key={item.hopDongCongTrinh}>
                  <td>{item.hopDongCongTrinh}</td>
                  <td>{formatCurrency(item.laiLoTamTinh)}</td>
                  <td>{formatPercent(item.bienLoiNhuan)}</td>
                </tr>
              ))}
            </tbody>
          </ResponsiveTable>
        </Panel>
      </div>
    </section>
  );
}

function DashboardMetricGroup({ title, children }) {
  return (
    <Panel title={title}>
      <div className="dashboard-metrics">{children}</div>
    </Panel>
  );
}

function DashboardMetric({ label, value }) {
  return (
    <div className="dashboard-metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function ExcelPanelActions({ addLabel, downloadLabel, isFormOpen, onAdd, onDownload, onUpload, uploadLabel }) {
  return (
    <div className="excel-actions">
      {!isFormOpen && <button type="button" onClick={onAdd}>{addLabel}</button>}
      <button className="secondary" type="button" onClick={onDownload}>{downloadLabel}</button>
      <label className="button-like secondary">
        {uploadLabel}
        <input accept=".xlsx,.xls" onChange={onUpload} type="file" />
      </label>
    </div>
  );
}

function ImportErrorPanel({ errors, onClear }) {
  if (!errors.length) return null;

  return (
    <Panel
      title="Lỗi dữ liệu import Excel"
      action={<button className="secondary" type="button" onClick={onClear}>Đóng</button>}
    >
      <ResponsiveTable>
        <thead>
          <tr>
            <th>Dòng lỗi</th>
            <th>Cột lỗi</th>
            <th>Nội dung lỗi</th>
            <th>Gợi ý sửa</th>
          </tr>
        </thead>
        <tbody>
          {errors.map((error, index) => (
            <tr key={`${error.row}-${error.column}-${index}`}>
              <td>{error.row}</td>
              <td>{error.column}</td>
              <td>{error.message}</td>
              <td>{error.suggestion}</td>
            </tr>
          ))}
        </tbody>
      </ResponsiveTable>
    </Panel>
  );
}

function ContractsPage({ contracts, form, isFormOpen, onCancel, onChange, onDownloadTemplate, onOpenForm, onSubmit, onUploadExcel }) {
  return (
    <>
      <Panel
        title="Danh sách hợp đồng"
        action={
          <ExcelPanelActions
            addLabel="Thêm mới"
            downloadLabel="Tải mẫu Hợp đồng"
            isFormOpen={isFormOpen}
            onAdd={onOpenForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload Hợp đồng"
          />
        }
      >
        {isFormOpen && (
          <form className="work-form" onSubmit={onSubmit}>
            <Field label="Mã hợp đồng">
              <input name="maHopDong" value={form.maHopDong} onChange={onChange} placeholder="HD-2026-001" />
            </Field>
            <Field label="Tên công trình">
              <input name="tenCongTrinh" value={form.tenCongTrinh} onChange={onChange} placeholder="Biệt thự Thảo Điền" />
            </Field>
            <Field label="Khách hàng">
              <input name="khachHang" value={form.khachHang} onChange={onChange} placeholder="Tên khách hàng/chủ đầu tư" />
            </Field>
            <Field label="Địa chỉ công trình">
              <input name="diaChiCongTrinh" value={form.diaChiCongTrinh} onChange={onChange} placeholder="Địa chỉ thi công" />
            </Field>
            <Field label="Giá trị hợp đồng">
              <input name="giaTriHopDong" type="number" min="0" value={form.giaTriHopDong} onChange={onChange} />
            </Field>
            <Field label="Ngày ký">
              <input name="ngayKy" type="date" value={form.ngayKy} onChange={onChange} />
            </Field>
            <Field label="Ngày khởi công">
              <input name="ngayKhoiCong" type="date" value={form.ngayKhoiCong} onChange={onChange} />
            </Field>
            <Field label="Ngày hoàn thành kế hoạch">
              <input name="ngayHoanThanhKeHoach" type="date" value={form.ngayHoanThanhKeHoach} onChange={onChange} />
            </Field>
            <Field label="PM phụ trách">
              <input name="pmPhuTrach" value={form.pmPhuTrach} onChange={onChange} placeholder="Quản lý dự án" />
            </Field>
            <Field label="Giám sát phụ trách">
              <input name="giamSatPhuTrach" value={form.giamSatPhuTrach} onChange={onChange} placeholder="Giám sát công trình" />
            </Field>
            <Field label="Trạng thái">
              <select name="trangThai" value={form.trangThai} onChange={onChange}>
                <option>Đang chuẩn bị</option>
                <option>Đang thi công</option>
                <option>Tạm dừng</option>
                <option>Chờ nghiệm thu</option>
                <option>Hoàn thành</option>
              </select>
            </Field>

            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Mã hợp đồng</th>
              <th>Tên công trình</th>
              <th>Khách hàng</th>
              <th>Địa chỉ</th>
              <th>Giá trị</th>
              <th>Ngày ký</th>
              <th>Khởi công</th>
              <th>Hoàn thành KH</th>
              <th>PM</th>
              <th>Giám sát</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {contracts.map((contract) => (
              <tr key={contract.id}>
                <td>{contract.maHopDong}</td>
                <td>{contract.tenCongTrinh}</td>
                <td>{contract.khachHang}</td>
                <td>{contract.diaChiCongTrinh || "-"}</td>
                <td>{formatCurrency(contract.giaTriHopDong)}</td>
                <td>{formatDate(contract.ngayKy)}</td>
                <td>{formatDate(contract.ngayKhoiCong)}</td>
                <td>{formatDate(contract.ngayHoanThanhKeHoach)}</td>
                <td>{contract.pmPhuTrach || "-"}</td>
                <td>{contract.giamSatPhuTrach || "-"}</td>
                <td>
                  <span className="status done">{contract.trangThai}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!contracts.length} text="Chưa có hợp đồng nào. Bấm Thêm mới để tạo hợp đồng đầu tiên." />
      </Panel>
    </>
  );
}

function BoqPage({ contracts, form, isFormOpen, items, onCancel, onChange, onDownloadTemplate, onOpenForm, onSubmit, onUploadExcel }) {
  const contractOptions = contracts.length
    ? contracts.map((contract) => ({
        id: contract.id,
        label: `${contract.maHopDong} - ${contract.tenCongTrinh}`,
      }))
    : projects.map((project) => ({
        id: project.ma,
        label: project.ten,
      }));

  const thanhTienHopDong = toNumber(form.khoiLuongHopDong) * toNumber(form.donGiaHopDong);
  const totals = items.reduce(
    (summary, item) => {
      summary.khoiLuong += toNumber(item.khoiLuongHopDong);
      summary.giaTriBoq += toNumber(item.thanhTienHopDong);
      summary.chiPhiKeHoach += toNumber(item.chiPhiKeHoach);
      return summary;
    },
    { khoiLuong: 0, giaTriBoq: 0, chiPhiKeHoach: 0 }
  );

  return (
    <section className="boq-layout">
      <div className="page-grid">
        <SummaryCard label="Tổng khối lượng" value={formatNumber(totals.khoiLuong)} tone="blue" />
        <SummaryCard label="Tổng giá trị BOQ" value={formatCurrency(totals.giaTriBoq)} tone="green" />
        <SummaryCard label="Tổng chi phí kế hoạch" value={formatCurrency(totals.chiPhiKeHoach)} tone="amber" />
      </div>

      <Panel
        title="Quản lý BOQ - Dự toán"
        action={
          <ExcelPanelActions
            addLabel="Thêm dòng BOQ"
            downloadLabel="Tải mẫu BOQ"
            isFormOpen={isFormOpen}
            onAdd={onOpenForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload BOQ"
          />
        }
      >
        {isFormOpen && (
          <form className="work-form" onSubmit={onSubmit}>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={form.hopDongCongTrinh} onChange={onChange}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>
                    {option.label}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Mã WBS">
              <input name="maWbs" value={form.maWbs} onChange={onChange} placeholder="WBS-01.01" />
            </Field>
            <Field label="Hạng mục công việc">
              <input name="hangMucCongViec" value={form.hangMucCongViec} onChange={onChange} placeholder="Sơn lót tường trong" />
            </Field>
            <Field label="ĐVT">
              <input name="dvt" value={form.dvt} onChange={onChange} placeholder="m2, md, bộ..." />
            </Field>
            <Field label="Khối lượng hợp đồng">
              <input name="khoiLuongHopDong" type="number" min="0" value={form.khoiLuongHopDong} onChange={onChange} />
            </Field>
            <Field label="Đơn giá hợp đồng">
              <input name="donGiaHopDong" type="number" min="0" value={form.donGiaHopDong} onChange={onChange} />
            </Field>
            <Field label="Thành tiền hợp đồng">
              <input value={formatCurrency(thanhTienHopDong)} readOnly />
            </Field>
            <Field label="Định mức vật tư">
              <input name="dinhMucVatTu" type="number" min="0" value={form.dinhMucVatTu} onChange={onChange} />
            </Field>
            <Field label="Định mức nhân công">
              <input name="dinhMucNhanCong" type="number" min="0" value={form.dinhMucNhanCong} onChange={onChange} />
            </Field>
            <Field label="Định mức MMTB">
              <input name="dinhMucMmtb" type="number" min="0" value={form.dinhMucMmtb} onChange={onChange} />
            </Field>
            <Field label="Chi phí kế hoạch">
              <input name="chiPhiKeHoach" type="number" min="0" value={form.chiPhiKeHoach} onChange={onChange} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={form.ghiChu} onChange={onChange} placeholder="Ghi chú kỹ thuật, phạm vi hoặc điều kiện nghiệm thu" />
            </Field>

            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục công việc</th>
              <th>ĐVT</th>
              <th>KL hợp đồng</th>
              <th>Đơn giá HĐ</th>
              <th>Thành tiền HĐ</th>
              <th>ĐM vật tư</th>
              <th>ĐM nhân công</th>
              <th>ĐM MMTB</th>
              <th>Chi phí kế hoạch</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {items.map((item) => (
              <tr key={item.id}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{item.dvt}</td>
                <td>{formatNumber(item.khoiLuongHopDong)}</td>
                <td>{formatCurrency(item.donGiaHopDong)}</td>
                <td>{formatCurrency(item.thanhTienHopDong)}</td>
                <td>{formatNumber(item.dinhMucVatTu)}</td>
                <td>{formatNumber(item.dinhMucNhanCong)}</td>
                <td>{formatNumber(item.dinhMucMmtb)}</td>
                <td>{formatCurrency(item.chiPhiKeHoach)}</td>
                <td>{item.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!items.length} text="Chưa có dòng BOQ nào. Bấm Thêm dòng BOQ để nhập dự toán đầu tiên." />
      </Panel>
    </section>
  );
}

function WbsPlanPage({ boqItems, contracts, form, isFormOpen, onCancel, onChange, onDownloadTemplate, onOpenForm, onSubmit, onUploadExcel, plans }) {
  const contractOptions = contracts.length
    ? contracts.map((contract) => ({
        id: contract.id,
        label: `${contract.maHopDong} - ${contract.tenCongTrinh}`,
      }))
    : projects.map((project) => ({
        id: project.ma,
        label: project.ten,
      }));
  const uniqueWbsItems = getUniqueWbsItems(boqItems.length ? boqItems : getCatalogSeed("wbsStandards"));
  const currentPlanDays = calculatePlanDays(form.ngayBatDauKeHoach, form.ngayKetThucKeHoach);
  const totals = plans.reduce(
    (summary, plan) => {
      summary.khoiLuongKeHoach += toNumber(plan.khoiLuongKeHoach);
      summary.byStatus[plan.trangThai] = (summary.byStatus[plan.trangThai] || 0) + 1;
      return summary;
    },
    { khoiLuongKeHoach: 0, byStatus: {} }
  );

  return (
    <section className="wbs-layout">
      <div className="page-grid">
        <SummaryCard label="Tổng khối lượng kế hoạch" value={formatNumber(totals.khoiLuongKeHoach)} tone="blue" />
        {WBS_STATUSES.map((status) => (
          <SummaryCard key={status} label={status} value={totals.byStatus[status] || 0} tone={getStatusTone(status)} />
        ))}
      </div>

      <Panel
        title="Quản lý kế hoạch thi công WBS"
        action={
          <ExcelPanelActions
            addLabel="Thêm kế hoạch"
            downloadLabel="Tải mẫu Tiến độ WBS"
            isFormOpen={isFormOpen}
            onAdd={onOpenForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload WBS"
          />
        }
      >
        {isFormOpen && (
          <form className="work-form" onSubmit={onSubmit}>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={form.hopDongCongTrinh} onChange={onChange}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>
                    {option.label}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Mã WBS">
              {uniqueWbsItems.length ? (
                <select name="maWbs" value={form.maWbs} onChange={onChange}>
                  <option value="">Chọn mã WBS từ BOQ</option>
                  {uniqueWbsItems.map((item) => (
                    <option key={`${item.maWbs}-${item.hangMucCongViec}`} value={item.maWbs}>
                      {item.maWbs} - {item.hangMucCongViec}
                    </option>
                  ))}
                </select>
              ) : (
                <input name="maWbs" value={form.maWbs} onChange={onChange} placeholder="WBS-01.01" />
              )}
            </Field>
            <Field label="Hạng mục công việc">
              <input name="hangMucCongViec" value={form.hangMucCongViec} onChange={onChange} placeholder="Sơn lót tường trong" />
            </Field>
            <Field label="Khối lượng kế hoạch">
              <input name="khoiLuongKeHoach" type="number" min="0" value={form.khoiLuongKeHoach} onChange={onChange} />
            </Field>
            <Field label="Ngày bắt đầu kế hoạch">
              <input name="ngayBatDauKeHoach" type="date" value={form.ngayBatDauKeHoach} onChange={onChange} />
            </Field>
            <Field label="Ngày kết thúc kế hoạch">
              <input name="ngayKetThucKeHoach" type="date" value={form.ngayKetThucKeHoach} onChange={onChange} />
            </Field>
            <Field label="Số ngày kế hoạch">
              <input value={currentPlanDays ? `${currentPlanDays} ngày` : "Chưa đủ ngày"} readOnly />
            </Field>
            <Field label="Người phụ trách">
              <input name="nguoiPhuTrach" value={form.nguoiPhuTrach} onChange={onChange} placeholder="PM hoặc giám sát phụ trách" />
            </Field>
            <Field label="Trạng thái">
              <select name="trangThai" value={form.trangThai} onChange={onChange}>
                {WBS_STATUSES.map((status) => (
                  <option key={status}>{status}</option>
                ))}
              </select>
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={form.ghiChu} onChange={onChange} placeholder="Ràng buộc tiến độ, điều kiện thi công, ghi chú phối hợp..." />
            </Field>

            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục công việc</th>
              <th>Bắt đầu KH</th>
              <th>Kết thúc KH</th>
              <th>Số ngày KH</th>
              <th>Khối lượng KH</th>
              <th>Người phụ trách</th>
              <th>Trạng thái</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {plans.map((plan) => (
              <tr key={plan.id}>
                <td>{plan.hopDongCongTrinh}</td>
                <td>{plan.maWbs}</td>
                <td>{plan.hangMucCongViec}</td>
                <td>{formatDate(plan.ngayBatDauKeHoach)}</td>
                <td>{formatDate(plan.ngayKetThucKeHoach)}</td>
                <td>{formatNumber(plan.soNgayKeHoach)} ngày</td>
                <td>{formatNumber(plan.khoiLuongKeHoach)}</td>
                <td>{plan.nguoiPhuTrach || "-"}</td>
                <td>
                  <span className={plan.trangThai === "Hoàn thành" ? "status done" : "status"}>{plan.trangThai}</span>
                </td>
                <td>{plan.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!plans.length} text="Chưa có kế hoạch WBS nào. Bấm Thêm kế hoạch để tạo công việc đầu tiên." />
      </Panel>
    </section>
  );
}

function MaterialsPage({ contracts, dailyLogs, form, ins, isFormOpen, onCancel, onChange, onDownloadTemplate, onOpenForm, onSubmit, onUploadExcel }) {
  const contractOptions = getDailyContractOptions(contracts);
  const materialCatalog = getCatalogSeed("materials");
  const materialOuts = getMaterialOutsFromDailyLogs(dailyLogs);
  const inventoryRows = getMaterialInventoryRows(ins, materialOuts);
  const totalIn = ins.reduce((sum, item) => sum + toNumber(item.soLuongNhap), 0);
  const totalOut = materialOuts.reduce((sum, item) => sum + toNumber(item.soLuongSuDung), 0);
  const negativeRows = inventoryRows.filter((item) => item.tonKho < 0).length;
  const materialCodes = new Set(
    inventoryRows.map((item) => `${item.hopDongCongTrinh}|${item.maVatTu || item.tenVatTu}|${item.dvt}`)
  );

  return (
    <section className="materials-layout">
      <div className="page-grid materials-summary">
        <SummaryCard label="Tổng số mã vật tư" value={materialCodes.size} tone="blue" />
        <SummaryCard label="Tổng lượng nhập" value={formatNumber(totalIn)} tone="green" />
        <SummaryCard label="Tổng lượng đã sử dụng" value={formatNumber(totalOut)} tone="amber" />
        <SummaryCard label="Số dòng âm kho" value={negativeRows} tone="red" />
      </div>

      <Panel
        title="Nhập vật tư về công trình"
        action={
          <ExcelPanelActions
            addLabel="Thêm phiếu nhập"
            downloadLabel="Tải mẫu Nhập vật tư"
            isFormOpen={isFormOpen}
            onAdd={onOpenForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload Nhập vật tư"
          />
        }
      >
        {isFormOpen && (
          <form className="work-form" onSubmit={onSubmit}>
            <Field label="Ngày nhập">
              <input name="ngayNhap" type="date" value={form.ngayNhap} onChange={onChange} />
            </Field>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={form.hopDongCongTrinh} onChange={onChange}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>
                    {option.label}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Mã vật tư">
              <input name="maVatTu" value={form.maVatTu} onChange={onChange} list="material-catalog-codes" placeholder="VT-001" />
              <datalist id="material-catalog-codes">
                {materialCatalog.map((item) => (
                  <option key={item.id} value={item.maVatTu}>{item.tenVatTu}</option>
                ))}
              </datalist>
            </Field>
            <Field label="Tên vật tư">
              <input name="tenVatTu" value={form.tenVatTu} onChange={onChange} list="material-catalog-names" placeholder="Sơn lót, bột trét..." />
              <datalist id="material-catalog-names">
                {materialCatalog.map((item) => (
                  <option key={item.id} value={item.tenVatTu}>{item.maVatTu}</option>
                ))}
              </datalist>
            </Field>
            <Field label="ĐVT">
              <input name="dvt" value={form.dvt} onChange={onChange} />
            </Field>
            <Field label="Số lượng nhập">
              <input name="soLuongNhap" type="number" min="0" value={form.soLuongNhap} onChange={onChange} />
            </Field>
            <Field label="Nguồn cấp">
              <select name="nguonCap" value={form.nguonCap} onChange={onChange}>
                <option>Nhà máy</option>
                <option>Kho công ty</option>
                <option>Điều chuyển</option>
              </select>
            </Field>
            <Field label="Người giao">
              <input name="nguoiGiao" value={form.nguoiGiao} onChange={onChange} />
            </Field>
            <Field label="Người nhận">
              <input name="nguoiNhan" value={form.nguoiNhan} onChange={onChange} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={form.ghiChu} onChange={onChange} />
            </Field>

            <div className="form-actions">
              <button type="submit">Lưu phiếu nhập</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}
        <EmptyState show={!isFormOpen} text="Bấm Thêm phiếu nhập để ghi nhận vật tư nhập về công trình." />

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày nhập</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã vật tư</th>
              <th>Tên vật tư</th>
              <th>ĐVT</th>
              <th>Số lượng nhập</th>
              <th>Nguồn cấp</th>
              <th>Người giao</th>
              <th>Người nhận</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {ins.map((item) => (
              <tr key={item.id}>
                <td>{formatDate(item.ngayNhap)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maVatTu}</td>
                <td>{item.tenVatTu}</td>
                <td>{item.dvt}</td>
                <td>{formatNumber(item.soLuongNhap)}</td>
                <td>{item.nguonCap}</td>
                <td>{item.nguoiGiao || "-"}</td>
                <td>{item.nguoiNhan || "-"}</td>
                <td>{item.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!ins.length} text="Chưa có phiếu nhập vật tư nào." />
      </Panel>

      <Panel title="Xuất dùng vật tư từ Nhật ký thi công">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Tên vật tư</th>
              <th>ĐVT</th>
              <th>Số lượng sử dụng</th>
              <th>Giám sát</th>
            </tr>
          </thead>
          <tbody>
            {materialOuts.map((item) => (
              <tr key={item.id}>
                <td>{formatDate(item.ngay)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs || "-"}</td>
                <td>{item.tenVatTu}</td>
                <td>{item.dvt}</td>
                <td>{formatNumber(item.soLuongSuDung)}</td>
                <td>{item.giamSat || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!materialOuts.length} text="Chưa có dữ liệu vật tư xuất dùng từ Nhật ký thi công." />
      </Panel>

      <Panel title="Tồn kho vật tư công trình">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã vật tư</th>
              <th>Tên vật tư</th>
              <th>ĐVT</th>
              <th>Tổng nhập</th>
              <th>Tổng sử dụng</th>
              <th>Tồn kho</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {inventoryRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maVatTu || "-"}</td>
                <td>{item.tenVatTu}</td>
                <td>{item.dvt}</td>
                <td>{formatNumber(item.tongNhap)}</td>
                <td>{formatNumber(item.tongSuDung)}</td>
                <td>{formatNumber(item.tonKho)}</td>
                <td>
                  <span className={item.trangThai === "Đủ" ? "status done" : "status"}>{item.trangThai}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!inventoryRows.length} text="Chưa có dữ liệu để tính tồn kho vật tư." />
      </Panel>
    </section>
  );
}

function OtherCostsPage({ contracts, form, isFormOpen, onCancel, onChange, onOpenForm, onSubmit, otherCosts }) {
  const contractOptions = getDailyContractOptions(contracts);
  const reportRows = getOtherCostReportRows(otherCosts);
  const approvedTotal = otherCosts
    .filter((item) => item.trangThai === "Đã duyệt" && toNumber(item.soTien) > 0)
    .reduce((sum, item) => sum + toNumber(item.soTien), 0);
  const pendingTotal = otherCosts
    .filter((item) => item.trangThai === "Chờ duyệt" && toNumber(item.soTien) > 0)
    .reduce((sum, item) => sum + toNumber(item.soTien), 0);
  const highestProject = getHighestOtherCostProject(otherCosts);

  return (
    <section className="other-costs-layout">
      <div className="page-grid other-costs-summary">
        <SummaryCard label="Tổng chi phí khác đã duyệt" value={formatCurrency(approvedTotal)} tone="blue" />
        <SummaryCard label="Chi phí chờ duyệt" value={formatCurrency(pendingTotal)} tone="amber" />
        <SummaryCard label="Số khoản chi" value={otherCosts.length} tone="green" />
        <SummaryCard label="Công trình chi phí cao nhất" value={highestProject} tone="red" />
      </div>

      <Panel
        title="Nhập chi phí khác"
        action={!isFormOpen && <button type="button" onClick={onOpenForm}>Thêm chi phí</button>}
      >
        {isFormOpen && (
          <form className="work-form" onSubmit={onSubmit}>
            <Field label="Ngày phát sinh">
              <input name="ngayPhatSinh" type="date" value={form.ngayPhatSinh} onChange={onChange} />
            </Field>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={form.hopDongCongTrinh} onChange={onChange}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>{option.label}</option>
                ))}
              </select>
            </Field>
            <Field label="Nhóm chi phí">
              <select name="nhomChiPhi" value={form.nhomChiPhi} onChange={onChange}>
                <option>Vận chuyển</option>
                <option>Xăng xe</option>
                <option>Ăn ở</option>
                <option>Bốc xếp</option>
                <option>Vệ sinh công trình</option>
                <option>Vật tư phụ</option>
                <option>Phát sinh xử lý hiện trường</option>
                <option>Khác</option>
              </select>
            </Field>
            <Field label="Nội dung chi phí">
              <input name="noiDungChiPhi" value={form.noiDungChiPhi} onChange={onChange} />
            </Field>
            <Field label="Số tiền">
              <input name="soTien" type="number" value={form.soTien} onChange={onChange} />
            </Field>
            <Field label="Người đề nghị / Người chi">
              <input name="nguoiDeNghi" value={form.nguoiDeNghi} onChange={onChange} />
            </Field>
            <Field label="Người duyệt">
              <input name="nguoiDuyet" value={form.nguoiDuyet} onChange={onChange} />
            </Field>
            <Field label="Trạng thái">
              <select name="trangThai" value={form.trangThai} onChange={onChange}>
                <option>Chờ duyệt</option>
                <option>Đã duyệt</option>
                <option>Từ chối</option>
              </select>
            </Field>
            <Field label="Hình thức thanh toán">
              <select name="hinhThucThanhToan" value={form.hinhThucThanhToan} onChange={onChange}>
                <option>Tiền mặt</option>
                <option>Chuyển khoản</option>
                <option>Tạm ứng</option>
                <option>Khác</option>
              </select>
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={form.ghiChu} onChange={onChange} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu chi phí</button>
              <button type="button" className="secondary" onClick={onCancel}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isFormOpen} text="Bấm Thêm chi phí để ghi nhận chi phí phát sinh ngoài vật tư, nhân công và thiết bị." />
      </Panel>

      <Panel title="Danh sách chi phí khác">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày</th>
              <th>Hợp đồng / Công trình</th>
              <th>Nhóm chi phí</th>
              <th>Nội dung</th>
              <th>Số tiền</th>
              <th>Người chi</th>
              <th>Người duyệt</th>
              <th>Trạng thái</th>
              <th>Hình thức thanh toán</th>
              <th>Ghi chú</th>
              <th>Cảnh báo</th>
            </tr>
          </thead>
          <tbody>
            {otherCosts.map((item) => {
              const warning = getOtherCostWarning(item);

              return (
                <tr key={item.id}>
                  <td>{formatDate(item.ngayPhatSinh)}</td>
                  <td>{item.hopDongCongTrinh}</td>
                  <td>{item.nhomChiPhi}</td>
                  <td>{item.noiDungChiPhi}</td>
                  <td>{formatCurrency(item.soTien)}</td>
                  <td>{item.nguoiDeNghi || "-"}</td>
                  <td>{item.nguoiDuyet || "-"}</td>
                  <td>{item.trangThai}</td>
                  <td>{item.hinhThucThanhToan}</td>
                  <td>{item.ghiChu || "-"}</td>
                  <td><span className={warning === "Hợp lệ" ? "status done" : "status"}>{warning}</span></td>
                </tr>
              );
            })}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!otherCosts.length} text="Chưa có chi phí khác nào." />
      </Panel>

      <Panel title="Báo cáo tổng hợp chi phí khác theo công trình và nhóm chi phí">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Nhóm chi phí</th>
              <th>Tổng số tiền</th>
              <th>Số khoản chi</th>
              <th>Trạng thái tổng hợp</th>
            </tr>
          </thead>
          <tbody>
            {reportRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.nhomChiPhi}</td>
                <td>{formatCurrency(item.tongSoTien)}</td>
                <td>{item.soKhoanChi}</td>
                <td><span className={item.trangThaiTongHop === "Đã duyệt" ? "status done" : "status"}>{item.trangThaiTongHop}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!reportRows.length} text="Chưa có dữ liệu để tổng hợp chi phí khác." />
      </Panel>
    </section>
  );
}

function InvoicesPage({
  acceptanceBatches,
  acceptances,
  collectionForm,
  collections,
  contracts,
  invoiceForm,
  invoices,
  isCollectionFormOpen,
  isInvoiceFormOpen,
  onCancelCollection,
  onCancelInvoice,
  onChangeCollection,
  onChangeInvoice,
  onOpenCollectionForm,
  onOpenInvoiceForm,
  onSubmitCollection,
  onSubmitInvoice,
}) {
  const contractOptions = getDailyContractOptions(contracts);
  const debtRows = getDebtRows(contracts, acceptances, invoices, collections);
  const totalContractValue = debtRows.reduce((sum, item) => sum + item.giaTriHopDong, 0);
  const totalAcceptanceValue = debtRows.reduce((sum, item) => sum + item.giaTriNghiemThu, 0);
  const totalInvoiceValue = debtRows.reduce((sum, item) => sum + item.giaTriDaXuatHoaDon, 0);
  const totalCollectionValue = debtRows.reduce((sum, item) => sum + item.giaTriDaThuTien, 0);
  const totalDebt = debtRows.reduce((sum, item) => sum + item.congNoConPhaiThu, 0);
  const giaTriTruocVat = toNumber(invoiceForm.giaTriNghiemThu) * (toNumber(invoiceForm.tyLeXuatHoaDon) / 100);
  const tienVat = giaTriTruocVat * (toNumber(invoiceForm.vatPercent) / 100);
  const tongGiaTriHoaDon = giaTriTruocVat + tienVat;

  return (
    <section className="invoices-layout">
      <div className="page-grid invoices-summary">
        <SummaryCard label="Tổng giá trị hợp đồng" value={formatCurrency(totalContractValue)} tone="blue" />
        <SummaryCard label="Tổng nghiệm thu" value={formatCurrency(totalAcceptanceValue)} tone="green" />
        <SummaryCard label="Tổng đã xuất hóa đơn" value={formatCurrency(totalInvoiceValue)} tone="amber" />
        <SummaryCard label="Tổng đã thu tiền" value={formatCurrency(totalCollectionValue)} tone="green" />
        <SummaryCard label="Tổng công nợ" value={formatCurrency(totalDebt)} tone="red" />
      </div>

      <Panel
        title="Xuất hóa đơn"
        action={!isInvoiceFormOpen && <button type="button" onClick={onOpenInvoiceForm}>Thêm hóa đơn</button>}
      >
        {isInvoiceFormOpen && (
          <form className="work-form" onSubmit={onSubmitInvoice}>
            <Field label="Số hóa đơn">
              <input name="soHoaDon" value={invoiceForm.soHoaDon} onChange={onChangeInvoice} />
            </Field>
            <Field label="Ngày xuất hóa đơn">
              <input name="ngayXuatHoaDon" type="date" value={invoiceForm.ngayXuatHoaDon} onChange={onChangeInvoice} />
            </Field>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={invoiceForm.hopDongCongTrinh} onChange={onChangeInvoice}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>{option.label}</option>
                ))}
              </select>
            </Field>
            <Field label="Số biên bản nghiệm thu">
              <select name="soBienBan" value={invoiceForm.soBienBan} onChange={onChangeInvoice}>
                <option value="">Chọn biên bản nghiệm thu</option>
                {acceptanceBatches.map((batch) => (
                  <option key={batch.id} value={batch.soBienBan}>{batch.soBienBan}</option>
                ))}
              </select>
            </Field>
            <Field label="Giá trị nghiệm thu">
              <input name="giaTriNghiemThu" type="number" min="0" value={invoiceForm.giaTriNghiemThu} onChange={onChangeInvoice} />
            </Field>
            <Field label="Tỷ lệ xuất hóa đơn %">
              <input name="tyLeXuatHoaDon" type="number" min="0" value={invoiceForm.tyLeXuatHoaDon} onChange={onChangeInvoice} />
            </Field>
            <Field label="Giá trị trước VAT">
              <input value={formatCurrency(giaTriTruocVat)} readOnly />
            </Field>
            <Field label="VAT %">
              <input name="vatPercent" type="number" min="0" value={invoiceForm.vatPercent} onChange={onChangeInvoice} />
            </Field>
            <Field label="Tiền VAT">
              <input value={formatCurrency(tienVat)} readOnly />
            </Field>
            <Field label="Tổng giá trị hóa đơn">
              <input value={formatCurrency(tongGiaTriHoaDon)} readOnly />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={invoiceForm.ghiChu} onChange={onChangeInvoice} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelInvoice}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isInvoiceFormOpen} text="Bấm Thêm hóa đơn để ghi nhận hóa đơn xuất cho nghiệm thu." />
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Số hóa đơn</th>
              <th>Ngày xuất</th>
              <th>Hợp đồng / Công trình</th>
              <th>Số biên bản</th>
              <th>Giá trị nghiệm thu</th>
              <th>Tỷ lệ %</th>
              <th>Trước VAT</th>
              <th>VAT</th>
              <th>Tổng hóa đơn</th>
              <th>Cảnh báo</th>
            </tr>
          </thead>
          <tbody>
            {invoices.map((item) => {
              const warning = getInvoiceWarning(item);

              return (
                <tr key={item.id}>
                  <td>{item.soHoaDon}</td>
                  <td>{formatDate(item.ngayXuatHoaDon)}</td>
                  <td>{item.hopDongCongTrinh}</td>
                  <td>{item.soBienBan}</td>
                  <td>{formatCurrency(item.giaTriNghiemThu)}</td>
                  <td>{formatNumber(item.tyLeXuatHoaDon)}%</td>
                  <td>{formatCurrency(item.giaTriTruocVat)}</td>
                  <td>{formatCurrency(item.tienVat)}</td>
                  <td>{formatCurrency(item.tongGiaTriHoaDon)}</td>
                  <td><span className={warning === "Hợp lệ" ? "status done" : "status"}>{warning}</span></td>
                </tr>
              );
            })}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!invoices.length} text="Chưa có hóa đơn nào." />
      </Panel>

      <Panel
        title="Thu tiền"
        action={!isCollectionFormOpen && <button type="button" onClick={onOpenCollectionForm}>Thêm phiếu thu</button>}
      >
        {isCollectionFormOpen && (
          <form className="work-form" onSubmit={onSubmitCollection}>
            <Field label="Số phiếu thu">
              <input name="soPhieuThu" value={collectionForm.soPhieuThu} onChange={onChangeCollection} />
            </Field>
            <Field label="Ngày thu tiền">
              <input name="ngayThuTien" type="date" value={collectionForm.ngayThuTien} onChange={onChangeCollection} />
            </Field>
            <Field label="Hợp đồng / Công trình">
              <input name="hopDongCongTrinh" value={collectionForm.hopDongCongTrinh} onChange={onChangeCollection} />
            </Field>
            <Field label="Số hóa đơn">
              <select name="soHoaDon" value={collectionForm.soHoaDon} onChange={onChangeCollection}>
                <option value="">Chọn hóa đơn</option>
                {invoices.map((invoice) => (
                  <option key={invoice.id} value={invoice.soHoaDon}>{invoice.soHoaDon}</option>
                ))}
              </select>
            </Field>
            <Field label="Số tiền thu">
              <input name="soTienThu" type="number" min="0" value={collectionForm.soTienThu} onChange={onChangeCollection} />
            </Field>
            <Field label="Hình thức thu">
              <select name="hinhThucThu" value={collectionForm.hinhThucThu} onChange={onChangeCollection}>
                <option>Chuyển khoản</option>
                <option>Tiền mặt</option>
              </select>
            </Field>
            <Field label="Người nộp">
              <input name="nguoiNop" value={collectionForm.nguoiNop} onChange={onChangeCollection} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={collectionForm.ghiChu} onChange={onChangeCollection} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelCollection}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isCollectionFormOpen} text="Bấm Thêm phiếu thu để ghi nhận khoản thu tiền." />
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Số phiếu thu</th>
              <th>Ngày thu</th>
              <th>Hợp đồng / Công trình</th>
              <th>Số hóa đơn</th>
              <th>Số tiền thu</th>
              <th>Hình thức</th>
              <th>Người nộp</th>
              <th>Cảnh báo</th>
            </tr>
          </thead>
          <tbody>
            {collections.map((item) => (
              <tr key={item.id}>
                <td>{item.soPhieuThu}</td>
                <td>{formatDate(item.ngayThuTien)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.soHoaDon}</td>
                <td>{formatCurrency(item.soTienThu)}</td>
                <td>{item.hinhThucThu}</td>
                <td>{item.nguoiNop || "-"}</td>
                <td>{item.canhBao ? <span className="status">{item.canhBao}</span> : <span className="status done">Hợp lệ</span>}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!collections.length} text="Chưa có phiếu thu nào." />
      </Panel>

      <Panel title="Theo dõi công nợ">
        <DebtTable rows={debtRows} />
      </Panel>

      <Panel title="Báo cáo doanh thu theo công trình">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Giá trị hợp đồng</th>
              <th>Giá trị nghiệm thu</th>
              <th>Đã xuất hóa đơn</th>
              <th>Đã thu tiền</th>
              <th>Công nợ</th>
              <th>Tỷ lệ nghiệm thu / HĐ</th>
              <th>Tỷ lệ thu tiền / HĐ</th>
            </tr>
          </thead>
          <tbody>
            {debtRows.map((item) => (
              <tr key={`revenue-${item.hopDongCongTrinh}`}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{formatCurrency(item.giaTriHopDong)}</td>
                <td>{formatCurrency(item.giaTriNghiemThu)}</td>
                <td>{formatCurrency(item.giaTriDaXuatHoaDon)}</td>
                <td>{formatCurrency(item.giaTriDaThuTien)}</td>
                <td>{formatCurrency(item.congNoConPhaiThu)}</td>
                <td>{formatNumber(item.tyLeNghiemThuHopDong)}%</td>
                <td>{formatNumber(item.tyLeThuTienHoaDon)}%</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!debtRows.length} text="Chưa có dữ liệu doanh thu để báo cáo." />
      </Panel>
    </section>
  );
}

function DebtTable({ rows }) {
  return (
    <>
      <ResponsiveTable>
        <thead>
          <tr>
            <th>Hợp đồng / Công trình</th>
            <th>Giá trị hợp đồng</th>
            <th>Giá trị nghiệm thu</th>
            <th>Đã xuất hóa đơn</th>
            <th>Đã thu tiền</th>
            <th>Công nợ còn phải thu</th>
            <th>Tỷ lệ thu tiền</th>
            <th>Trạng thái</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((item) => (
            <tr key={item.hopDongCongTrinh}>
              <td>{item.hopDongCongTrinh}</td>
              <td>{formatCurrency(item.giaTriHopDong)}</td>
              <td>{formatCurrency(item.giaTriNghiemThu)}</td>
              <td>{formatCurrency(item.giaTriDaXuatHoaDon)}</td>
              <td>{formatCurrency(item.giaTriDaThuTien)}</td>
              <td>{formatCurrency(item.congNoConPhaiThu)}</td>
              <td>{formatNumber(item.tyLeThuTienHoaDon)}%</td>
              <td><span className={item.trangThai === "Đã thu đủ" ? "status done" : "status"}>{item.trangThai}</span></td>
            </tr>
          ))}
        </tbody>
      </ResponsiveTable>
      <EmptyState show={!rows.length} text="Chưa có dữ liệu công nợ." />
    </>
  );
}

function AcceptancePage({
  acceptanceBatchForm,
  acceptanceBatches,
  acceptanceForm,
  acceptances,
  boqItems,
  contracts,
  dailyLogs,
  isAcceptanceBatchFormOpen,
  isAcceptanceFormOpen,
  onCancelAcceptance,
  onCancelAcceptanceBatch,
  onChangeAcceptance,
  onChangeAcceptanceBatch,
  onDownloadTemplate,
  onOpenAcceptanceBatchForm,
  onOpenAcceptanceForm,
  onSubmitAcceptance,
  onSubmitAcceptanceBatch,
  onUploadExcel,
}) {
  const contractOptions = getDailyContractOptions(contracts);
  const reconciliationRows = getAcceptanceReconciliationRows(boqItems, dailyLogs, acceptances);
  const totalExecuted = reconciliationRows.reduce((sum, item) => sum + item.khoiLuongThucHien, 0);
  const totalAccepted = reconciliationRows.reduce((sum, item) => sum + item.khoiLuongNghiemThu, 0);
  const totalValue = acceptances.reduce((sum, item) => sum + toNumber(item.thanhTienNghiemThuKyNay), 0);
  const notAccepted = totalExecuted - totalAccepted;
  const previousAccepted = getAcceptedQuantity(acceptances, acceptanceForm.hopDongCongTrinh, acceptanceForm.maWbs);
  const acceptedThisPeriod = toNumber(acceptanceForm.khoiLuongNghiemThuKyNay);
  const acceptedCumulativePreview = previousAccepted + acceptedThisPeriod;
  const remainingPreview = toNumber(acceptanceForm.khoiLuongHopDong) - acceptedCumulativePreview;
  const acceptanceValuePreview = acceptedThisPeriod * toNumber(acceptanceForm.donGiaHopDong);

  return (
    <section className="acceptance-layout">
      <div className="page-grid acceptance-summary">
        <SummaryCard label="Tổng số đợt nghiệm thu" value={acceptanceBatches.length} tone="blue" />
        <SummaryCard label="Tổng khối lượng thực hiện" value={formatNumber(totalExecuted)} tone="green" />
        <SummaryCard label="Tổng khối lượng nghiệm thu" value={formatNumber(totalAccepted)} tone="amber" />
        <SummaryCard label="Tổng giá trị nghiệm thu" value={formatCurrency(totalValue)} tone="red" />
        <SummaryCard label="Khối lượng chưa nghiệm thu" value={formatNumber(notAccepted)} tone="blue" />
      </div>

      <Panel
        title="Tạo đợt nghiệm thu"
        action={!isAcceptanceBatchFormOpen && <button type="button" onClick={onOpenAcceptanceBatchForm}>Thêm đợt nghiệm thu</button>}
      >
        {isAcceptanceBatchFormOpen && (
          <form className="work-form" onSubmit={onSubmitAcceptanceBatch}>
            <Field label="Số biên bản nghiệm thu">
              <input name="soBienBan" value={acceptanceBatchForm.soBienBan} onChange={onChangeAcceptanceBatch} />
            </Field>
            <Field label="Ngày nghiệm thu">
              <input name="ngayNghiemThu" type="date" value={acceptanceBatchForm.ngayNghiemThu} onChange={onChangeAcceptanceBatch} />
            </Field>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={acceptanceBatchForm.hopDongCongTrinh} onChange={onChangeAcceptanceBatch}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>{option.label}</option>
                ))}
              </select>
            </Field>
            <Field label="Khách hàng / Chủ đầu tư">
              <input name="khachHangChuDauTu" value={acceptanceBatchForm.khachHangChuDauTu} onChange={onChangeAcceptanceBatch} />
            </Field>
            <Field label="Người đại diện bên A">
              <input name="daiDienBenA" value={acceptanceBatchForm.daiDienBenA} onChange={onChangeAcceptanceBatch} />
            </Field>
            <Field label="Người đại diện bên B">
              <input name="daiDienBenB" value={acceptanceBatchForm.daiDienBenB} onChange={onChangeAcceptanceBatch} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={acceptanceBatchForm.ghiChu} onChange={onChangeAcceptanceBatch} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelAcceptanceBatch}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isAcceptanceBatchFormOpen} text="Bấm Thêm đợt nghiệm thu để tạo biên bản nghiệm thu." />
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Số biên bản</th>
              <th>Ngày nghiệm thu</th>
              <th>Hợp đồng / Công trình</th>
              <th>Khách hàng / Chủ đầu tư</th>
              <th>Đại diện bên A</th>
              <th>Đại diện bên B</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {acceptanceBatches.map((item) => (
              <tr key={item.id}>
                <td>{item.soBienBan}</td>
                <td>{formatDate(item.ngayNghiemThu)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.khachHangChuDauTu || "-"}</td>
                <td>{item.daiDienBenA || "-"}</td>
                <td>{item.daiDienBenB || "-"}</td>
                <td>{item.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!acceptanceBatches.length} text="Chưa có đợt nghiệm thu nào." />
      </Panel>

      <Panel
        title="Chi tiết khối lượng nghiệm thu theo WBS"
        action={
          <ExcelPanelActions
            addLabel="Thêm dòng nghiệm thu"
            downloadLabel="Tải mẫu Nghiệm thu"
            isFormOpen={isAcceptanceFormOpen}
            onAdd={onOpenAcceptanceForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload Nghiệm thu"
          />
        }
      >
        {isAcceptanceFormOpen && (
          <form className="work-form" onSubmit={onSubmitAcceptance}>
            <Field label="Số biên bản nghiệm thu">
              <select name="soBienBan" value={acceptanceForm.soBienBan} onChange={onChangeAcceptance}>
                <option value="">Chọn số biên bản</option>
                {acceptanceBatches.map((batch) => (
                  <option key={batch.id} value={batch.soBienBan}>{batch.soBienBan}</option>
                ))}
              </select>
            </Field>
            <Field label="Hợp đồng / Công trình">
              <input name="hopDongCongTrinh" value={acceptanceForm.hopDongCongTrinh} onChange={onChangeAcceptance} />
            </Field>
            <Field label="Mã WBS">
              <select name="maWbs" value={acceptanceForm.maWbs} onChange={onChangeAcceptance}>
                <option value="">Chọn mã WBS từ BOQ</option>
                {boqItems.map((item) => (
                  <option key={item.id} value={item.maWbs}>{item.maWbs} - {item.hangMucCongViec}</option>
                ))}
              </select>
            </Field>
            <Field label="Hạng mục công việc">
              <input name="hangMucCongViec" value={acceptanceForm.hangMucCongViec} onChange={onChangeAcceptance} />
            </Field>
            <Field label="ĐVT">
              <input name="dvt" value={acceptanceForm.dvt} onChange={onChangeAcceptance} />
            </Field>
            <Field label="Khối lượng hợp đồng">
              <input name="khoiLuongHopDong" type="number" min="0" value={acceptanceForm.khoiLuongHopDong} onChange={onChangeAcceptance} />
            </Field>
            <Field label="Khối lượng đã thực hiện lũy kế">
              <input name="khoiLuongDaThucHienLuyKe" type="number" min="0" value={acceptanceForm.khoiLuongDaThucHienLuyKe} onChange={onChangeAcceptance} />
            </Field>
            <Field label="Khối lượng nghiệm thu kỳ này">
              <input name="khoiLuongNghiemThuKyNay" type="number" min="0" value={acceptanceForm.khoiLuongNghiemThuKyNay} onChange={onChangeAcceptance} />
            </Field>
            <Field label="Khối lượng nghiệm thu lũy kế">
              <input value={formatNumber(acceptedCumulativePreview)} readOnly />
            </Field>
            <Field label="Khối lượng còn lại">
              <input value={formatNumber(remainingPreview)} readOnly />
            </Field>
            <Field label="Đơn giá hợp đồng">
              <input name="donGiaHopDong" type="number" min="0" value={acceptanceForm.donGiaHopDong} onChange={onChangeAcceptance} />
            </Field>
            <Field label="Thành tiền nghiệm thu kỳ này">
              <input value={formatCurrency(acceptanceValuePreview)} readOnly />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={acceptanceForm.ghiChu} onChange={onChangeAcceptance} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelAcceptance}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isAcceptanceFormOpen} text="Bấm Thêm dòng nghiệm thu để ghi nhận khối lượng nghiệm thu theo WBS." />
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Số biên bản</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>ĐVT</th>
              <th>KL hợp đồng</th>
              <th>KL thực hiện lũy kế</th>
              <th>KL nghiệm thu kỳ này</th>
              <th>KL nghiệm thu lũy kế</th>
              <th>KL còn lại</th>
              <th>Đơn giá</th>
              <th>Thành tiền kỳ này</th>
              <th>Cảnh báo</th>
            </tr>
          </thead>
          <tbody>
            {acceptances.map((item) => (
              <tr key={item.id}>
                <td>{item.soBienBan}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{item.dvt}</td>
                <td>{formatNumber(item.khoiLuongHopDong)}</td>
                <td>{formatNumber(item.khoiLuongDaThucHienLuyKe)}</td>
                <td>{formatNumber(item.khoiLuongNghiemThuKyNay)}</td>
                <td>{formatNumber(item.khoiLuongNghiemThuLuyKe)}</td>
                <td>{formatNumber(item.khoiLuongConLai)}</td>
                <td>{formatCurrency(item.donGiaHopDong)}</td>
                <td>{formatCurrency(item.thanhTienNghiemThuKyNay)}</td>
                <td>{item.canhBao ? <span className="status">{item.canhBao}</span> : <span className="status done">Hợp lệ</span>}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!acceptances.length} text="Chưa có chi tiết nghiệm thu nào." />
      </Panel>

      <Panel title="Đối chiếu khối lượng thực hiện - nghiệm thu - còn lại">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>KL hợp đồng</th>
              <th>KL thực hiện</th>
              <th>KL nghiệm thu</th>
              <th>KL chưa nghiệm thu</th>
              <th>KL còn lại HĐ</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reconciliationRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{formatNumber(item.khoiLuongHopDong)}</td>
                <td>{formatNumber(item.khoiLuongThucHien)}</td>
                <td>{formatNumber(item.khoiLuongNghiemThu)}</td>
                <td>{formatNumber(item.khoiLuongChuaNghiemThu)}</td>
                <td>{formatNumber(item.khoiLuongConLaiHopDong)}</td>
                <td><span className={item.trangThai === "Còn nghiệm thu" ? "status" : "status done"}>{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!reconciliationRows.length} text="Chưa có dữ liệu BOQ để đối chiếu khối lượng." />
      </Panel>

      <Panel title="Báo cáo giá trị nghiệm thu">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Số biên bản</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>Khối lượng nghiệm thu</th>
              <th>Đơn giá</th>
              <th>Giá trị nghiệm thu</th>
            </tr>
          </thead>
          <tbody>
            {acceptances.map((item) => (
              <tr key={`value-${item.id}`}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.soBienBan}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{formatNumber(item.khoiLuongNghiemThuKyNay)}</td>
                <td>{formatCurrency(item.donGiaHopDong)}</td>
                <td>{formatCurrency(item.thanhTienNghiemThuKyNay)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!acceptances.length} text="Chưa có dữ liệu giá trị nghiệm thu." />
      </Panel>
    </section>
  );
}

function EquipmentPage({
  contracts,
  dailyLogs,
  equipmentForm,
  equipmentList,
  equipmentTransferForm,
  equipmentTransfers,
  isEquipmentFormOpen,
  isEquipmentTransferFormOpen,
  onCancelEquipment,
  onCancelEquipmentTransfer,
  onChangeEquipment,
  onChangeEquipmentTransfer,
  onDownloadTemplate,
  onOpenEquipmentForm,
  onOpenEquipmentTransferForm,
  onSubmitEquipment,
  onSubmitEquipmentTransfer,
  onUploadExcel,
}) {
  const contractOptions = getDailyContractOptions(contracts);
  const equipmentUsages = getEquipmentUsagesFromDailyLogs(dailyLogs);
  const reportRows = getEquipmentReportRows(equipmentUsages, equipmentTransfers, equipmentList);
  const totalHours = reportRows.reduce((sum, item) => sum + toNumber(item.tongGioHoatDong), 0);
  const totalCost = reportRows.reduce((sum, item) => sum + item.chiPhiThietBi, 0);

  return (
    <section className="equipment-layout">
      <div className="page-grid equipment-summary">
        <SummaryCard label="Tổng số thiết bị" value={equipmentList.length} tone="blue" />
        <SummaryCard label="Thiết bị đang tại công trình" value={reportRows.length} tone="green" />
        <SummaryCard label="Tổng giờ hoạt động" value={formatNumber(totalHours)} tone="amber" />
        <SummaryCard label="Chi phí thiết bị tạm tính" value={formatCurrency(totalCost)} tone="red" />
      </div>

      <Panel
        title="Danh mục máy móc thiết bị"
        action={
          <ExcelPanelActions
            addLabel="Thêm thiết bị"
            downloadLabel="Tải mẫu Thiết bị"
            isFormOpen={isEquipmentFormOpen}
            onAdd={onOpenEquipmentForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload Thiết bị"
          />
        }
      >
        {isEquipmentFormOpen && (
          <form className="work-form" onSubmit={onSubmitEquipment}>
            <Field label="Mã thiết bị">
              <input name="maThietBi" value={equipmentForm.maThietBi} onChange={onChangeEquipment} placeholder="TB-001" />
            </Field>
            <Field label="Tên thiết bị">
              <input name="tenThietBi" value={equipmentForm.tenThietBi} onChange={onChangeEquipment} />
            </Field>
            <Field label="Loại thiết bị">
              <select name="loaiThietBi" value={equipmentForm.loaiThietBi} onChange={onChangeEquipment}>
                <option>Máy phun</option>
                <option>Máy nén khí</option>
                <option>Giàn giáo</option>
                <option>Máy khoan</option>
                <option>Xe vận chuyển</option>
                <option>Khác</option>
              </select>
            </Field>
            <Field label="Nguồn sở hữu">
              <select name="nguonSoHuu" value={equipmentForm.nguonSoHuu} onChange={onChangeEquipment}>
                <option>Công ty</option>
                <option>Thuê ngoài</option>
              </select>
            </Field>
            <Field label="Đơn vị sở hữu hoặc bên cho thuê">
              <input name="donViSoHuu" value={equipmentForm.donViSoHuu} onChange={onChangeEquipment} />
            </Field>
            <Field label="Đơn giá thuê/ngày nếu có">
              <input name="donGiaThueNgay" type="number" min="0" value={equipmentForm.donGiaThueNgay} onChange={onChangeEquipment} />
            </Field>
            <Field label="Đơn giá thuê/giờ nếu có">
              <input name="donGiaThueGio" type="number" min="0" value={equipmentForm.donGiaThueGio} onChange={onChangeEquipment} />
            </Field>
            <Field label="Tình trạng hiện tại">
              <input name="tinhTrangHienTai" value={equipmentForm.tinhTrangHienTai} onChange={onChangeEquipment} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={equipmentForm.ghiChu} onChange={onChangeEquipment} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelEquipment}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isEquipmentFormOpen} text="Bấm Thêm thiết bị để tạo danh mục máy móc thiết bị." />

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Mã thiết bị</th>
              <th>Tên thiết bị</th>
              <th>Loại</th>
              <th>Nguồn sở hữu</th>
              <th>Đơn vị sở hữu / cho thuê</th>
              <th>Đơn giá/ngày</th>
              <th>Đơn giá/giờ</th>
              <th>Tình trạng</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {equipmentList.map((item) => (
              <tr key={item.id}>
                <td>{item.maThietBi}</td>
                <td>{item.tenThietBi}</td>
                <td>{item.loaiThietBi}</td>
                <td>{item.nguonSoHuu}</td>
                <td>{item.donViSoHuu || "-"}</td>
                <td>{formatCurrency(item.donGiaThueNgay)}</td>
                <td>{formatCurrency(item.donGiaThueGio)}</td>
                <td>{item.tinhTrangHienTai || "-"}</td>
                <td>{item.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!equipmentList.length} text="Chưa có thiết bị nào trong danh mục." />
      </Panel>

      <Panel
        title="Điều chuyển / nhận thiết bị về công trình"
        action={!isEquipmentTransferFormOpen && <button type="button" onClick={onOpenEquipmentTransferForm}>Thêm phiếu nhận</button>}
      >
        {isEquipmentTransferFormOpen && (
          <form className="work-form" onSubmit={onSubmitEquipmentTransfer}>
            <Field label="Ngày nhận">
              <input name="ngayNhan" type="date" value={equipmentTransferForm.ngayNhan} onChange={onChangeEquipmentTransfer} />
            </Field>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={equipmentTransferForm.hopDongCongTrinh} onChange={onChangeEquipmentTransfer}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>{option.label}</option>
                ))}
              </select>
            </Field>
            <Field label="Mã thiết bị">
              {equipmentList.length ? (
                <select name="maThietBi" value={equipmentTransferForm.maThietBi} onChange={onChangeEquipmentTransfer}>
                  <option value="">Chọn thiết bị</option>
                  {equipmentList.map((item) => (
                    <option key={item.id} value={item.maThietBi}>{item.maThietBi} - {item.tenThietBi}</option>
                  ))}
                </select>
              ) : (
                <input name="maThietBi" value={equipmentTransferForm.maThietBi} onChange={onChangeEquipmentTransfer} />
              )}
            </Field>
            <Field label="Tên thiết bị">
              <input name="tenThietBi" value={equipmentTransferForm.tenThietBi} onChange={onChangeEquipmentTransfer} />
            </Field>
            <Field label="Nguồn">
              <select name="nguon" value={equipmentTransferForm.nguon} onChange={onChangeEquipmentTransfer}>
                <option>Công ty</option>
                <option>Thuê ngoài</option>
              </select>
            </Field>
            <Field label="Người giao">
              <input name="nguoiGiao" value={equipmentTransferForm.nguoiGiao} onChange={onChangeEquipmentTransfer} />
            </Field>
            <Field label="Người nhận">
              <input name="nguoiNhan" value={equipmentTransferForm.nguoiNhan} onChange={onChangeEquipmentTransfer} />
            </Field>
            <Field label="Tình trạng khi nhận">
              <input name="tinhTrangKhiNhan" value={equipmentTransferForm.tinhTrangKhiNhan} onChange={onChangeEquipmentTransfer} />
            </Field>
            <Field label="Ngày trả dự kiến">
              <input name="ngayTraDuKien" type="date" value={equipmentTransferForm.ngayTraDuKien} onChange={onChangeEquipmentTransfer} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={equipmentTransferForm.ghiChu} onChange={onChangeEquipmentTransfer} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelEquipmentTransfer}>Hủy</button>
            </div>
          </form>
        )}
        <EmptyState show={!isEquipmentTransferFormOpen} text="Bấm Thêm phiếu nhận để điều chuyển hoặc nhận thiết bị về công trình." />

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày nhận</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã thiết bị</th>
              <th>Tên thiết bị</th>
              <th>Nguồn</th>
              <th>Người giao</th>
              <th>Người nhận</th>
              <th>Tình trạng khi nhận</th>
              <th>Ngày trả dự kiến</th>
            </tr>
          </thead>
          <tbody>
            {equipmentTransfers.map((item) => (
              <tr key={item.id}>
                <td>{formatDate(item.ngayNhan)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maThietBi}</td>
                <td>{item.tenThietBi}</td>
                <td>{item.nguon}</td>
                <td>{item.nguoiGiao || "-"}</td>
                <td>{item.nguoiNhan || "-"}</td>
                <td>{item.tinhTrangKhiNhan || "-"}</td>
                <td>{formatDate(item.ngayTraDuKien)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!equipmentTransfers.length} text="Chưa có phiếu nhận thiết bị nào." />
      </Panel>

      <Panel title="Nhật ký sử dụng thiết bị">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>Mã thiết bị</th>
              <th>Tên thiết bị</th>
              <th>Số giờ hoạt động</th>
              <th>Nguồn</th>
              <th>Giám sát</th>
            </tr>
          </thead>
          <tbody>
            {equipmentUsages.map((item) => (
              <tr key={item.id}>
                <td>{formatDate(item.ngay)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs || "-"}</td>
                <td>{item.hangMuc || "-"}</td>
                <td>{item.maThietBi || "-"}</td>
                <td>{item.tenThietBi}</td>
                <td>{formatNumber(item.soGioHoatDong)}</td>
                <td>{item.nguon || "-"}</td>
                <td>{item.giamSat || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!equipmentUsages.length} text="Chưa có dữ liệu sử dụng thiết bị từ Nhật ký thi công" />
      </Panel>

      <Panel title="Báo cáo thiết bị theo công trình">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã thiết bị</th>
              <th>Tên thiết bị</th>
              <th>Tổng giờ hoạt động</th>
              <th>Số ngày tại công trình</th>
              <th>Đơn giá thuê/giờ</th>
              <th>Đơn giá thuê/ngày</th>
              <th>Chi phí thiết bị tạm tính</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maThietBi || "-"}</td>
                <td>{item.tenThietBi}</td>
                <td>{formatNumber(item.tongGioHoatDong)}</td>
                <td>{formatNumber(item.soNgayTaiCongTrinh)}</td>
                <td>{item.donGiaThueGio > 0 ? formatCurrency(item.donGiaThueGio) : "-"}</td>
                <td>{item.donGiaThueNgay > 0 ? formatCurrency(item.donGiaThueNgay) : "-"}</td>
                <td>{formatCurrency(item.chiPhiThietBi)}</td>
                <td><span className={item.trangThai === "Thiếu đơn giá thuê" ? "status" : "status done"}>{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!reportRows.length} text="Chưa có dữ liệu để tổng hợp báo cáo thiết bị." />
      </Panel>
    </section>
  );
}

function SubcontractorsPage({
  contracts,
  dailyLogs,
  isSubContractFormOpen,
  isSubcontractorFormOpen,
  onCancelSubContract,
  onCancelSubcontractor,
  onChangeSubContract,
  onChangeSubcontractor,
  onOpenSubContractForm,
  onOpenSubcontractorForm,
  onSubmitSubContract,
  onSubmitSubcontractor,
  subContractForm,
  subContracts,
  subcontractorForm,
  subcontractors,
  wbsPlans,
}) {
  const contractOptions = getDailyContractOptions(contracts);
  const wbsOptions = wbsPlans.length ? getUniqueWbsPlanItems(wbsPlans) : sampleWbsPlans;
  const reportRows = getSubcontractorReportRows(subContracts, dailyLogs);
  const totalSubContractValue = subContracts.reduce((sum, item) => sum + toNumber(item.thanhTienKhoan), 0);
  const totalAcceptedValue = reportRows.reduce((sum, item) => sum + item.giaTriNghiemThu, 0);
  const totalPayable = reportRows.reduce((sum, item) => sum + item.conPhaiTra, 0);
  const subContractAmount = toNumber(subContractForm.khoiLuongGiaoKhoan) * toNumber(subContractForm.donGiaKhoan);

  return (
    <section className="subcontractors-layout">
      <div className="page-grid subcontractors-summary">
        <SummaryCard label="Tổng số nhà thầu phụ / tổ đội" value={subcontractors.length} tone="blue" />
        <SummaryCard label="Tổng giá trị giao khoán" value={formatCurrency(totalSubContractValue)} tone="green" />
        <SummaryCard label="Tổng giá trị nghiệm thu" value={formatCurrency(totalAcceptedValue)} tone="amber" />
        <SummaryCard label="Tổng còn phải trả" value={formatCurrency(totalPayable)} tone="red" />
      </div>

      <Panel
        title="Danh mục nhà thầu phụ / tổ đội"
        action={
          !isSubcontractorFormOpen && (
            <button type="button" onClick={onOpenSubcontractorForm}>
              Thêm đối tác
            </button>
          )
        }
      >
        {isSubcontractorFormOpen && (
          <form className="work-form" onSubmit={onSubmitSubcontractor}>
            <Field label="Mã đối tác">
              <input name="maDoiTac" value={subcontractorForm.maDoiTac} onChange={onChangeSubcontractor} placeholder="DT-001" />
            </Field>
            <Field label="Tên nhà thầu phụ / tổ đội">
              <input name="tenDoiTac" value={subcontractorForm.tenDoiTac} onChange={onChangeSubcontractor} />
            </Field>
            <Field label="Người đại diện">
              <input name="nguoiDaiDien" value={subcontractorForm.nguoiDaiDien} onChange={onChangeSubcontractor} />
            </Field>
            <Field label="Số điện thoại">
              <input name="soDienThoai" value={subcontractorForm.soDienThoai} onChange={onChangeSubcontractor} />
            </Field>
            <Field label="Loại đối tác">
              <select name="loaiDoiTac" value={subcontractorForm.loaiDoiTac} onChange={onChangeSubcontractor}>
                <option>Tổ đội thi công</option>
                <option>Nhà thầu phụ</option>
                <option>Nhà cung cấp dịch vụ</option>
              </select>
            </Field>
            <Field label="Hạng mục chuyên môn">
              <select name="hangMucChuyenMon" value={subcontractorForm.hangMucChuyenMon} onChange={onChangeSubcontractor}>
                <option>Sơn đá</option>
                <option>Sơn lót</option>
                <option>Giàn giáo</option>
                <option>Vệ sinh</option>
                <option>Vận chuyển</option>
                <option>Khác</option>
              </select>
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={subcontractorForm.ghiChu} onChange={onChangeSubcontractor} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelSubcontractor}>
                Hủy
              </button>
            </div>
          </form>
        )}
        <EmptyState show={!isSubcontractorFormOpen} text="Bấm Thêm đối tác để tạo danh mục nhà thầu phụ / tổ đội." />

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Mã đối tác</th>
              <th>Tên nhà thầu phụ / tổ đội</th>
              <th>Người đại diện</th>
              <th>Số điện thoại</th>
              <th>Loại đối tác</th>
              <th>Hạng mục chuyên môn</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {subcontractors.map((item) => (
              <tr key={item.id}>
                <td>{item.maDoiTac}</td>
                <td>{item.tenDoiTac}</td>
                <td>{item.nguoiDaiDien || "-"}</td>
                <td>{item.soDienThoai || "-"}</td>
                <td>{item.loaiDoiTac}</td>
                <td>{item.hangMucChuyenMon}</td>
                <td>{item.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!subcontractors.length} text="Chưa có đối tác nào trong danh mục." />
      </Panel>

      <Panel
        title="Hợp đồng giao khoán / phân giao khối lượng"
        action={
          !isSubContractFormOpen && (
            <button type="button" onClick={onOpenSubContractForm}>
              Thêm giao khoán
            </button>
          )
        }
      >
        {isSubContractFormOpen && (
          <form className="work-form" onSubmit={onSubmitSubContract}>
            <Field label="Hợp đồng / Công trình">
              <select name="hopDongCongTrinh" value={subContractForm.hopDongCongTrinh} onChange={onChangeSubContract}>
                <option value="">Chọn hợp đồng / công trình</option>
                {contractOptions.map((option) => (
                  <option key={option.id} value={option.label}>{option.label}</option>
                ))}
              </select>
            </Field>
            <Field label="Mã WBS">
              <select name="maWbs" value={subContractForm.maWbs} onChange={onChangeSubContract}>
                <option value="">Chọn mã WBS</option>
                {wbsOptions.map((item) => (
                  <option key={`${item.maWbs}-${item.hangMucCongViec}`} value={item.maWbs}>
                    {item.maWbs} - {item.hangMucCongViec}
                  </option>
                ))}
              </select>
            </Field>
            <Field label="Hạng mục công việc">
              <input name="hangMucCongViec" value={subContractForm.hangMucCongViec} onChange={onChangeSubContract} />
            </Field>
            <Field label="Nhà thầu phụ / Tổ đội">
              <select name="doiTac" value={subContractForm.doiTac} onChange={onChangeSubContract}>
                <option value="">Chọn đối tác</option>
                {subcontractors.map((item) => (
                  <option key={item.id} value={item.tenDoiTac}>{item.maDoiTac} - {item.tenDoiTac}</option>
                ))}
              </select>
            </Field>
            <Field label="ĐVT">
              <input name="dvt" value={subContractForm.dvt} onChange={onChangeSubContract} />
            </Field>
            <Field label="Khối lượng giao khoán">
              <input name="khoiLuongGiaoKhoan" type="number" min="0" value={subContractForm.khoiLuongGiaoKhoan} onChange={onChangeSubContract} />
            </Field>
            <Field label="Đơn giá khoán">
              <input name="donGiaKhoan" type="number" min="0" value={subContractForm.donGiaKhoan} onChange={onChangeSubContract} />
            </Field>
            <Field label="Thành tiền khoán">
              <input value={formatCurrency(subContractAmount)} readOnly />
            </Field>
            <Field label="Ngày bắt đầu">
              <input name="ngayBatDau" type="date" value={subContractForm.ngayBatDau} onChange={onChangeSubContract} />
            </Field>
            <Field label="Ngày kết thúc">
              <input name="ngayKetThuc" type="date" value={subContractForm.ngayKetThuc} onChange={onChangeSubContract} />
            </Field>
            <Field label="Trạng thái">
              <select name="trangThai" value={subContractForm.trangThai} onChange={onChangeSubContract}>
                <option>Đang giao khoán</option>
                <option>Đang thực hiện</option>
                <option>Tạm dừng</option>
                <option>Hoàn thành</option>
              </select>
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={subContractForm.ghiChu} onChange={onChangeSubContract} />
            </Field>
            <div className="form-actions">
              <button type="submit">Lưu</button>
              <button type="button" className="secondary" onClick={onCancelSubContract}>
                Hủy
              </button>
            </div>
          </form>
        )}
        <EmptyState show={!isSubContractFormOpen} text="Bấm Thêm giao khoán để phân giao khối lượng cho đối tác." />

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>Đối tác</th>
              <th>ĐVT</th>
              <th>Khối lượng</th>
              <th>Đơn giá</th>
              <th>Thành tiền</th>
              <th>Bắt đầu</th>
              <th>Kết thúc</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {subContracts.map((item) => (
              <tr key={item.id}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{item.doiTac}</td>
                <td>{item.dvt}</td>
                <td>{formatNumber(item.khoiLuongGiaoKhoan)}</td>
                <td>{formatCurrency(item.donGiaKhoan)}</td>
                <td>{formatCurrency(item.thanhTienKhoan)}</td>
                <td>{formatDate(item.ngayBatDau)}</td>
                <td>{formatDate(item.ngayKetThuc)}</td>
                <td><span className="status done">{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!subContracts.length} text="Chưa có dòng giao khoán nào." />
      </Panel>

      <Panel title="Báo cáo thực hiện - nghiệm thu - thanh toán tổ đội">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Nhà thầu phụ / Tổ đội</th>
              <th>Hạng mục</th>
              <th>KL giao khoán</th>
              <th>KL thực hiện</th>
              <th>KL nghiệm thu</th>
              <th>Giá trị nghiệm thu</th>
              <th>Đã thanh toán</th>
              <th>Còn phải trả</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.doiTac}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{formatNumber(item.khoiLuongGiaoKhoan)}</td>
                <td>{formatNumber(item.khoiLuongThucHien)}</td>
                <td>{formatNumber(item.khoiLuongNghiemThu)}</td>
                <td>{formatCurrency(item.giaTriNghiemThu)}</td>
                <td>{formatCurrency(item.daThanhToan)}</td>
                <td>{formatCurrency(item.conPhaiTra)}</td>
                <td><span className={item.conPhaiTra > 0 ? "status" : "status done"}>{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!reportRows.length} text="Chưa có dữ liệu giao khoán để lập báo cáo." />
      </Panel>
    </section>
  );
}

function LaborPage({ dailyLogs, form, isFormOpen, onCancel, onChange, onDownloadTemplate, onOpenForm, onSubmit, onUploadExcel, teams }) {
  const laborUsages = getLaborUsagesFromDailyLogs(dailyLogs);
  const reportRows = getLaborReportRows(laborUsages, teams);
  const totalWorkers = laborUsages.reduce((sum, item) => sum + toNumber(item.soNguoi), 0);
  const totalHours = laborUsages.reduce((sum, item) => sum + toNumber(item.soGio), 0);
  const totalCost = reportRows.reduce((sum, item) => sum + (item.hasLaborRate ? item.chiPhiNhanCong : 0), 0);

  return (
    <section className="labor-layout">
      <div className="page-grid labor-summary">
        <SummaryCard label="Tổng số tổ đội" value={teams.length} tone="blue" />
        <SummaryCard label="Tổng nhân công sử dụng" value={formatNumber(totalWorkers)} tone="green" />
        <SummaryCard label="Tổng giờ công" value={formatNumber(totalHours)} tone="amber" />
        <SummaryCard label="Chi phí nhân công tạm tính" value={formatCurrency(totalCost)} tone="red" />
      </div>

      <Panel
        title="Danh mục tổ đội"
        action={
          <ExcelPanelActions
            addLabel="Thêm tổ đội"
            downloadLabel="Tải mẫu Tổ đội"
            isFormOpen={isFormOpen}
            onAdd={onOpenForm}
            onDownload={onDownloadTemplate}
            onUpload={onUploadExcel}
            uploadLabel="Upload Tổ đội"
          />
        }
      >
        {isFormOpen && (
          <form className="work-form" onSubmit={onSubmit}>
            <Field label="Mã tổ đội">
              <input name="maToDoi" value={form.maToDoi} onChange={onChange} placeholder="TD-001" />
            </Field>
            <Field label="Tên tổ đội">
              <input name="tenToDoi" value={form.tenToDoi} onChange={onChange} placeholder="Tổ sơn đá A" />
            </Field>
            <Field label="Người đại diện">
              <input name="nguoiDaiDien" value={form.nguoiDaiDien} onChange={onChange} />
            </Field>
            <Field label="Số điện thoại">
              <input name="soDienThoai" value={form.soDienThoai} onChange={onChange} />
            </Field>
            <Field label="Loại tổ đội">
              <select name="loaiToDoi" value={form.loaiToDoi} onChange={onChange}>
                <option>Sơn đá</option>
                <option>Sơn lót</option>
                <option>Phụ trợ</option>
                <option>Giàn giáo</option>
                <option>Khác</option>
              </select>
            </Field>
            <Field label="Đơn giá công/ngày">
              <input name="donGiaCongNgay" type="number" min="0" value={form.donGiaCongNgay} onChange={onChange} />
            </Field>
            <Field label="Ghi chú" wide>
              <textarea name="ghiChu" rows="3" value={form.ghiChu} onChange={onChange} />
            </Field>

            <div className="form-actions">
              <button type="submit">Lưu tổ đội</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}
        <EmptyState show={!isFormOpen} text="Bấm Thêm tổ đội để tạo danh mục tổ đội thi công." />

        <ResponsiveTable>
          <thead>
            <tr>
              <th>Mã tổ đội</th>
              <th>Tên tổ đội</th>
              <th>Người đại diện</th>
              <th>Số điện thoại</th>
              <th>Loại tổ đội</th>
              <th>Đơn giá công/ngày</th>
              <th>Ghi chú</th>
            </tr>
          </thead>
          <tbody>
            {teams.map((team) => (
              <tr key={team.id}>
                <td>{team.maToDoi}</td>
                <td>{team.tenToDoi}</td>
                <td>{team.nguoiDaiDien || "-"}</td>
                <td>{team.soDienThoai || "-"}</td>
                <td>{team.loaiToDoi}</td>
                <td>{formatCurrency(team.donGiaCongNgay)}</td>
                <td>{team.ghiChu || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!teams.length} text="Chưa có tổ đội nào trong danh mục." />
      </Panel>

      <Panel title="Nhân công sử dụng từ Nhật ký thi công">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>Tổ đội</th>
              <th>Số người</th>
              <th>Số giờ</th>
              <th>Giám sát</th>
            </tr>
          </thead>
          <tbody>
            {laborUsages.map((item) => (
              <tr key={item.id}>
                <td>{formatDate(item.ngay)}</td>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs || "-"}</td>
                <td>{item.hangMuc || "-"}</td>
                <td>{item.toDoi}</td>
                <td>{formatNumber(item.soNguoi)}</td>
                <td>{formatNumber(item.soGio)}</td>
                <td>{item.giamSat || "-"}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!laborUsages.length} text="Chưa có dữ liệu nhân công từ Nhật ký thi công." />
      </Panel>

      <Panel title="Báo cáo nhân công theo công trình">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Tổ đội</th>
              <th>Tổng số người</th>
              <th>Tổng số giờ</th>
              <th>Tổng công quy đổi</th>
              <th>Đơn giá công/ngày</th>
              <th>Chi phí nhân công tạm tính</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.toDoi}</td>
                <td>{formatNumber(item.tongSoNguoi)}</td>
                <td>{formatNumber(item.tongSoGio)}</td>
                <td>{formatNumber(item.tongCongQuyDoi)}</td>
                <td>{item.hasLaborRate ? formatCurrency(item.donGiaCongNgay) : "Chưa có đơn giá"}</td>
                <td>{item.hasLaborRate ? formatCurrency(item.chiPhiNhanCong) : "Chưa có đơn giá"}</td>
                <td>
                  <span className={item.status === "Đủ đơn giá" ? "status done" : "status"}>{item.status}</span>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!reportRows.length} text="Chưa có dữ liệu để tổng hợp báo cáo nhân công." />
      </Panel>
    </section>
  );
}

function DailyLogPage({
  allLogs,
  contracts,
  form,
  imageMeta,
  images,
  isFormOpen,
  logs,
  onCancel,
  onChange,
  onChangeImageMeta,
  onDelete,
  onOpenForm,
  onRemoveImage,
  onSubmit,
  onToggleImages,
  onUploadImages,
  visibleImageLogId,
  wbsPlans,
}) {
  const contractOptions = getDailyContractOptions(contracts);
  const wbsOptions = wbsPlans.length ? getUniqueWbsPlanItems(wbsPlans) : sampleWbsPlans;
  const previewCumulative =
    allLogs
      .filter((log) => log.hopDongCongTrinh === form.hopDongCongTrinh && log.maWbs === form.maWbs)
      .reduce((sum, log) => sum + toNumber(log.khoiLuongThucHienNgay), 0) + toNumber(form.khoiLuongThucHienNgay);
  const previewCompletionRate = toNumber(form.khoiLuongKeHoach)
    ? (previewCumulative / toNumber(form.khoiLuongKeHoach)) * 100
    : 0;
  const totals = logs.reduce(
    (summary, log) => {
      summary.nhanCong += toNumber(log.soNguoi);
      summary.vatTu += toNumber(log.soLuongVatTu);
      summary.gioThietBi += toNumber(log.soGioHoatDong);
      return summary;
    },
    { nhanCong: 0, vatTu: 0, gioThietBi: 0 }
  );

  return (
    <>
      <section className="page-grid daily-summary">
        <SummaryCard label="Tổng nhân công" value={formatNumber(totals.nhanCong)} tone="blue" />
        <SummaryCard label="Tổng vật tư sử dụng" value={formatNumber(totals.vatTu)} tone="green" />
        <SummaryCard label="Tổng giờ thiết bị" value={formatNumber(totals.gioThietBi)} tone="amber" />
      </section>

      <Panel
        title="Nhập nhật ký thi công hằng ngày"
        action={
          !isFormOpen && (
            <button type="button" onClick={onOpenForm}>
              Thêm nhật ký
            </button>
          )
        }
      >
        {isFormOpen && (
        <form className="work-form" onSubmit={onSubmit}>
          <SectionLabel title="Thông tin chung" />
          <Field label="Ngày">
            <input name="ngay" type="date" value={form.ngay} onChange={onChange} />
          </Field>
          <Field label="Hợp đồng / Công trình">
            <select name="hopDongCongTrinh" value={form.hopDongCongTrinh} onChange={onChange}>
              <option value="">Chọn hợp đồng / công trình</option>
              {contractOptions.map((option) => (
                <option key={option.id} value={option.label}>
                  {option.label}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Mã WBS">
            <select name="maWbs" value={form.maWbs} onChange={onChange}>
              <option value="">Chọn mã WBS</option>
              {wbsOptions.map((wbs) => (
                <option key={`${wbs.maWbs}-${wbs.hangMucCongViec}`} value={wbs.maWbs}>
                  {wbs.maWbs} - {wbs.hangMucCongViec}
                </option>
              ))}
            </select>
          </Field>
          <Field label="Hạng mục công việc">
            <input name="hangMucCongViec" value={form.hangMucCongViec} onChange={onChange} />
          </Field>
          <Field label="Giám sát">
            <input name="giamSat" value={form.giamSat} onChange={onChange} placeholder="Tên cán bộ giám sát" />
          </Field>
          <Field label="Thời tiết">
            <select name="thoiTiet" value={form.thoiTiet} onChange={onChange}>
              <option>Nắng</option>
              <option>Mưa</option>
              <option>Âm u</option>
              <option>Gián đoạn do thời tiết</option>
            </select>
          </Field>
          <Field label="Nội dung công việc" wide>
            <textarea name="noiDung" rows="4" value={form.noiDung} onChange={onChange} />
          </Field>
          <Field label="Vấn đề phát sinh" wide>
            <textarea name="vanDe" rows="3" value={form.vanDe} onChange={onChange} />
          </Field>
          <Field label="Ghi chú" wide>
            <textarea name="ghiChu" rows="3" value={form.ghiChu} onChange={onChange} />
          </Field>

          <SectionLabel title="Khối lượng thực hiện" />
          <Field label="ĐVT">
            <input name="dvtKhoiLuong" value={form.dvtKhoiLuong} onChange={onChange} />
          </Field>
          <Field label="Khối lượng kế hoạch">
            <input name="khoiLuongKeHoach" type="number" min="0" value={form.khoiLuongKeHoach} onChange={onChange} />
          </Field>
          <Field label="Khối lượng thực hiện trong ngày">
            <input name="khoiLuongThucHienNgay" type="number" min="0" value={form.khoiLuongThucHienNgay} onChange={onChange} />
          </Field>
          <Field label="Khối lượng lũy kế thực hiện">
            <input value={formatNumber(previewCumulative)} readOnly />
          </Field>
          <Field label="Tỷ lệ hoàn thành %">
            <input value={`${formatNumber(Math.min(previewCompletionRate, 100))}%`} readOnly />
          </Field>

          <SectionLabel title="Vật tư sử dụng" />
          <Field label="Mã vật tư">
            <input name="maVatTu" value={form.maVatTu} onChange={onChange} placeholder="VT-001" />
          </Field>
          <Field label="Tên vật tư">
            <input name="tenVatTu" value={form.tenVatTu} onChange={onChange} placeholder="Sơn lót, bột trét..." />
          </Field>
          <Field label="ĐVT">
            <input name="dvtVatTu" value={form.dvtVatTu} onChange={onChange} />
          </Field>
          <Field label="Số lượng sử dụng">
            <input name="soLuongVatTu" type="number" min="0" value={form.soLuongVatTu} onChange={onChange} />
          </Field>

          <SectionLabel title="Nhân công" />
          <Field label="Tổ đội">
            <input name="toDoi" value={form.toDoi} onChange={onChange} placeholder="Tổ sơn nước" />
          </Field>
          <Field label="Số người">
            <input name="soNguoi" type="number" min="0" value={form.soNguoi} onChange={onChange} />
          </Field>
          <Field label="Số giờ">
            <input name="soGio" type="number" min="0" value={form.soGio} onChange={onChange} />
          </Field>
          <Field label="Ghi chú nhân công">
            <input name="ghiChuNhanCong" value={form.ghiChuNhanCong} onChange={onChange} />
          </Field>

          <SectionLabel title="Máy móc thiết bị" />
          <Field label="Mã thiết bị">
            <input name="maThietBi" value={form.maThietBi} onChange={onChange} placeholder="TB-001" />
          </Field>
          <Field label="Tên thiết bị">
            <input name="tenThietBi" value={form.tenThietBi} onChange={onChange} />
          </Field>
          <Field label="Số giờ hoạt động">
            <input name="soGioHoatDong" type="number" min="0" value={form.soGioHoatDong} onChange={onChange} />
          </Field>
          <Field label="Nguồn">
            <select name="nguonThietBi" value={form.nguonThietBi} onChange={onChange}>
              <option>Công ty</option>
              <option>Thuê ngoài</option>
            </select>
          </Field>

          <SectionLabel title="Hình ảnh hiện trường" />
          <Field label="Loại ảnh">
            <select name="loaiAnh" value={imageMeta.loaiAnh} onChange={onChangeImageMeta}>
              {DAILY_IMAGE_TYPES.map((type) => (
                <option key={type}>{type}</option>
              ))}
            </select>
          </Field>
          <Field label="Người tải lên">
            <input name="nguoiTaiLen" value={imageMeta.nguoiTaiLen} onChange={onChangeImageMeta} />
          </Field>
          <Field label="Ghi chú ảnh">
            <input name="ghiChuAnh" value={imageMeta.ghiChuAnh} onChange={onChangeImageMeta} />
          </Field>
          <Field label="Upload ảnh">
            <input type="file" accept="image/*" multiple onChange={onUploadImages} />
          </Field>
          <div className="daily-image-preview field wide">
            <span>Ảnh đã chọn ({images.length})</span>
            <div className="image-grid">
              {images.map((image) => (
                <figure className="image-thumb" key={image.id}>
                  <img src={image.fileAnh} alt={image.tenFile} />
                  <figcaption>
                    <strong>{image.loaiAnh}</strong>
                    <span>{image.tenFile}</span>
                    <span>{image.ghiChuAnh || "-"}</span>
                  </figcaption>
                  <button type="button" className="text-danger" onClick={() => onRemoveImage(image.id)}>Xóa ảnh</button>
                </figure>
              ))}
            </div>
            <EmptyState show={!images.length} text="Chưa chọn ảnh hiện trường. Ảnh sẽ tự giảm kích thước trước khi lưu." />
          </div>

          <div className="form-actions">
            <button type="submit">Lưu nhật ký</button>
            <button type="button" className="secondary" onClick={onCancel}>
              Hủy
            </button>
          </div>
        </form>
        )}
        <EmptyState show={!isFormOpen} text="Bấm Thêm nhật ký để nhập nhật ký thi công hằng ngày." />
      </Panel>

      <Panel title="Nhật ký đã lưu">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Ngày</th>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>Nội dung</th>
              <th>KL ngày</th>
              <th>Lũy kế</th>
              <th>Hoàn thành</th>
              <th>Nhân công</th>
              <th>Vật tư</th>
              <th>Thiết bị</th>
              <th>Ảnh</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {logs.map((log) => {
              const logImages = getDailyLogImages(log);
              const isViewingImages = visibleImageLogId === log.id;

              return (
                <Fragment key={log.id}>
                  <tr>
                    <td>{formatDate(log.ngay)}</td>
                    <td>{log.hopDongCongTrinh || log.congTrinh || "-"}</td>
                    <td>{log.maWbs || "-"}</td>
                    <td>{log.hangMucCongViec || log.hangMuc || "-"}</td>
                    <td className="description-cell">{log.noiDung}</td>
                    <td>{formatNumber(log.khoiLuongThucHienNgay || log.khoiLuong)} {log.dvtKhoiLuong || ""}</td>
                    <td>{formatNumber(log.khoiLuongLuyKe)} {log.dvtKhoiLuong || ""}</td>
                    <td>{formatNumber(log.tyLeHoanThanh)}%</td>
                    <td>{formatNumber(log.soNguoi)} người / {formatNumber(log.soGio)} giờ</td>
                    <td>{log.tenVatTu || "-"} {formatNumber(log.soLuongVatTu)} {log.dvtVatTu || ""}</td>
                    <td>{log.tenThietBi || "-"} {formatNumber(log.soGioHoatDong)} giờ</td>
                    <td>
                      <div className="table-actions">
                        <span className="status done">{logImages.length} ảnh</span>
                        <button type="button" className="secondary" disabled={!logImages.length} onClick={() => onToggleImages(isViewingImages ? null : log.id)}>
                          Xem ảnh
                        </button>
                      </div>
                    </td>
                    <td>
                      <button className="text-danger" type="button" onClick={() => onDelete(log.id)}>
                        Xóa
                      </button>
                    </td>
                  </tr>
                  {isViewingImages && (
                    <tr>
                      <td colSpan="13">
                        <DailyLogImageGallery images={logImages} />
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!logs.length} text="Chưa có nhật ký nào. Dữ liệu lưu trong dailyLogs sẽ hiển thị tại đây." />
      </Panel>
    </>
  );
}

function DailyLogImageGallery({ images }) {
  return (
    <div className="daily-image-gallery">
      {images.map((image) => (
        <figure className="image-thumb saved" key={image.id || image.tenFile}>
          <img src={image.fileAnh} alt={image.tenFile} />
          <figcaption>
            <strong>{image.loaiAnh}</strong>
            <span>{image.tenFile}</span>
            <span>{formatDate(image.ngayTaiLen)} - {image.nguoiTaiLen || "Chưa ghi người tải"}</span>
            <span>{image.ghiChuAnh || "-"}</span>
          </figcaption>
        </figure>
      ))}
    </div>
  );
}

function ProjectDocumentsPage({
  activeTab,
  contracts,
  documents,
  form,
  projectFilter,
  search,
  onCancel,
  onChange,
  onChangeFile,
  onChangeProjectFilter,
  onChangeSearch,
  onChangeTab,
  onSubmit,
}) {
  const contractOptions = getDailyContractOptions(contracts);
  const filteredDocuments = getFilteredProjectDocuments(documents, activeTab, projectFilter, search);
  const dashboard = getProjectDocumentDashboard(documents);

  return (
    <section className="documents-layout">
      <div className="page-grid documents-summary">
        <SummaryCard label="Tổng hồ sơ" value={dashboard.total} tone="blue" />
        <SummaryCard label="Đã tải lên" value={dashboard.uploaded} tone="green" />
        <SummaryCard label="Thiếu hồ sơ" value={dashboard.missing} tone="red" />
        <SummaryCard label="Dung lượng lưu trữ" value={formatBytes(dashboard.storageSize)} tone="amber" />
      </div>

      <Panel title="Bộ lọc hồ sơ">
        <div className="work-form">
          <Field label="Tìm kiếm">
            <input value={search} onChange={(event) => onChangeSearch(event.target.value)} placeholder="Mã, tên, loại hồ sơ, người tải..." />
          </Field>
          <Field label="Lọc theo công trình">
            <select value={projectFilter} onChange={(event) => onChangeProjectFilter(event.target.value)}>
              <option>{ALL_PROJECTS}</option>
              {contractOptions.map((option) => (
                <option key={option.id} value={option.label}>{option.label}</option>
              ))}
            </select>
          </Field>
        </div>
      </Panel>

      <Panel title="Nhóm hồ sơ">
        <div className="document-tabs">
          {DOCUMENT_TABS.map((tab) => (
            <button
              className={tab === activeTab ? "tab-button active" : "tab-button"}
              key={tab}
              type="button"
              onClick={() => onChangeTab(tab)}
            >
              {tab}
            </button>
          ))}
        </div>
      </Panel>

      <Panel title={`Upload hồ sơ - ${activeTab}`}>
        <form className="work-form" onSubmit={onSubmit}>
          <Field label="Hợp đồng / Công trình">
            <select name="hopDongCongTrinh" value={form.hopDongCongTrinh} onChange={onChange}>
              <option value="">Chọn hợp đồng / công trình</option>
              {contractOptions.map((option) => (
                <option key={option.id} value={option.label}>{option.label}</option>
              ))}
            </select>
          </Field>
          <Field label="Mã hồ sơ">
            <input name="maHoSo" value={form.maHoSo} onChange={onChange} />
          </Field>
          <Field label="Tên hồ sơ">
            <input name="tenHoSo" value={form.tenHoSo} onChange={onChange} />
          </Field>
          <Field label="Loại hồ sơ">
            <input name="loaiHoSo" value={form.loaiHoSo} onChange={onChange} placeholder={activeTab} />
          </Field>
          <Field label="Ngày tải lên">
            <input name="ngayTaiLen" type="date" value={form.ngayTaiLen} onChange={onChange} />
          </Field>
          <Field label="Người tải">
            <input name="nguoiTai" value={form.nguoiTai} onChange={onChange} />
          </Field>
          <Field label="Trạng thái">
            <select name="trangThai" value={form.trangThai} onChange={onChange}>
              <option>Đã tải lên</option>
              <option>Thiếu hồ sơ</option>
              <option>Chờ cập nhật</option>
              <option>Đã duyệt</option>
            </select>
          </Field>
          <Field label="File đính kèm">
            <input type="file" onChange={onChangeFile} />
          </Field>
          <Field label="File đã chọn" wide>
            <input value={form.fileName ? `${form.fileName} (${formatBytes(form.fileSize)})` : "Chưa chọn file"} readOnly />
          </Field>
          <div className="form-actions">
            <button type="submit">Upload file</button>
            <button type="button" className="secondary" onClick={onCancel}>Hủy</button>
          </div>
        </form>
      </Panel>

      <Panel title={`Danh sách hồ sơ - ${activeTab}`}>
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Mã hồ sơ</th>
              <th>Tên hồ sơ</th>
              <th>Loại hồ sơ</th>
              <th>Ngày tải lên</th>
              <th>Người tải</th>
              <th>Trạng thái</th>
              <th>File đính kèm</th>
              <th>Thao tác</th>
            </tr>
          </thead>
          <tbody>
            {filteredDocuments.map((item) => (
              <tr key={item.id}>
                <td>{item.maHoSo}</td>
                <td>{item.tenHoSo}</td>
                <td>{item.loaiHoSo || item.nhomHoSo}</td>
                <td>{formatDate(item.ngayTaiLen)}</td>
                <td>{item.nguoiTai || "-"}</td>
                <td><span className={item.trangThai === "Đã tải lên" || item.trangThai === "Đã duyệt" ? "status done" : "status"}>{item.trangThai}</span></td>
                <td>{item.fileName || "Chưa có file"}</td>
                <td>
                  <div className="table-actions">
                    <button type="button" className="secondary" disabled={!item.fileData} onClick={() => viewProjectDocumentFile(item)}>Xem</button>
                    <button type="button" disabled={!item.fileData} onClick={() => downloadProjectDocumentFile(item)}>Download</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!filteredDocuments.length} text="Chưa có hồ sơ phù hợp với tab hoặc bộ lọc hiện tại." />
      </Panel>
    </section>
  );
}

function DocumentTemplatesPage({
  editingTemplateId,
  filters,
  form,
  onCancel,
  onChange,
  onChangeFile,
  onChangeFilter,
  onChangeSearch,
  onChangeStatus,
  onDownload,
  onEdit,
  onSubmit,
  onViewDescription,
  search,
  templates,
}) {
  const filteredTemplates = getFilteredDocumentTemplates(templates, search, filters);
  const dashboard = getDocumentTemplateDashboard(templates);

  return (
    <section className="document-templates-layout">
      <div className="page-grid document-templates-summary">
        <SummaryCard label="Tổng số tài liệu" value={dashboard.total} tone="blue" />
        <SummaryCard label="Đang hiệu lực" value={dashboard.active} tone="green" />
        <SummaryCard label="Hết hiệu lực" value={dashboard.expired} tone="red" />
        <SummaryCard label="Số lượt tải" value={dashboard.downloads} tone="amber" />
      </div>

      <Panel title="Bộ lọc tài liệu">
        <div className="work-form">
          <Field label="Tìm kiếm">
            <input value={search} onChange={(event) => onChangeSearch(event.target.value)} placeholder="Mã tài liệu, tên tài liệu, nhóm..." />
          </Field>
          <Field label="Nhóm tài liệu">
            <select name="nhomTaiLieu" value={filters.nhomTaiLieu} onChange={onChangeFilter}>
              <option>{ALL_PROJECTS}</option>
              {DOCUMENT_TEMPLATE_GROUPS.map((group) => <option key={group}>{group}</option>)}
            </select>
          </Field>
          <Field label="Loại file">
            <select name="loaiFile" value={filters.loaiFile} onChange={onChangeFilter}>
              <option>{ALL_PROJECTS}</option>
              {DOCUMENT_TEMPLATE_FILE_TYPES.map((type) => <option key={type}>{type}</option>)}
            </select>
          </Field>
          <Field label="Trạng thái">
            <select name="trangThai" value={filters.trangThai} onChange={onChangeFilter}>
              <option>{ALL_PROJECTS}</option>
              {DOCUMENT_TEMPLATE_STATUSES.map((status) => <option key={status}>{status}</option>)}
            </select>
          </Field>
        </div>
      </Panel>

      <Panel title={editingTemplateId ? "Chỉnh sửa tài liệu mẫu" : "Upload tài liệu mẫu"}>
        <form className="work-form" onSubmit={onSubmit}>
          <Field label="Mã tài liệu">
            <input name="maTaiLieu" value={form.maTaiLieu} onChange={onChange} placeholder="BM-DT-001" />
          </Field>
          <Field label="Tên tài liệu">
            <input name="tenTaiLieu" value={form.tenTaiLieu} onChange={onChange} placeholder="Mẫu dự toán Hodastone" />
          </Field>
          <Field label="Nhóm tài liệu">
            <select name="nhomTaiLieu" value={form.nhomTaiLieu} onChange={onChange}>
              {DOCUMENT_TEMPLATE_GROUPS.map((group) => <option key={group}>{group}</option>)}
            </select>
          </Field>
          <Field label="Loại file">
            <select name="loaiFile" value={form.loaiFile} onChange={onChange}>
              {DOCUMENT_TEMPLATE_FILE_TYPES.map((type) => <option key={type}>{type}</option>)}
            </select>
          </Field>
          <Field label="Phiên bản">
            <input name="phienBan" value={form.phienBan} onChange={onChange} placeholder="1.0" />
          </Field>
          <Field label="Ngày hiệu lực">
            <input name="ngayHieuLuc" type="date" value={form.ngayHieuLuc} onChange={onChange} />
          </Field>
          <Field label="Ngày hết hiệu lực">
            <input name="ngayHetHieuLuc" type="date" value={form.ngayHetHieuLuc} onChange={onChange} />
          </Field>
          <Field label="Người ban hành">
            <input name="nguoiBanHanh" value={form.nguoiBanHanh} onChange={onChange} />
          </Field>
          <Field label="Phòng ban phụ trách">
            <input name="phongBanPhuTrach" value={form.phongBanPhuTrach} onChange={onChange} />
          </Field>
          <Field label="Trạng thái">
            <select name="trangThai" value={form.trangThai} onChange={onChange}>
              {DOCUMENT_TEMPLATE_STATUSES.map((status) => <option key={status}>{status}</option>)}
            </select>
          </Field>
          <Field label="File đính kèm">
            <input type="file" onChange={onChangeFile} />
          </Field>
          <Field label="File đã chọn">
            <input value={form.fileName ? `${form.fileName} (${formatBytes(form.fileSize)})` : "Chưa chọn file"} readOnly />
          </Field>
          <Field label="Mô tả / Hướng dẫn sử dụng" wide>
            <textarea name="moTaHuongDan" rows="3" value={form.moTaHuongDan} onChange={onChange} />
          </Field>
          <div className="form-actions">
            <button type="submit">{editingTemplateId ? "Lưu chỉnh sửa" : "Upload tài liệu"}</button>
            <button type="button" className="secondary" onClick={onCancel}>Hủy</button>
          </div>
        </form>
      </Panel>

      <Panel title="Danh sách biểu mẫu & tài liệu thi công">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Mã tài liệu</th>
              <th>Tên tài liệu</th>
              <th>Nhóm</th>
              <th>Loại file</th>
              <th>Phiên bản</th>
              <th>Ngày hiệu lực</th>
              <th>Trạng thái</th>
              <th>Lượt tải</th>
              <th>Thao tác</th>
            </tr>
          </thead>
          <tbody>
            {filteredTemplates.map((item) => (
              <tr key={item.id}>
                <td>{item.maTaiLieu}</td>
                <td>{item.tenTaiLieu}</td>
                <td>{item.nhomTaiLieu}</td>
                <td>{item.loaiFile}</td>
                <td>{item.phienBan || "-"}</td>
                <td>{formatDate(item.ngayHieuLuc)}</td>
                <td><span className={item.trangThai === "Đang hiệu lực" ? "status done" : "status"}>{item.trangThai}</span></td>
                <td>{formatNumber(item.downloadCount)}</td>
                <td>
                  <div className="table-actions">
                    <button type="button" disabled={!item.fileData} onClick={() => onDownload(item)}>Tải xuống</button>
                    <button type="button" className="secondary" onClick={() => onViewDescription(item)}>Xem mô tả</button>
                    <button type="button" className="secondary" onClick={() => onEdit(item)}>Chỉnh sửa</button>
                    <button type="button" className="secondary" onClick={() => onChangeStatus(item.id, "Đang hiệu lực")}>Đang hiệu lực</button>
                    <button type="button" className="text-danger" onClick={() => onChangeStatus(item.id, "Hết hiệu lực")}>Hết hiệu lực</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!filteredTemplates.length} text="Chưa có tài liệu phù hợp với bộ lọc hiện tại." />
      </Panel>
    </section>
  );
}

function ReportsPage({
  acceptanceBatches,
  acceptances,
  boqItems,
  collections,
  contracts,
  dailyLogs,
  equipmentList,
  equipmentTransfers,
  filters,
  invoices,
  laborTeams,
  materialIns,
  onFilterChange,
  otherCosts,
  subContracts,
  wbsPlans,
}) {
  const reportData = getConstructionReportsData({
    acceptanceBatches,
    acceptances,
    boqItems,
    collections,
    contracts,
    dailyLogs,
    equipmentList,
    equipmentTransfers,
    filters,
    invoices,
    laborTeams,
    materialIns,
    otherCosts,
    subContracts,
    wbsPlans,
  });

  return (
    <section className="reports-layout">
      <Panel title="Bộ lọc báo cáo">
        <div className="work-form">
          <Field label="Lọc theo công trình">
            <select name="project" value={filters.project} onChange={onFilterChange}>
              <option>{ALL_PROJECTS}</option>
              {reportData.projectOptions.map((project) => (
                <option key={project}>{project}</option>
              ))}
            </select>
          </Field>
          <Field label="Từ ngày">
            <input name="fromDate" type="date" value={filters.fromDate} onChange={onFilterChange} />
          </Field>
          <Field label="Đến ngày">
            <input name="toDate" type="date" value={filters.toDate} onChange={onFilterChange} />
          </Field>
        </div>
      </Panel>

      <Panel title="Cảnh báo điều hành">
        <div className="warning-list">
          {reportData.alerts.map((alert) => (
            <span className="status" key={alert}>{alert}</span>
          ))}
        </div>
        <EmptyState show={!reportData.alerts.length} text="Chưa có cảnh báo theo dữ liệu hiện tại." />
      </Panel>

      <Panel title="A. Báo cáo tổng quan dự án">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Giá trị hợp đồng</th>
              <th>Tổng giá trị BOQ</th>
              <th>Giá trị thực hiện</th>
              <th>KL kế hoạch</th>
              <th>Khối lượng thực hiện</th>
              <th>KL nghiệm thu</th>
              <th>Đã xuất hóa đơn</th>
              <th>Đã thu tiền</th>
              <th>Công nợ</th>
              <th>Chi phí thực tế</th>
              <th>Lãi/lỗ</th>
              <th>Biên LN</th>
            </tr>
          </thead>
          <tbody>
            {reportData.overviewRows.map((item) => (
              <tr key={item.hopDongCongTrinh}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{formatCurrency(item.giaTriHopDong)}</td>
                <td>{formatCurrency(item.tongGiaTriBoq)}</td>
                <td>{formatCurrency(item.tongGiaTriThucHien)}</td>
                <td>{formatNumber(item.tongKhoiLuongKeHoach)}</td>
                <td>{formatNumber(item.tongKhoiLuongThucHien)}</td>
                <td>{formatNumber(item.tongKhoiLuongNghiemThu)}</td>
                <td>{formatCurrency(item.tongDaXuatHoaDon)}</td>
                <td>{formatCurrency(item.tongDaThuTien)}</td>
                <td>{formatCurrency(item.tongCongNo)}</td>
                <td>{formatCurrency(item.tongChiPhiThucTe)}</td>
                <td>{formatCurrency(item.laiLoTamTinh)}</td>
                <td>{formatPercent(item.bienLoiNhuan)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="B. Báo cáo tiến độ">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>Bắt đầu KH</th>
              <th>Kết thúc KH</th>
              <th>KL kế hoạch</th>
              <th>KL thực hiện</th>
              <th>Hoàn thành</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportData.progressRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{formatDate(item.ngayBatDauKeHoach)}</td>
                <td>{formatDate(item.ngayKetThucKeHoach)}</td>
                <td>{formatNumber(item.khoiLuongKeHoach)}</td>
                <td>{formatNumber(item.khoiLuongThucHien)}</td>
                <td>{formatPercent(item.tyLeHoanThanh)}</td>
                <td><span className={item.trangThai === "Hoàn thành" ? "status done" : "status"}>{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!reportData.progressRows.length} text="Chưa có dữ liệu tiến độ WBS." />
      </Panel>

      <Panel title="C. Báo cáo khối lượng thực hiện">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã WBS</th>
              <th>Hạng mục</th>
              <th>ĐVT</th>
              <th>KL hợp đồng</th>
              <th>KL thực hiện</th>
              <th>KL nghiệm thu</th>
              <th>Giá trị thực hiện tạm tính</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportData.quantityRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maWbs}</td>
                <td>{item.hangMucCongViec}</td>
                <td>{item.dvt || "-"}</td>
                <td>{formatNumber(item.khoiLuongHopDong)}</td>
                <td>{formatNumber(item.khoiLuongThucHien)}</td>
                <td>{formatNumber(item.khoiLuongNghiemThu)}</td>
                <td>{formatCurrency(item.giaTriThucHienTamTinh)}</td>
                <td><span className={item.trangThai === "Đã nghiệm thu đủ" ? "status done" : "status"}>{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="D. Báo cáo vật tư">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã vật tư</th>
              <th>Tên vật tư</th>
              <th>Tổng nhập</th>
              <th>Tổng sử dụng</th>
              <th>Tồn kho</th>
              <th>Định mức BOQ</th>
              <th>Chênh lệch</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportData.materialRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maVatTu || "-"}</td>
                <td>{item.tenVatTu}</td>
                <td>{formatNumber(item.tongNhap)}</td>
                <td>{formatNumber(item.tongSuDung)}</td>
                <td>{formatNumber(item.tonKho)}</td>
                <td>{formatNumber(item.dinhMucBoq)}</td>
                <td>{formatNumber(item.chenhLechDinhMuc)}</td>
                <td><span className={item.trangThai === "Đủ" ? "status done" : "status"}>{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="E. Báo cáo nhân công">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Tổ đội</th>
              <th>Tổng số người</th>
              <th>Tổng số giờ</th>
              <th>Công quy đổi</th>
              <th>Đơn giá công/ngày</th>
              <th>Chi phí nhân công</th>
            </tr>
          </thead>
          <tbody>
            {reportData.laborRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.toDoi}</td>
                <td>{formatNumber(item.tongSoNguoi)}</td>
                <td>{formatNumber(item.tongSoGio)}</td>
                <td>{formatNumber(item.tongCongQuyDoi)}</td>
                <td>{item.hasLaborRate ? formatCurrency(item.donGiaCongNgay) : item.status}</td>
                <td>{item.hasLaborRate ? formatCurrency(item.chiPhiNhanCong) : item.status}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="F. Báo cáo máy móc thiết bị">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Mã thiết bị</th>
              <th>Tên thiết bị</th>
              <th>Tổng giờ hoạt động</th>
              <th>Số ngày tại công trình</th>
              <th>Chi phí thiết bị</th>
              <th>Trạng thái</th>
            </tr>
          </thead>
          <tbody>
            {reportData.equipmentRows.map((item) => (
              <tr key={item.key}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{item.maThietBi || "-"}</td>
                <td>{item.tenThietBi}</td>
                <td>{formatNumber(item.tongGioHoatDong)}</td>
                <td>{formatNumber(item.soNgayTaiCongTrinh)}</td>
                <td>{formatCurrency(item.chiPhiThietBi)}</td>
                <td><span className="status">{item.trangThai}</span></td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="G. Báo cáo nghiệm thu">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>KL thực hiện</th>
              <th>KL nghiệm thu</th>
              <th>Giá trị nghiệm thu</th>
            </tr>
          </thead>
          <tbody>
            {reportData.acceptanceRows.map((item) => (
              <tr key={item.hopDongCongTrinh}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{formatNumber(item.tongKhoiLuongThucHien)}</td>
                <td>{formatNumber(item.tongKhoiLuongNghiemThu)}</td>
                <td>{formatCurrency(item.giaTriNghiemThu)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="H. Báo cáo doanh thu - công nợ">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Giá trị hợp đồng</th>
              <th>Giá trị nghiệm thu</th>
              <th>Đã xuất hóa đơn</th>
              <th>Đã thu tiền</th>
              <th>Công nợ còn phải thu</th>
              <th>Tỷ lệ thu tiền</th>
            </tr>
          </thead>
          <tbody>
            {reportData.revenueRows.map((item) => (
              <tr key={item.hopDongCongTrinh}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{formatCurrency(item.giaTriHopDong)}</td>
                <td>{formatCurrency(item.giaTriNghiemThu)}</td>
                <td>{formatCurrency(item.giaTriDaXuatHoaDon)}</td>
                <td>{formatCurrency(item.giaTriDaThuTien)}</td>
                <td>{formatCurrency(item.congNoConPhaiThu)}</td>
                <td>{formatPercent(item.tyLeThuTienHoaDon)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="I. Báo cáo chi phí">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Chi phí vật tư</th>
              <th>Chi phí nhân công</th>
              <th>Chi phí thiết bị</th>
              <th>Chi phí khác đã duyệt</th>
              <th>Tổng chi phí thực tế</th>
            </tr>
          </thead>
          <tbody>
            {reportData.costRows.map((item) => (
              <tr key={item.hopDongCongTrinh}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{formatCurrency(item.chiPhiVatTu)}</td>
                <td>{formatCurrency(item.chiPhiNhanCong)}</td>
                <td>{formatCurrency(item.chiPhiThietBi)}</td>
                <td>{formatCurrency(item.chiPhiKhac)}</td>
                <td>{formatCurrency(item.tongChiPhiThucTe)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <Panel title="J. Báo cáo lãi/lỗ công trình">
        <ResponsiveTable>
          <thead>
            <tr>
              <th>Hợp đồng / Công trình</th>
              <th>Doanh thu ghi nhận</th>
              <th>Tổng chi phí thực tế</th>
              <th>Lãi/lỗ tạm tính</th>
              <th>Biên lợi nhuận</th>
            </tr>
          </thead>
          <tbody>
            {reportData.profitRows.map((item) => (
              <tr key={item.hopDongCongTrinh}>
                <td>{item.hopDongCongTrinh}</td>
                <td>{formatCurrency(item.doanhThuGhiNhan)}</td>
                <td>{formatCurrency(item.tongChiPhiThucTe)}</td>
                <td>{formatCurrency(item.laiLoTamTinh)}</td>
                <td>{formatPercent(item.bienLoiNhuan)}</td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
      </Panel>

      <EmptyState show={!reportData.overviewRows.length} text="Chưa có dữ liệu để lập báo cáo tổng hợp." />
    </section>
  );
}

function SystemCatalogPage({
  activeTab,
  catalogs,
  editingItemId,
  form,
  isFormOpen,
  onCancel,
  onChange,
  onChangeSearch,
  onChangeTab,
  onEdit,
  onOpenForm,
  onSubmit,
  search,
  tabs,
}) {
  const activeDefinition = tabs.find((tab) => tab.key === activeTab) ?? tabs[0];
  const items = catalogs[activeDefinition.key] || [];
  const normalizedSearch = normalizeText(search);
  const filteredItems = normalizedSearch
    ? items.filter((item) =>
        activeDefinition.columns.some((column) => normalizeText(item[column.key]).includes(normalizedSearch)),
      )
    : items;
  const activeItems = items.filter((item) => !isInactiveCatalogItem(item)).length;
  const updatedItems = items.filter((item) => item.updatedAt).length;

  return (
    <section className="system-catalog-layout">
      <div className="document-tabs system-catalog-tabs">
        {tabs.map((tab) => (
          <button
            className={activeTab === tab.key ? "tab-button active" : "tab-button"}
            key={tab.key}
            onClick={() => onChangeTab(tab.key)}
            type="button"
          >
            {tab.label}
          </button>
        ))}
      </div>

      <div className="page-grid system-catalog-summary">
        <SummaryCard label="Tổng số bản ghi" value={items.length} tone="blue" />
        <SummaryCard label="Đang sử dụng" value={activeItems} tone="green" />
        <SummaryCard label="Đã chỉnh sửa" value={updatedItems} tone="amber" />
        <SummaryCard label="Kết quả tìm kiếm" value={filteredItems.length} tone="red" />
      </div>

      <Panel
        title={`Danh mục ${activeDefinition.label}`}
        action={
          !isFormOpen && (
            <button type="button" onClick={onOpenForm}>
              Thêm mới
            </button>
          )
        }
      >
        <div className="catalog-toolbar">
          <label className="field catalog-search">
            <span>Tìm kiếm</span>
            <input value={search} onChange={(event) => onChangeSearch(event.target.value)} placeholder={`Tìm trong ${activeDefinition.label}`} />
          </label>
        </div>

        {isFormOpen && (
          <form className="work-form catalog-form" onSubmit={onSubmit}>
            {activeDefinition.columns.map((column) => (
              <Field key={column.key} label={`${column.label}${column.required ? " *" : ""}`} wide={column.wide}>
                {renderCatalogInput(column, form, onChange)}
              </Field>
            ))}

            <div className="form-actions">
              <button type="submit">{editingItemId ? "Lưu chỉnh sửa" : "Lưu"}</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}

        <ResponsiveTable>
          <thead>
            <tr>
              {activeDefinition.columns.map((column) => (
                <th key={column.key}>{column.label}</th>
              ))}
              <th>Cập nhật</th>
              <th>Thao tác</th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.map((item) => (
              <tr key={item.id}>
                {activeDefinition.columns.map((column) => (
                  <td key={column.key}>{formatCatalogCellValue(item[column.key], column)}</td>
                ))}
                <td>{item.updatedAt || "-"}</td>
                <td>
                  <button className="secondary" type="button" onClick={() => onEdit(item)}>
                    Chỉnh sửa
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!filteredItems.length} text="Chưa có dữ liệu phù hợp. Bấm Thêm mới để tạo bản ghi danh mục." />
      </Panel>
    </section>
  );
}

function SystemSettingsPage({
  activeTab,
  catalogs,
  editingItemId,
  form,
  isFormOpen,
  onCancel,
  onChange,
  onChangeSearch,
  onChangeTab,
  onEdit,
  onExport,
  onImport,
  onOpenForm,
  onRestore,
  onSubmit,
  onTogglePermission,
  search,
  settings,
  tabs,
}) {
  const activeDefinition = tabs.find((tab) => tab.key === activeTab) ?? tabs[0];
  const items = activeTab === "users" ? catalogs.users || [] : settings[activeTab] || [];
  const normalizedSearch = normalizeText(search);
  const filteredItems = activeDefinition.columns && normalizedSearch
    ? items.filter((item) =>
        activeDefinition.columns.some((column) => normalizeText(item[column.key]).includes(normalizedSearch)),
      )
    : items;

  return (
    <section className="system-settings-layout">
      <div className="document-tabs system-catalog-tabs">
        {tabs.map((tab) => (
          <button
            className={activeTab === tab.key ? "tab-button active" : "tab-button"}
            key={tab.key}
            onClick={() => onChangeTab(tab.key)}
            type="button"
          >
            {tab.label}
          </button>
        ))}
      </div>

      {activeDefinition.columns && (
        <CatalogEditorPanel
          definition={activeDefinition}
          editingItemId={editingItemId}
          filteredItems={filteredItems}
          form={form}
          isFormOpen={isFormOpen}
          items={items}
          onCancel={onCancel}
          onChange={onChange}
          onChangeSearch={onChangeSearch}
          onEdit={onEdit}
          onOpenForm={onOpenForm}
          onSubmit={onSubmit}
          search={search}
        />
      )}

      {activeTab === "permissions" && (
        <PermissionsPanel permissions={settings.permissions || []} onTogglePermission={onTogglePermission} />
      )}

      {activeTab === "approvalWorkflows" && (
        <ApprovalWorkflowsPanel workflows={settings.approvalWorkflows || []} />
      )}

      {activeTab === "systemLogs" && (
        <SystemLogsPanel logs={settings.systemLogs || []} />
      )}

      {activeTab === "backup" && (
        <BackupPanel onExport={onExport} onImport={onImport} onRestore={onRestore} />
      )}
    </section>
  );
}

function CatalogEditorPanel({
  definition,
  editingItemId,
  filteredItems,
  form,
  isFormOpen,
  items,
  onCancel,
  onChange,
  onChangeSearch,
  onEdit,
  onOpenForm,
  onSubmit,
  search,
}) {
  const activeItems = items.filter((item) => !isInactiveCatalogItem(item)).length;
  const updatedItems = items.filter((item) => item.updatedAt).length;

  return (
    <>
      <div className="page-grid system-catalog-summary">
        <SummaryCard label="Tổng số bản ghi" value={items.length} tone="blue" />
        <SummaryCard label="Đang sử dụng" value={activeItems} tone="green" />
        <SummaryCard label="Đã chỉnh sửa" value={updatedItems} tone="amber" />
        <SummaryCard label="Kết quả tìm kiếm" value={filteredItems.length} tone="red" />
      </div>

      <Panel
        title={definition.label}
        action={
          !isFormOpen && (
            <button type="button" onClick={onOpenForm}>
              Thêm mới
            </button>
          )
        }
      >
        <div className="catalog-toolbar">
          <label className="field catalog-search">
            <span>Tìm kiếm</span>
            <input value={search} onChange={(event) => onChangeSearch(event.target.value)} placeholder={`Tìm trong ${definition.label}`} />
          </label>
        </div>

        {isFormOpen && (
          <form className="work-form catalog-form" onSubmit={onSubmit}>
            {definition.columns.map((column) => (
              <Field key={column.key} label={`${column.label}${column.required ? " *" : ""}`} wide={column.wide}>
                {renderCatalogInput(column, form, onChange)}
              </Field>
            ))}

            <div className="form-actions">
              <button type="submit">{editingItemId ? "Lưu chỉnh sửa" : "Lưu"}</button>
              <button type="button" className="secondary" onClick={onCancel}>
                Hủy
              </button>
            </div>
          </form>
        )}

        <ResponsiveTable>
          <thead>
            <tr>
              {definition.columns.map((column) => (
                <th key={column.key}>{column.label}</th>
              ))}
              <th>Cập nhật</th>
              <th>Thao tác</th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.map((item) => (
              <tr key={item.id}>
                {definition.columns.map((column) => (
                  <td key={column.key}>{formatCatalogCellValue(item[column.key], column)}</td>
                ))}
                <td>{item.updatedAt || "-"}</td>
                <td>
                  <button className="secondary" type="button" onClick={() => onEdit(item)}>
                    Chỉnh sửa
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </ResponsiveTable>
        <EmptyState show={!filteredItems.length} text="Chưa có dữ liệu phù hợp. Bấm Thêm mới để tạo bản ghi." />
      </Panel>
    </>
  );
}

function PermissionsPanel({ permissions, onTogglePermission }) {
  const [selectedRoleCode, setSelectedRoleCode] = useState(constructionRoleDefinitions[0].maVaiTro);
  const permissionColumns = [
    { key: "xem", label: "Xem" },
    { key: "them", label: "Thêm" },
    { key: "sua", label: "Sửa" },
    { key: "xoa", label: "Xóa" },
    { key: "duyet", label: "Duyệt" },
  ];
  const selectedRolePermissions = permissions.filter((row) => row.roleCode === selectedRoleCode);

  return (
    <Panel title="Ma trận phân quyền">
      <div className="catalog-toolbar">
        <label className="field catalog-search">
          <span>Vai trò</span>
          <select value={selectedRoleCode} onChange={(event) => setSelectedRoleCode(event.target.value)}>
            {constructionRoleDefinitions.map((role) => (
              <option key={role.maVaiTro} value={role.maVaiTro}>
                {role.maVaiTro} - {role.tenVaiTro}
              </option>
            ))}
          </select>
        </label>
      </div>
      <ResponsiveTable>
        <thead>
          <tr>
            <th>Module</th>
            {permissionColumns.map((column) => (
              <th key={column.key}>{column.label}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {selectedRolePermissions.map((row) => (
            <tr key={row.id}>
              <td>{row.module}</td>
              {permissionColumns.map((column) => (
                <td key={column.key}>
                  <input
                    aria-label={`${row.module} ${column.label}`}
                    checked={Boolean(row[column.key])}
                    onChange={() => onTogglePermission(row.id, column.key)}
                    type="checkbox"
                  />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </ResponsiveTable>
    </Panel>
  );
}

function ApprovalWorkflowsPanel({ workflows }) {
  return (
    <Panel title="Quy trình phê duyệt">
      <div className="workflow-grid">
        {workflows.map((workflow) => (
          <article className="workflow-card" key={workflow.id}>
            <strong>{workflow.name}</strong>
            <div className="workflow-steps">
              {workflow.steps.map((step, index) => (
                <Fragment key={`${workflow.id}-${step}`}>
                  <span>{step}</span>
                  {index < workflow.steps.length - 1 && <b>↓</b>}
                </Fragment>
              ))}
            </div>
          </article>
        ))}
      </div>
    </Panel>
  );
}

function SystemLogsPanel({ logs }) {
  return (
    <Panel title="Nhật ký hệ thống">
      <ResponsiveTable>
        <thead>
          <tr>
            <th>Người dùng</th>
            <th>Thời gian</th>
            <th>Chức năng</th>
            <th>Hành động</th>
          </tr>
        </thead>
        <tbody>
          {logs.map((log) => (
            <tr key={log.id}>
              <td>{log.nguoiDung}</td>
              <td>{log.thoiGian}</td>
              <td>{log.chucNang}</td>
              <td>{log.hanhDong}</td>
            </tr>
          ))}
        </tbody>
      </ResponsiveTable>
    </Panel>
  );
}

function BackupPanel({ onExport, onImport, onRestore }) {
  return (
    <Panel title="Sao lưu dữ liệu">
      <div className="backup-actions">
        <button type="button" onClick={onExport}>Xuất dữ liệu JSON</button>
        <label className="button-like secondary">
          Nhập dữ liệu JSON
          <input accept="application/json" onChange={onImport} type="file" />
        </label>
        <button className="secondary" type="button" onClick={onRestore}>Khôi phục dữ liệu</button>
      </div>
    </Panel>
  );
}

function LoginPage({ error, form, onChange, onSubmit }) {
  return (
    <main className="login-shell">
      <form className="login-card" onSubmit={onSubmit}>
        <div className="login-brand">
          <span className="brand-mark">TC</span>
          <div>
            <p className="eyebrow">Sơn Hòa Bình</p>
            <h1>Quản Lý Thi Công</h1>
          </div>
        </div>

        <label className="field">
          <span>Email</span>
          <input
            autoComplete="email"
            name="email"
            onChange={onChange}
            placeholder="admin@sonhoabinh.vn"
            type="email"
            value={form.email}
          />
        </label>

        <label className="field">
          <span>Mật khẩu</span>
          <input
            autoComplete="current-password"
            name="matKhau"
            onChange={onChange}
            placeholder="123456"
            type="password"
            value={form.matKhau}
          />
        </label>

        {error && <div className="login-error">{error}</div>}

        <button type="submit">Đăng nhập</button>
      </form>
    </main>
  );
}

function renderCatalogInput(column, form, onChange) {
  if (column.type === "roles") {
    const selectedRoles = Array.isArray(form[column.key]) ? form[column.key] : normalizeUserRoles(form);

    return (
      <div className="role-picker">
        {constructionRoleDefinitions.map((role) => (
          <label className="role-check" key={role.maVaiTro}>
            <input
              checked={selectedRoles.includes(role.maVaiTro)}
              onChange={() => {
                const nextRoles = selectedRoles.includes(role.maVaiTro)
                  ? selectedRoles.filter((roleCode) => roleCode !== role.maVaiTro)
                  : [...selectedRoles, role.maVaiTro];
                onChange({ target: { name: column.key, value: nextRoles } });
              }}
              type="checkbox"
            />
            <span>{role.maVaiTro}</span>
          </label>
        ))}
      </div>
    );
  }

  if (column.type === "select") {
    return (
      <select name={column.key} value={form[column.key] || ""} onChange={onChange}>
        {(column.options || []).map((option) => (
          <option key={option}>{option}</option>
        ))}
      </select>
    );
  }

  if (column.wide) {
    return <textarea name={column.key} rows="3" value={form[column.key] || ""} onChange={onChange} />;
  }

  return <input name={column.key} type={column.type || "text"} value={form[column.key] || ""} onChange={onChange} />;
}

function formatCatalogCellValue(value, column) {
  if (column.type === "roles") return <RoleBadges roles={value} />;
  if (column.type === "password") return value ? "••••••" : "-";
  if (column.type === "number") return formatNumber(value);
  if (column.type === "date") return formatDate(value);
  return value || "-";
}

function RoleBadges({ roles }) {
  const roleCodes = Array.isArray(roles) ? roles : roles ? [roles] : [];
  if (!roleCodes.length) return "-";

  return (
    <div className="role-badges">
      {roleCodes.map((roleCode) => (
        <span key={roleCode}>{roleCode}</span>
      ))}
    </div>
  );
}

function isInactiveCatalogItem(item) {
  return ["ngưng", "nghi viec", "nghi việc", "khóa", "khoa", "tạm ngưng", "tam ngung"].some((token) =>
    normalizeText(item.trangThai).includes(normalizeText(token)),
  );
}

function PlaceholderPage({ title }) {
  return (
    <Panel title={title}>
      <p className="placeholder-text">Chức năng chi tiết sẽ được phát triển ở bước tiếp theo.</p>
    </Panel>
  );
}

function SummaryCard({ label, value, tone }) {
  return (
    <article className={`summary-card ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function Panel({ title, action = null, className = "", children }) {
  return (
    <section className={`panel ${className}`}>
      <div className="panel-heading">
        <h2>{title}</h2>
        {action}
      </div>
      {children}
    </section>
  );
}

function Field({ label, wide = false, children }) {
  return (
    <label className={wide ? "field wide" : "field"}>
      <span>{label}</span>
      {children}
    </label>
  );
}

function SectionLabel({ title }) {
  return <h3 className="section-label">{title}</h3>;
}

function ResponsiveTable({ children }) {
  return (
    <div className="table-wrap">
      <table>{children}</table>
    </div>
  );
}

function EmptyState({ show, text }) {
  if (!show) return null;
  return <p className="empty-state">{text}</p>;
}

function downloadExcelTemplate(templateKey) {
  if (templateKey === "boq") {
    downloadBoqEstimateTemplate();
    return;
  }

  const workbook = XLSX.utils.book_new();
  const template = getExcelTemplate(templateKey);

  template.sheets.forEach((sheet) => {
    const worksheet = XLSX.utils.aoa_to_sheet([sheet.headers, ...(sheet.rows || [])]);
    (sheet.formulas || []).forEach(({ cell, formula }) => {
      worksheet[cell] = worksheet[cell] || { t: "n" };
      worksheet[cell].f = formula;
    });
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  });

  XLSX.writeFile(workbook, template.fileName);
}

function getExcelTemplate(templateKey) {
  const templates = {
    contracts: {
      fileName: "mau-hop-dong.xlsx",
      sheets: [
        {
          name: "Hop dong",
          headers: ["Mã hợp đồng", "Tên công trình", "Khách hàng", "Địa chỉ", "Giá trị hợp đồng", "Ngày ký", "Ngày khởi công", "Ngày hoàn thành kế hoạch", "PM phụ trách", "Giám sát phụ trách", "Trạng thái"],
          rows: [["HD-2026-001", "Biệt thự Thảo Điền", "Công ty TNHH Minh Phú", "Quận 2, TP.HCM", 1200000000, "2026-06-01", "2026-06-10", "2026-09-30", "Nguyễn Văn A", "Trần Văn B", "Đang thi công"]],
        },
      ],
    },
    boq: {
      fileName: "mau-boq-du-toan.xlsx",
      sheets: [
        {
          name: "Thông tin dự toán",
          headers: ["Mã hợp đồng", "Tên công trình", "Địa chỉ", "Ngày làm dự toán", "Chủng loại", "Khối lượng m2", "Đơn giá bán", "Doanh số"],
          rows: [["HD-2026-001", "Biệt thự Thảo Điền", "Quận 2, TP.HCM", "2026-06-11", "Sơn đá", 1000, 350000, ""]],
          formulas: [{ cell: "H2", formula: "F2*G2" }],
        },
        {
          name: "BOQ vật tư chính",
          headers: ["Mã WBS", "Hạng mục công việc", "Mã vật tư", "Tên vật tư", "ĐVT", "Định mức vật tư", "Khối lượng thi công", "Số lượng vật tư", "Đơn giá", "Thành tiền"],
          rows: [["WBS-01.01", "Sơn lót tường trong", "VT-001", "Sơn lót kháng kiềm", "kg", 0.12, 1000, "", 85000, ""]],
          formulas: [{ cell: "H2", formula: "F2*G2" }, { cell: "J2", formula: "H2*I2" }],
        },
        {
          name: "Chi phí khác",
          headers: ["Mã chi phí", "Nhóm chi phí", "Nội dung chi phí", "Tỷ lệ %", "Khối lượng", "Đơn giá", "Thành tiền"],
          rows: [["CP-001", "Vận chuyển", "Vận chuyển vật tư", 0, 1, 5000000, ""]],
          formulas: [{ cell: "G2", formula: "E2*F2" }],
        },
        {
          name: "Ngân lưu",
          headers: ["Đợt", "Nội dung", "% giá trị hoàn thành", "Ngày dự kiến thu", "Khối lượng hoàn thành lũy kế", "Thành tiền"],
          rows: [["Đợt 1", "Tạm ứng/thu theo tiến độ", 30, "2026-07-15", 300, 105000000]],
        },
        {
          name: "Dự trù khấu trừ",
          headers: ["Đợt", "Nội dung", "Tổng giá trị phải khấu trừ", "% khấu trừ", "Ngày khấu trừ", "Thành tiền"],
          rows: [["Đợt 1", "Khấu trừ tạm ứng", 100000000, 20, "2026-08-15", 20000000]],
        },
        {
          name: "Dự trù chi nhân công",
          headers: ["Đợt", "Nội dung", "ĐVT", "Khối lượng theo hợp đồng", "Đơn giá", "% khối lượng hoàn thành", "Ngày chi", "Thành tiền"],
          rows: [["Đợt 1", "Chi nhân công sơn lót", "m2", 1000, 45000, 30, "2026-07-20", 13500000]],
        },
      ],
    },
    wbs: {
      fileName: "mau-tien-do-wbs.xlsx",
      sheets: [
        {
          name: "Tiến độ WBS",
          headers: ["Mã hợp đồng", "Tên công trình", "Mã WBS cha", "Mã WBS", "Nội dung công việc", "Khu vực / Tòa / Hạng mục", "Số công nhân", "Ngày bắt đầu", "Ngày kết thúc", "Số ngày", "Trạng thái", "Ghi chú"],
          rows: [["HD-2026-001", "Biệt thự Thảo Điền", "WBS-01", "WBS-01.01", "Sơn lót tường trong", "Tầng 1", 6, "2026-06-12", "2026-06-18", "", "Đang thi công", ""]],
          formulas: [{ cell: "J2", formula: "I2-H2+1" }],
        },
      ],
    },
    materials: {
      fileName: "mau-nhap-vat-tu.xlsx",
      sheets: [
        {
          name: "Nhập vật tư",
          headers: ["Mã hợp đồng", "Tên công trình", "Ngày nhập", "Mã vật tư", "Tên vật tư", "ĐVT", "Số lượng nhập", "Nguồn cấp", "Người giao", "Người nhận", "Ghi chú"],
          rows: [["HD-2026-001", "Biệt thự Thảo Điền", "2026-06-12", "VT-001", "Sơn lót kháng kiềm", "kg", 120, "Nhà máy", "Chị Mai", "Anh Tuấn", ""]],
        },
      ],
    },
    labor: {
      fileName: "mau-to-doi.xlsx",
      sheets: [
        {
          name: "Tổ đội",
          headers: ["Mã tổ đội", "Tên tổ đội", "Người đại diện", "Số điện thoại", "Loại tổ đội", "Đơn giá công/ngày", "Ghi chú"],
          rows: [["TD-001", "Tổ sơn đá A", "Anh Tuấn", "0909555666", "Sơn đá", 550000, ""]],
        },
      ],
    },
    equipment: {
      fileName: "mau-thiet-bi.xlsx",
      sheets: [
        {
          name: "Thiết bị",
          headers: ["Mã thiết bị", "Tên thiết bị", "Loại thiết bị", "Nguồn sở hữu", "Đơn vị sở hữu", "Đơn giá thuê/ngày", "Đơn giá thuê/giờ", "Tình trạng hiện tại", "Ghi chú"],
          rows: [["TB-001", "Máy phun sơn Graco", "Máy phun", "Công ty", "Sơn Hòa Bình", 0, 0, "Sẵn sàng", ""]],
        },
      ],
    },
    acceptance: {
      fileName: "mau-nghiem-thu.xlsx",
      sheets: [
        {
          name: "Nghiệm thu",
          headers: ["Số biên bản", "Mã hợp đồng", "Tên công trình", "Mã WBS", "Hạng mục công việc", "ĐVT", "Khối lượng hợp đồng", "Khối lượng đã thực hiện lũy kế", "Khối lượng nghiệm thu kỳ này", "Đơn giá hợp đồng", "Ghi chú"],
          rows: [["NT-001", "HD-2026-001", "Biệt thự Thảo Điền", "WBS-01.01", "Sơn lót tường trong", "m2", 1000, 300, 250, 350000, ""]],
        },
      ],
    },
  };

  return templates[templateKey];
}

function downloadBoqEstimateTemplate() {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ["CÔNG TY CP SƠN VÀ CHẤT PHỦ HÒA BÌNH", "", "", "", "", "Mã số", "DT-HB-01", ""],
    ["Hodastone", "", "", "", "", "Soát xét", "01", ""],
    ["", "", "", "", "", "Ngày hiệu lực", "11/06/2026", ""],
    [],
    ["BẢNG DỰ TOÁN CÔNG TRÌNH", "", "", "", "", "", "", ""],
    ["Hodastone", "", "", "", "", "", "", ""],
    [],
    ["Mã hợp đồng", "HD-2026-001", "", "", "Chủng loại", "Sơn đá", "", ""],
    ["Công trình", "Biệt thự Thảo Điền", "", "", "Doanh số (VNĐ)", 350000000, "", ""],
    ["Địa chỉ", "Quận 2, TP.HCM", "", "", "Khối lượng (m2)", 1000, "", ""],
    ["Ngày làm dự toán", "11/06/2026", "", "", "Đơn giá (đ/m2)", 350000, "", ""],
    [],
    ["STT", "Tên công việc", "ĐVT", "ĐMVT", "KLTC", "SLVT", "Đơn giá", "Thành tiền"],
    ["I", "Vật tư chính", "", "", "", "", "", ""],
    [1, "KEO LÓT (HCP)", "kg", 0.12, "$F$10", "", 85000, ""],
    [2, "HODA MASTIC", "kg", 1.2, "$F$10", "", 12000, ""],
    [3, "HGM", "kg", 0.35, "$F$10", "", 45000, ""],
    [4, "KEO PHỦ UV", "kg", 0.1, "$F$10", "", 125000, ""],
    ["", "Tổng vật tư chính", "", "", "", "", "", ""],
    ["II", "Các chi phí khác", "", "", "", "", "", ""],
    [1, "Chi phí nhân công, vật tư phụ, tháo lắp giàn", "", 0.18, "", "", "", ""],
    [2, "Vận chuyển", "", 0.02, "", "", "", ""],
    [3, "Thiết bị, giàn giáo", "", 0.04, "", "", "", ""],
    [4, "Chi phí quản lý trực tiếp + lương GS", "", 0.03, "", "", "", ""],
    [5, "Chi phí nghiệm thu và thiết kế", "", 0.01, "", "", "", ""],
    [6, "Chi phí công trình", "", 0.02, "", "", "", ""],
    [7, "Chi phí thưởng công trình", "", 0.01, "", "", "", ""],
    [8, "Chi phí thuê khác công trình", "", 0.01, "", "", "", ""],
    [9, "Chi phí dự phòng, bảo lưu", "", 0.02, "", "", "", ""],
    ["", "Tổng chi phí khác", "", "", "", "", "", ""],
    ["", "Tổng chi phí", "", "", "", "", "", ""],
    ["", "Lãi/Lỗ/Doanh số", "", "", "", "", "", ""],
    [],
    ["NGÂN LƯU", "", "", "", "", "", "", ""],
    ["STT", "Nội dung", "% giá trị hoàn thành so với HĐ", "Ngày dự kiến thu", "Khối lượng hoàn thành lũy kế (m2)", "Thành tiền", "", ""],
    [1, "Đợt 1", 0.3, "15/07/2026", 300, ""],
    [2, "Đợt 2", 0.4, "15/08/2026", 700, ""],
    [3, "Đợt 3", 0.3, "15/09/2026", 1000, ""],
    [],
    ["DỰ TRÙ KHẤU TRỪ CHI PHÍ VẬT TƯ, VẬN CHUYỂN", "", "", "", "", "", "", ""],
    ["Đợt", "Nội dung", "Tổng giá trị phải khấu trừ", "% khấu trừ", "Ngày khấu trừ", "Thành tiền", "", ""],
    ["Đợt 1", "Khấu trừ vật tư", 100000000, 0.2, "15/08/2026", ""],
    ["Đợt 2", "Khấu trừ vận chuyển", 20000000, 0.5, "15/09/2026", ""],
    [],
    ["DỰ TRÙ CHI NHÂN CÔNG", "", "", "", "", "", "", ""],
    ["Đợt", "Nội dung", "ĐVT", "Khối lượng theo hợp đồng", "Đơn giá", "% khối lượng hoàn thành", "Ngày chi", "Thành tiền"],
    ["Đợt 1", "Chi nhân công đợt 1", "m2", 1000, 45000, 0.3, "20/07/2026", ""],
    ["Đợt 2", "Chi nhân công đợt 2", "m2", 1000, 45000, 0.4, "20/08/2026", ""],
    ["Đợt 3", "Chi nhân công đợt 3", "m2", 1000, 45000, 0.3, "20/09/2026", ""],
    [],
    ["DUYỆT (TGĐ)", "", "DUYỆT (PTGĐ)", "", "", "PT. DỰ TOÁN", "", ""],
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(rows);

  [
    ["E15", "$F$10", 1000], ["F15", "D15*E15", 120], ["H15", "F15*G15", 10200000],
    ["E16", "$F$10", 1000], ["F16", "D16*E16", 1200], ["H16", "F16*G16", 14400000],
    ["E17", "$F$10", 1000], ["F17", "D17*E17", 350], ["H17", "F17*G17", 15750000],
    ["E18", "$F$10", 1000], ["F18", "D18*E18", 100], ["H18", "F18*G18", 12500000],
    ["H19", "SUM(H15:H18)", 52850000],
    ["H21", "IF(D21>0,$F$9*D21,G21)", 63000000],
    ["H22", "IF(D22>0,$F$9*D22,G22)", 7000000],
    ["H23", "IF(D23>0,$F$9*D23,G23)", 14000000],
    ["H24", "IF(D24>0,$F$9*D24,G24)", 10500000],
    ["H25", "IF(D25>0,$F$9*D25,G25)", 3500000],
    ["H26", "IF(D26>0,$F$9*D26,G26)", 7000000],
    ["H27", "IF(D27>0,$F$9*D27,G27)", 3500000],
    ["H28", "IF(D28>0,$F$9*D28,G28)", 3500000],
    ["H29", "IF(D29>0,$F$9*D29,G29)", 7000000],
    ["H30", "SUM(H21:H29)", 119000000],
    ["H31", "H19+H30", 171850000],
    ["H32", "IF($F$9>0,($F$9-H31)/$F$9,0)", 0.509],
    ["F36", "$F$9*C36", 105000000], ["F37", "$F$9*C37", 140000000], ["F38", "$F$9*C38", 105000000],
    ["F42", "C42*D42", 20000000], ["F43", "C43*D43", 10000000],
    ["H47", "D47*E47*F47", 13500000], ["H48", "D48*E48*F48", 18000000], ["H49", "D49*E49*F49", 13500000],
  ].forEach(([cell, formula, value]) => {
    worksheet[cell] = worksheet[cell] || { t: "n" };
    worksheet[cell].t = "n";
    worksheet[cell].f = formula;
    if (value !== undefined) worksheet[cell].v = value;
  });

  worksheet["!merges"] = [
    XLSX.utils.decode_range("A1:D1"),
    XLSX.utils.decode_range("A5:H5"),
    XLSX.utils.decode_range("A6:H6"),
    XLSX.utils.decode_range("B8:D8"),
    XLSX.utils.decode_range("B9:D9"),
    XLSX.utils.decode_range("B10:D10"),
    XLSX.utils.decode_range("B11:D11"),
    XLSX.utils.decode_range("A14:H14"),
    XLSX.utils.decode_range("A20:H20"),
    XLSX.utils.decode_range("A34:H34"),
    XLSX.utils.decode_range("A40:H40"),
    XLSX.utils.decode_range("A45:H45"),
    XLSX.utils.decode_range("A51:B51"),
    XLSX.utils.decode_range("C51:D51"),
    XLSX.utils.decode_range("F51:H51"),
  ];
  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 42 },
    { wch: 16 },
    { wch: 16 },
    { wch: 18 },
    { wch: 18 },
    { wch: 16 },
    { wch: 18 },
  ];
  worksheet["!margins"] = { left: 0.3, right: 0.3, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  worksheet["!pageSetup"] = { paperSize: 9, orientation: "portrait", fitToWidth: 1, fitToHeight: 0 };
  applyDuToanWorksheetFormat(worksheet);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Du toan");
  workbook.Workbook = { Views: [{ RTL: false }] };
  XLSX.writeFile(workbook, "mau-du-toan-hodastone.xlsx");
}

function applyDuToanWorksheetFormat(worksheet) {
  const moneyFormat = "#,##0";
  const percentFormat = "0.00%";
  const dateFormat = "dd/mm/yyyy";
  const borderStyle = {
    top: { style: "thin", color: { rgb: "999999" } },
    bottom: { style: "thin", color: { rgb: "999999" } },
    left: { style: "thin", color: { rgb: "999999" } },
    right: { style: "thin", color: { rgb: "999999" } },
  };
  const headerRows = new Set([13, 35, 41, 46]);
  const sectionRows = new Set([14, 20, 34, 40, 45]);
  const totalRows = new Set([19, 30, 31, 32]);

  Object.keys(worksheet).forEach((cellAddress) => {
    if (cellAddress.startsWith("!")) return;
    const cellRef = XLSX.utils.decode_cell(cellAddress);
    const row = cellRef.r + 1;
    const col = cellRef.c + 1;
    const cell = worksheet[cellAddress];
    cell.s = {
      alignment: { vertical: "center", wrapText: true },
      border: row >= 13 && row <= 49 ? borderStyle : undefined,
    };

    if (row === 5) cell.s = { ...cell.s, font: { bold: true, sz: 16 }, alignment: { horizontal: "center", vertical: "center" } };
    if (row === 6) cell.s = { ...cell.s, font: { bold: true, italic: true }, alignment: { horizontal: "center", vertical: "center" } };
    if (headerRows.has(row)) cell.s = { ...cell.s, font: { bold: true }, fill: { fgColor: { rgb: "E5E7EB" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: borderStyle };
    if (sectionRows.has(row)) cell.s = { ...cell.s, font: { bold: true }, fill: { fgColor: { rgb: "F3F4F6" } }, border: borderStyle };
    if (totalRows.has(row)) cell.s = { ...cell.s, font: { bold: true }, fill: { fgColor: { rgb: "FEF3C7" } }, border: borderStyle };
    if ([6, 7, 8].includes(col) && [9, 10, 11, 15, 16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 36, 37, 38, 42, 43, 47, 48, 49].includes(row)) cell.z = moneyFormat;
    if ((col === 4 && row >= 21 && row <= 29) || (col === 3 && row >= 36 && row <= 38) || (col === 4 && row >= 42 && row <= 43) || (col === 6 && row >= 47 && row <= 49) || cellAddress === "H32") cell.z = percentFormat;
    if ((col === 2 && row === 11) || (col === 4 && row >= 36 && row <= 38) || (col === 5 && row >= 42 && row <= 43) || (col === 7 && row >= 47 && row <= 49)) cell.z = dateFormat;
  });
}

function readExcelWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        resolve(XLSX.read(reader.result, { type: "array", cellDates: true }));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseExcelImport(workbook, templateKey) {
  if (templateKey === "boq") return parseBoqWorkbook(workbook);
  if (templateKey === "wbs") return parseWbsWorkbook(workbook);
  return parseSimpleWorkbook(workbook, templateKey);
}

function parseBoqWorkbook(workbook) {
  const worksheet = workbook.Sheets["Du toan"];
  const errors = [];

  if (!worksheet) {
    return {
      errors: [{ row: "-", column: "Sheet", message: "Không tìm thấy sheet Du toan", suggestion: "Dùng đúng file mẫu Dự toán Excel mới" }],
      boqItems: [],
      otherCosts: [],
      cashFlowPlans: [],
      deductionPlans: [],
      laborPaymentPlans: [],
    };
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
  const maHopDong = getAoAValue(rows, 8, 2);
  const tenCongTrinh = getAoAValue(rows, 9, 2);
  const hopDongCongTrinh = getProjectLabel(maHopDong, tenCongTrinh);
  const doanhSo = toExcelNumber(getAoAValue(rows, 9, 6));

  validateRequiredValue(maHopDong, 8, "Mã hợp đồng", errors);
  validateRequiredValue(tenCongTrinh, 9, "Công trình", errors);
  validateDateValue(getAoAValue(rows, 11, 2), 11, "Ngày làm dự toán", errors, false);
  validateNumberValue(getAoAValue(rows, 9, 6), 9, "Doanh số (VNĐ)", errors, false);
  validateNumberValue(getAoAValue(rows, 10, 6), 10, "Khối lượng (m2)", errors, false);
  validateNumberValue(getAoAValue(rows, 11, 6), 11, "Đơn giá (đ/m2)", errors, false);

  const materialRows = getRowsBetweenSection(rows, "Vật tư chính", "Tổng vật tư chính");
  const boqItems = materialRows.map(({ row, rowNumber }) => {
    const dmvt = toExcelNumber(row[3]);
    const kltc = toExcelNumber(row[4]);
    const slvt = toExcelNumber(row[5]) || dmvt * kltc;
    const unitPrice = toExcelNumber(row[6]);
    const amount = toExcelNumber(row[7]) || slvt * unitPrice;

    validateSharedContract(maHopDong, tenCongTrinh, rowNumber, errors);
    validateRequiredValue(row[1], rowNumber, "Tên công việc", errors);
    validateNumberValue(row[3], rowNumber, "ĐMVT", errors);
    validateNumberValue(row[4], rowNumber, "KLTC", errors);
    validateNumberValue(row[6], rowNumber, "Đơn giá", errors);
    validateNumberValue(amount, rowNumber, "Thành tiền", errors);

    return {
      id: crypto.randomUUID(),
      maHopDong,
      hopDongCongTrinh,
      maWbs: `DT-I-${String(row[0] || rowNumber).padStart(2, "0")}`,
      hangMucCongViec: row[1],
      maVatTu: `VT-${String(row[0] || rowNumber).padStart(3, "0")}`,
      tenVatTu: row[1],
      dvt: row[2],
      dinhMucVatTu: dmvt,
      khoiLuongHopDong: kltc,
      soLuongVatTu: slvt,
      donGiaHopDong: unitPrice,
      thanhTienHopDong: amount,
      createdAt: new Date().toLocaleString("vi-VN"),
    };
  });

  const otherRows = getRowsBetweenSection(rows, "Các chi phí khác", "Tổng chi phí khác");
  const otherCosts = otherRows.map(({ row, rowNumber }) => {
    const ratio = toExcelNumber(row[3]);
    const directAmount = toExcelNumber(row[7]);
    const amount = directAmount || (ratio > 0 ? doanhSo * ratio : 0);
    validateNumberValue(row[3], rowNumber, "Tỷ lệ %", errors, false);
    validateNumberValue(amount, rowNumber, "Thành tiền", errors, false);

    return {
      id: crypto.randomUUID(),
      maHopDong,
      hopDongCongTrinh,
      maChiPhi: `CP-${String(row[0] || rowNumber).padStart(3, "0")}`,
      nhomChiPhi: "Chi phí khác",
      noiDungChiPhi: row[1],
      tyLe: ratio,
      soTien: amount,
      trangThai: "Chờ duyệt",
      createdAt: new Date().toLocaleString("vi-VN"),
    };
  });

  const cashFlowPlans = getRowsBelowHeader(rows, "NGÂN LƯU", "DỰ TRÙ KHẤU TRỪ").map(({ row, rowNumber }) => {
    const ratio = toExcelNumber(row[2]);
    const amount = toExcelNumber(row[5]) || doanhSo * ratio;
    validateDateValue(row[3], rowNumber, "Ngày dự kiến thu", errors, false);
    validateNumberValue(row[2], rowNumber, "% giá trị hoàn thành so với HĐ", errors, false);
    validateNumberValue(amount, rowNumber, "Thành tiền", errors, false);

    return {
      id: crypto.randomUUID(),
      maHopDong,
      hopDongCongTrinh,
      dot: row[0],
      noiDung: row[1],
      phanTramGiaTriHoanThanh: ratio,
      ngayDuKienThu: normalizeDateValue(row[3]),
      khoiLuongHoanThanhLuyKe: row[4],
      thanhTien: amount,
    };
  });

  const deductionPlans = getRowsBelowHeader(rows, "DỰ TRÙ KHẤU TRỪ", "DỰ TRÙ CHI NHÂN CÔNG").map(({ row, rowNumber }) => {
    const totalDeduct = toExcelNumber(row[2]);
    const ratio = toExcelNumber(row[3]);
    const amount = toExcelNumber(row[5]) || totalDeduct * ratio;
    validateDateValue(row[4], rowNumber, "Ngày khấu trừ", errors, false);
    validateNumberValue(row[2], rowNumber, "Tổng giá trị phải khấu trừ", errors, false);
    validateNumberValue(row[3], rowNumber, "% khấu trừ", errors, false);
    validateNumberValue(amount, rowNumber, "Thành tiền", errors, false);

    return {
      id: crypto.randomUUID(),
      maHopDong,
      hopDongCongTrinh,
      dot: row[0],
      noiDung: row[1],
      tongGiaTriPhaiKhauTru: totalDeduct,
      phanTramKhauTru: ratio,
      ngayKhauTru: normalizeDateValue(row[4]),
      thanhTien: amount,
    };
  });

  const laborPaymentPlans = getRowsBelowHeader(rows, "DỰ TRÙ CHI NHÂN CÔNG", "DUYỆT").map(({ row, rowNumber }) => {
    const contractQuantity = toExcelNumber(row[3]);
    const unitPrice = toExcelNumber(row[4]);
    const ratio = toExcelNumber(row[5]);
    const amount = toExcelNumber(row[7]) || contractQuantity * unitPrice * ratio;
    validateDateValue(row[6], rowNumber, "Ngày chi", errors, false);
    validateNumberValue(row[3], rowNumber, "Khối lượng theo hợp đồng", errors, false);
    validateNumberValue(row[4], rowNumber, "Đơn giá", errors, false);
    validateNumberValue(row[5], rowNumber, "% khối lượng hoàn thành", errors, false);
    validateNumberValue(amount, rowNumber, "Thành tiền", errors, false);

    return {
      id: crypto.randomUUID(),
      maHopDong,
      hopDongCongTrinh,
      dot: row[0],
      noiDung: row[1],
      dvt: row[2],
      khoiLuongTheoHopDong: contractQuantity,
      donGia: unitPrice,
      phanTramKhoiLuongHoanThanh: ratio,
      ngayChi: normalizeDateValue(row[6]),
      thanhTien: amount,
    };
  });

  return { errors, boqItems, otherCosts, cashFlowPlans, deductionPlans, laborPaymentPlans };
}

function parseWbsWorkbook(workbook) {
  const errors = [];
  const wbsPlans = getSheetRows(workbook, "Tiến độ WBS").map((row, index) => {
    const rowNumber = index + 2;
    validateRequired(row, rowNumber, "Mã hợp đồng", errors);
    validateRequired(row, rowNumber, "Tên công trình", errors);
    validateRequired(row, rowNumber, "Mã WBS", errors);
    validateDateCell(row, rowNumber, "Ngày bắt đầu", errors);
    validateDateCell(row, rowNumber, "Ngày kết thúc", errors);
    validateNumberCell(row, rowNumber, "Số công nhân", errors, false);
    validateNumberCell(row, rowNumber, "Số ngày", errors, false);

    const maHopDong = getCellValue(row, "Mã hợp đồng");
    const tenCongTrinh = getCellValue(row, "Tên công trình");
    const startDate = normalizeDateValue(getCellValue(row, "Ngày bắt đầu"));
    const endDate = normalizeDateValue(getCellValue(row, "Ngày kết thúc"));
    return {
      id: crypto.randomUUID(),
      maHopDong,
      hopDongCongTrinh: getProjectLabel(maHopDong, tenCongTrinh),
      maWbsCha: getCellValue(row, "Mã WBS cha"),
      maWbs: getCellValue(row, "Mã WBS"),
      hangMucCongViec: getCellValue(row, "Nội dung công việc"),
      khuVucHangMuc: getCellValue(row, "Khu vực / Tòa / Hạng mục"),
      soCongNhan: getCellValue(row, "Số công nhân"),
      ngayBatDauKeHoach: startDate,
      ngayKetThucKeHoach: endDate,
      soNgayKeHoach: toNumber(getCellValue(row, "Số ngày")) || calculatePlanDays(startDate, endDate),
      trangThai: getCellValue(row, "Trạng thái") || WBS_STATUSES[0],
      ghiChu: getCellValue(row, "Ghi chú"),
      createdAt: new Date().toLocaleString("vi-VN"),
    };
  });

  return { errors, wbsPlans };
}

function parseSimpleWorkbook(workbook, templateKey) {
  const errors = [];
  const firstSheetName = workbook.SheetNames[0];
  const rows = getSheetRows(workbook, firstSheetName);
  const parsers = {
    contracts: parseContractExcelRow,
    materials: parseMaterialExcelRow,
    labor: parseLaborExcelRow,
    equipment: parseEquipmentExcelRow,
    acceptance: parseAcceptanceExcelRow,
  };
  const resultKey = {
    contracts: "contracts",
    materials: "materialIns",
    labor: "laborTeams",
    equipment: "equipmentList",
    acceptance: "acceptances",
  }[templateKey];

  return {
    errors,
    [resultKey]: rows.map((row, index) => parsers[templateKey](row, index + 2, errors)),
  };
}

function parseContractExcelRow(row, rowNumber, errors) {
  validateRequired(row, rowNumber, "Mã hợp đồng", errors);
  validateRequired(row, rowNumber, "Tên công trình", errors);
  validateNumberCell(row, rowNumber, "Giá trị hợp đồng", errors, false);
  validateDateCell(row, rowNumber, "Ngày ký", errors, false);
  validateDateCell(row, rowNumber, "Ngày khởi công", errors, false);
  validateDateCell(row, rowNumber, "Ngày hoàn thành kế hoạch", errors, false);
  return {
    id: crypto.randomUUID(),
    maHopDong: getCellValue(row, "Mã hợp đồng"),
    tenCongTrinh: getCellValue(row, "Tên công trình"),
    khachHang: getCellValue(row, "Khách hàng"),
    diaChiCongTrinh: getCellValue(row, "Địa chỉ"),
    giaTriHopDong: getCellValue(row, "Giá trị hợp đồng"),
    ngayKy: normalizeDateValue(getCellValue(row, "Ngày ký")),
    ngayKhoiCong: normalizeDateValue(getCellValue(row, "Ngày khởi công")),
    ngayHoanThanhKeHoach: normalizeDateValue(getCellValue(row, "Ngày hoàn thành kế hoạch")),
    pmPhuTrach: getCellValue(row, "PM phụ trách"),
    giamSatPhuTrach: getCellValue(row, "Giám sát phụ trách"),
    trangThai: getCellValue(row, "Trạng thái") || "Đang chuẩn bị",
    createdAt: new Date().toLocaleString("vi-VN"),
  };
}

function parseMaterialExcelRow(row, rowNumber, errors) {
  validateRequired(row, rowNumber, "Mã hợp đồng", errors);
  validateRequired(row, rowNumber, "Tên công trình", errors);
  validateDateCell(row, rowNumber, "Ngày nhập", errors);
  validateNumberCell(row, rowNumber, "Số lượng nhập", errors);
  const maHopDong = getCellValue(row, "Mã hợp đồng");
  return {
    id: crypto.randomUUID(),
    maHopDong,
    hopDongCongTrinh: getProjectLabel(maHopDong, getCellValue(row, "Tên công trình")),
    ngayNhap: normalizeDateValue(getCellValue(row, "Ngày nhập")),
    maVatTu: getCellValue(row, "Mã vật tư"),
    tenVatTu: getCellValue(row, "Tên vật tư"),
    dvt: getCellValue(row, "ĐVT"),
    soLuongNhap: getCellValue(row, "Số lượng nhập"),
    nguonCap: getCellValue(row, "Nguồn cấp") || "Nhà máy",
    nguoiGiao: getCellValue(row, "Người giao"),
    nguoiNhan: getCellValue(row, "Người nhận"),
    ghiChu: getCellValue(row, "Ghi chú"),
    createdAt: new Date().toLocaleString("vi-VN"),
  };
}

function parseLaborExcelRow(row, rowNumber, errors) {
  validateRequired(row, rowNumber, "Mã tổ đội", errors);
  validateRequired(row, rowNumber, "Tên tổ đội", errors);
  validateNumberCell(row, rowNumber, "Đơn giá công/ngày", errors, false);
  return {
    id: crypto.randomUUID(),
    maToDoi: getCellValue(row, "Mã tổ đội"),
    tenToDoi: getCellValue(row, "Tên tổ đội"),
    nguoiDaiDien: getCellValue(row, "Người đại diện"),
    soDienThoai: getCellValue(row, "Số điện thoại"),
    loaiToDoi: getCellValue(row, "Loại tổ đội") || "Sơn đá",
    donGiaCongNgay: getCellValue(row, "Đơn giá công/ngày"),
    ghiChu: getCellValue(row, "Ghi chú"),
    createdAt: new Date().toLocaleString("vi-VN"),
  };
}

function parseEquipmentExcelRow(row, rowNumber, errors) {
  validateRequired(row, rowNumber, "Mã thiết bị", errors);
  validateRequired(row, rowNumber, "Tên thiết bị", errors);
  validateNumberCell(row, rowNumber, "Đơn giá thuê/ngày", errors, false);
  validateNumberCell(row, rowNumber, "Đơn giá thuê/giờ", errors, false);
  return {
    id: crypto.randomUUID(),
    maThietBi: getCellValue(row, "Mã thiết bị"),
    tenThietBi: getCellValue(row, "Tên thiết bị"),
    loaiThietBi: getCellValue(row, "Loại thiết bị") || "Máy phun",
    nguonSoHuu: getCellValue(row, "Nguồn sở hữu") || "Công ty",
    donViSoHuu: getCellValue(row, "Đơn vị sở hữu"),
    donGiaThueNgay: getCellValue(row, "Đơn giá thuê/ngày"),
    donGiaThueGio: getCellValue(row, "Đơn giá thuê/giờ"),
    tinhTrangHienTai: getCellValue(row, "Tình trạng hiện tại") || "Sẵn sàng",
    ghiChu: getCellValue(row, "Ghi chú"),
    createdAt: new Date().toLocaleString("vi-VN"),
  };
}

function parseAcceptanceExcelRow(row, rowNumber, errors) {
  validateRequired(row, rowNumber, "Mã hợp đồng", errors);
  validateRequired(row, rowNumber, "Tên công trình", errors);
  validateRequired(row, rowNumber, "Mã WBS", errors);
  validateNumberCell(row, rowNumber, "Khối lượng hợp đồng", errors);
  validateNumberCell(row, rowNumber, "Khối lượng đã thực hiện lũy kế", errors);
  validateNumberCell(row, rowNumber, "Khối lượng nghiệm thu kỳ này", errors);
  validateNumberCell(row, rowNumber, "Đơn giá hợp đồng", errors);
  const maHopDong = getCellValue(row, "Mã hợp đồng");
  const acceptedThisPeriod = toNumber(getCellValue(row, "Khối lượng nghiệm thu kỳ này"));
  const unitPrice = toNumber(getCellValue(row, "Đơn giá hợp đồng"));
  return {
    id: crypto.randomUUID(),
    maHopDong,
    soBienBan: getCellValue(row, "Số biên bản"),
    hopDongCongTrinh: getProjectLabel(maHopDong, getCellValue(row, "Tên công trình")),
    maWbs: getCellValue(row, "Mã WBS"),
    hangMucCongViec: getCellValue(row, "Hạng mục công việc"),
    dvt: getCellValue(row, "ĐVT"),
    khoiLuongHopDong: getCellValue(row, "Khối lượng hợp đồng"),
    khoiLuongDaThucHienLuyKe: getCellValue(row, "Khối lượng đã thực hiện lũy kế"),
    khoiLuongNghiemThuKyNay: acceptedThisPeriod,
    khoiLuongNghiemThuLuyKe: acceptedThisPeriod,
    khoiLuongConLai: toNumber(getCellValue(row, "Khối lượng hợp đồng")) - acceptedThisPeriod,
    donGiaHopDong: unitPrice,
    thanhTienNghiemThuKyNay: acceptedThisPeriod * unitPrice,
    ghiChu: getCellValue(row, "Ghi chú"),
    createdAt: new Date().toLocaleString("vi-VN"),
  };
}

function getSheetRows(workbook, sheetName) {
  const worksheet = workbook.Sheets[sheetName] || workbook.Sheets[workbook.SheetNames[0]];
  if (!worksheet) return [];
  return XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false }).filter((row) =>
    Object.values(row).some((value) => String(value || "").trim()),
  );
}

function getAoAValue(rows, rowNumber, columnNumber) {
  return rows[rowNumber - 1]?.[columnNumber - 1] ?? "";
}

function getRowsBetweenSection(rows, startText, endText) {
  const startIndex = rows.findIndex((row) => row.some((cell) => normalizeText(cell).includes(normalizeText(startText))));
  if (startIndex < 0) return [];
  const endIndex = rows.findIndex((row, index) => index > startIndex && row.some((cell) => normalizeText(cell).includes(normalizeText(endText))));
  return rows
    .slice(startIndex + 1, endIndex > -1 ? endIndex : rows.length)
    .map((row, index) => ({ row, rowNumber: startIndex + 2 + index }))
    .filter(({ row }) => isMeaningfulImportRow(row) && !isTableHeaderRow(row));
}

function getRowsBelowHeader(rows, sectionText, nextSectionText) {
  const sectionIndex = rows.findIndex((row) => row.some((cell) => normalizeText(cell).includes(normalizeText(sectionText))));
  if (sectionIndex < 0) return [];
  const headerIndex = sectionIndex + 1;
  const nextSectionIndex = nextSectionText
    ? rows.findIndex((row, index) => index > headerIndex && row.some((cell) => normalizeText(cell).includes(normalizeText(nextSectionText))))
    : rows.length;
  return rows
    .slice(headerIndex + 1, nextSectionIndex > -1 ? nextSectionIndex : rows.length)
    .map((row, index) => ({ row, rowNumber: headerIndex + 2 + index }))
    .filter(({ row }) => isMeaningfulImportRow(row) && !isTableHeaderRow(row));
}

function isMeaningfulImportRow(row) {
  return row.some((cell) => String(cell || "").trim());
}

function isTableHeaderRow(row) {
  const normalized = row.map(normalizeText).join(" ");
  return normalized.includes("stt ten cong viec") || normalized.includes("noi dung") && normalized.includes("thanh tien");
}

function getCellValue(row, columnName) {
  return row[columnName] ?? "";
}

function validateRequiredValue(value, rowNumber, column, errors) {
  if (!String(value || "").trim()) {
    errors.push({ row: rowNumber, column, message: "Thiếu dữ liệu bắt buộc", suggestion: `Nhập giá trị cho ${column}` });
  }
}

function validateNumberValue(value, rowNumber, column, errors, required = true) {
  if (!required && (value === "" || value === null || value === undefined)) return;
  if (required && (value === "" || value === null || value === undefined)) {
    errors.push({ row: rowNumber, column, message: "Thiếu số liệu", suggestion: `Nhập số cho ${column}` });
    return;
  }
  if (Number.isNaN(toExcelNumber(value))) {
    errors.push({ row: rowNumber, column, message: "Giá trị không phải là số", suggestion: "Chỉ nhập số, không nhập chữ hoặc ký hiệu tiền tệ" });
  }
}

function validateDateValue(value, rowNumber, column, errors, required = true) {
  if (!required && !value) return;
  if (!normalizeDateValue(value)) {
    errors.push({ row: rowNumber, column, message: "Ngày không đúng định dạng", suggestion: "Dùng định dạng dd/mm/yyyy hoặc yyyy-mm-dd" });
  }
}

function toExcelNumber(value) {
  if (typeof value === "number") return value;
  if (value === "" || value === null || value === undefined) return 0;
  const normalized = String(value)
    .replace(/\s/g, "")
    .replace(/%$/, "")
    .replaceAll(",", "");
  const numberValue = Number(normalized);
  if (String(value).includes("%") && numberValue > 1) return numberValue / 100;
  return numberValue;
}

function validateSharedContract(maHopDong, tenCongTrinh, row, errors) {
  if (!maHopDong) errors.push({ row, column: "Mã hợp đồng", message: "Thiếu mã hợp đồng", suggestion: "Nhập Mã hợp đồng tại sheet Thông tin dự toán" });
  if (!tenCongTrinh) errors.push({ row, column: "Tên công trình", message: "Thiếu tên công trình", suggestion: "Nhập Tên công trình tại sheet Thông tin dự toán" });
}

function validateRequired(row, rowNumber, column, errors) {
  if (!String(getCellValue(row, column) || "").trim()) {
    errors.push({ row: rowNumber, column, message: "Thiếu dữ liệu bắt buộc", suggestion: `Nhập giá trị cho cột ${column}` });
  }
}

function validateNumberCell(row, rowNumber, column, errors, required = true) {
  const value = getCellValue(row, column);
  if (!required && value === "") return;
  if (required && value === "") {
    errors.push({ row: rowNumber, column, message: "Thiếu số liệu", suggestion: `Nhập số cho cột ${column}` });
    return;
  }
  if (Number.isNaN(Number(String(value).replaceAll(",", "")))) {
    errors.push({ row: rowNumber, column, message: "Giá trị không phải là số", suggestion: "Chỉ nhập số, không nhập chữ hoặc ký hiệu tiền tệ" });
  }
}

function validateDateCell(row, rowNumber, column, errors, required = true) {
  const value = getCellValue(row, column);
  if (!required && value === "") return;
  if (!normalizeDateValue(value)) {
    errors.push({ row: rowNumber, column, message: "Ngày không đúng định dạng", suggestion: "Dùng định dạng yyyy-mm-dd hoặc dd/mm/yyyy" });
  }
}

function normalizeDateValue(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (match) return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return "";
}

function getProjectLabel(maHopDong, tenCongTrinh) {
  return `${maHopDong || ""}${maHopDong ? " - " : ""}${tenCongTrinh || ""}`.trim();
}

function getContractCodeFromText(value) {
  return String(value || "").split(" - ")[0] || "";
}

function chooseExcelImportAction() {
  const choice = prompt("Chọn cách import:\n1. Thêm mới\n2. Ghi đè theo Mã hợp đồng\n3. Hủy import", "1");
  if (choice === "2") return "replace";
  if (choice === "3" || choice === null) return "cancel";
  return "append";
}

function mergeImportedRows(currentRows, importedRows, action, getKey) {
  if (action === "replace") {
    const importedKeys = new Set(importedRows.map(getKey).filter(Boolean));
    return [...importedRows, ...currentRows.filter((row) => !importedKeys.has(getKey(row)))];
  }
  return [...importedRows, ...currentRows];
}

function buildDefaultPermissionMatrix() {
  return constructionRoleDefinitions.flatMap((role) =>
    constructionModulePermissions.map((moduleName, moduleIndex) => {
      const roleDefaults = rolePermissionDefaults[role.maVaiTro] || {};
      const permissions = roleDefaults.all || roleDefaults[moduleName] || [];

      return {
        id: `PERM-${role.maVaiTro}-${moduleIndex + 1}`,
        roleCode: role.maVaiTro,
        roleName: role.tenVaiTro,
        module: moduleName,
        xem: permissions.includes("xem"),
        them: permissions.includes("them"),
        sua: permissions.includes("sua"),
        xoa: permissions.includes("xoa"),
        duyet: permissions.includes("duyet"),
      };
    }),
  );
}

function normalizeRolesCatalog(savedRoles) {
  const savedRoleCodes = new Set((savedRoles || []).map((role) => role.maVaiTro));
  const hasAllStandardRoles = constructionRoleDefinitions.every((role) => savedRoleCodes.has(role.maVaiTro));

  if (!hasAllStandardRoles || savedRoleCodes.size !== constructionRoleDefinitions.length) {
    return constructionRoleDefinitions;
  }

  return constructionRoleDefinitions.map((standardRole) => {
    const savedRole = savedRoles.find((role) => role.maVaiTro === standardRole.maVaiTro);
    return { ...standardRole, ...savedRole };
  });
}

function normalizePermissionMatrix(savedPermissions) {
  const hasRoleBasedRows = Array.isArray(savedPermissions) && savedPermissions.some((row) => row.roleCode);
  const hasAllModules = hasRoleBasedRows && constructionRoleDefinitions.every((role) =>
    constructionModulePermissions.every((moduleName) =>
      savedPermissions.some((row) => row.roleCode === role.maVaiTro && row.module === moduleName),
    ),
  );

  if (!hasAllModules) return defaultPermissionMatrix;
  return savedPermissions;
}

function normalizeUserItem(user) {
  return {
    ...user,
    // Demo only - production must hash password and use backend auth.
    matKhau: user.matKhau || user.password || "",
    sdt: user.sdt || user.soDienThoai || "",
    chucDanh: user.chucDanh || user.chucVu || "",
    congTrinhPhanCong: user.congTrinhPhanCong || user.congTrinh || user.projectName || "",
    roles: normalizeUserRoles(user),
    trangThai: user.trangThai || "Hoạt động",
  };
}

function getDefaultAuthUser() {
  return normalizeUserItem({
    id: "CAT-USR-001",
    maNguoiDung: "ADMIN",
    hoTen: "Admin hệ thống",
    email: "admin@sonhoabinh.vn",
    matKhau: "123456",
    sdt: "",
    chucDanh: "Quản trị hệ thống",
    congTrinhPhanCong: "Tất cả",
    roles: ["ADMIN"],
    trangThai: "Hoạt động",
  });
}

const DEFAULT_AUTH_USERS = [
  getDefaultAuthUser(),
  { id: "CAT-USR-TGD", maNguoiDung: "TGD", hoTen: "Tong Giam doc", email: "tgd@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Tong Giam doc", congTrinhPhanCong: "Tat ca", roles: ["TGD"], trangThai: "Hoat dong" },
  { id: "CAT-USR-GDTC", maNguoiDung: "GDTC", hoTen: "Giam doc Thi cong", email: "gdtc@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Giam doc Thi cong", congTrinhPhanCong: "Tat ca", roles: ["GDTC"], trangThai: "Hoat dong" },
  { id: "CAT-USR-PM", maNguoiDung: "PM", hoTen: "Giam doc Du an / PM", email: "pm@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Giam doc Du an / PM", congTrinhPhanCong: "Tat ca", roles: ["PM"], trangThai: "Hoat dong" },
  { id: "CAT-USR-CHT", maNguoiDung: "CHT", hoTen: "Chi huy truong", email: "cht@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Chi huy truong", congTrinhPhanCong: "Cong trinh duoc giao", roles: ["CHT"], trangThai: "Hoat dong" },
  { id: "CAT-USR-QS", maNguoiDung: "QS", hoTen: "QS", email: "qs@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "QS", congTrinhPhanCong: "Cong trinh duoc giao", roles: ["QS"], trangThai: "Hoat dong" },
  { id: "CAT-USR-QC", maNguoiDung: "QC", hoTen: "QC", email: "qc@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "QC", congTrinhPhanCong: "Cong trinh duoc giao", roles: ["QC"], trangThai: "Hoat dong" },
  { id: "CAT-USR-TKDA", maNguoiDung: "TKDA", hoTen: "Thu ky du an", email: "tkda@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Thu ky du an", congTrinhPhanCong: "Cong trinh duoc giao", roles: ["TKDA"], trangThai: "Hoat dong" },
  { id: "CAT-USR-KTCT", maNguoiDung: "KTCT", hoTen: "Ke toan cong trinh", email: "ktct@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Ke toan cong trinh", congTrinhPhanCong: "Cong trinh duoc giao", roles: ["KTCT"], trangThai: "Hoat dong" },
  { id: "CAT-USR-KHOCT", maNguoiDung: "KHOCT", hoTen: "Kho cong trinh", email: "khoct@sonhoabinh.vn", matKhau: "123456", sdt: "", chucDanh: "Kho cong trinh", congTrinhPhanCong: "Cong trinh duoc giao", roles: ["KHOCT"], trangThai: "Hoat dong" },
].map(normalizeUserItem);

function dedupeUsersByEmail(users) {
  const seenEmails = new Set();

  return (users || []).map(normalizeUserItem).filter((user) => {
    const emailKey = normalizeText(user.email);
    if (!emailKey) return true;
    if (seenEmails.has(emailKey)) return false;
    seenEmails.add(emailKey);
    return true;
  });
}

function ensureDefaultAdmin(users) {
  const uniqueUsers = dedupeUsersByEmail(users);
  const existingEmails = new Set(uniqueUsers.map((user) => normalizeText(user.email)).filter(Boolean));
  const missingDefaultUsers = DEFAULT_AUTH_USERS.filter((user) => !existingEmails.has(normalizeText(user.email)));

  return [...uniqueUsers, ...missingDefaultUsers];
}

function readStoredUsers(fallbackUsers = []) {
  const savedUsers = readStorageArray(USERS_STORAGE_KEY);
  const seedUsers = savedUsers.length ? savedUsers : fallbackUsers;
  const users = ensureDefaultAdmin(seedUsers.length ? seedUsers : DEFAULT_AUTH_USERS);
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users));
  return users;
}

function readCurrentUser() {
  const saved = localStorage.getItem(CURRENT_USER_STORAGE_KEY);
  if (!saved) return null;

  try {
    const parsed = JSON.parse(saved);
    return parsed?.email ? parsed : null;
  } catch {
    localStorage.removeItem(CURRENT_USER_STORAGE_KEY);
    return null;
  }
}

function getCurrentUser() {
  return readCurrentUser();
}

function normalizeMenuKey(menuKey) {
  return MENU_KEY_ALIASES[menuKey] || menuKey;
}

function canUserViewMenu(user, menuKey) {
  if (!user) return false;

  const normalizedMenuKey = normalizeMenuKey(menuKey);
  const roles = Array.isArray(user.roles) ? user.roles : user.roles ? [user.roles] : [];
  const normalizedRoles = roles.map((role) => String(role).toUpperCase());

  if (normalizedRoles.includes("ADMIN")) return true;
  return normalizedRoles.some((role) => {
    const allowed = ROLE_MENU_PERMISSIONS[role] || [];
    return allowed.includes("*") || allowed.includes(normalizedMenuKey);
  });
}

function canViewMenu(menuKey) {
  try {
    const currentUser = JSON.parse(localStorage.getItem(CURRENT_USER_STORAGE_KEY) || "null");
    return canUserViewMenu(currentUser, menuKey);
  } catch {
    return false;
  }
}

function getPreferredMenuKey(user = getCurrentUser(), preferredMenuKey = "dashboard") {
  if (canUserViewMenu(user, preferredMenuKey)) return preferredMenuKey;

  const dashboardMenu = mainMenus.find((menu) => menu.key === "dashboard" && canUserViewMenu(user, menu.key));
  const firstPermittedMenu = mainMenus.find((menu) => canUserViewMenu(user, menu.key));
  return dashboardMenu?.key || firstPermittedMenu?.key || preferredMenuKey;
}

function stripSensitiveUser(user) {
  const sessionUser = { ...user };
  delete sessionUser.matKhau;
  delete sessionUser.password;
  return sessionUser;
}

function isLockedUser(user) {
  return normalizeText(user.trangThai).includes("khoa") || normalizeText(user.trangThai).includes("khóa");
}

function formatUserRoles(roles) {
  const roleCodes = Array.isArray(roles) ? roles : roles ? [roles] : [];
  return roleCodes.length ? roleCodes.join(" / ") : "-";
}

function normalizeUserRoles(user) {
  const sourceRoles = Array.isArray(user.roles)
    ? user.roles
    : [user.role, user.vaiTro, user.maVaiTro].filter(Boolean);

  return sourceRoles
    .flatMap((role) => String(role).split(/[,+|/]/))
    .map((role) => mapLegacyRoleToCode(role))
    .filter(Boolean)
    .filter((role, index, roles) => roles.indexOf(role) === index);
}

function mapLegacyRoleToCode(role) {
  const normalizedRole = normalizeText(role);
  const exactRole = constructionRoleDefinitions.find((item) => normalizeText(item.maVaiTro) === normalizedRole);
  if (exactRole) return exactRole.maVaiTro;

  if (normalizedRole.includes("admin")) return "ADMIN";
  if (normalizedRole.includes("tgd") || normalizedRole.includes("tong giam doc")) return "TGD";
  if (normalizedRole.includes("giam doc thi cong")) return "GDTC";
  if (normalizedRole.includes("pm") || normalizedRole.includes("quan ly du an") || normalizedRole.includes("giam doc du an")) return "PM";
  if (normalizedRole.includes("chi huy") || normalizedRole.includes("giam sat")) return "CHT";
  if (normalizedRole.includes("qs")) return "QS";
  if (normalizedRole.includes("qc")) return "QC";
  if (normalizedRole.includes("thu ky")) return "TKDA";
  if (normalizedRole.includes("ke toan")) return "KTCT";
  if (normalizedRole.includes("kho")) return "KHOCT";

  return role;
}

function readSavedLogs() {
  const saved = localStorage.getItem(DAILY_LOG_STORAGE_KEY) || localStorage.getItem(STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    return [];
  }
}

function readStorageArray(key) {
  const saved = localStorage.getItem(key);
  if (!saved) return [];

  try {
    const parsed = JSON.parse(saved);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function readSavedSystemCatalogs() {
  const saved = localStorage.getItem(SYSTEM_CATALOG_STORAGE_KEY);
  if (!saved) {
    const initialCatalogs = { ...defaultSystemCatalogs, users: readStoredUsers(defaultSystemCatalogs.users) };
    publishSystemCatalogs(initialCatalogs);
    return initialCatalogs;
  }

  try {
    const parsed = JSON.parse(saved);
    const mergedCatalogs = dataCatalogTabs.reduce((catalogs, tab) => {
      catalogs[tab.key] = Array.isArray(parsed[tab.key]) ? parsed[tab.key] : defaultSystemCatalogs[tab.key];
      return catalogs;
    }, {});
    mergedCatalogs.users = readStoredUsers(mergedCatalogs.users || []);
    publishSystemCatalogs(mergedCatalogs);
    return mergedCatalogs;
  } catch {
    localStorage.removeItem(SYSTEM_CATALOG_STORAGE_KEY);
    const initialCatalogs = { ...defaultSystemCatalogs, users: readStoredUsers(defaultSystemCatalogs.users) };
    publishSystemCatalogs(initialCatalogs);
    return initialCatalogs;
  }
}

function readSavedSystemSettings() {
  const saved = localStorage.getItem(SYSTEM_SETTINGS_STORAGE_KEY);
  const savedRoles = readStorageArray("rolesCatalog");
  const savedPermissions = readStorageArray("permissionMatrix");
  const savedWorkflows = readStorageArray("approvalWorkflows");
  const savedLogs = readStorageArray("systemLogs");

  if (!saved) {
    const initialSettings = {
      roles: normalizeRolesCatalog(savedRoles),
      permissions: normalizePermissionMatrix(savedPermissions),
      approvalWorkflows: savedWorkflows.length ? savedWorkflows : defaultApprovalWorkflows,
      systemLogs: savedLogs.length ? savedLogs : defaultSystemLogs,
    };
    localStorage.setItem(SYSTEM_SETTINGS_STORAGE_KEY, JSON.stringify(initialSettings));
    return initialSettings;
  }

  try {
    const parsed = JSON.parse(saved);
    return {
      roles: normalizeRolesCatalog(Array.isArray(parsed.roles) ? parsed.roles : savedRoles),
      permissions: normalizePermissionMatrix(Array.isArray(parsed.permissions) ? parsed.permissions : savedPermissions),
      approvalWorkflows: Array.isArray(parsed.approvalWorkflows) ? parsed.approvalWorkflows : savedWorkflows.length ? savedWorkflows : defaultApprovalWorkflows,
      systemLogs: Array.isArray(parsed.systemLogs) ? parsed.systemLogs : savedLogs.length ? savedLogs : defaultSystemLogs,
    };
  } catch {
    localStorage.removeItem(SYSTEM_SETTINGS_STORAGE_KEY);
    return defaultSystemSettings;
  }
}

function publishSystemCatalogs(catalogs) {
  dataCatalogTabs.forEach((tab) => {
    localStorage.setItem(tab.storageKey, JSON.stringify(catalogs[tab.key] || []));
  });
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify((catalogs.users || []).map(normalizeUserItem)));
}

function getCatalogSeed(catalogKey) {
  const savedCatalogs = localStorage.getItem(SYSTEM_CATALOG_STORAGE_KEY);

  if (savedCatalogs) {
    try {
      const parsedCatalogs = JSON.parse(savedCatalogs);
      if (Array.isArray(parsedCatalogs[catalogKey])) return parsedCatalogs[catalogKey];
    } catch {
      return defaultSystemCatalogs[catalogKey] || [];
    }
  }

  return readStorageArray(dataCatalogTabs.find((tab) => tab.key === catalogKey)?.storageKey || "") || defaultSystemCatalogs[catalogKey] || [];
}

function readSavedContracts() {
  const saved = localStorage.getItem("contracts") || localStorage.getItem(CONTRACT_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(CONTRACT_STORAGE_KEY);
    return [];
  }
}

function readSavedBoqItems() {
  const saved = localStorage.getItem(BOQ_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(BOQ_STORAGE_KEY);
    return [];
  }
}

function readSavedWbsPlans() {
  const saved = localStorage.getItem(WBS_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(WBS_STORAGE_KEY);
    return [];
  }
}

function readSavedMaterialIns() {
  const saved = localStorage.getItem(MATERIAL_IN_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(MATERIAL_IN_STORAGE_KEY);
    return [];
  }
}

function readSavedLaborTeams() {
  const saved = localStorage.getItem(LABOR_TEAM_STORAGE_KEY);
  if (!saved) return getCatalogSeed("teams");

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(LABOR_TEAM_STORAGE_KEY);
    return [];
  }
}

function readSavedSubcontractors() {
  const saved = localStorage.getItem(SUBCONTRACTOR_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(SUBCONTRACTOR_STORAGE_KEY);
    return [];
  }
}

function readSavedSubContracts() {
  const saved = localStorage.getItem(SUBCONTRACT_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(SUBCONTRACT_STORAGE_KEY);
    return [];
  }
}

function readSavedEquipmentList() {
  const saved = localStorage.getItem(EQUIPMENT_LIST_STORAGE_KEY);
  if (!saved) return getCatalogSeed("equipment");

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(EQUIPMENT_LIST_STORAGE_KEY);
    return [];
  }
}

function readSavedEquipmentTransfers() {
  const saved = localStorage.getItem(EQUIPMENT_TRANSFER_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(EQUIPMENT_TRANSFER_STORAGE_KEY);
    return [];
  }
}

function readSavedAcceptanceBatches() {
  const saved = localStorage.getItem(ACCEPTANCE_BATCH_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(ACCEPTANCE_BATCH_STORAGE_KEY);
    return [];
  }
}

function readSavedAcceptances() {
  const saved = localStorage.getItem(ACCEPTANCE_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(ACCEPTANCE_STORAGE_KEY);
    return [];
  }
}

function readSavedInvoices() {
  const saved = localStorage.getItem(INVOICE_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(INVOICE_STORAGE_KEY);
    return [];
  }
}

function readSavedCollections() {
  const saved = localStorage.getItem(COLLECTION_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(COLLECTION_STORAGE_KEY);
    return [];
  }
}

function readSavedOtherCosts() {
  const saved = localStorage.getItem(OTHER_COST_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(OTHER_COST_STORAGE_KEY);
    return [];
  }
}

function readSavedProjectDocuments() {
  const saved = localStorage.getItem(PROJECT_DOCUMENT_STORAGE_KEY);
  if (!saved) return [];

  try {
    return JSON.parse(saved);
  } catch {
    localStorage.removeItem(PROJECT_DOCUMENT_STORAGE_KEY);
    return [];
  }
}

function readSavedDocumentTemplates() {
  const saved = localStorage.getItem(DOCUMENT_TEMPLATE_STORAGE_KEY);
  if (!saved) return [];

  try {
    const parsed = JSON.parse(saved);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    localStorage.removeItem(DOCUMENT_TEMPLATE_STORAGE_KEY);
    return [];
  }
}

function getFilteredProjectDocuments(documents, activeTab, projectFilter, search) {
  const normalizedSearch = normalizeText(search);

  return documents.filter((item) => {
    const matchesTab = item.nhomHoSo === activeTab;
    const matchesProject = projectFilter === ALL_PROJECTS || item.hopDongCongTrinh === projectFilter;
    const searchableText = normalizeText([
      item.maHoSo,
      item.tenHoSo,
      item.loaiHoSo,
      item.nguoiTai,
      item.trangThai,
      item.fileName,
      item.hopDongCongTrinh,
    ].join(" "));
    const matchesSearch = !normalizedSearch || searchableText.includes(normalizedSearch);

    return matchesTab && matchesProject && matchesSearch;
  });
}

function getProjectDocumentDashboard(documents) {
  return {
    total: documents.length,
    uploaded: documents.filter((item) => item.fileData || item.trangThai === "Đã tải lên" || item.trangThai === "Đã duyệt").length,
    missing: documents.filter((item) => item.trangThai === "Thiếu hồ sơ" || !item.fileData).length,
    storageSize: documents.reduce((sum, item) => sum + toNumber(item.fileSize), 0),
  };
}

function getDocumentTemplateDashboard(templates) {
  return {
    total: templates.length,
    active: templates.filter((item) => item.trangThai === "Đang hiệu lực").length,
    expired: templates.filter((item) => item.trangThai === "Hết hiệu lực").length,
    downloads: templates.reduce((sum, item) => sum + toNumber(item.downloadCount), 0),
  };
}

function getFilteredDocumentTemplates(templates, search, filters) {
  const normalizedSearch = normalizeText(search);

  return templates.filter((item) => {
    const matchesSearch =
      !normalizedSearch ||
      [item.maTaiLieu, item.tenTaiLieu, item.nhomTaiLieu].some((value) => normalizeText(value).includes(normalizedSearch));
    const matchesGroup = filters.nhomTaiLieu === ALL_PROJECTS || item.nhomTaiLieu === filters.nhomTaiLieu;
    const matchesType = filters.loaiFile === ALL_PROJECTS || item.loaiFile === filters.loaiFile;
    const matchesStatus = filters.trangThai === ALL_PROJECTS || item.trangThai === filters.trangThai;

    return matchesSearch && matchesGroup && matchesType && matchesStatus;
  });
}

function detectDocumentTemplateFileType(file) {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (["xls", "xlsx", "xlsm", "csv"].includes(extension)) return "Excel";
  if (["doc", "docx"].includes(extension)) return "Word";
  if (extension === "pdf") return "PDF";
  if (["jpg", "jpeg", "png", "gif", "webp", "bmp"].includes(extension)) return "Image";
  return "Khác";
}

function viewProjectDocumentFile(item) {
  if (!item.fileData) return;

  const fileWindow = window.open();
  if (fileWindow) {
    fileWindow.document.write(
      `<iframe src="${item.fileData}" title="${item.fileName || item.tenHoSo}" style="border:0;width:100%;height:100vh;"></iframe>`,
    );
  }
}

function downloadProjectDocumentFile(item) {
  if (!item.fileData) return;

  const link = document.createElement("a");
  link.href = item.fileData;
  link.download = item.fileName || `${item.maHoSo || "ho-so-du-an"}.bin`;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function getDailyLogImages(log) {
  return Array.isArray(log.hinhAnhHienTruong) ? log.hinhAnhHienTruong : [];
}

function compressImageFile(file, maxWidth = 800, quality = 0.7) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const image = new Image();
      image.onload = () => {
        const scale = Math.min(maxWidth / image.width, 1);
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(Math.round(image.width * scale), 1);
        canvas.height = Math.max(Math.round(image.height * scale), 1);

        const context = canvas.getContext("2d");
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      image.onerror = reject;
      image.src = reader.result;
    };

    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function getOtherCostWarning(item) {
  if (toNumber(item.soTien) <= 0) return "Số tiền không hợp lệ";
  if (item.trangThai === "Đã duyệt") return "Hợp lệ";
  return "Chờ duyệt";
}

function getOtherCostReportRows(otherCosts) {
  const groups = new Map();

  otherCosts.forEach((item) => {
    const key = `${item.hopDongCongTrinh || "-"}|${item.nhomChiPhi || "Khác"}`;
    const row = groups.get(key) || {
      key,
      hopDongCongTrinh: item.hopDongCongTrinh || "-",
      nhomChiPhi: item.nhomChiPhi || "Khác",
      tongSoTien: 0,
      soKhoanChi: 0,
      pendingCount: 0,
      rejectedCount: 0,
      invalidCount: 0,
    };

    row.soKhoanChi += 1;

    if (toNumber(item.soTien) <= 0) {
      row.invalidCount += 1;
    } else if (item.trangThai === "Đã duyệt") {
      row.tongSoTien += toNumber(item.soTien);
    } else if (item.trangThai === "Chờ duyệt") {
      row.pendingCount += 1;
    } else if (item.trangThai === "Từ chối") {
      row.rejectedCount += 1;
    }

    groups.set(key, row);
  });

  return Array.from(groups.values()).map((row) => ({
    ...row,
    trangThaiTongHop: getOtherCostSummaryStatus(row),
  }));
}

function getOtherCostSummaryStatus(row) {
  if (row.invalidCount > 0) return "Có khoản không hợp lệ";
  if (row.pendingCount > 0) return "Có khoản chờ duyệt";
  if (row.rejectedCount > 0) return "Có khoản từ chối";
  return "Đã duyệt";
}

function getHighestOtherCostProject(otherCosts) {
  const totals = new Map();

  otherCosts
    .filter((item) => item.trangThai === "Đã duyệt" && toNumber(item.soTien) > 0)
    .forEach((item) => {
      const project = item.hopDongCongTrinh || "-";
      totals.set(project, (totals.get(project) || 0) + toNumber(item.soTien));
    });

  if (!totals.size) return "-";

  return Array.from(totals.entries()).sort((first, second) => second[1] - first[1])[0][0];
}

function getConstructionReportsData({
  acceptanceBatches,
  acceptances,
  boqItems,
  collections,
  contracts,
  dailyLogs,
  equipmentList,
  equipmentTransfers,
  filters,
  invoices,
  laborTeams,
  materialIns,
  otherCosts,
  subContracts,
  wbsPlans,
  useProvidedData = false,
}) {
  const getSourceArray = (key, fallback) => {
    const savedItems = readStorageArray(key);
    return !useProvidedData && savedItems.length ? savedItems : fallback;
  };
  const source = {
    acceptanceBatches: getSourceArray("acceptanceBatches", acceptanceBatches),
    acceptances: getSourceArray("acceptances", acceptances),
    boqItems: getSourceArray("boqItems", boqItems),
    collections: getSourceArray("collections", collections),
    contracts: getSourceArray("contracts", contracts),
    dailyLogs: getSourceArray("dailyLogs", dailyLogs),
    equipmentList: getSourceArray("equipmentList", equipmentList),
    equipmentTransfers: getSourceArray("equipmentTransfers", equipmentTransfers),
    invoices: getSourceArray("invoices", invoices),
    laborTeams: getSourceArray("laborTeams", laborTeams),
    materialIns: getSourceArray("materialIns", materialIns),
    otherCosts: getSourceArray("otherCosts", otherCosts),
    subContracts: getSourceArray("subContracts", subContracts),
    subPayments: readStorageArray("subPayments"),
    wbsPlans: getSourceArray("wbsPlans", wbsPlans),
  };
  const filtered = filterReportSource(source, filters);
  const projectOptions = getReportProjectOptions(source);
  const allProjects = getReportProjects(filtered);
  const materialOuts = getMaterialOutsFromDailyLogs(filtered.dailyLogs);
  const materialRows = getReportMaterialRows(filtered.materialIns, materialOuts, filtered.boqItems);
  const laborRows = getLaborReportRows(getLaborUsagesFromDailyLogs(filtered.dailyLogs), filtered.laborTeams);
  const equipmentRows = getEquipmentReportRows(
    getEquipmentUsagesFromDailyLogs(filtered.dailyLogs),
    filtered.equipmentTransfers,
    filtered.equipmentList,
  );
  const quantityRows = getReportQuantityRows(filtered.boqItems, filtered.wbsPlans, filtered.dailyLogs, filtered.acceptances);
  const progressRows = getReportProgressRows(filtered.wbsPlans, filtered.boqItems, filtered.dailyLogs);
  const acceptanceRows = getReportAcceptanceRows(allProjects, filtered.dailyLogs, filtered.acceptances);
  const revenueRows = getDebtRows(filtered.contracts, filtered.acceptances, filtered.invoices, filtered.collections);
  const costRows = getReportCostRows(allProjects, filtered.materialIns, laborRows, equipmentRows, filtered.otherCosts);
  const profitRows = getReportProfitRows(allProjects, filtered.acceptances, filtered.invoices, costRows);
  const overviewRows = getReportOverviewRows({
    projects: allProjects,
    contracts: filtered.contracts,
    boqItems: filtered.boqItems,
    wbsPlans: filtered.wbsPlans,
    dailyLogs: filtered.dailyLogs,
    acceptances: filtered.acceptances,
    invoices: filtered.invoices,
    collections: filtered.collections,
    costRows,
    quantityRows,
  });
  const dashboard = getReportDashboard(overviewRows, filtered.invoices);
  const alerts = getReportAlerts(progressRows, materialRows, revenueRows, profitRows, dashboard);

  void acceptanceBatches;
  void source.acceptanceBatches;
  void source.subContracts;
  void source.subPayments;

  return {
    acceptanceRows,
    alerts,
    costRows,
    dashboard,
    equipmentRows,
    laborRows,
    materialRows,
    overviewRows,
    profitRows,
    progressRows,
    projectOptions,
    quantityRows,
    revenueRows,
  };
}

function filterReportSource(source, filters) {
  return {
    ...source,
    acceptances: filterReportItems(source.acceptances, filters, ["ngayNghiemThu", "createdAt"]),
    boqItems: filterReportItems(source.boqItems, filters),
    collections: filterReportItems(source.collections, filters, ["ngayThuTien"]),
    contracts: filterReportItems(source.contracts, filters),
    dailyLogs: filterReportItems(source.dailyLogs, filters, ["ngay"]),
    equipmentTransfers: filterReportItems(source.equipmentTransfers, filters, ["ngayNhan"]),
    invoices: filterReportItems(source.invoices, filters, ["ngayXuatHoaDon"]),
    materialIns: filterReportItems(source.materialIns, filters, ["ngayNhap"]),
    otherCosts: filterReportItems(source.otherCosts, filters, ["ngayPhatSinh"]),
    subContracts: filterReportItems(source.subContracts, filters, ["ngayBatDau", "ngayKetThuc"]),
    wbsPlans: filterReportItems(source.wbsPlans, filters, ["ngayBatDauKeHoach", "ngayKetThucKeHoach"]),
  };
}

function filterReportItems(items, filters, dateFields = []) {
  return items.filter((item) => isReportProjectMatched(item, filters.project) && isReportDateMatched(item, filters, dateFields));
}

function isReportProjectMatched(item, projectFilter) {
  if (!projectFilter || projectFilter === ALL_PROJECTS) return true;
  return getReportProjectName(item) === projectFilter;
}

function isReportDateMatched(item, filters, dateFields) {
  if (!filters.fromDate && !filters.toDate) return true;
  if (!dateFields.length) return true;

  const dateValue = dateFields.map((field) => item[field]).find(Boolean);
  if (!dateValue) return true;

  const itemDate = new Date(dateValue);
  if (Number.isNaN(itemDate.getTime())) return true;

  if (filters.fromDate && itemDate < new Date(filters.fromDate)) return false;
  if (filters.toDate && itemDate > new Date(filters.toDate)) return false;
  return true;
}

function getReportProjectOptions(source) {
  return getReportProjects(source).filter((project) => project !== "-");
}

function getReportProjects(source) {
  const projectsSet = new Set();

  source.contracts.forEach((item) => projectsSet.add(getContractProjectName(item)));
  [
    source.boqItems,
    source.collections,
    source.dailyLogs,
    source.equipmentTransfers,
    source.invoices,
    source.materialIns,
    source.otherCosts,
    source.subContracts,
    source.wbsPlans,
    source.acceptances,
  ].forEach((items) => {
    items.forEach((item) => projectsSet.add(getReportProjectName(item)));
  });

  return Array.from(projectsSet).filter(Boolean);
}

function getContractProjectName(contract) {
  const code = contract.maHopDong || contract.code || contract.contractCode || "";
  const name = contract.tenCongTrinh || contract.projectName || contract.name || contract.ten || contract.hopDongCongTrinh || "-";
  return `${code}${code ? " - " : ""}${name}`;
}

function getReportProjectName(item) {
  if (!item) return "-";
  if (item.tenCongTrinh || item.maHopDong) return getContractProjectName(item);
  return item.hopDongCongTrinh || item.congTrinh || item.projectName || item.project || "-";
}

function getReportMaterialRows(materialIns, materialOuts, boqItems) {
  const rows = getMaterialInventoryRows(materialIns, materialOuts);

  return rows.map((row) => {
    const dinhMucBoq = getBoqMaterialNorm(boqItems, row);
    const chenhLechDinhMuc = dinhMucBoq > 0 ? row.tongSuDung - dinhMucBoq : 0;
    const trangThai = dinhMucBoq > 0 && row.tongSuDung > dinhMucBoq ? "Vượt định mức" : row.trangThai;

    return {
      ...row,
      dinhMucBoq,
      chenhLechDinhMuc,
      trangThai,
    };
  });
}

function getBoqMaterialNorm(boqItems, materialRow) {
  return boqItems
    .filter((item) => item.hopDongCongTrinh === materialRow.hopDongCongTrinh)
    .filter((item) => {
      if (!item.maVatTu && !item.tenVatTu) return false;
      return (
        normalizeText(item.maVatTu) === normalizeText(materialRow.maVatTu) ||
        normalizeText(item.tenVatTu) === normalizeText(materialRow.tenVatTu)
      );
    })
    .reduce((sum, item) => sum + toNumber(item.dinhMucVatTu), 0);
}

function getReportQuantityRows(boqItems, wbsPlans, dailyLogs, acceptances) {
  const itemMap = new Map();

  boqItems.forEach((item) => {
    const key = `${item.hopDongCongTrinh}|${item.maWbs || item.hangMucCongViec}`;
    itemMap.set(key, {
      key,
      hopDongCongTrinh: item.hopDongCongTrinh || "-",
      maWbs: item.maWbs || "-",
      hangMucCongViec: item.hangMucCongViec || "-",
      dvt: item.dvt || "",
      khoiLuongHopDong: toNumber(item.khoiLuongHopDong),
      donGiaHopDong: toNumber(item.donGiaHopDong),
    });
  });

  wbsPlans.forEach((plan) => {
    const key = `${plan.hopDongCongTrinh}|${plan.maWbs || plan.hangMucCongViec}`;
    if (!itemMap.has(key)) {
      itemMap.set(key, {
        key,
        hopDongCongTrinh: plan.hopDongCongTrinh || "-",
        maWbs: plan.maWbs || "-",
        hangMucCongViec: plan.hangMucCongViec || "-",
        dvt: plan.dvt || "",
        khoiLuongHopDong: toNumber(plan.khoiLuongKeHoach),
        donGiaHopDong: 0,
      });
    }
  });

  return Array.from(itemMap.values()).map((row) => {
    const khoiLuongThucHien = getExecutedQuantity(dailyLogs, row.hopDongCongTrinh, row.maWbs);
    const khoiLuongNghiemThu = getAcceptedQuantity(acceptances, row.hopDongCongTrinh, row.maWbs);

    return {
      ...row,
      khoiLuongThucHien,
      khoiLuongNghiemThu,
      giaTriThucHienTamTinh: khoiLuongThucHien * row.donGiaHopDong,
      trangThai: getReportExecutedQuantityStatus(khoiLuongThucHien, khoiLuongNghiemThu),
    };
  });
}

function getReportExecutedQuantityStatus(executedQuantity, acceptedQuantity) {
  if (executedQuantity > acceptedQuantity) return "Có khối lượng chờ nghiệm thu";
  if (executedQuantity === acceptedQuantity) return "Đã nghiệm thu đủ";
  return "Nghiệm thu vượt thực hiện";
}

function getReportProgressRows(wbsPlans, boqItems, dailyLogs) {
  const plans = wbsPlans.length
    ? wbsPlans
    : boqItems.map((item) => ({
      hopDongCongTrinh: item.hopDongCongTrinh,
      maWbs: item.maWbs,
      hangMucCongViec: item.hangMucCongViec,
      khoiLuongKeHoach: item.khoiLuongHopDong,
    }));

  return plans.map((plan) => {
    const khoiLuongKeHoach = toNumber(plan.khoiLuongKeHoach);
    const khoiLuongThucHien = getExecutedQuantity(dailyLogs, plan.hopDongCongTrinh, plan.maWbs);
    const tyLeHoanThanh = khoiLuongKeHoach ? (khoiLuongThucHien / khoiLuongKeHoach) * 100 : 0;

    return {
      key: plan.id || `${plan.hopDongCongTrinh}|${plan.maWbs}|${plan.hangMucCongViec}`,
      hopDongCongTrinh: plan.hopDongCongTrinh || "-",
      maWbs: plan.maWbs || "-",
      hangMucCongViec: plan.hangMucCongViec || "-",
      ngayBatDauKeHoach: plan.ngayBatDauKeHoach || "",
      ngayKetThucKeHoach: plan.ngayKetThucKeHoach || "",
      khoiLuongKeHoach,
      khoiLuongThucHien,
      tyLeHoanThanh,
      trangThai: getReportProgressStatus(plan, tyLeHoanThanh, khoiLuongThucHien),
    };
  });
}

function getReportProgressStatus(plan, completionRate, executedQuantity) {
  if (completionRate >= 100) return "Hoàn thành";
  if (plan.ngayKetThucKeHoach && new Date(plan.ngayKetThucKeHoach) < new Date()) return "Chậm tiến độ";
  if (executedQuantity > 0) return "Đang thi công";
  return "Chưa bắt đầu";
}

function getReportAcceptanceRows(projectsList, dailyLogs, acceptances) {
  return projectsList.map((project) => {
    const tongKhoiLuongThucHien = dailyLogs
      .filter((log) => getReportProjectName(log) === project)
      .reduce((sum, log) => sum + toNumber(log.khoiLuongThucHienNgay || log.khoiLuong), 0);
    const projectAcceptances = acceptances.filter((item) => getReportProjectName(item) === project);
    const tongKhoiLuongNghiemThu = projectAcceptances.reduce(
      (sum, item) => sum + toNumber(item.khoiLuongNghiemThuKyNay || item.khoiLuongNghiemThu || item.khoiLuong),
      0,
    );

    return {
      hopDongCongTrinh: project,
      tongKhoiLuongThucHien,
      tongKhoiLuongNghiemThu,
      giaTriNghiemThu: projectAcceptances.reduce((sum, item) => sum + toNumber(item.thanhTienNghiemThuKyNay), 0),
      khoiLuongChuaNghiemThu: tongKhoiLuongThucHien - tongKhoiLuongNghiemThu,
    };
  });
}

function getReportCostRows(projectsList, materialIns, laborRows, equipmentRows, otherCosts) {
  return projectsList.map((project) => {
    const chiPhiVatTu = materialIns
      .filter((item) => getReportProjectName(item) === project)
      .reduce((sum, item) => sum + getMaterialInputCost(item), 0);
    const chiPhiNhanCong = laborRows
      .filter((item) => item.hopDongCongTrinh === project)
      .reduce((sum, item) => sum + toNumber(item.chiPhiNhanCong), 0);
    const chiPhiThietBi = equipmentRows
      .filter((item) => item.hopDongCongTrinh === project)
      .reduce((sum, item) => sum + toNumber(item.chiPhiThietBi), 0);
    const chiPhiKhac = otherCosts
      .filter((item) => getReportProjectName(item) === project && item.trangThai === "Đã duyệt" && toNumber(item.soTien) > 0)
      .reduce((sum, item) => sum + toNumber(item.soTien), 0);

    return {
      hopDongCongTrinh: project,
      chiPhiVatTu,
      chiPhiNhanCong,
      chiPhiThietBi,
      chiPhiKhac,
      tongChiPhiThucTe: chiPhiVatTu + chiPhiNhanCong + chiPhiThietBi + chiPhiKhac,
    };
  });
}

function getMaterialInputCost(item) {
  if (toNumber(item.thanhTien) > 0) return toNumber(item.thanhTien);
  if (toNumber(item.giaTri) > 0) return toNumber(item.giaTri);
  return toNumber(item.soLuongNhap) * toNumber(item.donGiaNhap || item.donGia);
}

function getReportProfitRows(projectsList, acceptances, invoices, costRows) {
  return projectsList.map((project) => {
    const acceptanceValue = acceptances
      .filter((item) => getReportProjectName(item) === project)
      .reduce((sum, item) => sum + toNumber(item.thanhTienNghiemThuKyNay), 0);
    const invoiceBeforeVat = invoices
      .filter((item) => getReportProjectName(item) === project)
      .reduce((sum, item) => sum + getInvoiceBeforeVatValue(item), 0);
    const doanhThuGhiNhan = acceptanceValue > 0 ? acceptanceValue : invoiceBeforeVat;
    const costRow = costRows.find((item) => item.hopDongCongTrinh === project);
    const tongChiPhiThucTe = toNumber(costRow?.tongChiPhiThucTe);
    const laiLoTamTinh = doanhThuGhiNhan - tongChiPhiThucTe;

    return {
      hopDongCongTrinh: project,
      doanhThuGhiNhan,
      tongChiPhiThucTe,
      laiLoTamTinh,
      bienLoiNhuan: doanhThuGhiNhan ? (laiLoTamTinh / doanhThuGhiNhan) * 100 : 0,
    };
  });
}

function getReportOverviewRows({
  projects,
  contracts,
  boqItems,
  wbsPlans,
  dailyLogs,
  acceptances,
  invoices,
  collections,
  costRows,
  quantityRows,
}) {
  return projects.map((project) => {
    const contract = contracts.find((item) => getContractProjectName(item) === project);
    const projectAcceptances = acceptances.filter((item) => getReportProjectName(item) === project);
    const tongGiaTriNghiemThu = projectAcceptances.reduce((sum, item) => sum + toNumber(item.thanhTienNghiemThuKyNay), 0);
    const tongDaXuatHoaDon = invoices
      .filter((item) => getReportProjectName(item) === project)
      .reduce((sum, item) => sum + toNumber(item.tongGiaTriHoaDon), 0);
    const tongDaThuTien = collections
      .filter((item) => getReportProjectName(item) === project)
      .reduce((sum, item) => sum + toNumber(item.soTienThu), 0);
    const costRow = costRows.find((item) => item.hopDongCongTrinh === project);
    const tongChiPhiThucTe = toNumber(costRow?.tongChiPhiThucTe);
    const doanhThuGhiNhan = tongGiaTriNghiemThu || invoices
      .filter((item) => getReportProjectName(item) === project)
      .reduce((sum, item) => sum + getInvoiceBeforeVatValue(item), 0);
    const laiLoTamTinh = doanhThuGhiNhan - tongChiPhiThucTe;
    const tongGiaTriThucHien = quantityRows
      .filter((item) => item.hopDongCongTrinh === project)
      .reduce((sum, item) => sum + toNumber(item.giaTriThucHienTamTinh), 0);

    return {
      hopDongCongTrinh: project,
      giaTriHopDong: toNumber(contract?.giaTriHopDong || contract?.contractValue),
      tongGiaTriThucHien,
      tongGiaTriNghiemThu,
      tongGiaTriBoq: boqItems
        .filter((item) => getReportProjectName(item) === project)
        .reduce((sum, item) => sum + getBoqContractValue(item), 0),
      tongKhoiLuongKeHoach: wbsPlans
        .filter((item) => getReportProjectName(item) === project)
        .reduce((sum, item) => sum + toNumber(item.khoiLuongKeHoach), 0),
      tongKhoiLuongThucHien: dailyLogs
        .filter((item) => getReportProjectName(item) === project)
        .reduce((sum, item) => sum + toNumber(item.khoiLuongThucHienNgay || item.khoiLuong), 0),
      tongKhoiLuongNghiemThu: projectAcceptances.reduce(
        (sum, item) => sum + toNumber(item.khoiLuongNghiemThuKyNay || item.khoiLuongNghiemThu || item.khoiLuong),
        0,
      ),
      tongDaXuatHoaDon,
      tongDaThuTien,
      tongCongNo: tongDaXuatHoaDon - tongDaThuTien,
      tongChiPhiThucTe,
      laiLoTamTinh,
      bienLoiNhuan: doanhThuGhiNhan ? (laiLoTamTinh / doanhThuGhiNhan) * 100 : 0,
    };
  });
}

function getReportDashboard(overviewRows, invoices) {
  const totalExecutedValue = overviewRows.reduce((sum, item) => sum + item.tongGiaTriThucHien, 0);
  const totalAcceptanceValue = overviewRows.reduce((sum, item) => sum + item.tongGiaTriNghiemThu, 0);
  const totalInvoiceBeforeVat = invoices.reduce((sum, item) => sum + getInvoiceBeforeVatValue(item), 0);
  const totalCollections = overviewRows.reduce((sum, item) => sum + item.tongDaThuTien, 0);

  return {
    totalContractValue: overviewRows.reduce((sum, item) => sum + item.giaTriHopDong, 0),
    totalExecutedValue,
    totalAcceptanceValue,
    totalInvoiceBeforeVat,
    totalCollections,
    totalDebt: overviewRows.reduce((sum, item) => sum + item.tongCongNo, 0),
    totalActualCost: overviewRows.reduce((sum, item) => sum + item.tongChiPhiThucTe, 0),
    totalProfit: overviewRows.reduce((sum, item) => sum + item.laiLoTamTinh, 0),
    pendingAcceptanceValue: totalExecutedValue - totalAcceptanceValue,
    acceptanceNotInvoicedValue: totalAcceptanceValue - totalInvoiceBeforeVat,
    invoiceNotCollectedValue: totalInvoiceBeforeVat - totalCollections,
  };
}

function getDashboardSummaryData(reportData, source) {
  const totalPlanValue = source.boqItems.reduce((sum, item) => sum + toNumber(item.chiPhiKeHoach), 0);
  const totalPlannedQuantity = reportData.overviewRows.reduce((sum, item) => sum + item.tongKhoiLuongKeHoach, 0) ||
    source.wbsPlans.reduce((sum, item) => sum + toNumber(item.khoiLuongKeHoach), 0);
  const totalExecutedQuantity = reportData.overviewRows.reduce((sum, item) => sum + item.tongKhoiLuongThucHien, 0);
  const totalAcceptedQuantity = reportData.overviewRows.reduce((sum, item) => sum + item.tongKhoiLuongNghiemThu, 0);
  const totalWorkers = source.dailyLogs.reduce((sum, item) => sum + toNumber(item.soNguoi), 0);
  const totalLaborHours = source.dailyLogs.reduce((sum, item) => sum + toNumber(item.soGio), 0);
  const totalEquipmentHours = source.dailyLogs.reduce((sum, item) => sum + toNumber(item.soGioHoatDong), 0);
  const totalMaterialUsage = source.dailyLogs.reduce((sum, item) => sum + toNumber(item.soLuongVatTu), 0);
  const recognizedRevenue = reportData.dashboard.totalAcceptanceValue || reportData.dashboard.totalInvoiceBeforeVat;
  const projectProgressRows = reportData.overviewRows.map((item) => ({
    hopDongCongTrinh: item.hopDongCongTrinh,
    khoiLuongKeHoach: item.tongKhoiLuongKeHoach,
    khoiLuongThucHien: item.tongKhoiLuongThucHien,
    tyLeHoanThanh: item.tongKhoiLuongKeHoach ? (item.tongKhoiLuongThucHien / item.tongKhoiLuongKeHoach) * 100 : 0,
  }));
  const topRevenueProjects = reportData.overviewRows
    .map((item) => ({
      hopDongCongTrinh: item.hopDongCongTrinh,
      doanhThu: item.tongGiaTriNghiemThu || item.tongDaXuatHoaDon,
      daThuTien: item.tongDaThuTien,
    }))
    .sort((first, second) => second.doanhThu - first.doanhThu)
    .slice(0, 5);
  const topProfitProjects = reportData.overviewRows
    .map((item) => ({
      hopDongCongTrinh: item.hopDongCongTrinh,
      laiLoTamTinh: item.laiLoTamTinh,
      bienLoiNhuan: item.bienLoiNhuan,
    }))
    .sort((first, second) => second.laiLoTamTinh - first.laiLoTamTinh)
    .slice(0, 5);

  return {
    ...reportData.dashboard,
    completionRate: totalPlannedQuantity ? (totalExecutedQuantity / totalPlannedQuantity) * 100 : 0,
    profitMargin: recognizedRevenue ? (reportData.dashboard.totalProfit / recognizedRevenue) * 100 : 0,
    projectCount: source.contracts.length,
    projectProgressRows,
    topProfitProjects,
    topRevenueProjects,
    totalAcceptedQuantity,
    totalBoqValue: reportData.overviewRows.reduce((sum, item) => sum + item.tongGiaTriBoq, 0),
    totalEquipmentHours,
    totalExecutedQuantity,
    totalLaborHours,
    totalMaterialUsage,
    totalPlanValue,
    totalPlannedQuantity,
    totalWorkers,
  };
}

function getReportAlerts(progressRows, materialRows, revenueRows, profitRows, dashboard) {
  const alerts = [];

  if (progressRows.some((item) => item.trangThai === "Chậm tiến độ")) alerts.push("Có công trình chậm tiến độ");
  if (materialRows.some((item) => item.trangThai === "Âm kho")) alerts.push("Có vật tư âm kho");
  if (materialRows.some((item) => item.trangThai === "Vượt định mức")) alerts.push("Có vật tư vượt định mức");
  if (revenueRows.some((item) => item.congNoConPhaiThu > 0)) alerts.push("Có công nợ còn phải thu");
  if (profitRows.some((item) => item.laiLoTamTinh < 0)) alerts.push("Có công trình lỗ");
  if (dashboard.pendingAcceptanceValue > 0) alerts.push("Có khối lượng chờ nghiệm thu");
  if (dashboard.acceptanceNotInvoicedValue > 0) alerts.push("Có nghiệm thu chưa xuất hóa đơn");
  if (dashboard.invoiceNotCollectedValue > 0) alerts.push("Có hóa đơn chưa thu tiền");
  if (hasAbnormalReportFlow(dashboard)) alerts.push("Dữ liệu bất thường - kiểm tra nguồn dữ liệu");

  return alerts;
}

function hasAbnormalReportFlow(dashboard) {
  return (
    dashboard.totalExecutedValue < dashboard.totalAcceptanceValue ||
    dashboard.totalAcceptanceValue < dashboard.totalInvoiceBeforeVat ||
    dashboard.totalInvoiceBeforeVat < dashboard.totalCollections
  );
}

function getBoqContractValue(item) {
  if (toNumber(item.thanhTienHopDong) > 0) return toNumber(item.thanhTienHopDong);
  return toNumber(item.khoiLuongHopDong) * toNumber(item.donGiaHopDong);
}

function getInvoiceBeforeVatValue(item) {
  if (toNumber(item.giaTriTruocVat) > 0) return toNumber(item.giaTriTruocVat);
  if (toNumber(item.valueBeforeVat) > 0) return toNumber(item.valueBeforeVat);
  if (toNumber(item.beforeVat) > 0) return toNumber(item.beforeVat);
  return toNumber(item.giaTriNghiemThu) * (toNumber(item.tyLeXuatHoaDon || 100) / 100);
}

function getMaterialOutsFromDailyLogs(logs) {
  return logs
    .filter((log) => log.tenVatTu || log.maVatTu || toNumber(log.soLuongVatTu) > 0)
    .map((log) => ({
      id: log.id,
      ngay: log.ngay,
      hopDongCongTrinh: log.hopDongCongTrinh || log.congTrinh || "-",
      maWbs: log.maWbs || "",
      maVatTu: log.maVatTu || "",
      tenVatTu: log.tenVatTu || log.vatTu || "Vật tư chưa đặt tên",
      dvt: log.dvtVatTu || "",
      soLuongSuDung: toNumber(log.soLuongVatTu),
      giamSat: log.giamSat || "",
    }));
}

function getMaterialInventoryRows(ins, outs) {
  const groups = new Map();

  ins.forEach((item) => {
    const key = getMaterialGroupKey(item.hopDongCongTrinh, item.maVatTu, item.tenVatTu, item.dvt);
    const row = groups.get(key) || createMaterialInventoryRow(key, item);
    row.tongNhap += toNumber(item.soLuongNhap);
    groups.set(key, row);
  });

  outs.forEach((item) => {
    const key = getMaterialGroupKey(item.hopDongCongTrinh, item.maVatTu, item.tenVatTu, item.dvt);
    const row = groups.get(key) || createMaterialInventoryRow(key, item);
    row.tongSuDung += toNumber(item.soLuongSuDung);
    groups.set(key, row);
  });

  return Array.from(groups.values()).map((row) => {
    const tonKho = row.tongNhap - row.tongSuDung;
    return {
      ...row,
      tonKho,
      trangThai: tonKho > 0 ? "Đủ" : tonKho === 0 ? "Hết" : "Âm kho",
    };
  });
}

function createMaterialInventoryRow(key, item) {
  return {
    key,
    hopDongCongTrinh: item.hopDongCongTrinh || "-",
    maVatTu: item.maVatTu || "",
    tenVatTu: item.tenVatTu || "Vật tư chưa đặt tên",
    dvt: item.dvt || "",
    tongNhap: 0,
    tongSuDung: 0,
    tonKho: 0,
    trangThai: "Hết",
  };
}

function getMaterialGroupKey(project, code, name, unit) {
  return [project || "-", code || name || "unknown", unit || ""].join("|");
}

function getLaborUsagesFromDailyLogs(logs) {
  return logs
    .filter((log) => log.toDoi || toNumber(log.soNguoi) > 0 || toNumber(log.soGio) > 0)
    .map((log) => ({
      id: log.id,
      ngay: log.ngay,
      hopDongCongTrinh: log.hopDongCongTrinh || log.congTrinh || "-",
      maWbs: log.maWbs || "",
      hangMuc: log.hangMucCongViec || log.hangMuc || "",
      maToDoi: log.maToDoi || "",
      toDoi: log.toDoi || "Chưa chọn tổ đội",
      soNguoi: toNumber(log.soNguoi),
      soGio: toNumber(log.soGio),
      giamSat: log.giamSat || "",
    }));
}

function getExecutedQuantity(logs, project, wbsCode) {
  return logs
    .filter((log) => (log.hopDongCongTrinh || log.congTrinh) === project && log.maWbs === wbsCode)
    .reduce((sum, log) => sum + toNumber(log.khoiLuongThucHienNgay || log.khoiLuong), 0);
}

function getAcceptedQuantity(acceptances, project, wbsCode) {
  return acceptances
    .filter((item) => item.hopDongCongTrinh === project && item.maWbs === wbsCode)
    .reduce((sum, item) => sum + toNumber(item.khoiLuongNghiemThuKyNay || item.khoiLuongNghiemThu || item.khoiLuong), 0);
}

function getAcceptanceReconciliationRows(boqItems, dailyLogs, acceptances) {
  return boqItems.map((item) => {
    const khoiLuongHopDong = toNumber(item.khoiLuongHopDong);
    const khoiLuongThucHien = getExecutedQuantity(dailyLogs, item.hopDongCongTrinh, item.maWbs);
    const khoiLuongNghiemThu = getAcceptedQuantity(acceptances, item.hopDongCongTrinh, item.maWbs);
    const khoiLuongChuaNghiemThu = khoiLuongThucHien - khoiLuongNghiemThu;
    const khoiLuongConLaiHopDong = khoiLuongHopDong - khoiLuongNghiemThu;

    return {
      key: item.id || `${item.hopDongCongTrinh}|${item.maWbs}`,
      hopDongCongTrinh: item.hopDongCongTrinh,
      maWbs: item.maWbs,
      hangMucCongViec: item.hangMucCongViec,
      khoiLuongHopDong,
      khoiLuongThucHien,
      khoiLuongNghiemThu,
      khoiLuongChuaNghiemThu,
      khoiLuongConLaiHopDong,
      trangThai: khoiLuongChuaNghiemThu > 0 ? "Còn nghiệm thu" : "Đã nghiệm thu hết",
    };
  });
}

function getAcceptanceValueByBatch(acceptances, batchNo) {
  const batchAcceptances = acceptances.filter((item) => item.soBienBan === batchNo);
  if (!batchAcceptances.length) return null;

  return {
    hopDongCongTrinh: batchAcceptances[0].hopDongCongTrinh,
    giaTriNghiemThu: batchAcceptances.reduce((sum, item) => sum + toNumber(item.thanhTienNghiemThuKyNay), 0),
  };
}

function getInvoiceWarning(invoice) {
  const acceptanceValue = toNumber(invoice.acceptanceValue || invoice.giaTriNghiemThu);
  const beforeVat = toNumber(invoice.beforeVat || invoice.valueBeforeVat || invoice.giaTriTruocVat);

  if (beforeVat > acceptanceValue) {
    return "Xuất hóa đơn vượt nghiệm thu";
  }

  return "Hợp lệ";
}

function getDebtRows(contracts, acceptances, invoices, collections) {
  const projectMap = new Map();

  contracts.forEach((contract) => {
    const projectName = `${contract.maHopDong || ""}${contract.maHopDong ? " - " : ""}${contract.tenCongTrinh}`;
    projectMap.set(projectName, {
      hopDongCongTrinh: projectName,
      giaTriHopDong: toNumber(contract.giaTriHopDong),
      giaTriNghiemThu: 0,
      giaTriDaXuatHoaDon: 0,
      giaTriDaThuTien: 0,
    });
  });

  acceptances.forEach((item) => {
    const row = getOrCreateDebtRow(projectMap, item.hopDongCongTrinh);
    row.giaTriNghiemThu += toNumber(item.thanhTienNghiemThuKyNay);
  });

  invoices.forEach((item) => {
    const row = getOrCreateDebtRow(projectMap, item.hopDongCongTrinh);
    row.giaTriDaXuatHoaDon += toNumber(item.tongGiaTriHoaDon);
  });

  collections.forEach((item) => {
    const row = getOrCreateDebtRow(projectMap, item.hopDongCongTrinh);
    row.giaTriDaThuTien += toNumber(item.soTienThu);
  });

  return Array.from(projectMap.values()).map((row) => {
    const congNoConPhaiThu = row.giaTriDaXuatHoaDon - row.giaTriDaThuTien;
    const tyLeThuTienHoaDon = row.giaTriDaXuatHoaDon ? (row.giaTriDaThuTien / row.giaTriDaXuatHoaDon) * 100 : 0;
    const tyLeNghiemThuHopDong = row.giaTriHopDong ? (row.giaTriNghiemThu / row.giaTriHopDong) * 100 : 0;

    return {
      ...row,
      congNoConPhaiThu,
      tyLeThuTienHoaDon,
      tyLeNghiemThuHopDong,
      trangThai: getDebtStatus(row.giaTriDaXuatHoaDon, row.giaTriDaThuTien),
    };
  });
}

function getOrCreateDebtRow(projectMap, projectName) {
  const normalizedProject = projectName || "-";

  if (!projectMap.has(normalizedProject)) {
    projectMap.set(normalizedProject, {
      hopDongCongTrinh: normalizedProject,
      giaTriHopDong: 0,
      giaTriNghiemThu: 0,
      giaTriDaXuatHoaDon: 0,
      giaTriDaThuTien: 0,
    });
  }

  return projectMap.get(normalizedProject);
}

function getDebtStatus(invoiceValue, collectedValue) {
  if (invoiceValue <= 0) return "Chưa xuất hóa đơn";
  if (collectedValue <= 0) return "Chưa thu tiền";
  if (collectedValue < invoiceValue) return "Còn công nợ";
  if (collectedValue === invoiceValue) return "Đã thu đủ";
  return "Thu vượt";
}

function getSubcontractorReportRows(subContracts, dailyLogs) {
  const acceptances = readStorageArray("acceptances");
  const payments = readStorageArray("subPayments");

  return subContracts.map((item) => {
    const khoiLuongThucHien = dailyLogs
      .filter((log) => isSameProjectWbsWork(log, item))
      .reduce((sum, log) => sum + toNumber(log.khoiLuongThucHienNgay || log.khoiLuong), 0);
    const khoiLuongNghiemThu = acceptances
      .filter((acceptance) => isSameProjectWbsWork(acceptance, item))
      .reduce((sum, acceptance) => sum + toNumber(acceptance.khoiLuongNghiemThu || acceptance.khoiLuong || acceptance.quantity), 0);
    const daThanhToan = payments
      .filter((payment) => isSamePartnerPayment(payment, item))
      .reduce((sum, payment) => sum + toNumber(payment.soTien || payment.amount || payment.giaTriThanhToan), 0);
    const giaTriNghiemThu = khoiLuongNghiemThu * toNumber(item.donGiaKhoan);
    const conPhaiTra = giaTriNghiemThu - daThanhToan;

    return {
      key: item.id,
      hopDongCongTrinh: item.hopDongCongTrinh,
      doiTac: item.doiTac,
      hangMucCongViec: item.hangMucCongViec,
      khoiLuongGiaoKhoan: toNumber(item.khoiLuongGiaoKhoan),
      khoiLuongThucHien,
      khoiLuongNghiemThu,
      giaTriNghiemThu,
      daThanhToan,
      conPhaiTra,
      trangThai: conPhaiTra > 0 ? "Còn phải trả" : giaTriNghiemThu > 0 ? "Đã thanh toán đủ" : "Chưa nghiệm thu",
    };
  });
}

function isSameProjectWbsWork(source, target) {
  const sourceProject = source.hopDongCongTrinh || source.congTrinh || source.project || "";
  const sourceWbs = source.maWbs || source.wbsCode || "";
  const sourceWork = source.hangMucCongViec || source.hangMuc || source.workItem || "";

  return (
    sourceProject === target.hopDongCongTrinh &&
    sourceWbs === target.maWbs &&
    (!sourceWork || !target.hangMucCongViec || normalizeText(sourceWork) === normalizeText(target.hangMucCongViec))
  );
}

function isSamePartnerPayment(payment, target) {
  const paymentProject = payment.hopDongCongTrinh || payment.congTrinh || payment.project || "";
  const paymentPartner = payment.doiTac || payment.nhaThauPhu || payment.toDoi || payment.partner || "";
  const paymentWbs = payment.maWbs || payment.wbsCode || "";

  return (
    (!paymentProject || paymentProject === target.hopDongCongTrinh) &&
    (!paymentPartner || normalizeText(paymentPartner) === normalizeText(target.doiTac)) &&
    (!paymentWbs || paymentWbs === target.maWbs)
  );
}

function getEquipmentUsagesFromDailyLogs(logs) {
  return logs
    .filter((log) => log.tenThietBi || log.maThietBi || toNumber(log.soGioHoatDong) > 0)
    .map((log) => ({
      id: log.id,
      ngay: log.ngay,
      hopDongCongTrinh: log.hopDongCongTrinh || log.congTrinh || "-",
      maWbs: log.maWbs || "",
      hangMuc: log.hangMucCongViec || log.hangMuc || "",
      maThietBi: log.maThietBi || "",
      tenThietBi: log.tenThietBi || "Thiết bị chưa đặt tên",
      soGioHoatDong: toNumber(log.soGioHoatDong),
      nguon: log.nguonThietBi || "",
      giamSat: log.giamSat || "",
    }));
}

function getEquipmentReportRows(usages, transfers, equipmentList) {
  const groups = new Map();

  usages.forEach((usage) => {
    const equipment = findEquipment(usage, equipmentList);
    const normalized = normalizeEquipmentIdentity(usage, equipment);
    const key = getEquipmentGroupKey(usage.hopDongCongTrinh, normalized.maThietBi, normalized.tenThietBi);
    const row = groups.get(key) || {
      key,
      hopDongCongTrinh: usage.hopDongCongTrinh,
      maThietBi: normalized.maThietBi,
      tenThietBi: normalized.tenThietBi,
      nguonSoHuu: normalized.nguonSoHuu,
      tongGioHoatDong: 0,
      transferDays: 0,
      donGiaThueGio: normalized.donGiaThueGio,
      donGiaThueNgay: normalized.donGiaThueNgay,
      chiPhiThietBi: 0,
      trangThai: "",
    };

    row.tongGioHoatDong += usage.soGioHoatDong;
    groups.set(key, row);
  });

  transfers.forEach((transfer) => {
    const equipment = findEquipment(transfer, equipmentList);
    const normalized = normalizeEquipmentIdentity(transfer, equipment);
    const key = getEquipmentGroupKey(transfer.hopDongCongTrinh, normalized.maThietBi, normalized.tenThietBi);
    const row = groups.get(key) || {
      key,
      hopDongCongTrinh: transfer.hopDongCongTrinh,
      maThietBi: normalized.maThietBi,
      tenThietBi: normalized.tenThietBi,
      nguonSoHuu: normalized.nguonSoHuu,
      tongGioHoatDong: 0,
      transferDays: 0,
      donGiaThueGio: normalized.donGiaThueGio,
      donGiaThueNgay: normalized.donGiaThueNgay,
      chiPhiThietBi: 0,
      trangThai: "",
    };

    row.transferDays += calculateEquipmentTransferDays(transfer);
    groups.set(key, row);
  });

  return Array.from(groups.values()).map((row) => {
    const soNgayTaiCongTrinh = Math.max(row.transferDays, row.tongGioHoatDong > 0 ? 1 : 0);
    const chiPhiThietBi =
      row.donGiaThueGio > 0
        ? row.tongGioHoatDong * row.donGiaThueGio
        : row.donGiaThueNgay > 0
          ? soNgayTaiCongTrinh * row.donGiaThueNgay
          : 0;
    const trangThai = getEquipmentCostStatus(row);

    return {
      ...row,
      soNgayTaiCongTrinh,
      chiPhiThietBi,
      trangThai,
      transferDays: undefined,
    };
  });
}

function findEquipment(source, equipmentList) {
  const sourceCode = normalizeText(source.maThietBi);
  const sourceName = normalizeText(source.tenThietBi);

  if (sourceCode) {
    const byCode = equipmentList.find((item) => normalizeText(item.maThietBi) === sourceCode);
    if (byCode) return byCode;
  }

  if (sourceName) {
    const byExactName = equipmentList.find((item) => normalizeText(item.tenThietBi) === sourceName);
    if (byExactName) return byExactName;

    const byIncludedName = equipmentList.find((item) => {
      const itemName = normalizeText(item.tenThietBi);
      return itemName.includes(sourceName) || sourceName.includes(itemName);
    });
    if (byIncludedName) return byIncludedName;
  }

  return null;
}

function getEquipmentGroupKey(project, code, name) {
  return [project || "-", code || name || "unknown"].join("|");
}

function normalizeEquipmentIdentity(source, equipment) {
  return {
    maThietBi: equipment?.maThietBi || source.maThietBi || "",
    tenThietBi: equipment?.tenThietBi || source.tenThietBi || "Thiết bị chưa đặt tên",
    nguonSoHuu: equipment?.nguonSoHuu || source.nguon || source.nguonSoHuu || "",
    donGiaThueGio: toNumber(equipment?.donGiaThueGio),
    donGiaThueNgay: toNumber(equipment?.donGiaThueNgay),
  };
}

function calculateEquipmentTransferDays(transfer) {
  if (!transfer.ngayNhan) return 1;

  const start = new Date(transfer.ngayNhan);
  const end = new Date(transfer.ngayTraThucTe || transfer.ngayTraDuKien || new Date().toISOString().slice(0, 10));
  const diff = end.getTime() - start.getTime();

  if (Number.isNaN(diff) || diff < 0) return 1;
  return Math.max(Math.floor(diff / 86400000) + 1, 1);
}

function getEquipmentCostStatus(row) {
  if (row.donGiaThueGio > 0 || row.donGiaThueNgay > 0) return "Đủ đơn giá";
  if (normalizeText(row.nguonSoHuu).includes("cong ty")) return "Thiết bị công ty";
  if (normalizeText(row.nguonSoHuu).includes("thue ngoai")) return "Thiếu đơn giá thuê";
  return "Thiết bị công ty / chưa tính chi phí";
}

function getLaborReportRows(usages, teams) {
  const groups = new Map();

  usages.forEach((usage) => {
    const key = `${usage.hopDongCongTrinh}|${usage.toDoi}`;
    const laborRateInfo = getLaborRateInfo(usage, teams);
    const row = groups.get(key) || {
      key,
      hopDongCongTrinh: usage.hopDongCongTrinh,
      toDoi: usage.toDoi,
      tongSoNguoi: 0,
      tongSoGio: 0,
      tongCongQuyDoi: 0,
      donGiaCongNgay: laborRateInfo.dayRate,
      hasLaborRate: laborRateInfo.hasLaborRate,
      matchedTeam: laborRateInfo.matchedTeam,
      status: laborRateInfo.status,
      chiPhiNhanCong: 0,
    };

    row.tongSoNguoi += usage.soNguoi;
    row.tongSoGio += usage.soGio;
    row.tongCongQuyDoi = row.tongSoGio / 8;
    row.chiPhiNhanCong = row.hasLaborRate ? row.tongCongQuyDoi * row.donGiaCongNgay : 0;
    groups.set(key, row);
  });

  return Array.from(groups.values());
}

function getLaborRateInfo(usage, teams) {
  const matchedTeam = findLaborTeamForUsage(usage, teams);

  if (!matchedTeam) {
    return {
      dayRate: null,
      hasLaborRate: false,
      matchedTeam: null,
      status: "Thiếu danh mục tổ đội",
    };
  }

  const dayRate = toNumber(matchedTeam.donGiaCongNgay);

  if (dayRate <= 0) {
    return {
      dayRate: null,
      hasLaborRate: false,
      matchedTeam,
      status: "Chưa có đơn giá",
    };
  }

  return {
    dayRate,
    hasLaborRate: true,
    matchedTeam,
    status: "Đủ đơn giá",
  };
}

function findLaborTeamForUsage(usage, teams) {
  const usageCode = normalizeText(usage.maToDoi);
  const usageName = normalizeText(usage.toDoi);

  if (usageCode) {
    const byCode = teams.find((team) => normalizeText(team.maToDoi) === usageCode);
    if (byCode) return byCode;
  }

  const byExactName = teams.find((team) => normalizeText(team.tenToDoi) === usageName);
  if (byExactName) return byExactName;

  const byIncludedName = teams.find((team) => {
    const teamName = normalizeText(team.tenToDoi);
    return teamName.includes(usageName) || usageName.includes(teamName);
  });
  if (byIncludedName) return byIncludedName;

  return teams.find((team) => hasMeaningfulTokenMatch(usage.toDoi, team.tenToDoi));
}

function hasMeaningfulTokenMatch(sourceName, candidateName) {
  const sourceTokens = getLaborNameTokens(sourceName);
  const candidateTokens = getLaborNameTokens(candidateName);

  if (!sourceTokens.length || !candidateTokens.length) return false;
  return (
    sourceTokens.every((token) => candidateTokens.includes(token)) ||
    candidateTokens.every((token) => sourceTokens.includes(token))
  );
}

function getLaborNameTokens(value) {
  const ignoredTokens = new Set(["to", "doi", "nhom", "team", "son", "tho"]);
  return normalizeText(value)
    .split(" ")
    .filter((token) => token && !ignoredTokens.has(token));
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "d")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function getUniqueWbsItems(items) {
  const seen = new Set();

  return items.filter((item) => {
    if (!item.maWbs || seen.has(item.maWbs)) return false;
    seen.add(item.maWbs);
    return true;
  });
}

function getUniqueWbsPlanItems(plans) {
  const seen = new Set();

  return plans.filter((plan) => {
    if (!plan.maWbs || seen.has(plan.maWbs)) return false;
    seen.add(plan.maWbs);
    return true;
  });
}

function getDailyContractOptions(contracts) {
  const preferredContracts = readStorageArray("contracts");
  const projectCatalog = getCatalogSeed("projects");
  const sourceContracts = preferredContracts.length ? preferredContracts : contracts;

  if (sourceContracts.length) {
    return sourceContracts.map((contract, index) => {
      const code = contract.maHopDong || contract.code || contract.contractCode || `HD-${index + 1}`;
      const name = contract.tenCongTrinh || contract.projectName || contract.name || contract.ten || "Công trình";
      return {
        id: contract.id || code,
        label: `${code} - ${name}`,
      };
    });
  }

  if (projectCatalog.length) {
    return projectCatalog.map((project, index) => {
      const code = project.maCongTrinh || project.ma || `CT-${index + 1}`;
      const name = project.tenCongTrinh || project.ten || project.name || "Công trình";

      return {
        id: project.id || code,
        label: `${code} - ${name}`,
      };
    });
  }

  return projects.map((project) => ({
    id: project.ma,
    label: project.ten,
  }));
}

function getContractFilterOptions(contracts) {
  const preferredContracts = readStorageArray("contracts");
  const projectCatalog = getCatalogSeed("projects");
  const sourceContracts = preferredContracts.length ? preferredContracts : contracts;

  if (!sourceContracts.length) {
    return projectCatalog.map((project, index) => {
      const code = project.maCongTrinh || project.ma || `CT-${index + 1}`;
      const name = project.tenCongTrinh || project.ten || project.name || "CÃ´ng trÃ¬nh";

      return {
        id: code,
        label: `${code} - ${name}`,
        name,
      };
    });
  }

  return sourceContracts.map((contract, index) => {
    const code = contract.contractId || contract.maHopDong || contract.contractCode || contract.hopDong || contract.id || `HD-${index + 1}`;
    const name = contract.tenCongTrinh || contract.projectName || contract.congTrinh || contract.name || contract.ten || "Công trình";

    return {
      id: code,
      label: `${code} - ${name}`,
      name,
    };
  });
}

function isMatchContractFilter(item, selectedContractId, selectedContractName, selectedContractLabel) {
  if (selectedContractId === ALL_CONTRACTS) return true;

  return (
    item.contractId === selectedContractId ||
    item.maHopDong === selectedContractId ||
    item.contractCode === selectedContractId ||
    item.hopDong === selectedContractId ||
    item.id === selectedContractId ||
    item.congTrinh === selectedContractName ||
    item.projectName === selectedContractName ||
    item.tenCongTrinh === selectedContractName ||
    item.hopDongCongTrinh === selectedContractName ||
    item.hopDongCongTrinh === selectedContractLabel ||
    isLegacyContractTextMatch(item, selectedContractId, selectedContractName, selectedContractLabel)
  );
}

function isLegacyContractTextMatch(item, selectedContractId, selectedContractName, selectedContractLabel) {
  const candidateTexts = [
    item.hopDongCongTrinh,
    item.congTrinh,
    item.projectName,
    item.tenCongTrinh,
    item.hopDong,
    item.contractName,
  ]
    .filter(Boolean)
    .map((value) => normalizeText(value));
  const selectedTexts = [selectedContractId, selectedContractName, selectedContractLabel]
    .filter(Boolean)
    .map((value) => normalizeText(value));

  return candidateTexts.some((candidate) =>
    selectedTexts.some((selected) => candidate === selected || candidate.includes(selected) || selected.includes(candidate)),
  );
}

function calculatePlanDays(startDate, endDate) {
  if (!startDate || !endDate) return 0;

  const start = new Date(startDate);
  const end = new Date(endDate);
  const diff = end.getTime() - start.getTime();

  if (Number.isNaN(diff) || diff < 0) return 0;
  return Math.floor(diff / 86400000) + 1;
}

function getStatusTone(status) {
  if (status === "Hoàn thành") return "green";
  if (status === "Đang thi công") return "blue";
  if (status === "Chậm tiến độ") return "red";
  return "amber";
}

function toNumber(value) {
  return Number(value || 0);
}

function formatNumber(value) {
  return new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 1 }).format(toNumber(value));
}

function formatPercent(value) {
  return `${new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 1 }).format(toNumber(value))}%`;
}

function formatBytes(value) {
  const bytes = toNumber(value);
  if (bytes <= 0) return "0 KB";

  const units = ["B", "KB", "MB", "GB"];
  const unitIndex = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const size = bytes / (1024 ** unitIndex);

  return `${new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 1 }).format(size)} ${units[unitIndex]}`;
}

function formatCurrency(value) {
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0,
  }).format(toNumber(value));
}

function formatDate(value) {
  if (!value) return "-";
  return new Intl.DateTimeFormat("vi-VN").format(new Date(value));
}

export default App;
