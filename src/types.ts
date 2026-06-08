export interface CounselingDiary {
  // Section 1: Thông tin người học
  ho_ten: string;
  ngay_sinh: string;
  gt_nam: string; // "X" or ""
  gt_nu: string;  // "X" or ""
  gt_kb: string;  // "X" or ""
  ho_ten_cha: string;
  tuoi_cha: string;
  nghe_nghiep_cha: string;
  ho_ten_me: string;
  tuoi_me: string;
  nghe_nghiep_me: string;
  hoan_canh_gia_dinh: string;
  nguoi_cham_soc: string;

  // Section 2: Nội dung tư vấn
  ut_hoc_tap: string;
  ut_quan_he_xa_hoi: string;
  ut_tam_ly: string;
  ut_ky_nang_song: string;
  ut_huong_nghiep: string;
  ut_chinh_sach: string;
  ut_dich_vu_ctxh: string;

  // Section 3: Hình thức tư vấn
  tt_dia_diem: string;
  tt_thoi_gian: string;
  tt_thoi_luong: string;
  on_kenh: string;
  on_thoi_gian: string;
  on_thoi_luong: string;

  // Section 4-8: Nhật ký tự luận
  muc_4_kho_khan_nhu_cau: string;
  muc_5_tom_tat_thong_tin: string;
  muc_6_nhan_dinh_so_bo: string;
  muc_7_hinh_thuc_da_ap_dung: string;
  muc_8_danh_gia_hieu_qua: string;

  // Section 9: Kết thúc
  kt_dung_theo_doi: string;    // "X" or ""
  kt_len_ke_hoach: string;     // "X" or ""
  kt_thuc_hien_chuyen: string; // "X" or ""
  kt_chuyen_gui_noi: string;

  // Section 10: Lịch hẹn & Người thực hiện
  lich_hen_tiep_theo: string;
  ten_nguoi_thuc_hien: string;
}

export const INITIAL_DIARY_STATE: CounselingDiary = {
  ho_ten: "",
  ngay_sinh: "",
  gt_nam: "",
  gt_nu: "",
  gt_kb: "",
  ho_ten_cha: "",
  tuoi_cha: "",
  nghe_nghiep_cha: "",
  ho_ten_me: "",
  tuoi_me: "",
  nghe_nghiep_me: "",
  hoan_canh_gia_dinh: "",
  nguoi_cham_soc: "",

  ut_hoc_tap: "",
  ut_quan_he_xa_hoi: "",
  ut_tam_ly: "",
  ut_ky_nang_song: "",
  ut_huong_nghiep: "",
  ut_chinh_sach: "",
  ut_dich_vu_ctxh: "",

  tt_dia_diem: "",
  tt_thoi_gian: "",
  tt_thoi_luong: "",
  on_kenh: "",
  on_thoi_gian: "",
  on_thoi_luong: "",

  muc_4_kho_khan_nhu_cau: "",
  muc_5_tom_tat_thong_tin: "",
  muc_6_nhan_dinh_so_bo: "",
  muc_7_hinh_thuc_da_ap_dung: "",
  muc_8_danh_gia_hieu_qua: "",

  kt_dung_theo_doi: "",
  kt_len_ke_hoach: "",
  kt_thuc_hien_chuyen: "",
  kt_chuyen_gui_noi: "",

  lich_hen_tiep_theo: "",
  ten_nguoi_thuc_hien: ""
};
