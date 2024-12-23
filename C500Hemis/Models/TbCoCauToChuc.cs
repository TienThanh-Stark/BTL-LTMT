﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using C500Hemis.Models.DM;

namespace C500Hemis.Models;

public partial class TbCoCauToChuc
{
    [Display(Name = "STT")]
    public int IdCoCauToChuc { get; set; }

    [Display(Name = "MÃ PHÒNG BAN ĐƠN VỊ")]
    public string? MaPhongBanDonVi { get; set; }

    [Display(Name = "LOẠI PHÒNG BAN")]
    public int? IdLoaiPhongBan { get; set; }

    [Display(Name = "MÃ PHÒNG BAN ĐƠN VỊ CHA")]
    public string? MaPhongBanDonViCha { get; set; }

    [Display(Name = "TÊN PHÒNG BAN ĐƠN VỊ")]
    public string? TenPhongBanDonVi { get; set; }

    [Display(Name = "SỐ QUYẾT ĐỊNH THÀNH LẬP")]
    public string? SoQuyetDinhThanhLap { get; set; }

    [Display(Name = "NGÀY RA QUYẾT ĐỊNH")]
    [DataType(DataType.Date)]
    [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy}")]
    public DateOnly? NgayRaQuyetDinh { get; set; }

    [Display(Name = "TRẠNG THÁI")]
    public int? IdTrangThai { get; set; }

    [Display(Name = "LOẠI PHÒNG BAN")]
    public virtual DmLoaiPhongBan? IdLoaiPhongBanNavigation { get; set; }

    [Display(Name = "TRẠNG THÁI")]
    public virtual DmTrangThaiCoSoGd? IdTrangThaiNavigation { get; set; }
}