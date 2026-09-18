from functools import wraps
import datetime
import json
import os
import sqlite3
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
import pandas as pd

app = Flask(__name__)
app.secret_key = 'sgk_phu_tho_super_secret_key_2026'

DRIVE_SAVE_DIR = r'G:\My Drive\PHẠM MAI PHONG\SGK_2026-2027\Bao cao nhanh_SGK_ hàng ngày'
DB_FILE = 'sgk_2026.db'
MASTER_EXCEL = 'Danh sach cac truong_Phổ thông.xlsx'

# ---------------------------------------------------------
# DANH MỤC SÁCH GDPT 2018 (3 CẤP HỌC)
# ---------------------------------------------------------
CATALOG_TH = [
    * [
        {'lop': f'Lớp {l}', 'mon': 'Toán', 'ma': f'TO{l}01', 'ten': f'Toán {l} (Tập 1)'}
        for l in range(1, 6)
    ],
    * [
        {'lop': f'Lớp {l}', 'mon': 'Toán', 'ma': f'TO{l}02', 'ten': f'Toán {l} (Tập 2)'}
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Việt',
            'ma': f'TV{l}01',
            'ten': f'Tiếng Việt {l} (Tập 1)',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Việt',
            'ma': f'TV{l}02',
            'ten': f'Tiếng Việt {l} (Tập 2)',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Anh',
            'ma': f'TA{l}01',
            'ten': f'Tiếng Anh {l} Global Success',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Đạo đức',
            'ma': f'DD{l}01',
            'ten': f'Đạo đức {l}',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tự nhiên và Xã hội',
            'ma': f'TN{l}01',
            'ten': f'Tự nhiên và Xã hội {l}',
        }
        for l in [1, 2]
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Khoa học',
            'ma': f'KH{l}01',
            'ten': f'Khoa học {l}',
        }
        for l in [3, 4, 5]
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Lịch sử và Địa lí',
            'ma': f'LS{l}01',
            'ten': f'Lịch sử và Địa lí {l}',
        }
        for l in [3, 4, 5]
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tin học',
            'ma': f'TH{l}01',
            'ten': f'Tin học {l}',
        }
        for l in [3, 4, 5]
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Công nghệ',
            'ma': f'CN{l}01',
            'ten': f'Công nghệ {l}',
        }
        for l in [3, 4, 5]
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Âm nhạc',
            'ma': f'AN{l}01',
            'ten': f'Âm nhạc {l}',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Mĩ thuật',
            'ma': f'MT{l}01',
            'ten': f'Mĩ thuật {l}',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Giáo dục thể chất',
            'ma': f'TC{l}01',
            'ten': f'Giáo dục thể chất {l}',
        }
        for l in range(1, 6)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Hoạt động trải nghiệm',
            'ma': f'HD{l}01',
            'ten': f'Hoạt động trải nghiệm {l}',
        }
        for l in range(1, 6)
    ],
]

CATALOG_THCS = [
    * [
        {'lop': f'Lớp {l}', 'mon': 'Toán', 'ma': f'TO{l}01', 'ten': f'Toán {l} (Tập 1)'}
        for l in range(6, 10)
    ],
    * [
        {'lop': f'Lớp {l}', 'mon': 'Toán', 'ma': f'TO{l}02', 'ten': f'Toán {l} (Tập 2)'}
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Ngữ văn',
            'ma': f'NV{l}01',
            'ten': f'Ngữ văn {l} (Tập 1)',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Ngữ văn',
            'ma': f'NV{l}02',
            'ten': f'Ngữ văn {l} (Tập 2)',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Anh',
            'ma': f'TA{l}01',
            'ten': f'Tiếng Anh {l} Global Success (Tập 1)',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Anh',
            'ma': f'TA{l}02',
            'ten': f'Tiếng Anh {l} Global Success (Tập 2)',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Khoa học tự nhiên',
            'ma': f'KH{l}01',
            'ten': f'Khoa học tự nhiên {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Lịch sử và Địa lí',
            'ma': f'LS{l}01',
            'ten': f'Lịch sử và Địa lí {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Giáo dục công dân',
            'ma': f'GD{l}01',
            'ten': f'Giáo dục công dân {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tin học',
            'ma': f'TH{l}01',
            'ten': f'Tin học {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Công nghệ',
            'ma': f'CN{l}01',
            'ten': f'Công nghệ {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Âm nhạc',
            'ma': f'AN{l}01',
            'ten': f'Âm nhạc {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Mĩ thuật',
            'ma': f'MT{l}01',
            'ten': f'Mĩ thuật {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Giáo dục thể chất',
            'ma': f'TC{l}01',
            'ten': f'Giáo dục thể chất {l}',
        }
        for l in range(6, 10)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Hoạt động trải nghiệm',
            'ma': f'HD{l}01',
            'ten': f'Hoạt động trải nghiệm, hướng nghiệp {l}',
        }
        for l in range(6, 10)
    ],
]

CATALOG_THPT = [
    * [
        {'lop': f'Lớp {l}', 'mon': 'Toán', 'ma': f'TO{l}01', 'ten': f'Toán {l} (Tập 1)'}
        for l in range(10, 13)
    ],
    * [
        {'lop': f'Lớp {l}', 'mon': 'Toán', 'ma': f'TO{l}02', 'ten': f'Toán {l} (Tập 2)'}
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Ngữ văn',
            'ma': f'NV{l}01',
            'ten': f'Ngữ văn {l} (Tập 1)',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Ngữ văn',
            'ma': f'NV{l}02',
            'ten': f'Ngữ văn {l} (Tập 2)',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Anh',
            'ma': f'TA{l}01',
            'ten': f'Tiếng Anh {l} Global Success (Tập 1)',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tiếng Anh',
            'ma': f'TA{l}02',
            'ten': f'Tiếng Anh {l} Global Success (Tập 2)',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Vật lí',
            'ma': f'VL{l}01',
            'ten': f'Vật lí {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Hóa học',
            'ma': f'HH{l}01',
            'ten': f'Hóa học {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Sinh học',
            'ma': f'SH{l}01',
            'ten': f'Sinh học {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Lịch sử',
            'ma': f'LS{l}01',
            'ten': f'Lịch sử {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Địa lí',
            'ma': f'DL{l}01',
            'ten': f'Địa lí {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Giáo dục KT & PL',
            'ma': f'PL{l}01',
            'ten': f'Giáo dục kinh tế và pháp luật {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Tin học',
            'ma': f'TH{l}01',
            'ten': f'Tin học {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Công nghệ',
            'ma': f'CN{l}01',
            'ten': f'Công nghệ {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Âm nhạc',
            'ma': f'AN{l}01',
            'ten': f'Âm nhạc {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Mĩ thuật',
            'ma': f'MT{l}01',
            'ten': f'Mĩ thuật {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Giáo dục thể chất',
            'ma': f'TC{l}01',
            'ten': f'Giáo dục thể chất {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Giáo dục QP-AN',
            'ma': f'QP{l}01',
            'ten': f'Giáo dục quốc phòng và an ninh {l}',
        }
        for l in range(10, 13)
    ],
    * [
        {
            'lop': f'Lớp {l}',
            'mon': 'Hoạt động trải nghiệm',
            'ma': f'HD{l}01',
            'ten': f'Hoạt động trải nghiệm, hướng nghiệp {l}',
        }
        for l in range(10, 13)
    ],
]

FULL_CATALOG_JSON = json.dumps(
    {'TH': CATALOG_TH, 'THCS': CATALOG_THCS, 'THPT': CATALOG_THPT},
    ensure_ascii=False,
)


def create_fresh_templates():
  os.makedirs('templates', exist_ok=True)

  # 1. login.html
  with open('templates/login.html', 'w', encoding='utf-8') as f:
    f.write('''<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Đăng nhập - Hệ Thống SGK 2026-2027</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #eef2f7; min-height: 100vh; }
        .school-list-box { max-height: 480px; overflow-y: auto; font-size: 13px; }
    </style>
</head>
<body>
<div class="container-fluid py-4 px-md-5">
    <div class="text-center mb-4 bg-primary text-white p-3 rounded shadow-sm">
        <h3 class="mb-0 fw-bold">SỞ GIÁO DỤC VÀ ĐÀO TẠO PHÚ THỌ</h3>
        <p class="mb-0 text-white-50 fs-6">Hệ Thống Báo Cáo & Cung Ứng SGK Năm Học 2026 - 2027</p>
    </div>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }} alert-dismissible fade show fw-bold mb-3" role="alert">
            {{ message }}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
          </div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    <div class="row g-4">
        <div class="col-lg-5">
            <div class="card shadow-lg border-0 rounded-3">
                <div class="card-header bg-dark text-white text-center py-3">
                    <h5 class="mb-0 fw-bold">ĐĂNG NHẬP HỆ THỐNG</h5>
                </div>
                <div class="card-body p-4">
                    <ul class="nav nav-pills nav-justified mb-3" id="loginTab">
                        <li class="nav-item">
                            <button class="nav-link active fw-bold" id="school-tab" data-bs-toggle="pill" data-bs-target="#school-login">Trường Báo Cáo</button>
                        </li>
                        <li class="nav-item">
                            <button class="nav-link fw-bold text-danger" id="admin-tab" data-bs-toggle="pill" data-bs-target="#admin-login">Sở GD&ĐT / Quản Trị</button>
                        </li>
                    </ul>

                    <div class="tab-content">
                        <div class="tab-pane fade show active" id="school-login">
                            <form action="{{ url_for('login') }}" method="POST">
                                <input type="hidden" name="login_type" value="school">
                                
                                <div class="mb-3 bg-light p-2 rounded border">
                                    <div class="form-check form-check-inline">
                                        <input class="form-check-input" type="radio" name="school_mode" id="modeSelect" value="select" checked onclick="toggleSchoolInput('select')">
                                        <label class="form-check-label fw-bold" for="modeSelect">Chọn từ danh sách Sở</label>
                                    </div>
                                    <div class="form-check form-check-inline mt-1">
                                        <input class="form-check-input" type="radio" name="school_mode" id="modeCustom" value="custom" onclick="toggleSchoolInput('custom')">
                                        <label class="form-check-label fw-bold text-primary" for="modeCustom">➕ Khai báo đơn vị mới</label>
                                    </div>
                                </div>

                                <div id="selectSchoolBox" class="mb-3">
                                    <label class="form-label fw-bold">Chọn Tên Trường / Đơn vị (*)</label>
                                    <select class="form-select form-select-lg" name="ten_truong_select" id="ten_truong_select">
                                        <option value="">-- Click vào đây để chọn trường --</option>
                                        {% for truong in truong_list %}
                                            <option value="{{ truong }}">{{ truong }}</option>
                                        {% endfor %}
                                    </select>
                                </div>

                                <div id="customSchoolBox" class="mb-3 d-none">
                                    <label class="form-label fw-bold text-primary">Tên Trường / Cơ sở Giáo Dục mới (*)</label>
                                    <input type="text" class="form-control form-control-lg border-primary mb-2" name="ten_truong_custom" id="ten_truong_custom" placeholder="Nhập đầy đủ tên trường mới...">
                                    
                                    <label class="form-label fw-bold">Cấp học của đơn vị (*)</label>
                                    <select class="form-select" name="cap_hoc_custom" id="cap_hoc_custom">
                                        <option value="TH">Tiểu học (Lớp 1 - Lớp 5)</option>
                                        <option value="THCS">Trung học cơ sở (Lớp 6 - Lớp 9)</option>
                                        <option value="THPT">Trung học phổ thông (Lớp 10 - Lớp 12)</option>
                                    </select>
                                </div>

                                <div class="mb-3">
                                    <label class="form-label fw-bold">Mật khẩu <small class="text-muted fw-normal">(Lần đầu bỏ trống)</small></label>
                                    <input type="password" class="form-control form-control-lg" name="password" placeholder="Nhập mật khẩu (nếu đã tạo)...">
                                </div>
                                <button type="submit" class="btn btn-primary w-100 fw-bold py-2 fs-6">🚀 ĐĂNG NHẬP HỆ THỐNG</button>
                            </form>
                        </div>

                        <div class="tab-pane fade" id="admin-login">
                            <form action="{{ url_for('login') }}" method="POST">
                                <input type="hidden" name="login_type" value="admin">
                                <div class="mb-3">
                                    <label class="form-label fw-bold">Cấp độ Quản Trị</label>
                                    <select class="form-select mb-3" name="admin_level">
                                        <option value="admin">🔑 Cán Bộ Quản Trị (Thực thi / Reset Mật Khẩu)</option>
                                        <option value="super_admin">👑 Chủ Tài Khoản Cao Nhất (Super Admin)</option>
                                    </select>
                                </div>
                                <div class="mb-3">
                                    <label class="form-label fw-bold">Mật khẩu Quản trị</label>
                                    <input type="password" class="form-control form-control-lg" name="password" placeholder="Nhập mật khẩu tương ứng..." required>
                                </div>
                                <button type="submit" class="btn btn-danger w-100 fw-bold py-2 fs-6">🔑 ĐĂNG NHẬP QUẢN TRỊ</button>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="col-lg-7">
            <div class="card shadow-lg border-0 rounded-3">
                <div class="card-header bg-white py-3 d-flex justify-content-between align-items-center">
                    <h5 class="mb-0 fw-bold text-primary">TIẾN ĐỘ BÁO CÁO TOÀN TỈNH ({{ da_bao_cao_list|length }}/{{ tong_so_truong }})</h5>
                    <input type="text" id="searchReportStatus" class="form-control form-control-sm w-50" placeholder="🔍 Gõ tên trường để lọc nhanh...">
                </div>
                <div class="card-body p-3">
                    {% set percent = ((da_bao_cao_list|length / (tong_so_truong or 1)) * 100)|round(1) %}
                    <div class="progress mb-3" style="height: 22px;">
                        <div class="progress-bar bg-success progress-bar-striped progress-bar-animated fw-bold" role="progressbar" style="width: {{ percent }}%;">
                            {{ percent }}% Đã hoàn thành
                        </div>
                    </div>

                    <div class="row g-3">
                        <div class="col-md-6">
                            <div class="border border-success rounded p-2 bg-light">
                                <h6 class="fw-bold text-success border-bottom pb-2 mb-2 d-flex justify-content-between">
                                    <span>✅ ĐÃ CHỐT BÁO CÁO</span>
                                    <span class="badge bg-success">{{ da_bao_cao_list|length }}</span>
                                </h6>
                                <div class="school-list-box">
                                    <ul class="list-group list-group-flush" id="submittedList">
                                        {% for school in da_bao_cao_list %}
                                            <li class="list-group-item bg-transparent text-success fw-bold py-1 px-2 school-item">
                                                ✓ {{ school.name }}
                                                {% if school.so_lan_sua > 0 %}
                                                    <span class="badge bg-warning text-dark float-end ms-1" style="font-size:10px;">Sửa lần {{ school.so_lan_sua }}</span>
                                                {% endif %}
                                                {% if school.is_new == 1 %}
                                                    <span class="badge bg-info text-dark float-end" style="font-size:10px;">🟢 Khai báo mới</span>
                                                {% endif %}
                                            </li>
                                        {% else %}
                                            <li class="list-group-item bg-transparent text-muted py-2">Chưa có trường nào chốt.</li>
                                        {% endfor %}
                                    </ul>
                                </div>
                            </div>
                        </div>

                        <div class="col-md-6">
                            <div class="border border-danger rounded p-2 bg-light">
                                <h6 class="fw-bold text-danger border-bottom pb-2 mb-2 d-flex justify-content-between">
                                    <span>⏳ CHƯA CHỐT BÁO CÁO</span>
                                    <span class="badge bg-danger">{{ chua_bao_cao_list|length }}</span>
                                </h6>
                                <div class="school-list-box">
                                    <ul class="list-group list-group-flush" id="pendingList">
                                        {% for school in chua_bao_cao_list %}
                                            <li class="list-group-item bg-transparent text-secondary py-1 px-2 school-item">
                                                • {{ school.name }}
                                                {% if school.is_new == 1 %}
                                                    <span class="badge bg-info text-dark float-end" style="font-size:10px;">🟢 Khai báo mới</span>
                                                {% endif %}
                                            </li>
                                        {% else %}
                                            <li class="list-group-item bg-transparent text-success py-2">100% các trường đã chốt!</li>
                                        {% endfor %}
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    function toggleSchoolInput(mode) {
        const selectBox = document.getElementById('selectSchoolBox');
        const customBox = document.getElementById('customSchoolBox');
        const selectElem = document.getElementById('ten_truong_select');
        const customElem = document.getElementById('ten_truong_custom');

        if (mode === 'custom') {
            selectBox.classList.add('d-none');
            customBox.classList.remove('d-none');
            selectElem.required = false;
            customElem.required = true;
        } else {
            selectBox.classList.remove('d-none');
            customBox.classList.add('d-none');
            selectElem.required = true;
            customElem.required = false;
        }
    }

    document.getElementById('searchReportStatus')?.addEventListener('keyup', function() {
        const filter = this.value.toLowerCase();
        const items = document.querySelectorAll('.school-item');
        items.forEach(item => {
            const text = item.textContent.toLowerCase();
            item.style.display = text.includes(filter) ? '' : 'none';
        });
    });
</script>
</body>
</html>''')

  # 2. set_password.html
  with open('templates/set_password.html', 'w', encoding='utf-8') as f:
    f.write('''<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Đặt mật khẩu mới - {{ ten_truong }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light d-flex align-items-center justify-content-center vh-100">
<div class="card shadow-lg border-0 p-4" style="max-width: 460px; width: 100%;">
    <div class="text-center mb-3">
        <h4 class="text-primary fw-bold mb-1">ĐẶT MẬT KHẨU BẢO MẬT</h4>
        <span class="badge bg-info text-dark fs-6">{{ ten_truong }}</span>
    </div>
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }} alert-dismissible fade show small" role="alert">
            {{ message }}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
          </div>
        {% endfor %}
      {% endif %}
    {% endwith %}
    <form action="{{ url_for('set_password') }}" method="POST">
        <div class="mb-3">
            <label class="form-label fw-bold">Tạo mật khẩu mới (*)</label>
            <input type="password" class="form-control form-control-lg" name="new_password" placeholder="Nhập mật khẩu tự chọn..." required minlength="4">
        </div>
        <div class="mb-3">
            <label class="form-label fw-bold">Xác nhận lại mật khẩu (*)</label>
            <input type="password" class="form-control form-control-lg" name="confirm_password" placeholder="Gõ lại mật khẩu..." required minlength="4">
        </div>
        <button type="submit" class="btn btn-success w-100 fw-bold py-2 fs-6">💾 LƯU MẬT KHẨU & BÁO CÁO</button>
    </form>
</div>
</body>
</html>''')

  # 3. index.html
  with open('templates/index.html', 'w', encoding='utf-8') as f:
    f.write('''<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Biểu Mẫu SGK - {{ session['don_vi'] }}</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .table-input-header { background-color: #0d6efd; color: #ffffff; text-align: center; vertical-align: middle; font-size: 13px; }
        .table-input-body td { vertical-align: middle; padding: 4px; }
        .table-input-body input, .table-input-body select { font-size: 13px; padding: 4px 6px; }
    </style>
</head>
<body class="bg-light">
<div class="container-fluid py-3">
    <div class="d-flex justify-content-between align-items-center mb-3 bg-white p-3 rounded shadow-sm border-start border-4 border-primary">
        <h4 class="m-0 fw-bold text-primary">SỞ GD&ĐT PHÚ THỌ - HỆ THỐNG CUNG ỨNG SGK 2026-2027</h4>
        <div>
            <span class="me-3 fw-bold text-dark">
                Tài khoản: 
                {% if session['role'] == 'super_admin' %}
                    <span class="badge bg-danger fs-6">👑 CHỦ TÀI KHOẢN (Super Admin)</span>
                {% elif session['role'] == 'admin' %}
                    <span class="badge bg-warning text-dark fs-6">🔑 CÁN BỘ QUẢN TRỊ (Admin)</span>
                {% else %}
                    <span class="text-primary fw-bold">{{ session['don_vi'] }}</span>
                {% endif %}
            </span>
            <a href="{{ url_for('logout') }}" class="btn btn-sm btn-outline-danger fw-bold">Đăng xuất</a>
        </div>
    </div>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        {% for category, message in messages %}
          <div class="alert alert-{{ category }} alert-dismissible fade show fw-bold" role="alert">
            {{ message }}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
          </div>
        {% endfor %}
      {% endif %}
    {% endwith %}

    {% if session['role'] in ['admin', 'super_admin'] %}
    <div class="card mb-4 shadow-sm border-danger">
        <div class="card-header bg-danger text-white fw-bold fs-5">
            {% if session['role'] == 'super_admin' %}
                👑 QUẢN TRỊ CAO NHẤT (SUPER ADMIN)
            {% else %}
                🔑 BÀN TRỰC BAN CÁN BỘ QUẢN TRỊ
            {% endif %}
        </div>
        <div class="card-body text-center py-4">
            <a href="{{ url_for('export') }}" class="btn btn-success btn-lg fw-bold px-4 py-2">
                📊 XUẤT BÁO CÁO TỔNG HỢP EXCEL (Tự động lưu về Google Drive)
            </a>
            <p class="text-muted small mt-2 mb-0">File lưu trực tiếp tại: <code>G:\\My Drive\\PHẠM MAI PHONG\\SGK_2026-2027\\Bao cao nhanh_SGK_ hàng ngày</code></p>
        </div>
    </div>

    <div class="card shadow-sm mb-4 border-danger">
        <div class="card-header bg-dark text-white fw-bold d-flex justify-content-between align-items-center">
            <span>🔑 DANG SÁCH TÀI KHOẢN & HỖ TRỢ RESET MẬT KHẨU CÁC TRƯỜNG</span>
            <input type="text" id="searchSchool" class="form-control form-control-sm w-25" placeholder="🔍 Gõ tên trường để tìm...">
        </div>
        <div class="card-body p-0">
            <div class="table-responsive" style="max-height: 380px; overflow-y: auto;">
                <table class="table table-striped table-hover align-middle mb-0 text-center" id="schoolAccountTable" style="font-size: 13px;">
                    <thead class="table-secondary sticky-top">
                        <tr>
                            <th style="width: 50px;">STT</th>
                            <th class="text-start">Tên Trường / Cơ sở Giáo Dục</th>
                            <th>Loại Đơn Vị</th>
                            <th>Cấp Học</th>
                            <th>Mật Khẩu</th>
                            <th>Trạng Thái Báo Cáo</th>
                            <th>Thao Tác Thực Thi</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for school in school_accounts %}
                        <tr>
                            <td>{{ loop.index }}</td>
                            <td class="fw-bold text-start school-name">{{ school.ten_truong }}</td>
                            <td>
                                {% if school.is_new_unit == 1 %}
                                    <span class="badge bg-warning text-dark">🟢 Khai báo mới</span>
                                {% else %}
                                    <span class="badge bg-secondary">Danh mục Sở</span>
                                {% endif %}
                            </td>
                            <td><span class="badge bg-info text-dark">{{ school.cap_hoc or 'Chưa rõ' }}</span></td>
                            <td>
                                {% if school.password %}
                                    <span class="badge bg-success">Đã tạo mật khẩu</span>
                                {% else %}
                                    <span class="badge bg-warning text-dark">Chưa có (Hoặc đã Reset)</span>
                                {% endif %}
                            </td>
                            <td>
                                {% if school.is_locked == 1 %}
                                    <span class="badge bg-danger">🔒 Đã chốt</span>
                                    {% if school.so_lan_sua > 0 %}
                                        <span class="badge bg-warning text-dark">Sửa lần {{ school.so_lan_sua }}</span>
                                    {% endif %}
                                {% else %}
                                    <span class="badge bg-info text-dark">🔓 Đang mở</span>
                                {% endif %}
                            </td>
                            <td>
                                <a href="{{ url_for('admin_reset_school', ten_truong=school.ten_truong) }}" 
                                   class="btn btn-sm btn-outline-danger fw-bold"
                                   onclick="return confirm('Xác nhận Reset mật khẩu và Mở khóa cho trường này?');">
                                    🔑 Reset Mật Khẩu
                                </a>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    {% else %}

    <div class="card shadow-sm mb-4 border-primary">
        <div class="card-header bg-primary text-white fw-bold d-flex justify-content-between align-items-center">
            <span>BẢNG KÊ KHAI ĐĂNG KÝ SÁCH GIÁO KHOA 2026-2027</span>
            <small class="text-white-50">Tự động nạp toàn bộ danh mục sách GDPT 2018 theo cấp học của đơn vị</small>
        </div>
        <div class="card-body p-3">
            <form action="{{ url_for('add_batch') }}" method="POST" id="batchForm">
                <div class="row g-3 mb-3 bg-light p-2 rounded border">
                    <div class="col-md-3">
                        <label class="form-label fw-bold">Xã/Phường/Đặc khu (*)</label>
                        <select class="form-select form-select-sm" name="xa_phuong" id="xa_phuong" required>
                            <option value="">-- Chọn Xã/Phường --</option>
                            {% for xa in xa_list %}
                                <option value="{{ xa }}">{{ xa }}</option>
                            {% endfor %}
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label fw-bold">Cấp học (*)</label>
                        <select class="form-select form-select-sm" name="cap_hoc" id="cap_hoc" required>
                            <option value="TH" {% if session['cap_hoc'] == 'TH' %}selected{% endif %}>Tiểu Học (Lớp 1 - Lớp 5)</option>
                            <option value="THCS" {% if session['cap_hoc'] == 'THCS' %}selected{% endif %}>THCS (Lớp 6 - Lớp 9)</option>
                            <option value="THPT" {% if session['cap_hoc'] == 'THPT' %}selected{% endif %}>THPT (Lớp 10 - Lớp 12)</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label fw-bold text-primary">Lọc theo Khối Lớp</label>
                        <select class="form-select form-select-sm border-primary fw-bold" id="filter_lop">
                            <option value="ALL">-- Tất cả các khối lớp --</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label class="form-label fw-bold">Đơn vị cung ứng mặc định</label>
                        <select class="form-select form-select-sm" name="default_supplier" id="default_supplier">
                            <option value="Công ty Sách Phú Thọ">Công ty Sách Phú Thọ</option>
                            <option value="Công ty Sách Vĩnh Phúc">Công ty Sách Vĩnh Phúc</option>
                            <option value="Công ty Sách Hòa Bình">Công ty Sách Hòa Bình</option>
                        </select>
                    </div>
                </div>

                <div class="table-responsive">
                    <table class="table table-bordered table-hover align-middle mb-3" id="inputTable">
                        <thead class="table-input-header">
                            <tr>
                                <th style="width: 40px;">STT</th>
                                <th style="width: 90px;">Lớp</th>
                                <th style="width: 140px;">Môn</th>
                                <th style="width: 90px;">Mã sách</th>
                                <th>Tên sách</th>
                                <th style="width: 140px;">Tổng HS (chính + phân hiệu + Điểm trường)</th>
                                <th style="width: 120px;">Số HS đã có SGK</th>
                                <th style="width: 120px;" class="bg-warning text-dark">Số HS chưa có SGK</th>
                                <th style="width: 130px;" class="bg-success text-white">Đề nghị ĐK mua BS</th>
                            </tr>
                        </thead>
                        <tbody class="table-input-body" id="bookTableBody">
                        </tbody>
                    </table>
                </div>

                <div class="d-flex justify-content-between align-items-center">
                    <button type="button" class="btn btn-sm btn-outline-primary fw-bold" id="btnAddRow">
                        ➕ Bổ sung 1 dòng sách ngoài danh mục
                    </button>
                    <button type="submit" class="btn btn-success fw-bold px-4 py-2" id="btnSaveBatch">
                        💾 LƯU DANH MỤC SÁCH
                    </button>
                </div>
            </form>
        </div>
    </div>

    <div class="card shadow-sm mb-4 border-warning">
        <div class="card-body text-center bg-light py-3">
            <h5 class="text-danger fw-bold m-0 mb-1">Hoàn thành biểu mẫu báo cáo?</h5>
            <p class="text-muted small mb-2">Sau khi cơ sở nhập/sửa xong số liệu, bấm nút chốt để hoàn tất gửi về Sở GD&ĐT.</p>
            <form action="{{ url_for('finalize') }}" method="POST" onsubmit="return confirm('Bạn có chắc chắn muốn chốt số liệu và gửi Sở?');">
                <button type="submit" class="btn btn-warning fw-bold px-5 py-2">✅ CHỐT SỐ LIỆU & GỬI SỞ</button>
            </form>
        </div>
    </div>
    {% endif %}

    <div class="card shadow-sm">
        <div class="card-header bg-dark text-white fw-bold">Danh sách Báo cáo đã lưu trong hệ thống</div>
        <div class="card-body p-0">
            <div class="table-responsive">
                <table class="table table-striped table-hover m-0 align-middle text-center" style="font-size: 13px;">
                    <thead class="table-dark">
                        <tr>
                            <th>STT</th>
                            <th>Xã/Phường</th>
                            <th class="text-start">Trường</th>
                            <th>Cấp</th>
                            <th>Lớp</th>
                            <th>Môn</th>
                            <th>Mã sách</th>
                            <th class="text-start">Tên sách</th>
                            <th>Tổng HS</th>
                            <th>Số HS đã có SGK</th>
                            <th class="text-danger fw-bold">Số HS chưa có SGK</th>
                            <th class="text-success fw-bold">Đề nghị ĐK mua BS</th>
                            <th>Đơn vị cung ứng</th>
                            <th style="width: 80px;">Thao tác</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in records %}
                        <tr>
                            <td>{{ loop.index }}</td>
                            <td>{{ row.xa_phuong }}</td>
                            <td class="fw-bold text-start">{{ row.ten_truong }}</td>
                            <td>{{ row.cap_hoc }}</td>
                            <td>{{ row.lop }}</td>
                            <td>{{ row.mon_hoc }}</td>
                            <td><code>{{ row.ma_sach }}</code></td>
                            <td class="text-start fw-bold">{{ row.ten_sach }}</td>
                            <td>{{ row.tong_hs }}</td>
                            <td>{{ row.so_hs_da_co }}</td>
                            <td class="text-danger fw-bold">{{ row.so_hs_chua_co }}</td>
                            <td class="text-success fw-bold">{{ row.de_nghi_dk }}</td>
                            <td>{{ row.don_vi_cung_cap }}</td>
                            <td>
                                <a href="{{ url_for('delete_record', record_id=row.id) }}" 
                                   class="btn btn-sm btn-outline-danger py-0 px-2 fw-bold"
                                   onclick="return confirm('Bạn có chắc muốn xóa dòng số liệu này?');">
                                    🗑️ Xóa
                                </a>
                            </td>
                        </tr>
                        {% else %}
                        <tr>
                            <td colspan="14" class="text-center text-muted py-4">Chưa có dữ liệu báo cáo nào được lưu.</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    const catalogData = {{ FULL_CATALOG_JSON|safe }};
    const existingRecords = {{ EXISTING_RECORDS_JSON|safe }};

    const capHocSelect = document.getElementById('cap_hoc');
    const filterLopSelect = document.getElementById('filter_lop');
    const tableBody = document.getElementById('bookTableBody');
    const xaPhuongSelect = document.getElementById('xa_phuong');

    function updateLopFilterOptions(cap) {
        if (!filterLopSelect) return;
        filterLopSelect.innerHTML = '<option value="ALL">-- Tất cả các khối lớp --</option>';
        if (cap === 'TH') {
            for (let i = 1; i <= 5; i++) filterLopSelect.innerHTML += `<option value="Lớp ${i}">Khối Lớp ${i}</option>`;
        } else if (cap === 'THCS') {
            for (let i = 6; i <= 9; i++) filterLopSelect.innerHTML += `<option value="Lớp ${i}">Khối Lớp ${i}</option>`;
        } else if (cap === 'THPT') {
            for (let i = 10; i <= 12; i++) filterLopSelect.innerHTML += `<option value="Lớp ${i}">Khối Lớp ${i}</option>`;
        }
    }

    function createRowHTML(index, item) {
        const saved = existingRecords[item.ma] || {};
        const tong = saved.tong_hs || 0;
        const daco = saved.so_hs_da_co || 0;
        const chuaco = saved.so_hs_chua_co || Math.max(0, tong - daco);
        const denghi = saved.de_nghi_dk || chuaco;

        if (saved.xa_phuong && xaPhuongSelect && !xaPhuongSelect.value) {
            xaPhuongSelect.value = saved.xa_phuong;
        }

        return `
            <td class="text-center fw-bold row-stt">${index}</td>
            <td class="text-center"><input type="text" class="form-control form-control-sm text-center" name="lop[]" value="${item.lop}"></td>
            <td><input type="text" class="form-control form-control-sm" name="mon_hoc[]" value="${item.mon}"></td>
            <td class="text-center"><input type="text" class="form-control form-control-sm text-center fw-bold text-primary" name="ma_sach[]" value="${item.ma}"></td>
            <td><input type="text" class="form-control form-control-sm fw-bold" name="ten_sach[]" value="${item.ten}"></td>
            <td><input type="number" class="form-control form-control-sm num-tong" name="tong_hs[]" min="0" value="${tong}"></td>
            <td><input type="number" class="form-control form-control-sm num-daco" name="so_hs_da_co[]" min="0" value="${daco}"></td>
            <td><input type="number" class="form-control form-control-sm num-chuaco bg-light text-danger fw-bold" name="so_hs_chua_co[]" value="${chuaco}" readonly></td>
            <td><input type="number" class="form-control form-control-sm num-denghi text-success fw-bold" name="de_nghi_dk[]" min="0" value="${denghi}"></td>
        `;
    }

    function renderTable() {
        if (!capHocSelect || !tableBody) return;
        const cap = capHocSelect.value;
        const filterLop = filterLopSelect ? filterLopSelect.value : 'ALL';
        tableBody.innerHTML = '';

        if (!cap || !catalogData[cap]) {
            tableBody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-4">Vui lòng chọn Cấp học để hiển thị danh mục sách.</td></tr>`;
            return;
        }

        let books = catalogData[cap];
        if (filterLop !== 'ALL') {
            books = books.filter(b => b.lop === filterLop);
        }

        books.forEach((b, index) => {
            const tr = document.createElement('tr');
            tr.innerHTML = createRowHTML(index + 1, b);
            tableBody.appendChild(tr);
        });

        bindCalcEvents();
    }

    function bindCalcEvents() {
        tableBody.querySelectorAll('tr').forEach(tr => {
            const inTong = tr.querySelector('.num-tong');
            const inDaCo = tr.querySelector('.num-daco');
            const inChuaCo = tr.querySelector('.num-chuaco');
            const inDeNghi = tr.querySelector('.num-denghi');

            if (!inTong || !inDaCo) return;

            function calcRow() {
                let tong = parseInt(inTong.value) || 0;
                let daco = parseInt(inDaCo.value) || 0;

                if (daco > tong) {
                    alert(`⚠️ Không hợp lệ! Số HS đã có SGK (${daco}) không được lớn hơn Tổng HS (${tong}).`);
                    inDaCo.value = tong;
                    daco = tong;
                }
                inDaCo.max = tong;

                const chuaco = Math.max(0, tong - daco);
                inChuaCo.value = chuaco;
                inDeNghi.value = chuaco;
            }

            inTong.oninput = calcRow;
            inDaCo.oninput = calcRow;
        });
    }

    document.getElementById('btnAddRow')?.addEventListener('click', function() {
        const rowsCount = tableBody.querySelectorAll('tr').length;
        const tr = document.createElement('tr');
        tr.innerHTML = createRowHTML(rowsCount + 1, { lop: "Khác", mon: "Môn mới", ma: "MA_SO", ten: "Tên sách bổ sung" });
        tableBody.appendChild(tr);
        bindCalcEvents();
    });

    document.getElementById('searchSchool')?.addEventListener('keyup', function() {
        const filter = this.value.toLowerCase();
        const rows = document.querySelectorAll('#schoolAccountTable tbody tr');
        rows.forEach(row => {
            const name = row.querySelector('.school-name').textContent.toLowerCase();
            row.style.display = name.includes(filter) ? '' : 'none';
        });
    });

    capHocSelect?.addEventListener('change', function() {
        updateLopFilterOptions(this.value);
        renderTable();
    });

    filterLopSelect?.addEventListener('change', renderTable);

    window.onload = function() {
        if (capHocSelect) updateLopFilterOptions(capHocSelect.value);
        renderTable();
    };
</script>
</body>
</html>''')


def get_db_connection():
  conn = sqlite3.connect(DB_FILE)
  conn.row_factory = sqlite3.Row
  return conn


def init_db():
  conn = get_db_connection()
  cursor = conn.cursor()

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS school_accounts (
            ten_truong TEXT PRIMARY KEY,
            password TEXT,
            is_locked INTEGER DEFAULT 0,
            cap_hoc TEXT,
            is_new_unit INTEGER DEFAULT 0,
            so_lan_sua INTEGER DEFAULT 0
        )
    """)

  cursor.execute('PRAGMA table_info(school_accounts)')
  cols = [col[1] for col in cursor.fetchall()]
  if 'is_new_unit' not in cols:
    cursor.execute(
        'ALTER TABLE school_accounts ADD COLUMN is_new_unit INTEGER DEFAULT 0'
    )
  if 'so_lan_sua' not in cols:
    cursor.execute(
        'ALTER TABLE school_accounts ADD COLUMN so_lan_sua INTEGER DEFAULT 0'
    )

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS du_lieu_sgk (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            xa_phuong TEXT,
            ten_truong TEXT,
            cap_hoc TEXT,
            lop TEXT,
            mon_hoc TEXT,
            ma_sach TEXT,
            ten_sach TEXT,
            tong_hs INTEGER DEFAULT 0,
            so_hs_da_co INTEGER DEFAULT 0,
            so_hs_chua_co INTEGER DEFAULT 0,
            de_nghi_dk INTEGER DEFAULT 0,
            don_vi_cung_cap TEXT,
            ngay_nhap TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
  conn.commit()
  conn.close()


def detect_cap_hoc(school_name):
  name_upper = school_name.upper()
  if any(k in name_upper for k in ['THCS', 'TRUNG HỌC CƠ SỞ', 'TH&THCS']):
    return 'THCS'
  elif any(k in name_upper for k in ['TIỂU HỌC', 'TH ']):
    return 'TH'
  elif any(
      k in name_upper
      for k in [
          'THPT',
          'TRUNG HỌC PHỔ THÔNG',
          'GDNN-GDTX',
          'GDTX',
          'PT DTNT',
          'TRUNG TÂM',
      ]
  ):
    return 'THPT'
  return 'THCS'


def load_master_data():
  try:
    xls = pd.ExcelFile(MASTER_EXCEL)
    df_truong = pd.read_excel(xls, sheet_name='Trường', header=None)
    df_xa = pd.read_excel(xls, sheet_name='Xã', header=None)

    truong_list = list(
        dict.fromkeys(df_truong[0].dropna().astype(str).str.strip().tolist())
    )
    xa_list = list(
        dict.fromkeys(df_xa[0].dropna().astype(str).str.strip().tolist())
    )
    return truong_list, xa_list
  except Exception as e:
    print('⚠️ Lỗi đọc Master Data Excel:', e)
    return [], []


def login_required(f):

  @wraps(f)
  def decorated_function(*args, **kwargs):
    if 'role' not in session:
      flash('Vui lòng đăng nhập hệ thống!', 'warning')
      return redirect(url_for('login'))
    return f(*args, **kwargs)

  return decorated_function


# KHỞI TẠO TẤT CẢ FILE KHI CHẠY APP
create_fresh_templates()
init_db()


@app.route('/login', methods=['GET', 'POST'])
def login():
  truong_list_master, _ = load_master_data()

  conn = get_db_connection()
  db_accounts = conn.execute('SELECT * FROM school_accounts').fetchall()
  db_account_dict = {a['ten_truong']: a for a in db_accounts}

  all_truongs_set = dict.fromkeys(truong_list_master)
  for a in db_accounts:
    if a['ten_truong'] not in all_truongs_set:
      all_truongs_set[a['ten_truong']] = True

  combined_truong_list = list(all_truongs_set.keys())

  da_bao_cao_list = []
  chua_bao_cao_list = []

  for t in combined_truong_list:
    acc = db_account_dict.get(t)
    is_new = acc['is_new_unit'] if acc else 0
    so_lan_sua = (
        acc['so_lan_sua'] if (acc and acc['so_lan_sua'] is not None) else 0
    )
    unit_info = {'name': t, 'is_new': is_new, 'so_lan_sua': so_lan_sua}

    if acc and acc['is_locked'] == 1:
      da_bao_cao_list.append(unit_info)
    else:
      chua_bao_cao_list.append(unit_info)

  conn.close()

  if request.method == 'POST':
    login_type = request.form.get('login_type')

    if login_type == 'admin':
      admin_level = request.form.get('admin_level', 'admin')
      password = request.form.get('password')

      if admin_level == 'super_admin' and password == 'SuperAdmin@2026':
        session['role'] = 'super_admin'
        session['don_vi'] = 'Chủ Tài Khoản Cao Nhất'
        session['cap_hoc'] = 'ALL'
        return redirect(url_for('index'))
      elif admin_level == 'admin' and password == 'Admin@2026':
        session['role'] = 'admin'
        session['don_vi'] = 'Cán Bộ Quản Trị'
        session['cap_hoc'] = 'ALL'
        return redirect(url_for('index'))
      else:
        flash('Mật khẩu Quản trị không chính xác!', 'danger')

    elif login_type == 'school':
      school_mode = request.form.get('school_mode', 'select')

      if school_mode == 'custom':
        ten_truong = request.form.get('ten_truong_custom', '').strip()
        cap_hoc_custom = request.form.get('cap_hoc_custom', 'THCS')
        is_new_unit = 1
      else:
        ten_truong = request.form.get('ten_truong_select', '').strip()
        cap_hoc_custom = None
        is_new_unit = 0

      password_input = request.form.get('password', '').strip()

      if not ten_truong:
        flash('Vui lòng chọn hoặc nhập tên đơn vị!', 'danger')
        return redirect(url_for('login'))

      conn = get_db_connection()
      account = conn.execute(
          'SELECT * FROM school_accounts WHERE ten_truong = ?', (ten_truong,)
      ).fetchone()

      if not account:
        cap_hoc = (
            cap_hoc_custom if cap_hoc_custom else detect_cap_hoc(ten_truong)
        )
        conn.execute(
            'INSERT INTO school_accounts (ten_truong, password, is_locked,'
            ' cap_hoc, is_new_unit, so_lan_sua) VALUES (?, NULL, 0, ?, ?, 0)',
            (ten_truong, cap_hoc, is_new_unit),
        )
        conn.commit()
        account = conn.execute(
            'SELECT * FROM school_accounts WHERE ten_truong = ?', (ten_truong,)
        ).fetchone()

      db_password = account['password']

      if not db_password or db_password.strip() == '':
        conn.close()
        session['role'] = 'school'
        session['don_vi'] = ten_truong
        session['cap_hoc'] = account['cap_hoc'] or detect_cap_hoc(ten_truong)
        session['must_set_password'] = True
        return redirect(url_for('set_password'))

      if db_password == password_input:
        conn.close()
        session['role'] = 'school'
        session['don_vi'] = ten_truong
        session['cap_hoc'] = account['cap_hoc'] or detect_cap_hoc(ten_truong)
        session['must_set_password'] = False
        return redirect(url_for('index'))
      else:
        conn.close()
        flash('Mật khẩu không đúng! Vui lòng liên hệ Sở để reset.', 'danger')

  return render_template(
      'login.html',
      truong_list=truong_list_master,
      tong_so_truong=len(combined_truong_list),
      da_bao_cao_list=da_bao_cao_list,
      chua_bao_cao_list=chua_bao_cao_list,
  )


@app.route('/set_password', methods=['GET', 'POST'])
@login_required
def set_password():
  if session.get('role') != 'school':
    return redirect(url_for('index'))

  ten_truong = session.get('don_vi')

  if request.method == 'POST':
    new_pass = request.form.get('new_password', '').strip()
    confirm_pass = request.form.get('confirm_password', '').strip()

    if not new_pass or len(new_pass) < 4:
      flash('Mật khẩu mới phải có ít nhất 4 ký tự!', 'danger')
      return redirect(url_for('set_password'))

    if new_pass != confirm_pass:
      flash('Mật khẩu xác nhận không khớp!', 'danger')
      return redirect(url_for('set_password'))

    conn = get_db_connection()
    conn.execute(
        'UPDATE school_accounts SET password = ? WHERE ten_truong = ?',
        (new_pass, ten_truong),
    )
    conn.commit()
    conn.close()

    session['must_set_password'] = False
    flash('✅ Thiết lập mật khẩu thành công!', 'success')
    return redirect(url_for('index'))

  return render_template('set_password.html', ten_truong=ten_truong)


@app.route('/logout')
def logout():
  session.clear()
  return redirect(url_for('login'))


@app.route('/')
@app.route('/school_dashboard')
@login_required
def index():
  if session.get('must_set_password'):
    return redirect(url_for('set_password'))

  conn = get_db_connection()
  role = session.get('role')
  don_vi = session.get('don_vi')
  _, xa_list = load_master_data()

  if role in ['admin', 'super_admin']:
    records = conn.execute(
        'SELECT * FROM du_lieu_sgk ORDER BY id DESC'
    ).fetchall()
    school_accounts = conn.execute(
        'SELECT * FROM school_accounts ORDER BY is_new_unit DESC, ten_truong'
        ' ASC'
    ).fetchall()
    existing_records_dict = {}
  else:
    records = conn.execute(
        'SELECT * FROM du_lieu_sgk WHERE ten_truong = ? ORDER BY id DESC',
        (don_vi,),
    ).fetchall()
    school_accounts = []
    existing_records_dict = {
        r['ma_sach']: {
            'tong_hs': r['tong_hs'],
            'so_hs_da_co': r['so_hs_da_co'],
            'so_hs_chua_co': r['so_hs_chua_co'],
            'de_nghi_dk': r['de_nghi_dk'],
            'xa_phuong': r['xa_phuong'],
        }
        for r in records
    }

  conn.close()

  return render_template(
      'index.html',
      records=records,
      xa_list=xa_list,
      school_accounts=school_accounts,
      EXISTING_RECORDS_JSON=json.dumps(
          existing_records_dict, ensure_ascii=False
      ),
      FULL_CATALOG_JSON=FULL_CATALOG_JSON,
  )


@app.route('/add_batch', methods=['POST'])
@login_required
def add_batch():
  if session.get('role') in ['admin', 'super_admin']:
    flash('Tài khoản Quản trị không tham gia kê khai dữ liệu!', 'danger')
    return redirect(url_for('index'))

  ten_truong = session.get('don_vi')
  xa_phuong = request.form.get('xa_phuong')
  cap_hoc = request.form.get('cap_hoc')
  don_vi_cung_cap = request.form.get('default_supplier', 'Công ty Sách Phú Thọ')

  if not xa_phuong or not cap_hoc:
    flash('Vui lòng chọn Xã/Phường và Cấp học!', 'danger')
    return redirect(url_for('index'))

  lops = request.form.getlist('lop[]')
  mons = request.form.getlist('mon_hoc[]')
  ma_sachs = request.form.getlist('ma_sach[]')
  ten_sachs = request.form.getlist('ten_sach[]')
  tong_hss = request.form.getlist('tong_hs[]')
  so_hs_da_cos = request.form.getlist('so_hs_da_co[]')
  de_nghi_dks = request.form.getlist('de_nghi_dk[]')

  records_to_insert = []

  for i in range(len(ma_sachs)):
    try:
      tong_hs = int(tong_hss[i] or 0)
      so_da_co = int(so_hs_da_cos[i] or 0)
      de_nghi = int(de_nghi_dks[i] or 0)
    except Exception:
      continue

    if tong_hs == 0 and so_da_co == 0 and de_nghi == 0:
      continue

    if so_da_co > tong_hs:
      flash(
          f'❌ Lỗi số liệu ở sách "{ten_sachs[i]}": [Số HS đã có SGK ({so_da_co})]'
          f' lớn hơn [Tổng HS ({tong_hs})]',
          'danger',
      )
      return redirect(url_for('index'))

    so_chua_co = max(0, tong_hs - so_da_co)

    records_to_insert.append((
        xa_phuong,
        ten_truong,
        cap_hoc,
        lops[i],
        mons[i],
        ma_sachs[i],
        ten_sachs[i],
        tong_hs,
        so_da_co,
        so_chua_co,
        de_nghi,
        don_vi_cung_cap,
    ))

  if not records_to_insert:
    flash('⚠️ Chưa có số liệu nào được điền!', 'warning')
    return redirect(url_for('index'))

  conn = get_db_connection()
  conn.execute('DELETE FROM du_lieu_sgk WHERE ten_truong = ?', (ten_truong,))
  conn.executemany(
      """
        INSERT INTO du_lieu_sgk (
            xa_phuong, ten_truong, cap_hoc, lop, mon_hoc, ma_sach, ten_sach,
            tong_hs, so_hs_da_co, so_hs_chua_co, de_nghi_dk, don_vi_cung_cap
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
      records_to_insert,
  )
  conn.commit()
  conn.close()

  flash(
      f'✅ Đã lưu {len(records_to_insert)} danh mục sách vào hệ thống!',
      'success',
  )
  return redirect(url_for('index'))


@app.route('/delete_record/<int:record_id>')
@login_required
def delete_record(record_id):
  conn = get_db_connection()
  if session.get('role') in ['admin', 'super_admin']:
    conn.execute('DELETE FROM du_lieu_sgk WHERE id = ?', (record_id,))
  else:
    conn.execute(
        'DELETE FROM du_lieu_sgk WHERE id = ? AND ten_truong = ?',
        (record_id, session.get('don_vi')),
    )
  conn.commit()
  conn.close()

  flash('🗑️ Đã xóa dòng số liệu thành công!', 'info')
  return redirect(url_for('index'))


@app.route('/finalize', methods=['POST'])
@login_required
def finalize():
  if session.get('role') == 'school':
    ten_truong = session.get('don_vi')
    conn = get_db_connection()

    count = conn.execute(
        'SELECT COUNT(*) as cnt FROM du_lieu_sgk WHERE ten_truong = ?',
        (ten_truong,),
    ).fetchone()['cnt']
    if count == 0:
      flash('❌ Bạn chưa nhập số liệu nào để chốt báo cáo!', 'danger')
      conn.close()
      return redirect(url_for('index'))

    acc = conn.execute(
        'SELECT * FROM school_accounts WHERE ten_truong = ?', (ten_truong,)
    ).fetchone()
    curr_so_lan = (
        acc['so_lan_sua'] if (acc and acc['so_lan_sua'] is not None) else 0
    )

    if acc and acc['is_locked'] == 1:
      new_so_lan = curr_so_lan + 1
    else:
      new_so_lan = curr_so_lan

    conn.execute(
        'UPDATE school_accounts SET is_locked = 1, so_lan_sua = ? WHERE'
        ' ten_truong = ?',
        (new_so_lan, ten_truong),
    )
    conn.commit()
    conn.close()

    session.clear()
    flash(
        f'✅ Trường "{ten_truong}" đã chốt và gửi báo cáo về Sở GD&ĐT thành'
        ' công!',
        'success',
    )
    return redirect(url_for('login'))
  return redirect(url_for('index'))


@app.route('/admin/reset_school/<path:ten_truong>')
@login_required
def admin_reset_school(ten_truong):
  if session.get('role') not in ['admin', 'super_admin']:
    flash('Không có quyền truy cập!', 'danger')
    return redirect(url_for('index'))

  conn = get_db_connection()
  conn.execute(
      'UPDATE school_accounts SET password = NULL, is_locked = 0, so_lan_sua ='
      ' 0 WHERE ten_truong = ?',
      (ten_truong,),
  )
  conn.commit()
  conn.close()

  flash(f'🔑 Đã Reset mật khẩu cho trường "{ten_truong}".', 'success')
  return redirect(url_for('index'))


@app.route('/export')
@login_required
def export():
  if session.get('role') not in ['admin', 'super_admin']:
    flash('Chỉ Quản trị mới có quyền xuất báo cáo Excel.', 'danger')
    return redirect(url_for('index'))

  conn = get_db_connection()
  df = pd.read_sql_query('SELECT * FROM du_lieu_sgk', conn)
  conn.close()

  df.rename(
      columns={
          'xa_phuong': 'Xã/Phường/Đặc khu',
          'ten_truong': 'Tên cơ sở giáo dục',
          'cap_hoc': 'Cấp học',
          'lop': 'Lớp',
          'mon_hoc': 'Môn học',
          'ma_sach': 'Mã sách',
          'ten_sach': 'Tên sách',
          'tong_hs': 'Tổng HS (chính + phân hiệu + Điểm trường)',
          'so_hs_da_co': 'Số HS đã có SGK',
          'so_hs_chua_co': 'Số HS chưa có SGK',
          'de_nghi_dk': 'Đề nghị ĐK mua bổ sung',
          'don_vi_cung_cap': 'Đơn vị cung ứng',
          'ngay_nhap': 'Ngày nhập',
      },
      inplace=True,
  )

  now_str = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
  file_name = f'Bao_cao_nhanh_SGK_{now_str}.xlsx'

  try:
    os.makedirs(DRIVE_SAVE_DIR, exist_ok=True)
    file_path = os.path.join(DRIVE_SAVE_DIR, file_name)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
      df.to_excel(writer, index=False, sheet_name='Bao_Cao_Tong_Hop')

    flash(f'✅ Đã lưu file báo cáo về Google Drive: {file_name}', 'success')
    return send_file(file_path, as_attachment=True, download_name=file_name)
  except Exception:
    temp_path = os.path.join(os.getcwd(), file_name)
    with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
      df.to_excel(writer, index=False, sheet_name='Bao_Cao_Tong_Hop')
    return send_file(temp_path, as_attachment=True, download_name=file_name)


if __name__ == '__main__':
  app.run(host='0.0.0.0', port=5000, debug=True)