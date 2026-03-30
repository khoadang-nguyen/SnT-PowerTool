import sys
import os
from pathlib import Path

# --- List lib for PYINSTALLER ---
import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import tqdm
import pyodbc
import sqlalchemy
# --------------------------------------------------------

# Import tools:
try:
    import _01_auto_concat
    import _02_profiling_tool
    import _03_target_vs_profiling
    import _04_HCP_vs_everything
except ImportError as e:
    print(f"❌ Lỗi Import: {e}")

# Mapping: Key -> (Hàm xử lý, Tên hiển thị)
TOOLS = {
    "2": (_01_auto_concat.run, "01_auto_concat"),
    "3": (_02_profiling_tool.run, "02_auto_profiling"),
    "4": (_03_target_vs_profiling.run, "03_auto_export"),
    "5": (_04_HCP_vs_everything.run, "04_HCP_vs_everything")
}

def run_script(tool_tuple):
    func = tool_tuple[0]
    name = tool_tuple[1]
    print(f"\n🚀 Running {name}...\n")
    try:
        func() 
        print(f"\n✅ Script {name} chạy xong.\n")
    except Exception as e:
        print(f"\n❌ Script {name} gặp lỗi: {e}")
        print("Trở về menu chính để bạn chọn lại.\n")

def run_all():
    for key in ["2", "3", "4", "5"]:
        run_script(TOOLS[key])

def menu():
    print("\n" + "="*20 + " TOOL MENU " + "="*20)
    print("1. Run tất cả")
    print("2. Concat các file LEO_HA8 và OB")
    print("3. Pivot file Profiling - Survey")
    print("4. Kiểm tra tiến độ Profiling vs Target List")
    print("5. Kiểm tra HCP List so với Target List và Survey")
    print("Gõ 'no' để thoát")

# --- Check your tỉnh táo ---
title = "Power tool for S&T Project V1.0"
width = 40
accepted_yes = ['y', 'yes', 'co', 'có']

while True:
    print("\n" + "+" + "-"*(width-2) + "+")
    print("|" + title.center(width-2) + "|")
    print("+" + "-"*(width-2) + "+")
    
    print("Trước khi tool bắt đầu chạy, mình trả lời câu hỏi chút nhé.")

    answer = input("\nBạn có đang tỉnh táo khi sử dụng tool không? (yes/no) ").strip().lower()
    if answer not in accepted_yes:
        print("Mình đi relax, rửa mặt gì đó xong làm nhé 😏")
        continue

    answer_2 = input("\nOk, vậy là tinh thần minh mẫn.\nBạn có đảm bảo data và folder yêu cầu theo file ReadMe đã chuẩn bị và format chuẩn chỉnh không? (yes/no) ").strip().lower()
    if answer_2 not in accepted_yes:
        print('Vui lòng kiểm tra đầy đủ trước khi chạy để tránh lỗi.')
        continue
    print("Good answer 😎 Chạy tool luôn!")
    break

# --- Menu loop ---
while True:
    menu()
    choice = input("👉 Chọn chức năng: ").strip()

    if choice.lower() == "no":
        print("👋 Bye!")
        break
    elif choice == "1":
        run_all()
    elif choice in TOOLS:
        run_script(TOOLS[choice])
    else:
        print("❌ Lựa chọn không hợp lệ!")

    again = input("\n👉 Tiếp tục? (yes/no): ").strip().lower()
    if again == "no":
        break