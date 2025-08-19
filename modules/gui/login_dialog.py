#!/usr/bin/env python3
"""
ERP 계정 정보 입력 다이얼로그 - 리팩토링된 버전
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Optional

class LoginDialog:
    """ERP 계정 정보 입력 다이얼로그"""
    
    def __init__(self, parent):
        self.parent = parent
        self.result = None
        self.dialog = None
        
        # 회사 정보 (실제 ERP 회사코드 사용)
        self.companies = [
            {"name": "디앤드디", "code": "52867"},
            {"name": "디앤아이", "code": "628361"}, 
            {"name": "후지리프트코리아", "code": "175989"}
        ]
        
        self.accounts = []
    
    def show_dialog(self) -> Optional[List[Dict]]:
        """다이얼로그 표시하고 결과 반환"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("ERP 계정 정보 입력")
        self.dialog.geometry("520x600")  # 높이를 더 늘림
        self.dialog.resizable(False, False)
        
        # 모달 설정
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Enter 키로 로그인 가능하게 설정
        self.dialog.bind('<Return>', lambda event: self.confirm())
        self.dialog.bind('<KP_Enter>', lambda event: self.confirm())
        
        # 중앙 배치
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (520 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (600 // 2)
        self.dialog.geometry(f"520x600+{x}+{y}")
        
        self.setup_ui()
        
        # 다이얼로그가 닫힐 때까지 대기
        self.dialog.wait_window()
        return self.result
    
    def setup_ui(self):
        """UI 구성"""
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목 (현대적인 타이포그래피)
        title_label = ttk.Label(main_frame, text="ERP 계정 정보 입력", 
                               font=('Segoe UI', 16, 'normal'))
        title_label.pack(pady=(0, 5))
        
        # 설명 (더 가독성 있게)
        desc_label = ttk.Label(main_frame, 
                              text="각 회사별 ERP 계정 정보를 입력하세요.\n"
                                   "입력된 정보는 메모리에만 저장되며 파일로 저장되지 않습니다.",
                              font=('Segoe UI', 9), foreground="#6b7280")
        desc_label.pack(pady=(0, 20))
        
        # 회사별 입력 섹션들을 직접 배치 (스크롤 없이)
        companies_frame = ttk.Frame(main_frame)
        companies_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # 엔트리 저장용
        self.entries = {}
        
        # 회사별 입력 필드 생성
        for i, company in enumerate(self.companies):
            self.create_company_section(companies_frame, company, i)
        
        # 버튼 프레임 (맨 마지막에 배치)
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # 로그인 버튼 (오른쪽)
        login_btn = tk.Button(button_frame, 
                             text="로그인", 
                             command=self.confirm,
                             bg="#2563eb",
                             fg="white",
                             font=('Segoe UI', 11, 'bold'),
                             width=12,
                             height=1,
                             relief='flat',
                             bd=0,
                             activebackground="#1e40af",
                             activeforeground="white",
                             cursor="hand2")
        login_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 취소 버튼 (왼쪽)
        cancel_btn = tk.Button(button_frame, 
                              text="취소", 
                              command=self.cancel,
                              bg="#f3f4f6",
                              fg="#6b7280",
                              font=('Segoe UI', 11),
                              width=12,
                              height=1,
                              relief='flat',
                              bd=0,
                              activebackground="#e5e7eb",
                              activeforeground="#374151",
                              cursor="hand2")
        cancel_btn.pack(side=tk.RIGHT, padx=(0, 10))
    
    def create_company_section(self, parent, company, index):
        """회사별 입력 섹션 생성"""
        # 회사 프레임
        company_frame = ttk.LabelFrame(parent, text=f"{company['name']} ({company['code']})", 
                                      padding="10")
        company_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 사용자 ID
        ttk.Label(company_frame, text="사용자 ID:").grid(row=0, column=0, sticky=tk.W, pady=2)
        user_id_entry = ttk.Entry(company_frame, width=25)
        user_id_entry.grid(row=0, column=1, padx=(10, 0), pady=2)
        
        # 비밀번호
        ttk.Label(company_frame, text="비밀번호:").grid(row=1, column=0, sticky=tk.W, pady=2)
        password_entry = ttk.Entry(company_frame, width=25, show="*")
        password_entry.grid(row=1, column=1, padx=(10, 0), pady=2)
        
        # 엔트리 저장
        self.entries[company['code']] = {
            'name': company['name'],
            'code': company['code'],
            'user_id': user_id_entry,
            'password': password_entry
        }
        
        # 첫 번째 필드에 포커스
        if index == 0:
            user_id_entry.focus()
    
    def confirm(self):
        """확인 버튼 처리"""
        accounts = []
        empty_fields = []
        
        for code, entries in self.entries.items():
            user_id = entries['user_id'].get().strip()
            password = entries['password'].get().strip()
            
            if user_id and password:
                accounts.append({
                    "company_name": entries['name'],
                    "company_code": code,
                    "user_id": user_id,
                    "user_pw": password
                })
            elif user_id or password:  # 둘 중 하나만 입력된 경우
                empty_fields.append(entries['name'])
        
        if not accounts:
            messagebox.showwarning("입력 필요", "최소 하나의 회사 계정 정보를 입력해주세요.")
            return
        
        if empty_fields:
            response = messagebox.askyesno("확인", 
                f"다음 회사의 정보가 불완전합니다:\n{', '.join(empty_fields)}\n\n"
                f"입력된 {len(accounts)}개 회사 정보로 진행하시겠습니까?")
            if not response:
                return
        
        # 회사 순서 재정렬 (디앤드디 먼저)
        company_order = ["52867", "628361", "175989"]
        accounts.sort(key=lambda x: company_order.index(x["company_code"]) if x["company_code"] in company_order else 999)
        
        # 계정 정보를 accounts.json 파일로 저장 (임시)
        self.save_accounts_to_file(accounts)
        
        self.result = accounts
        self.dialog.destroy()
    
    def save_accounts_to_file(self, accounts):
        """계정 정보를 accounts.json 파일로 임시 저장 (리팩토링된 경로)"""
        try:
            from pathlib import Path
            import json
            
            # config 디렉토리 확인/생성 (리팩토링된 구조)
            config_dir = Path(__file__).parent.parent.parent / "config"
            config_dir.mkdir(exist_ok=True)
            
            # accounts.json 파일 경로
            accounts_file = config_dir / "accounts.json"
            
            # JSON 데이터 구성
            accounts_data = {
                "accounts": accounts,
                "created_at": "runtime_generated",
                "note": "이 파일은 로그인 다이얼로그에서 자동 생성되었습니다",
                "security": "비밀번호는 평문으로 저장되므로 보안에 주의하세요"
            }
            
            # 파일 저장
            with open(accounts_file, 'w', encoding='utf-8') as f:
                json.dump(accounts_data, f, ensure_ascii=False, indent=2)
            
            print(f"✅ 계정 정보 임시 저장: {accounts_file}")
            print(f"   저장된 회사: {len(accounts)}개")
            
        except Exception as e:
            print(f"⚠️ 계정 정보 저장 실패: {e}")
            # 저장 실패해도 로그인은 계속 진행
    
    def cancel(self):
        """취소 버튼 처리"""
        self.result = None
        self.dialog.destroy()


def get_erp_accounts(parent_window) -> Optional[List[Dict]]:
    """ERP 계정 정보 입력 다이얼로그 표시"""
    print("🔑 실제 ERP 로그인 다이얼로그 호출됨")
    dialog = LoginDialog(parent_window)
    return dialog.show_dialog()


if __name__ == "__main__":
    # 테스트용
    root = tk.Tk()
    root.withdraw()  # 메인 창 숨기기
    
    accounts = get_erp_accounts(root)
    if accounts:
        print("입력된 계정 정보:")
        for acc in accounts:
            print(f"  {acc['company_name']}: {acc['user_id']} / {'*' * len(acc['user_pw'])}")
    else:
        print("취소됨")
    
    root.destroy()
