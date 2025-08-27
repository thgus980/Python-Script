"""
입사지원서 자동 작성 도구 - tkinter 기반 GUI
"""

import os
import sys
import threading
import time
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

# 핵심 로직이 작성된 스크립트 임포트
from excel_template_filler import ExcelTemplateFiller

class ProgressGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("입사지원서 자동 작성 도구")
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        
        # 윈도우를 화면 중앙에 배치
        self.center_window()
        
        # GUI 구성 요소
        self.setup_ui()
        
        # 진행 상태 변수
        self.current_step = 0
        self.total_steps = 0
        self.is_running = False
        
    def center_window(self):
        """윈도우를 화면 중앙에 배치"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_ui(self):
        """GUI 구성 요소 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="입사지원서 자동 작성 도구",
            font=("Arial", 14, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # 상태 메시지
        self.status_label = ttk.Label(
            main_frame,
            text="시작 준비 완료",
            font=("Arial", 10)
        )
        self.status_label.grid(row=1, column=0, columnspan=2, pady=(0, 10))
        
        # 진행률 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100,
            length=350,
            mode='determinate'
        )
        self.progress_bar.grid(row=2, column=0, columnspan=2, pady=(0, 8), sticky=(tk.W, tk.E))
        
        # 진행률 텍스트
        self.progress_text = ttk.Label(
            main_frame,
            text="0%",
            font=("Arial", 9)
        )
        self.progress_text.grid(row=3, column=0, columnspan=2, pady=(0, 15))
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(5, 40))
        
        # 시작 버튼
        self.start_button = ttk.Button(
            button_frame,
            text="지원서 생성 시작",
            command=self.start_processing,
            width=18
        )
        self.start_button.grid(row=0, column=0, padx=(0, 10))
        
        # 취소 버튼
        self.cancel_button = ttk.Button(
            button_frame,
            text="취소",
            command=self.cancel_processing,
            width=12,
            state="disabled"
        )
        self.cancel_button.grid(row=0, column=1, padx=(10, 0))
        
        # 그리드 가중치 설정
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def update_status(self, message):
        """상태 메시지 업데이트"""
        self.status_label.config(text=message)
        print(f"{message}")  # 콘솔에만 로그 출력
        
    def update_progress(self, current, total, message=""):
        """진행률 업데이트"""
        if total > 0:
            progress = (current / total) * 100
            self.progress_var.set(progress)
            self.progress_text.config(text=f"{progress:.1f}% ({current}/{total})")
            
            if message:
                self.update_status(message)
        
        self.root.update_idletasks()
        
    def start_processing(self):
        """처리 시작"""
        self.is_running = True
        self.start_button.config(state="disabled")
        self.cancel_button.config(state="normal")
        
        # 별도 스레드에서 처리 실행
        self.processing_thread = threading.Thread(target=self.run_processing)
        self.processing_thread.daemon = True
        self.processing_thread.start()
        
    def cancel_processing(self):
        """처리 취소"""
        self.is_running = False
        self.update_status("처리가 취소되었습니다")
        self.reset_ui()
        
    def reset_ui(self):
        """UI 초기 상태로 리셋"""
        self.start_button.config(state="normal")
        self.cancel_button.config(state="disabled")
        self.progress_var.set(0)
        self.progress_text.config(text="0%")
        
    def run_processing(self):
        """실제 처리 실행 (별도 스레드)"""
        try:
            # 1단계: 초기화 및 파일 확인
            self.update_status("파일 확인 중...")
            self.update_progress(1, 10, "설정 파일 로드 중...")
            
            filler = ExcelTemplateFiller()
            
            # 파일 경로 확인
            template_path = filler.config["template_file"]
            raw_data_path = filler.config["raw_data_file"]
            
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")
                
            if not os.path.exists(raw_data_path):
                raise FileNotFoundError(f"원천 데이터 파일이 없습니다: {raw_data_path}")
            
            self.update_progress(2, 10, "파일 확인 완료")
            
            # 2단계: 데이터 로드
            self.update_status("데이터 로드 중...")
            
            import pandas as pd
            df = pd.read_excel(
                raw_data_path, 
                sheet_name=filler.config["raw_data_sheet"],
                dtype=str
            ).fillna("")
            
            total_applicants = len(df)
            print(f"데이터 로드 완료: {total_applicants}명의 지원자")
            print(f"컬럼 목록: {list(df.columns)}")
            
            self.update_progress(3, 10, f"{total_applicants}명 지원자 데이터 로드 완료")
            
            if not self.is_running:
                return
                
            # 3단계: 출력 디렉토리 생성
            output_dir = filler.config["output_dir"]
            os.makedirs(output_dir, exist_ok=True)
            self.update_progress(4, 10, "출력 폴더 준비 완료")
            
            # 4단계: 각 지원자별 처리
            success_count = 0
            
            for index, row in df.iterrows():
                if not self.is_running:
                    break
                    
                try:
                    # 진행률 계산 (4~9단계를 지원자 처리에 할당)
                    progress_step = 4 + ((index + 1) / total_applicants) * 5
                    
                    context = row.to_dict()
                    applicant_name = context.get('이름', f'지원자{index+1}')
                    
                    self.update_progress(
                        progress_step, 10, 
                        f"{applicant_name} 지원서 생성 중... ({index+1}/{total_applicants})"
                    )
                    
                    # 파일명 생성
                    try:
                        filename = filler.config["filename_pattern"].format(**context)
                    except KeyError as e:
                        print(f"파일명 패턴 오류 (행 {index+1}): {e}")
                        filename = f"application_{index+1}.xlsx"
                    
                    output_path = os.path.join(output_dir, filename)
                    
                    # 지원서 생성
                    filler.fill_workbook(template_path, context, output_path)
                    success_count += 1
                    
                    print(f"{applicant_name} 지원서 완료: {filename}")
                    
                except Exception as e:
                    print(f"{applicant_name} 처리 실패: {e}")
                    continue
            
            if not self.is_running:
                return
                
            # 5단계: 완료
            self.update_progress(10, 10, "모든 처리 완료!")
            
            # 완료 팝업 표시
            self.show_completion_dialog(success_count, total_applicants, output_dir)
            
        except Exception as e:
            print(f"처리 중 오류 발생: {e}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n\n{e}")
            
        finally:
            self.reset_ui()
            
    def show_completion_dialog(self, success_count, total_count, output_dir):
        """완료 대화상자 표시"""
        # 완료 메시지 (경로 제외)
        if success_count == total_count:
            title = "완료!"
            status_message = f"모든 지원서 생성이 완료되었습니다!\n\n성공: {success_count}개"
        else:
            title = "부분 완료"
            status_message = f"지원서 생성이 완료되었습니다.\n\n성공: {success_count}개\n실패: {total_count - success_count}개"
        
        # 저장 위치 (별도 처리)
        full_path = os.path.abspath(output_dir)
        
        # 커스텀 대화상자
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.resizable(True, False)  # 가로만 크기 조정 가능
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 메인 프레임
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 상태 메시지
        status_label = ttk.Label(main_frame, text=status_message, justify=tk.CENTER)
        status_label.pack(pady=(0, 10))
        
        # 저장 위치 프레임
        path_frame = ttk.LabelFrame(main_frame, text="저장 위치", padding="10")
        path_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 저장 위치 텍스트 (자동 줄바꿈)
        path_label = ttk.Label(path_frame, text=full_path, wraplength=450, justify=tk.LEFT)
        path_label.pack()
        
        # 버튼 프레임
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack()
        
        # 폴더 열기 버튼
        def open_folder():
            try:
                os.startfile(os.path.abspath(output_dir))
            except:
                subprocess.run(['explorer', os.path.abspath(output_dir)])
            dialog.destroy()
        
        ttk.Button(btn_frame, text="폴더 열기", command=open_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="확인", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # 창 크기 자동 조정
        dialog.update_idletasks()
        
        # 최소 크기 설정
        min_width = 500
        min_height = 200
        
        # 실제 필요한 크기 계산
        req_width = max(min_width, dialog.winfo_reqwidth() + 40)
        req_height = max(min_height, dialog.winfo_reqheight() + 20)
        
        dialog.geometry(f"{req_width}x{req_height}")
        dialog.minsize(min_width, min_height)
        
        # 대화상자를 부모 창 중앙에 배치
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (req_width // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (req_height // 2)
        dialog.geometry(f"{req_width}x{req_height}+{x}+{y}")
        
        # 엔터키로 확인
        dialog.bind('<Return>', lambda e: dialog.destroy())
        dialog.focus()
        
    def run(self):
        """GUI 실행"""
        # 초기 상태 확인
        try:
            filler = ExcelTemplateFiller()
            template_path = filler.config["template_file"]
            raw_data_path = filler.config["raw_data_file"]
            
            if not os.path.exists(template_path):
                print(f"템플릿 파일이 없습니다: {template_path}")
                self.status_label.config(text="템플릿 파일을 확인해주세요")
                self.start_button.config(state="disabled")
            elif not os.path.exists(raw_data_path):
                print(f"원본 데이터 파일이 없습니다: {raw_data_path}")
                self.status_label.config(text="원본 데이터 파일을 확인해주세요")
                self.start_button.config(state="disabled")
            else:
                print("모든 파일이 준비되었습니다.")
                self.status_label.config(text="시작 준비 완료")
                
        except Exception as e:
            print(f"초기화 오류: {e}")
            self.status_label.config(text="설정 파일을 확인해주세요")
            self.start_button.config(state="disabled")
        
        # GUI 시작
        self.root.mainloop()


def main():
    """메인 실행 함수"""
    # GUI 모드로 실행
    app = ProgressGUI()
    app.run()


if __name__ == "__main__":
    main()
