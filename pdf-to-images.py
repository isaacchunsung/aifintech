import argparse
import os
import fitz  # PyMuPDF
import pikepdf
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
from PIL import Image
import io
import threading

def unlock_pdf(input_file, password, output_file=None):
    """비밀번호로 보호된 PDF 파일의 잠금을 해제합니다."""
    try:
        # 만약 출력 파일이 지정되지 않았다면, 임시 파일 생성
        if output_file is None:
            output_file = f"unlocked_{os.path.basename(input_file)}"
        
        # PDF 파일 열기 및 잠금 해제
        with pikepdf.open(input_file, password=password) as pdf:
            # 잠금 해제된 PDF 저장
            pdf.save(output_file)
        
        print(f"PDF 잠금 해제 성공: {output_file}")
        return output_file
    except Exception as e:
        print(f"PDF 잠금 해제 실패: {e}")
        return None

def convert_pdf_to_images(pdf_path, output_dir=None, password=None, dpi=300, image_format="png", start_page=1, end_page=None):
    """PDF 파일을 페이지별 이미지로 변환합니다."""
    try:
        unlocked_pdf = pdf_path
        temp_file_created = False
        
        # 비밀번호가 제공된 경우 PDF 잠금 해제
        if password:
            try:
                temp_unlocked = unlock_pdf(pdf_path, password)
                if temp_unlocked:
                    unlocked_pdf = temp_unlocked
                    temp_file_created = True
            except Exception as e:
                print(f"비밀번호 처리 중 오류 발생: {e}")
                # 비밀번호 오류가 있어도 계속 진행
                pass
        
        # 출력 디렉토리가 지정되지 않은 경우 기본값 설정
        if output_dir is None:
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = f"{base_name}_images"
        
        # 출력 디렉토리 생성
        os.makedirs(output_dir, exist_ok=True)
        
        # PDF 문서 열기
        try:
            doc = fitz.open(unlocked_pdf)
        except Exception as e:
            if "password" in str(e).lower():
                try:
                    # 빈 비밀번호로 시도
                    doc = fitz.open(unlocked_pdf, password="")
                except:
                    print(f"PDF 열기 실패: 올바른 비밀번호가 필요합니다.")
                    return False, 0
            else:
                print(f"PDF 열기 실패: {e}")
                return False, 0
        
        # 페이지 범위 확인
        if start_page < 1:
            start_page = 1
        
        page_count = len(doc)
        
        if end_page is None or end_page > page_count:
            end_page = page_count
        
        # 변환할 페이지 수
        total_pages = end_page - start_page + 1
        
        # 이미지로 변환
        success_count = 0
        for page_num in range(start_page-1, end_page):
            try:
                page = doc.load_page(page_num)
                
                # 픽셀맵으로 렌더링 (높은 해상도)
                pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                
                # 이미지 저장
                image_path = os.path.join(output_dir, f"page_{page_num+1}.{image_format.lower()}")
                
                # PNG 형식으로 저장
                if image_format.lower() == "png":
                    pix.save(image_path)
                # 다른 형식은 PIL을 사용하여 변환
                else:
                    img = Image.open(io.BytesIO(pix.tobytes()))
                    img.save(image_path, format=image_format.upper())
                
                success_count += 1
                
                # 진행 상황 출력
                print(f"페이지 {page_num+1}/{end_page} 변환 완료 ({success_count}/{total_pages})")
                
                # GUI에서 진행률 업데이트를 위한 콜백 함수
                if 'progress_callback' in locals() and progress_callback:
                    progress_callback(page_num + 1 - (start_page-1), total_pages)
                
            except Exception as e:
                print(f"페이지 {page_num+1} 변환 실패: {e}")
        
        # 임시 파일 삭제
        if temp_file_created and os.path.exists(unlocked_pdf):
            os.remove(unlocked_pdf)
        
        print(f"변환 완료: {success_count}개 페이지가 {output_dir}에 저장되었습니다.")
        return True, success_count
    
    except Exception as e:
        print(f"변환 중 오류 발생: {e}")
        return False, 0

class PDFToImageGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF를 이미지로 변환")
        self.root.geometry("600x500")
        
        # 변수 초기화
        self.input_file_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self.dpi_var = tk.IntVar(value=300)
        self.image_format_var = tk.StringVar(value="png")
        self.start_page_var = tk.IntVar(value=1)
        self.end_page_var = tk.StringVar(value="마지막")
        self.status_var = tk.StringVar(value="준비됨")
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 입력 파일 선택
        ttk.Label(main_frame, text="PDF 파일:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="찾아보기", command=self.select_input_file).grid(row=0, column=2, pady=5)
        
        # 출력 디렉토리 선택
        ttk.Label(main_frame, text="출력 폴더:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="찾아보기", command=self.select_output_dir).grid(row=1, column=2, pady=5)
        
        # 비밀번호 입력
        ttk.Label(main_frame, text="비밀번호:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.password_var, show="*", width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Label(main_frame, text="(필요시)").grid(row=2, column=2, pady=5)
        
        # 옵션 프레임
        options_frame = ttk.LabelFrame(main_frame, text="변환 옵션", padding=10)
        options_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        # DPI 설정
        ttk.Label(options_frame, text="해상도(DPI):").grid(row=0, column=0, sticky=tk.W, pady=5)
        dpi_combo = ttk.Combobox(options_frame, textvariable=self.dpi_var, width=10)
        dpi_combo['values'] = (72, 150, 300, 600)
        dpi_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(options_frame, text="높을수록 이미지 품질이 좋지만 파일 크기가 커집니다.").grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # 이미지 형식 설정
        ttk.Label(options_frame, text="이미지 형식:").grid(row=1, column=0, sticky=tk.W, pady=5)
        format_combo = ttk.Combobox(options_frame, textvariable=self.image_format_var, width=10)
        format_combo['values'] = ('png', 'jpg', 'tiff', 'bmp')
        format_combo.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(options_frame, text="PNG는 무손실, JPG는 손실 압축 형식입니다.").grid(row=1, column=2, sticky=tk.W, pady=5)
        
        # 페이지 범위 설정
        ttk.Label(options_frame, text="시작 페이지:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(options_frame, textvariable=self.start_page_var, width=10).grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        
        ttk.Label(options_frame, text="끝 페이지:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(options_frame, textvariable=self.end_page_var, width=10).grid(row=3, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(options_frame, text="'마지막'을 입력하면 PDF의 마지막 페이지까지 변환합니다.").grid(row=3, column=2, sticky=tk.W, pady=5)
        
        # 진행 상황 표시
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        # 변환 버튼
        self.convert_button = ttk.Button(main_frame, text="변환 시작", command=self.start_conversion)
        self.convert_button.grid(row=5, column=0, columnspan=3, pady=10)
        
        # 상태 표시줄
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def select_input_file(self):
        file = filedialog.askopenfilename(
            title="변환할 PDF 파일을 선택하세요",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
        )
        if file:
            self.input_file_var.set(file)
            # 기본 출력 디렉토리 이름 설정
            base_name = os.path.splitext(os.path.basename(file))[0]
            self.output_dir_var.set(os.path.join(os.path.dirname(file), f"{base_name}_images"))
    
    def select_output_dir(self):
        directory = filedialog.askdirectory(title="이미지를 저장할 폴더를 선택하세요")
        if directory:
            self.output_dir_var.set(directory)
    
    def update_progress(self, current, total):
        """진행 상황 업데이트"""
        progress_value = int((current / total) * 100)
        self.progress['value'] = progress_value
        self.status_var.set(f"변환 중... {current}/{total} 페이지 ({progress_value}%)")
        self.root.update_idletasks()
    
    def start_conversion(self):
        input_file = self.input_file_var.get()
        output_dir = self.output_dir_var.get()
        password = self.password_var.get() if self.password_var.get() else None
        dpi = self.dpi_var.get()
        image_format = self.image_format_var.get()
        start_page = self.start_page_var.get()
        
        # 끝 페이지 처리
        end_page_str = self.end_page_var.get()
        if end_page_str == "마지막" or not end_page_str:
            end_page = None
        else:
            try:
                end_page = int(end_page_str)
            except ValueError:
                messagebox.showerror("오류", "끝 페이지는 숫자 또는 '마지막'이어야 합니다.")
                return
        
        # 입력 값 검증
        if not input_file:
            messagebox.showerror("오류", "변환할 PDF 파일을 선택하세요.")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다: {input_file}")
            return
        
        if not output_dir:
            messagebox.showerror("오류", "출력 폴더를 지정하세요.")
            return
        
        # PDF가 잠겨있는지 확인
        try:
            fitz.open(input_file)
        except Exception as e:
            if "password" in str(e).lower():
                if not password:
                    # 비밀번호가 필요하지만 제공되지 않았을 때 물어보기
                    pwd = simpledialog.askstring("비밀번호 필요", 
                                               "이 PDF는 비밀번호로 보호되어 있습니다.\n비밀번호를 입력하세요:", 
                                               show='*')
                    if pwd:
                        password = pwd
                    else:
                        # 비밀번호 입력을 취소하면 계속 진행 (실패할 수 있음)
                        if not messagebox.askyesno("경고", 
                                                "비밀번호 없이 계속 진행하시겠습니까?\n변환이 실패할 수 있습니다."):
                            return
        
        # 기존 폴더가 있는지 확인하고 경고
        if os.path.exists(output_dir) and os.listdir(output_dir):
            if not messagebox.askyesno("경고", 
                                    f"출력 폴더 '{output_dir}'에 이미 파일이 있습니다.\n기존 파일을 덮어쓸 수 있습니다. 계속하시겠습니까?"):
                return
        
        # 버튼 비활성화 및 상태 업데이트
        self.convert_button['state'] = 'disabled'
        self.status_var.set("변환 준비 중...")
        self.progress['value'] = 0
        
        # 변환 함수에 진행률 업데이트 콜백 전달
        def conversion_thread():
            # 변환 진행
            success, count = convert_pdf_to_images(
                input_file, output_dir, password, dpi, image_format, 
                start_page, end_page
            )
            
            # UI 업데이트는 메인 스레드에서 수행
            self.root.after(0, lambda: self.conversion_completed(success, count, output_dir))
        
        # 변환 작업을 별도 스레드로 실행
        self.thread = threading.Thread(target=conversion_thread)
        self.thread.daemon = True
        self.thread.start()
        
        # 주기적으로 진행 상황을 폴링
        self.poll_progress(input_file, output_dir, start_page, end_page)
    
    def poll_progress(self, input_file, output_dir, start_page, end_page=None):
        """주기적으로 출력 폴더를 확인하여 진행 상황 업데이트"""
        if not self.thread.is_alive():
            return
        
        try:
            # 지정된 범위에서 총 페이지 수 계산
            if not hasattr(self, 'total_pages'):
                doc = fitz.open(input_file)
                total = len(doc)
                if end_page is None or end_page > total:
                    end_page = total
                self.total_pages = end_page - start_page + 1
            
            # 현재까지 생성된 이미지 파일 수 확인
            if os.path.exists(output_dir):
                files = [f for f in os.listdir(output_dir) if f.startswith("page_") and f.endswith(f".{self.image_format_var.get()}")]
                current = len(files)
                
                # 진행률 업데이트
                self.update_progress(current, self.total_pages)
        except:
            pass
        
        # 100ms 후 다시 확인
        self.root.after(100, lambda: self.poll_progress(input_file, output_dir, start_page, end_page))
    
    def conversion_completed(self, success, count, output_dir):
        """변환 완료 후 처리"""
        self.convert_button['state'] = 'normal'
        
        if success:
            self.status_var.set(f"변환 완료: {count}개 페이지가 변환되었습니다.")
            self.progress['value'] = 100
            
            if messagebox.askyesno("성공", 
                                f"PDF의 {count}개 페이지를 이미지로 변환했습니다.\n저장 위치: {output_dir}\n\n이미지 폴더를 열어보시겠습니까?"):
                # 운영체제별 폴더 열기
                import platform
                import subprocess
                
                if platform.system() == 'Windows':
                    os.startfile(output_dir)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', output_dir])
                else:  # Linux
                    subprocess.run(['xdg-open', output_dir])
        else:
            self.status_var.set("변환 실패")
            messagebox.showerror("실패", "PDF 변환에 실패했습니다.")

def main():
    parser = argparse.ArgumentParser(description='PDF 파일을 페이지별 이미지로 변환합니다.')
    parser.add_argument('input_file', nargs='?', help='변환할 PDF 파일 경로 (생략 시 GUI 모드로 실행)')
    parser.add_argument('-o', '--output', help='이미지를 저장할 디렉토리 (기본값: PDF 파일명_images)')
    parser.add_argument('-p', '--password', help='PDF 파일의 비밀번호')
    parser.add_argument('-d', '--dpi', type=int, default=300, help='이미지 해상도 (기본값: 300)')
    parser.add_argument('-f', '--format', default='png', choices=['png', 'jpg', 'tiff', 'bmp'], 
                      help='이미지 형식 (기본값: png)')
    parser.add_argument('-s', '--start', type=int, default=1, help='시작 페이지 (기본값: 1)')
    parser.add_argument('-e', '--end', type=int, help='끝 페이지 (기본값: 마지막 페이지)')
    parser.add_argument('-g', '--gui', action='store_true', help='GUI 모드로 실행')
    
    args = parser.parse_args()
    
    # GUI 모드로 실행
    if args.gui or args.input_file is None:
        root = tk.Tk()
        app = PDFToImageGUI(root)
        root.mainloop()
        return
    
    # 명령줄 모드로 실행
    if not os.path.exists(args.input_file):
        print(f"오류: 파일을 찾을 수 없습니다: {args.input_file}")
        return
    
    # PDF를 이미지로 변환
    success, count = convert_pdf_to_images(
        args.input_file, args.output, args.password, 
        args.dpi, args.format, args.start, args.end
    )
    
    if success:
        print(f"변환이 성공적으로 완료되었습니다. {count}개 페이지가 변환되었습니다.")
    else:
        print("변환에 실패했습니다.")

if __name__ == "__main__":
    main()
