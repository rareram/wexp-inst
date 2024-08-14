import os
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
import winreg
from tkinter import filedialog, messagebox, ttk, PhotoImage

import win32com.shell.shell as shell

# 관리자 권한 확인 및 실행
def is_admin():
    try:
        return shell.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        script = os.path.abspath(sys.argv[0])
        params = ' '.join([script] + sys.argv[1:])
        shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
        sys.exit(0)

class ServiceManagerApp:
    def __init__(self, master):
        self.master = master
        master.title('윈도우 익스포터 설치 도우미')
        master.geometry('500x500')

        self.version = '0.1.0'
        self.file_path = tk.StringVar()
        self.service_name = tk.StringVar(value='Prometheus Windows Exporter')
        self.service_description = tk.StringVar(value='Exports Windows metrics for Prometheus')

        self.load_images()
        self.create_title()
        self.create_widgets()

    def load_images(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # Github 아이콘 로드
        github_icon_path = os.path.join(current_dir, 'github_icon.png')
        self.github_icon = self.load_image(github_icon_path, 'Github icon')

        # 로고 이미지 로드
        logo_path = os.path.join(current_dir, 'logo.png')
        self.logo_image = self.load_image(logo_path, 'Logo')

    def load_image(self, path, description):
        try:
            return PhotoImage(file=path)
        except tk.TclError as e:
            print(f'Warning: Could not load {description} from {path}. Error: {e}')
            return None

    def create_title(self):
        title_frame = tk.Frame(self.master)
        title_frame.pack(fill='x', padx=20, pady=20)

        if self.logo_image:
            logo_label = tk.Label(title_frame, image=self.logo_image)
            logo_label.pack(side=tk.RIGHT, anchor='ne')

        title_font = tkfont.Font(family='Malgun Gothic, Helvetica', size=12, weight='bold')
        title_label = tk.Label(title_frame, text='윈도우 서버 모니터링 서비스 등록', font=title_font)
        title_label.pack(side=tk.TOP, anchor='nw')

        verdesc_font = tkfont.Font(family='Malgun Gothic, Helvetica', size=8)
        verdesc_label = tk.Label(title_frame, text='2024년  |  IT서비스품질관리팀  | ', font=verdesc_font)
        verdesc_label.pack(side=tk.LEFT)
        version_font = tkfont.Font(family='Helvetica', size=9, weight='bold', slant='italic')
        version_label = tk.Label(title_frame, text=f'v{self.version}', font=version_font, fg='#FF6347')
        version_label.pack(side=tk.LEFT)

    def create_widgets(self):
        notebook = ttk.Notebook(self.master)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)

        install_frame = ttk.Frame(notebook)
        uninstall_frame = ttk.Frame(notebook)

        notebook.add(install_frame, text='서비스 등록')
        notebook.add(uninstall_frame, text='서비스 제거')

        self.create_install_widgets(install_frame)
        self.create_uninstall_widgets(uninstall_frame)

    def create_install_widgets(self, parent):
        self.create_download_frame(parent)
        self.create_select_frame(parent)
        self.create_file_label(parent)
        self.create_service_frame(parent)
        tk.Button(parent, text='서비스 설치', command=self.install_service).pack(pady=20)

    def create_download_frame(self, parent):
        download_frame = tk.Frame(parent)
        download_frame.pack(fill='x', padx=10, pady=(10, 2), anchor='w')

        tk.Label(download_frame, text='① GitHub 링크 (windows_exporter): ', anchor='w').pack(side='left')
        if self.github_icon:
            github_icon_label = tk.Label(download_frame, image=self.github_icon)
            github_icon_label.pack(side='left', padx=(0, 5))
        github_link = tk.Label(download_frame, text='Windows Exporter Github Releases', fg='blue', cursor='hand2')
        github_link.pack(side='left')
        github_link.bind('<Button-1>', lambda e: self.open_github_link())

    def create_select_frame(self, parent):
        select_frame = tk.Frame(parent)
        select_frame.pack(fill='x', padx=10, pady=(2, 5), anchor='w')

        tk.Label(select_frame, text='② Windows Exporter 선택:', anchor='w').pack(side='left')
        tk.Button(select_frame, text='파일 선택', command=self.select_file).pack(side='left', padx=(5, 0))

    def create_file_label(self, parent):
        self.file_label = tk.Label(parent, text='아직 파일이 선택되지 않았습니다.', anchor='w')
        self.file_label.pack(fill='x', padx=50, pady=(0, 10), anchor='w')

    def create_service_frame(self, parent):
        service_frame = tk.Frame(parent)
        service_frame.pack(fill='x', padx=10, pady=5, anchor='w')

        tk.Label(service_frame, text='③ 서비스 정보 입력:', anchor='w').pack(anchor='w')

        name_frame = tk.Frame(service_frame)
        name_frame.pack(fill='x', pady=(5, 2))
        tk.Label(name_frame, text='서비스 이름:', width=15, anchor='w').pack(side='left')
        tk.Entry(name_frame, textvariable=self.service_name, width=50).pack(side='left', expand=True, fill='x')

        desc_frame = tk.Frame(service_frame)
        desc_frame.pack(fill='x', pady=(2, 5))
        tk.Label(desc_frame, text='서비스 설명:', width=15, anchor='w').pack(side='left')
        tk.Entry(desc_frame, textvariable=self.service_description, width=50).pack(side='left', expand=True, fill='x')

    def create_uninstall_widgets(self, parent):
        tk.Label(parent, text='제거할 서비스 선택:').pack(pady=10)

        self.service_listbox = tk.Listbox(parent, width=50, height=10)
        self.service_listbox.pack(pady=10)

        tk.Button(parent, text='서비스 목록 새로고침', command=self.refresh_service_list).pack()
        tk.Button(parent, text='선택한 서비스 제거', command=self.uninstall_service).pack(pady=20)

        self.refresh_service_list()

    def open_github_link(self):
        webbrowser.open_new('https://github.com/prometheus-community/windows_exporter/releases')

    def get_download_folder(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
                downloads_path = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
            return downloads_path
        except Exception:
            return os.path.join(os.path.expanduser('~'), 'Downloads')

    def select_file(self):
        filetypes = [('Windows Exporter', 'windows_exporter-*-amd64.exe')]
        downloads_folder = self.get_download_folder()
        filename = filedialog.askopenfilename(
            title='Select Windows Exporter file',
            filetypes=filetypes,
            initialdir=downloads_folder
        )
        if filename:
            self.file_path.set(filename)
            self.update_file_label()

    def update_file_label(self):
        if self.file_path.get():
            self.file_label.config(text=f'선택된 파일: {os.path.basename(self.file_path.get())}')
        else:
            self.file_label.config(text='파일이 선택되지 않았습니다.')

    def install_service(self):
        if not self.file_path.get():
            messagebox.showerror('Error', 'Windows Exporter 파일을 선택해주세요')
            return

        try:
            self.run_service_command('create', f'sc create "{self.service_name.get()}" binPath= "{self.file_path.get()}" start= auto DisplayName= "{self.service_name.get()}"')
            self.run_service_command('describe', f'sc description "{self.service_name.get()}" "{self.service_description.get()}"')
            self.run_service_command('failure', f'sc failure "{self.service_name.get()}" reset= 86400 actions= restart/60000/restart/60000/restart/60000')
            self.run_service_command('start', f'sc start "{self.service_name.get()}"')

            messagebox.showinfo('Success', '서비스가 성공적으로 설치되고 시작되었습니다!')
            self.refresh_service_list()
        except subprocess.CalledProcessError as e:
            messagebox.showerror('Error', f'서비스 설치 실패: {e}')

    def run_service_command(self, action, command):
        subprocess.run(command, check=True, shell=True)

    def refresh_service_list(self):
        self.service_listbox.delete(0, tk.END)
        services = self.get_services()
        for service in services:
            self.service_listbox.insert(tk.END, service)

    def get_services(self):
        services = []
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SYSTEM\CurrentControlSet\Services')
            index = 0
            while True:
                try:
                    service_name = winreg.EnumKey(key, index)
                    services.append(service_name)
                    index += 1
                except WindowsError:
                    break
        except WindowsError:
            messagebox.showerror('Error', '서비스 목록을 가져오는 데 실패했습니다')
        return services

    def uninstall_service(self):
        selected = self.service_listbox.curselection()
        if not selected:
            messagebox.showerror('Error', '제거할 서비스를 선택해주세요')
            return

        service_name = self.service_listbox.get(selected[0])
        try:
            self.run_service_command('stop', f'sc stop "{service_name}"')
            self.run_service_command('delete', f'sc delete "{service_name}"')

            messagebox.showinfo('Success', f'서비스 \'{service_name}\'가 성공적으로 제거되었습니다!')
            self.refresh_service_list()
        except subprocess.CalledProcessError as e:
            messagebox.showerror('Error', f'서비스 제거 실패: {e}')

if __name__ == '__main__':
    run_as_admin()
    root = tk.Tk()
    app = ServiceManagerApp(root)
    root.mainloop()