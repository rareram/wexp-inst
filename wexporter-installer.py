import tkinter as tk
from tkinter import filedialog, messagebox, ttk, PhotoImage
import os
import subprocess
import sys
import win32com.shell.shell as shell
import winreg
import webbrowser

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
        master.title("윈도우 익스포터 설치 도우미")
        master.geometry("500x500")

        self.file_path = tk.StringVar()
        self.service_name = tk.StringVar(value="Prometheus Windows Exporter")
        self.service_description = tk.StringVar(value="Exports Windows metrics for Prometheus")

        self.load_images()

        self.create_widgets()

    def load_images(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        image_path = os.path.join(current_dir, "github_icon.png")
        try:
            self.github_icon = PhotoImage(file=image_path)
        except tk.TclError:
            print(f"Warning: Could not load image from {image_path}")
            self.github_icon = None

    def create_widgets(self):
        notebook = ttk.Notebook(self.master)
        notebook.pack(expand=True, fill='both')

        install_frame = ttk.Frame(notebook)
        uninstall_frame = ttk.Frame(notebook)

        notebook.add(install_frame, text='서비스 등록')
        notebook.add(uninstall_frame, text='서비스 제거')

        self.create_install_widgets(install_frame)
        self.create_uninstall_widgets(uninstall_frame)

    def open_github_link(self):
        webbrowser.open_new("https://github.com/prometheus-community/windows_exporter/releases")

    def get_download_folder(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
                downloads_path = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
            return downloads_path
        except Exception:
            # 레지스트리에서 경로를 찾지 못한 경우 기본값 반환
            return os.path.join(os.path.expanduser('~'), 'Downloads')

    def select_file(self):
        filetypes = [("Windows Exporter", "windows_exporter-*-amd64.exe")]
        downloads_folder = self.get_download_folder()
        filename = filedialog.askopenfilename(
            title = "Select Windows Exporter file",
            filetypes = filetypes,
            #initialdir="/"
            initialdir = downloads_folder
        )
        if filename:
            self.file_path.set(filename)
            self.update_file_label()

    def update_file_label(self):
        if self.file_path.get():
            self.file_label.config(text=f"선택된 파일: {os.path.basename(self.file_path.get())}")
        else:
            self.file_label.config(text="파일이 선택되지 않았습니다.")

    def create_install_widgets(self, parent):
        # Windows Expoerter 다운로드 라인
        download_frame = tk.Frame(parent)
        download_frame.pack(fill='x', padx=10, pady=(10,2), anchor='w')

        tk.Label(download_frame, text="① Windows Exporter 다운로드: ", anchor="w").pack(side='left')
        if self.github_icon:
            github_icon_label = tk.Label(download_frame, image=self.github_icon)
            github_icon_label.pack(side='left', padx=(0, 5))
        github_link = tk.Label(download_frame, text="Windows Exporter Github Releases", fg="blue", cursor="hand2")
        github_link.pack(side='left')
        github_link.bind("<Button-1>", lambda e: self.open_github_link())

        # Windows Exporter 파일선택 라인
        select_frame = tk.Frame(parent)
        select_frame.pack(fill='x', padx=10, pady=(2,5), anchor='w')

        tk.Label(select_frame, text="② Windows Exporter 선택:", anchor="w").pack(side='left')
        tk.Button(select_frame, text="파일 선택", command=self.select_file).pack(side='left', padx=(5,0))


        # 선택된 파일 표시 라벨
        self.file_label = tk.Label(parent, text="아직 파일이 선택되지 않았습니다.", anchor="w")
        self.file_label.pack(fill='x', padx=10, pady=(0,10), anchor='w')

        # 서비스 이름 및 설명 수정/등록
        service_frame = tk.Frame(parent)
        service_frame.pack(fill='x', padx=10, pady=5, anchor='w')

        tk.Label(service_frame, text="③ 서비스 정보 입력:", anchor="w").pack(anchor='w')
    
        name_frame = tk.Frame(service_frame)
        name_frame.pack(fill='x', pady=(5,2))
        tk.Label(name_frame, text="서비스 이름:", width=15, anchor='w').pack(side='left')
        tk.Entry(name_frame, textvariable=self.service_name, width=50).pack(side='left', expand=True, fill='x')

        desc_frame = tk.Frame(service_frame)
        desc_frame.pack(fill='x', pady=(2,5))
        tk.Label(desc_frame, text="서비스 설명:", width=15, anchor='w').pack(side='left')
        tk.Entry(desc_frame, textvariable=self.service_description, width=50).pack(side='left', expand=True, fill='x')

        # 서비스 설치 버튼
        tk.Button(parent, text="서비스 설치", command=self.install_service).pack(pady=20)

    def create_uninstall_widgets(self, parent):
        tk.Label(parent, text="Select service to uninstall:").pack(pady=10)

        self.service_listbox = tk.Listbox(parent, width=50, height=10)
        self.service_listbox.pack(pady=10)

        tk.Button(parent, text="Refresh Service List", command=self.refresh_service_list).pack()
        tk.Button(parent, text="Uninstall Selected Service", command=self.uninstall_service).pack(pady=20)

        self.refresh_service_list()

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe")])
        if filename:
            self.file_path.set(filename)

    def install_service(self):
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select the Windows Exporter file")
            return

        try:
            create_cmd = f'sc create "{self.service_name.get()}" binPath= "{self.file_path.get()}" start= auto DisplayName= "{self.service_name.get()}"'
            subprocess.run(create_cmd, check=True, shell=True)

            describe_cmd = f'sc description "{self.service_name.get()}" "{self.service_description.get()}"'
            subprocess.run(describe_cmd, check=True, shell=True)

            failure_cmd = f'sc failure "{self.service_name.get()}" reset= 86400 actions= restart/60000/restart/60000/restart/60000'
            subprocess.run(failure_cmd, check=True, shell=True)

            start_cmd = f'sc start "{self.service_name.get()}"'
            subprocess.run(start_cmd, check=True, shell=True)

            messagebox.showinfo("Success", "Service installed and started successfully!")
            self.refresh_service_list()
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to install service: {e}")

    def refresh_service_list(self):
        self.service_listbox.delete(0, tk.END)
        services = self.get_services()
        for service in services:
            self.service_listbox.insert(tk.END, service)

    def get_services(self):
        services = []
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services")
            index = 0
            while True:
                try:
                    service_name = winreg.EnumKey(key, index)
                    services.append(service_name)
                    index += 1
                except WindowsError:
                    break
        except WindowsError:
            messagebox.showerror("Error", "Failed to retrieve services list")
        return services

    def uninstall_service(self):
        selected = self.service_listbox.curselection()
        if not selected:
            messagebox.showerror("Error", "Please select a service to uninstall")
            return

        service_name = self.service_listbox.get(selected[0])
        try:
            stop_cmd = f'sc stop "{service_name}"'
            subprocess.run(stop_cmd, check=True, shell=True)

            delete_cmd = f'sc delete "{service_name}"'
            subprocess.run(delete_cmd, check=True, shell=True)

            messagebox.showinfo("Success", f"Service '{service_name}' uninstalled successfully!")
            self.refresh_service_list()
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to uninstall service: {e}")

if __name__ == "__main__":
    run_as_admin()
    root = tk.Tk()
    app = ServiceManagerApp(root)
    root.mainloop()