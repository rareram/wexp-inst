import os
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
import winreg
from tkinter import filedialog, messagebox, ttk, PhotoImage
import requests
import socket
import win32com.shell.shell as shell
import win32service
import win32serviceutil
import pywintypes

# 관리자 권한 확인 및 실행
def is_admin():
    try:
        return shell.IsUserAnAdmin()
    except:
        return False

# def run_as_admin():
#     try:
#         if shell.IsUserAnAdmin():
#             return True
#         else:
#             script = os.path.abspath(sys.argv[0])
#             params = ' '.join([script] + sys.argv[1:])
#             shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
#             sys.exit(0)
#     except Exception as e:
#         messagebox.showerror("오류", f"관리자 권한으로 실행하는데 실패했습니다: {str(e)}")
#         return False

    # if not is_admin():
    #     script = os.path.abspath(sys.argv[0])
    #     params = ' '.join([script] + sys.argv[1:])
    #     shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
    #     sys.exit(0)

class ServiceManagerApp:
    def __init__(self, master):
        self.master = master
        master.title('윈도우 익스포터 설치 도우미')
        master.geometry('500x500')
        master.resizable(False, False)

        self.version = '0.4.0'
        self.file_path = tk.StringVar()
        self.service_name = tk.StringVar(value='Prometheus Windows Exporter')
        self.service_description = tk.StringVar(value='Exports Windows metrics for Prometheus')
        self.internal_ip = tk.StringVar()
        self.external_ip = tk.StringVar()

        self.load_images()
        self.create_title()
        self.create_widgets()

    def load_images(self):
        current_dir = os.path.dirname(os.path.abspath(__file__))

        # 로고 이미지 로드
        logo_path = os.path.join(current_dir, 'logo.png')
        self.logo_image = self.load_image(logo_path, 'Logo')

        # Github 아이콘 로드
        github_icon_path = os.path.join(current_dir, 'github_icon.png')
        self.github_icon = self.load_image(github_icon_path, 'Github icon')

        # # web 아이콘 로드
        # web_icon_path = os.path.join(current_dir, 'web_icon.png')
        # self.web_icon = self.load_image(web_icon_path, 'Web icon')

    def load_image(self, path, description):
        try:
            return PhotoImage(file=path)
        except tk.TclError as e:
            print(f'Warning: Could not load {description} from {path}. Error: {e}')
            return None

    def create_title(self):
        title_frame = tk.Frame(self.master)
        title_frame.pack(fill='x', padx=20, pady=(20, 10))

        if self.logo_image:
            logo_label = tk.Label(title_frame, image=self.logo_image)
            logo_label.pack(side=tk.RIGHT, anchor='ne')

        title_font = tkfont.Font(family='Malgun Gothic, Helvetica', size=12, weight='bold')
        title_label = tk.Label(title_frame, text='통합 모니터링 - 윈도우 서버 모니터링 서비스 등록', font=title_font)
        title_label.pack(side=tk.TOP, anchor='nw')

        verdesc_font = tkfont.Font(family='Malgun Gothic, Helvetica', size=8)
        verdesc_label = tk.Label(title_frame, text='2024년 8월  |  IT서비스품질관리팀  | ', font=verdesc_font)
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
        self.create_direct_install_frame(parent)
        self.create_manual_install_frame(parent)
        self.create_verification_frame(parent)
        self.create_service_frame(parent)
        self.create_ip_frame(parent)
        self.create_prometheus_frame(parent)

    def create_download_frame(self, parent):
        download_frame = tk.Frame(parent)
        download_frame.pack(fill='x', padx=10, pady=(10, 2), anchor='w')

        tk.Label(download_frame, text='① GitHub 링크 (windows_exporter): ', anchor='w').pack(side='left')
        if self.github_icon:
            github_icon_label = tk.Label(download_frame, image=self.github_icon)
            github_icon_label.pack(side='left', padx=(5, 0))
        github_link = tk.Label(download_frame, text='Windows Exporter Github Releases', fg='blue', cursor='hand2')
        github_link.pack(side='left')
        github_link.bind('<Button-1>', lambda e: self.open_github_link())

    def create_direct_install_frame(self, parent):
        direct_install_frame = tk.Frame(parent)
        direct_install_frame.pack(fill='x', padx=10, pady=(0, 2), anchor='w')

        tk.Label(direct_install_frame, text='② windows_exporter 직접 설치:', anchor='w').pack(side='left')
        if self.github_icon:
            github_icon_label = tk.Label(direct_install_frame, image=self.github_icon)
            github_icon_label.pack(side='left', padx=(5, 0))
        direct_install_link = tk.Label(direct_install_frame, text='windows_exporter-0.27.1-amd64.msi', fg='blue', cursor='hand2')
        direct_install_link.pack(side='left')
        direct_install_link.bind('<Button-1>', lambda e: self.download_and_install_msi())

    def create_manual_install_frame(self, parent):
        manual_install_frame = tk.Frame(parent)
        manual_install_frame.pack(fill='x', padx=10, pady=(0, 2), anchor='w')

        tk.Label(manual_install_frame, text='③ windows_exporter 수동 설치:', anchor='w').pack(side='left')
        tk.Button(manual_install_frame, text='파일 선택', command=self.select_and_move_file).pack(side='left', padx=(5, 0))

    def create_verification_frame(self, parent):
        verification_frame = tk.Frame(parent)
        verification_frame.pack(fill='x', padx=10, pady=(0, 2), anchor='w')

        tk.Label(verification_frame, text='④ windows_exporter 설치 확인:', anchor='w').pack(side='left')
        # if self.web_icon:
        #     web_icon_label = tk.Label(verification_frame, image=self.web_icon)
        #     web_icon_label.pack(side='left', padx=(5, 0))
        verify_link = tk.Label(verification_frame, text='localhost:9182/metrics', fg='blue', cursor='hand2')
        verify_link.pack(side='left', padx=(5, 0))
        verify_link.bind('<Button-1>', lambda e: webbrowser.open_new('http://localhost:9182/metrics'))

    def create_service_frame(self, parent):
        service_frame = tk.Frame(parent)
        service_frame.pack(fill='x', padx=10, pady=(0, 5), anchor='w')

        tk.Label(service_frame, text='⑤ 서비스 정보 입력:', anchor='w').pack(side='left')
        tk.Button(service_frame, text='서비스 설치', command=self.install_service).pack(side='right')

        # name_frame = tk.Frame(service_frame)
        name_frame = tk.Frame(parent)
        # name_frame.pack(fill='x', pady=(5, 2))
        name_frame.pack(fill='x', padx=(0, 5), pady=(0, 2))
        tk.Label(name_frame, text='      - 서비스 이름', width=15, anchor='w').pack(side='left')
        tk.Entry(name_frame, textvariable=self.service_name, width=40).pack(side='left', expand=True, fill='x')

        # desc_frame = tk.Frame(service_frame)
        desc_frame = tk.Frame(parent)
        # desc_frame.pack(fill='x', pady=(2, 5))
        desc_frame.pack(fill='x', padx=(0, 5), pady=(0, 5))
        tk.Label(desc_frame, text='      - 서비스 설명', width=15, anchor='w').pack(side='left')
        tk.Entry(desc_frame, textvariable=self.service_description, width=40).pack(side='left', expand=True, fill='x')

    def create_ip_frame(self, parent):
        ip_frame = tk.Frame(parent)
        ip_frame.pack(fill='x', padx=10, pady=(0, 2), anchor='w')

        tk.Label(ip_frame, text='⑥ IP address 정보:', anchor='w').pack(side='left')
        tk.Button(ip_frame, text='IP 읽어오기', command=self.update_ip_address).pack(side='right')

        internal_ip_frame = tk.Frame(parent)
        internal_ip_frame.pack(fill='x', padx=(0, 5), pady=(0, 2), anchor='w')
        tk.Label(internal_ip_frame, text='      - 내부 IP', width=15, anchor='w').pack(side='left')
        tk.Entry(internal_ip_frame, textvariable=self.internal_ip, state='readonly', width=20).pack(side='left', expand=True, fill='x')

        external_ip_frame = tk.Frame(parent)
        external_ip_frame.pack(fill='x', padx=(0, 5), pady=(0, 5), anchor='w')
        tk.Label(external_ip_frame, text='      - 외부 IP', width=15, anchor='w').pack(side='left')
        tk.Entry(external_ip_frame, textvariable=self.external_ip, state='readonly', width=20).pack(side='left', expand=True, fill='x')

    def create_prometheus_frame(self, parent):
        prometheus_frame = tk.Frame(parent)
        prometheus_frame.pack(fill='x', padx=10, pady=(0, 5), anchor='w')

        tk.Label(prometheus_frame, text='⑦ prometheus.yml 설정 가이드:', anchor='w').pack(anchor='w')

        config_text = tk.Text(prometheus_frame, height=6, width=50, font=('Consolas', 10))
        config_text.pack(fill='x', expand=True, padx=(20, 0), pady=(5, 0))
        config_text.insert(tk.END,
        """- job_name: 'windows_exporter'
  static_configs:
  - targets: ["{$ip_addr}:9182"]
 # 위의 ⑥ IP 정보를 참고하여 {$ip_addr}에 값을 입력하고
 # 저장한 다음 prometheus 프로세스를 재시작해주세요.""")
        config_text.config(state=tk.DISABLED)

    def create_uninstall_widgets(self, parent):
        tk.Label(parent, text='제거할 windows_exporter 서비스 선택:').pack(pady=(10, 0))

        self.service_listbox = tk.Listbox(parent, width=64, height=16, selectmode=tk.SINGLE)
        self.service_listbox.pack(pady=10)
        self.service_listbox.bind('<Double-1>', self.open_service_properties)

        button_frame = tk.Frame(parent)
        button_frame.pack(pady=10)

        button_width = 15

        open_services_button = tk.Button(button_frame, text='서비스 관리도구', command=self.open_services, width=button_width)
        open_services_button.pack(side=tk.LEFT, padx=3)

        properties_button = tk.Button(button_frame, text='속성 보기', command=self.open_service_properties, width=button_width)
        properties_button.pack(side=tk.LEFT, padx=3)

        refresh_button = tk.Button(button_frame, text='목록 새로고침', command=self.refresh_service_list, width=button_width)
        refresh_button.pack(side=tk.LEFT, padx=3)

        uninstall_button = tk.Button(button_frame, text='서비스 제거', command=self.uninstall_service, width=button_width)
        uninstall_button.pack(side=tk.LEFT, padx=3)

        self.refresh_service_list()

    def open_github_link(self):
        webbrowser.open_new('https://github.com/prometheus-community/windows_exporter/releases')

    def download_and_install_msi(self):
        # messagebox.showinfo('Direct Install', 'This function would download and install windows_exporter.msi')
        # webbrowser.open_new('https://github.com/prometheus-community/windows_exporter/releases/download/v0.27.1/windows_exporter-0.27.1-amd64.msi')
        url = 'https://github.com/prometheus-community/windows_exporter/releases/download/v0.27.1/windows_exporter-0.27.1-amd64.msi'
        filename = 'windows_exporter-0.27.1-amd64.msi'

        try:
            response = requests.get(url)
            response.raise_for_status()    # 오류 발생시 예외 처리
        
            download_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
            file_path = os.path.join(download_folder, filename)

            with open(file_path, 'wb') as file:
                file.write(response.content)
        
            subprocess.run(['msiexec', '/i', file_path], check=True)
            messagebox.showinfo('Success', f'{filename}이(가) 다운로드되었고, 설치가 시작되었습니다.')

            self.file_path.set(file_path)
            self.update_file_label()
    
        except requests.RequestException as e:
            messagebox.showerror('다운로드 오류', f'파일 다운로드에 실패했습니다: {str(e)}')
        except subprocess.CalledProcessError as e:
            messagebox.showerror('설치 오류', f'설치 시작에 실패했습니다: {str(e)}')
        except Exception as e:
            messagebox.showerror('오류', f'예상치 못한 오류가 발생했습니다: {str(e)}')

    def select_and_move_file(self):
        filename = filedialog.askopenfilename(
            title='Select windows_exporter-0.27.1-amd64.exe file',
            filetypes=[('Executable files', '*.exe')]
        )
        if filename:
            dest_dir = r'C:\Program Files\windows_exporter'
            os.makedirs(dest_dir, exist_ok=True)
            dest_file = os.path.join(dest_dir, 'windows_exporter.exe')
            try:
                os.replace(filename, dest_file)
                messagebox.showinfo('Success', f'File moved to {dest_file}')
                self.file_path.set(dest_file)
                self.update_file_label()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to move file: {str(e)}')

    def update_ip_address(self):
        self.internal_ip.set(self.get_internal_ip())
        self.external_ip.set(self.get_external_ip())

    def get_internal_ip(self):
        try:
            return socket.gethostbyname(socket.gethostname())
        except:
            return 'Unable to get internal IP address'

    def get_external_ip(self):
        try:
            return requests.get('https://api.ipify.org').text
        except:
            return 'Unable to get external IP address'

    def install_service(self):
        if not self.file_path.get():
            messagebox.showerror('Error', 'Windows Exporter 파일을 선택해주세요')
            return

        try:
            # 서비스 생성
            self.run_service_command('create', f'sc create "{self.service_name.get()}" binPath= "{self.file_path.get()}" start= auto DisplayName= "{self.service_name.get()}"')
            # 서비스 설명 설정
            self.run_service_command('describe', f'sc description "{self.service_name.get()}" "{self.service_description.get()}"')
            # 서비스 실패 시 동작 설정
            self.run_service_command('failure', f'sc failure "{self.service_name.get()}" reset= 86400 actions= restart/60000/restart/60000/restart/60000')
            # 서비스에 대한 권한 설정
            self.run_service_command('sdset', f'sc sdset "{self.service_name.get()}" D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)')
            # 서비스 시작
            self.run_service_command('start', f'sc start "{self.service_name.get()}"')

            messagebox.showinfo('Success', '서비스가 성공적으로 설치되고 시작되었습니다!')
            self.refresh_service_list()
        except subprocess.CalledProcessError as e:
            error_message = f'서비스 설치 실패: {e}\n\n'
            error_message += '다음 사항을 확인해 주세요:\n'
            error_message += '1. 관리자 권한으로 프로그램을 실행했는지 확인하세요.\n'
            error_message += '2. 안티바이러스 프로그램이 설치를 차단하고 있는지 확인하세요.\n'
            error_message += '3. windows_exporter 파일이 올바른 위치에 있는지 확인하세요.\n'
            error_message += '4. 이미 같은 이름의 서비스가 존재하지 않는지 확인하세요.'
            messagebox.showerror('Error', error_message)

    def open_services(self):
        try:
            subprocess.Popen(['mmc', 'services.msc'], creationflags=subprocess.CREATE_NEW_CONSOLE)
        except Excpetion as e:
            messagebox.showerror('Error', f'서비스 관리 도구를 열 수 없습니다: {str(e)}')

    # def open_service_properties(self, event=None):
    def open_service_properties(self):
        selection = self.service_listbox.curselection()
        if not selection:
            messagebox.showinfo("알림", "서비스를 선택해주세요.")
            return
        
        service_name = self.service_listbox.get(selection[0])
        try:
            # 서비스 상태 확인
            status = win32serviceutil.QueryServiceStatus(service_name)[1]
            status_str = {
                win32service.SERVICE_STOPPED: "중지됨",
                win32service.SERVICE_START_PENDING: "시작 중",
                win32service.SERVICE_STOP_PENDING: "중지 중",
                win32service.SERVICE_RUNNING: "실행 중",
            }.get(status, "알 수 없음")

            # 서비스 정보 가져오기
            service_info = win32serviceutil.GetServiceClassString(service_name)

            # 정보 표시
            info = f"서비스 이름: {service_name}\n"
            info += f"상태: {status_str}\n"
            info += f"서비스 정보: {service_info}\n"

            messagebox.showinfo("서비스 속성", info)

        except pywintypes.error as e:
            if e.winerror == 5:   # 액세스 거부
                messagebox.showerror("오류", "서비스 정보를 가져올 권한이 없습니다. 관리자 권한으로 실행해주세요.")
            else:
                messagebox.showerror("오류", f"서비스 정보를 가져오는 도중 오류가 발생했습니다: {str(e)}")


    def run_as_admin(cls):
        try:
            if shell.IsUserAnAdmin():
                return True
            else:
                script = os.path.abspath(sys.argv[0])
                params = ' '.join([script] + sys.argv[1:])
                shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
                sys.exit(0)
        except Exception as e:
            messagebox.showerror("오류", f"관리자 권한으로 실행하는데 실패했습니다: {str(e)}")
            return False

    def run_service_command(self, action, command):
        subprocess.run(command, check=True, shell=True)

    def refresh_service_list(self):
        self.service_listbox.delete(0, tk.END)
        services = self.get_services()
        for index, service in enumerate(services):
            self.service_listbox.insert(tk.END, service)
            if 'windows_exporter' in service.lower():
                self.service_listbox.selection_set(index)
                self.service_listbox.see(index)

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
            messagebox.showerror('Error', '서비스 목록을 가져오는데 실패했습니다')
        return sorted(services)

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
    # run_as_admin()
    root = tk.Tk()
    app = ServiceManagerApp(root)
    if not app.run_as_admin():
        messagebox.showwarning("경고", "일부 기능이 제한될 수 있습니다. 관리자 권한으로 다시 실행해 주세요.")
    root.mainloop()