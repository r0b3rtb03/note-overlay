import os
import sys
import re
import threading
import json
import shutil
import ctypes
import platform
import threading as _threading
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except Exception:
    TRAY_AVAILABLE = False
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from ctypes import wintypes


class NoteAssistantApp:
    def __init__(self, root, default_file='notes.txt'):
        self.root = root
        self.root.title('Study Notes Assistant')
        self.default_file = default_file
        self.visible = True
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'note_assistant_config.json')

        # If bundled by PyInstaller, resources are unpacked to _MEIPASS at runtime.
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        self.config_path = os.path.join(base_dir, 'note_assistant_config.json')

        # Window size and appearance
        self.root.geometry('640x360')
        self.root.minsize(400, 200)
        self.bg = '#1e1e1e'
        self.fg = '#dcdcdc'
        self.entry_bg = '#2b2b2b'
        self.highlight_bg = '#44475a'
        self.root.configure(bg=self.bg)

        # Top frame with controls
        top = tk.Frame(self.root, bg=self.bg)
        top.pack(fill='x', padx=8, pady=6)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(top, textvariable=self.search_var, bg=self.entry_bg, fg=self.fg, insertbackground=self.fg)
        self.search_entry.pack(side='left', fill='x', expand=True, padx=(0, 6))
        self.search_entry.bind('<Return>', lambda e: self.find_next())

        find_btn = tk.Button(top, text='Find', command=self.find_all, bg=self.entry_bg, fg=self.fg)
        find_btn.pack(side='left')

        next_btn = tk.Button(top, text='Next', command=self.find_next, bg=self.entry_bg, fg=self.fg)
        next_btn.pack(side='left', padx=(4, 0))

        open_btn = tk.Button(top, text='Open', command=self.open_file, bg=self.entry_bg, fg=self.fg)
        open_btn.pack(side='left', padx=(8, 0))

        import_btn = tk.Button(top, text='Import', command=self.import_and_copy, bg=self.entry_bg, fg=self.fg)
        import_btn.pack(side='left', padx=(6, 0))

        # Section filter dropdown
        tk.Label(top, text='Section:', bg=self.bg, fg=self.fg).pack(side='left', padx=(8, 2))
        self.section_var = tk.StringVar(value='All')
        self.section_menu = tk.OptionMenu(top, self.section_var, 'All', command=self._on_section_change)
        self.section_menu.config(bg=self.entry_bg, fg=self.fg, highlightthickness=0)
        self.section_menu['menu'].config(bg=self.entry_bg, fg=self.fg)
        self.section_menu.pack(side='left')

        self.topmost_var = tk.BooleanVar(value=False)
        topmost_cb = tk.Checkbutton(top, text='Always on Top', variable=self.topmost_var, command=self.set_topmost, bg=self.bg, fg=self.fg, selectcolor=self.bg, activebackground=self.bg)
        topmost_cb.pack(side='left', padx=(8, 0))

        # Option to hide from taskbar (Windows only)
        self.hide_taskbar_var = tk.BooleanVar(value=False)
        hide_cb = tk.Checkbutton(top, text='Hide from Taskbar', variable=self.hide_taskbar_var, command=self.apply_taskbar_setting, bg=self.bg, fg=self.fg, selectcolor=self.bg, activebackground=self.bg)
        hide_cb.pack(side='left', padx=(8, 0))

        # Minimize to tray option
        self.minimize_tray_var = tk.BooleanVar(value=False)
        tray_cb = tk.Checkbutton(top, text='Minimize to Tray', variable=self.minimize_tray_var, bg=self.bg, fg=self.fg, selectcolor=self.bg, activebackground=self.bg)
        tray_cb.pack(side='left', padx=(8, 0))
        
        # Stronger hide button (Windows only)
        strong_btn = tk.Button(top, text='Apply Stronger Hide', command=self.apply_stronger_hide, bg=self.entry_bg, fg=self.fg)
        strong_btn.pack(side='left', padx=(8, 0))

        # Text area for notes
        self.text = scrolledtext.ScrolledText(self.root, wrap='word', bg='#121212', fg=self.fg, insertbackground=self.fg)
        self.text.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self.text.tag_config('highlight', background=self.highlight_bg)
        # tag for showing the selected section heading
        self.text.tag_config('section', background=self.highlight_bg)

        # Status bar
        self.status = tk.Label(self.root, text='', anchor='w', bg=self.bg, fg=self.fg)
        self.status.pack(fill='x', padx=8, pady=(0, 6))

        # Internal state
        self.current_search = ''
        self.last_found_index = None
        self._full_text = ''
        self._sections = {}  # name -> text content

        # Load config (geometry, last file, topmost)
        self.load_config()

        # If no file loaded from config, load default file if present
        if not getattr(self, 'current_file', None):
            if os.path.exists(self.default_file):
                self.load_file(self.default_file)
            else:
                self.status.config(text=f'No {self.default_file} found — use Open to load notes')

        # Global hotkeys via Win32 RegisterHotKey (no admin, works in all apps)
        self._hotkey_thread = None
        if platform.system() == 'Windows':
            self._start_hotkey_listener()
        else:
            self.status.config(text='Global hotkeys only supported on Windows')

        # Ensure we unhook on exit
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)

    def _start_hotkey_listener(self):
        HOTKEY_TOGGLE = 1
        HOTKEY_QUIT = 2
        MOD_NONE = 0x0000
        MOD_CTRL = 0x0002
        MOD_SHIFT = 0x0004
        VK_F9 = 0x78
        WM_HOTKEY = 0x0312

        user32 = ctypes.windll.user32

        def listener():
            # RegisterHotKey is thread-bound — register in this thread
            if not user32.RegisterHotKey(None, HOTKEY_TOGGLE, MOD_NONE, VK_F9):
                return
            if not user32.RegisterHotKey(None, HOTKEY_QUIT, MOD_CTRL | MOD_SHIFT, VK_F9):
                user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
                return

            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == WM_HOTKEY:
                    if msg.wParam == HOTKEY_TOGGLE:
                        self.root.after(0, self.toggle_visibility)
                    elif msg.wParam == HOTKEY_QUIT:
                        self.root.after(0, self.on_close)

            user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
            user32.UnregisterHotKey(None, HOTKEY_QUIT)

        t = threading.Thread(target=listener, daemon=True)
        t.start()
        self._hotkey_thread = t

    def load_file(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror('Error', f'Could not open file: {e}')
            return
        self.text.delete('1.0', tk.END)
        self.text.insert('1.0', data)
        self._parse_sections(data)
        self.status.config(text=f'Loaded: {path}')
        self.current_file = path
        self.clear_search()

    def import_and_copy(self):
        path = filedialog.askopenfilename(title='Select notes file to import', filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')])
        if not path:
            return
        try:
            dest = os.path.join(os.path.dirname(os.path.abspath(__file__)), self.default_file)
            shutil.copy(path, dest)
            self.load_file(dest)
            self.status.config(text=f'Imported and copied to {self.default_file}')
        except Exception as e:
            messagebox.showerror('Import Error', f'Could not import file: {e}')

    def load_config(self):
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except Exception:
            return
        geom = cfg.get('geometry')
        if geom:
            try:
                self.root.geometry(geom)
            except Exception:
                pass
        last = cfg.get('last_file')
        if last and os.path.exists(last):
            try:
                self.load_file(last)
            except Exception:
                pass
        topmost = cfg.get('topmost')
        if topmost:
            try:
                self.topmost_var.set(True)
                self.set_topmost()
            except Exception:
                pass
        hide_taskbar = cfg.get('hide_from_taskbar')
        if hide_taskbar:
            self.hide_taskbar_var.set(True)
            # Defer the actual Win32 call until the window is mapped
            self.root.after(100, self.apply_taskbar_setting)
        minimize_tray = cfg.get('minimize_to_tray')
        if minimize_tray:
            try:
                self.minimize_tray_var.set(True)
            except Exception:
                pass

    def save_config(self):
        cfg = {
            'geometry': self.root.winfo_geometry(),
            'last_file': getattr(self, 'current_file', None),
            'topmost': bool(self.topmost_var.get()),
            'hide_from_taskbar': bool(self.hide_taskbar_var.get()),
            'minimize_to_tray': bool(self.minimize_tray_var.get()),
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass

    def open_file(self):
        path = filedialog.askopenfilename(title='Open notes file', filetypes=[('Text/Markdown', '*.txt *.md *.markdown'), ('All files', '*.*')])
        if path:
            self.load_file(path)

    def clear_search(self):
        self.text.tag_remove('highlight', '1.0', tk.END)
        self.current_search = ''
        self.last_found_index = None

    def _parse_sections(self, text):
        """Parse markdown-style # headings into sections."""
        self._full_text = text
        self._sections = {}
        # Find all lines starting with # (top-level headings)
        pattern = re.compile(r'^(#+\s+.+)$', re.MULTILINE)
        matches = list(pattern.finditer(text))
        if not matches:
            return
        for i, m in enumerate(matches):
            name = m.group(1).strip()
            start = m.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            self._sections[name] = text[start:end].rstrip()
        # Update the dropdown. We show a fixed set of desired section labels
        desired = ['# Multiple Choice', '# Short Answer', '# Essay']
        # Build a map from displayed label -> actual heading name (or None if missing)
        self._section_label_map = {}
        for d in desired:
            matched = None
            for h in self._sections:
                if h.lower().startswith(d.lower()):
                    matched = h
                    break
                # no alias mapping — require explicit '# Multiple Choice' heading
                
            self._section_label_map[d] = matched

        menu = self.section_menu['menu']
        menu.delete(0, 'end')
        menu.add_command(label='All', command=lambda: self._set_section('All'))
        for d in desired:
            menu.add_command(label=d, command=lambda l=d: self._set_section(l))
        self.section_var.set('All')

    def _set_section(self, name):
        self.section_var.set(name)
        self._on_section_change(name)

    def _on_section_change(self, value):
        # When a section (other than 'All') is chosen, show only that section's
        # content (hiding the others). Selecting 'All' restores the full text.
        self.text.config(state='normal')
        # remove previous section highlight
        try:
            self.text.tag_remove('section', '1.0', tk.END)
        except Exception:
            pass

        if value == 'All':
            self.text.delete('1.0', tk.END)
            self.text.insert('1.0', self._full_text)
        else:
            # Resolve the actual heading name from the label map
            mapped = getattr(self, '_section_label_map', {}).get(value)
            if not mapped:
                # If the heading doesn't exist yet, append it to the full text
                self._ensure_heading_exists(value)
                mapped = getattr(self, '_section_label_map', {}).get(value)

            content = self._sections.get(mapped, '') if mapped else ''
            self.text.delete('1.0', tk.END)
            self.text.insert('1.0', content)
            # highlight the heading line inside the section view if present
            idx = self.text.search(r'^(#+\s+.+)$', '1.0', regexp=True, stopindex=tk.END)
            if idx:
                line_num = idx.split('.')[0]
                line_start = f"{line_num}.0"
                line_end = f"{line_num}.end"
                self.text.tag_add('section', line_start, line_end)
                self.text.mark_set(tk.INSERT, line_start)
                self.text.see(line_start)

        # clear any previous search highlights
        self.clear_search()
        self.status.config(text=f'Section: {value}')

    # (Removed the Add MC helper — the app now auto-inserts the full Multiple Choice
    # block when a Multiple Choice/MC selection is chosen and the heading is missing.)

    def _ensure_heading_exists(self, display_label):
        """Ensure a heading matching `display_label` exists in the full text.
        If missing, append a simple placeholder section and re-parse."""
        # If it's the multiple choice canonical label or alias, append the
        # full Multiple Choice block provided by the user.
        if display_label.lower().startswith('# multiple'):
            mc_block = """# Multiple Choice
1. Which NIST cloud characteristic describes a multi-tenant environment where multiple customers share physical resources while remaining logically isolated?
    1. Rapid elasticity
    2. **Resource pooling**
    3. Measured service
    4. Location Independence

2. In Amazon S3 (Object Storage), if a single byte of a 5TB file needs to be changed, how does the system execute the update?
    1. It changes only the corresponding storage block.
    2. **It must update and replace the entire file.**
    3. It applies a delta patch via multipart upload.
    4. It alters the object's metadata pointer.

3. Which Route 53 routing policy calculates the optimal endpoint by analyzing geographic coordinates and allows an architect to expand or shrink a region's influence by applying a 'Bias' value?
    1. **Geoproximity routing**
    2. Latency-based routing
    3. Geolocation routing
    4. Weighted routing

4. When creating an Amazon EFS architecture for multiple EC2 instances, where should the EFS mount targets be placed?
    1. **In a private subnet, one per Availability Zone**
    2. In a public subnet attached to an Internet Gateway
    3. In an S3 bucket configured for VPC access
    4. Directly on the Transit Gateway

5. Which NIST service model describes a fully functional application, like Microsoft 365, where the user only needs to provide configuration?
    1. **Software as a Service (SaaS)**
    2. Platform as a Service (PaaS)
    3. Infrastructure as a Service (IaaS)
    4. Anything as a Service (XaaS)

6. What type of application architecture is specifically designed for the cloud, relying on automatic horizontal scaling and stateless execution?
    1. **Cloud native**
    2. Cloud enabled
    3. Lifted and shifted
    4. Monolithic

7. Which type of hypervisor runs directly on the server's bare metal hardware rather than running inside a host operating system?
    1. Type 2 Hypervisor
    2. **Type 1 Hypervisor**
    3. Virtual Machine Monitor
    4. Docker Engine

8. In the OSI Model, which layer handles host-to-host end-to-end addressing and routing?
    1. Layer 2 - Data Link
    2. **Layer 3 - Network**
    3. Layer 4 - Transport
    4. Layer 7 - Application

9. Which Transport Layer protocol drops error recovery and connection establishment features to optimize real-time voice and video streaming?
    1. TCP
    2. **UDP**
    3. IP
    4. DHCP

10. What is the mathematical formula used to calculate the number of physical links required to achieve a full mesh topology between 'n' nodes?
    1. 2^n - 2
    2. **n(n-1)/2**
    3. n(n+1)/2
    4. n^2

11. How many bits make up an IPv6 address?
    1. 32
    2. 48
    3. 64
    4. **128**

12. When creating a VPC in AWS, what is the largest allowable IPv4 CIDR block that can be assigned?
    1. /8
    2. **/16**
    3. /24
    4. /32

13. Which managed AWS service is deployed into a public subnet to enable instances in a private subnet to initiate IPv4 outbound traffic to the internet while preventing the internet from initiating a connection with those instances?
    1. Internet Gateway
    2. **NAT Gateway**
    3. Virtual Private Gateway
    4. VPC Peering Connection

14. Which AWS connectivity feature establishes a direct, non-transitive network relationship between two VPCs, allowing them to communicate using private IP addresses over the AWS global backbone?
    1. **VPC Peering**
    2. VPC Endpoints
    3. AWS PrivateLink
    4. Transit Gateway

15. What AWS networking component allows instances to securely connect to regional services like Amazon S3 and DynamoDB without routing traffic over the public internet?
    1. NAT Gateway
    2. Internet Gateway
    3. **VPC Endpoints**
    4. Network ACL

16. Which Route 53 policy distributes DNS responses based on assigned proportions, making it ideal for A/B testing?
    1. **Weighted routing**
    2. Latency based routing
    3. Simple routing
    4. Failover routing

17. Which Amazon EC2 instance family type is physically optimized specifically for in-memory databases (e.g., r4, r5)?
    1. Compute Optimized
    2. **Memory Optimized**
    3. General Purpose
    4. Accelerated Compute

18. What EC2 deployment feature allows you to pass a bash or PowerShell script to automate patching and software installations during the instance launch?
    1. **User Data**
    2. Instance Metadata
    3. Elastic Block Store
    4. IAM Instance Profile

19. Which EC2 storage option is physically attached to the underlying host computer and is considered ephemeral because its data does not persist if the instance transitions to a Stopped or Terminated state?
    1. Elastic Block Store (EBS)
    2. Amazon S3
    3. **Instance Store**
    4. Elastic File System (EFS)

20. Which special IP address can be queried from within a running EC2 instance to dynamically retrieve its metadata, such as its IAM role or public IP?
    1. http://192.168.1.1/latest/meta-data
    2. **http://169.254.169.254/latest/meta-data**
    3. http://10.0.0.1/latest/meta-data
    4. http://127.0.0.1/latest/meta-data

21. Which EC2 pricing model grants access to spare AWS compute capacity at a steep discount for fault-tolerant workloads, but may be reclaimed with a two-minute interruption notice?
    1. On Demand
    2. Reserved Instances
    3. **Spot Instances**
    4. Dedicated Hosts

22. Which Amazon ECS launch type abstracts the underlying infrastructure, allowing you to run containers without having to provision, configure, or scale the EC2 instances that form the cluster?
    1. **Fargate**
    2. Kubernetes
    3. Elastic Beanstalk
    4. Lightsail

23. Which serverless compute service executes code in response to triggers and maintains a sub-second billing model based on the exact duration of the execution and the memory allocated?
    1. Elastic Beanstalk
    2. Amazon EKS
    3. **AWS Lambda**
    4. AWS Batch

24. When provisioning an EBS volume, to what specific AWS infrastructure boundary is its availability natively restricted?
    1. Region
    2. **Availability Zone**
    3. Edge Location
    4. Virtual Private Cloud (entire VPC)

25. After establishing a full baseline backup of an EBS volume, how does AWS optimize subsequent storage usage for new snapshots?
    1. It compresses the full drive into a ZIP file.
    2. **It performs incremental backups by saving only the modified blocks.**
    3. It moves the snapshot directly to S3 Glacier Deep Archive.
    4. It duplicates the baseline to an alternative Region.

26. Which S3 Storage Class offers the lowest cost for long-term data retention (7-10 years) but imposes a retrieval time of up to 12 hours?
    1. S3 Standard-IA
    2. S3 One Zone-IA
    3. S3 Glacier Flexible Retrieval
    4. **S3 Glacier Deep Archive**

27. Amazon EFS provides scalable file storage. Which network protocol does EFS rely on to allow Linux EC2 instances to mount the file system?
    1. SMB
    2. iSCSI
    3. **NFSv4.1**
    4. FTP

28. When editing /etc/fstab to automatically mount a network drive in AWS, which parameter is critical to prevent the OS from hanging during boot if the storage isn't ready?
    1. defaults
    2. noauto
    3. rw
    4. **_netdev**

29. Under the AWS Shared Responsibility Model, which of the following is an example of AWS's responsibility 'of' the cloud?
    1. Managing S3 bucket policies
    2. **Protecting the physical security of data center facilities**
    3. Patching the guest OS on an EC2 instance
    4. Configuring IAM user passwords

30. Which AWS service is responsible for distributing incoming application traffic across multiple targets, such as EC2 instances, to ensure high availability?
    1. Amazon Route 53
    2. **Elastic Load Balancing (ELB)**
    3. Amazon CloudFront
    4. AWS Direct Connect

31. A company decides to 'Refactor' their application during a cloud migration. What does this process typically involve?
    1. **Modifying or redesigning the application to take advantage of cloud-native services (e.g., serverless).**
    2. Recreating the exact same on-premise architecture using VMs.
    3. Moving the physical servers into a colocation data center.
    4. Changing the billing model without changing the software.
"""
            if not self._full_text.endswith('\n'):
                self._full_text += '\n\n'
            self._full_text += mc_block + '\n\n'
            self._parse_sections(self._full_text)
            return

        # If already present in parsed sections, nothing to do
        for h in self._sections:
            if h.lower().startswith(display_label.lower()):
                return

        # Append a generic placeholder section
        if not self._full_text.endswith('\n'):
            self._full_text += '\n\n'
        if display_label.strip() == '# Short Answer':
            self._full_text += '# Short Answer\n\n- Prompt:\n'
        elif display_label.strip() == '# Essay':
            self._full_text += '# Essay\n\n- Topic:\n'
        else:
            # Fallback: append the display_label literally
            self._full_text += display_label + '\n\n'

        self._parse_sections(self._full_text)

    def find_all(self):
        pattern = self.search_var.get().strip()
        self.text.tag_remove('highlight', '1.0', tk.END)
        if not pattern:
            self.status.config(text='Empty search')
            return
        start = '1.0'
        count = 0
        while True:
            idx = self.text.search(pattern, start, nocase=1, stopindex=tk.END)
            if not idx:
                break
            end = f'{idx}+{len(pattern)}c'
            self.text.tag_add('highlight', idx, end)
            start = end
            count += 1
        self.current_search = pattern
        self.last_found_index = None
        self.status.config(text=f'Found {count} matches for "{pattern}"')

    def find_next(self):
        pattern = self.search_var.get().strip()
        if not pattern:
            self.status.config(text='Empty search')
            return
        if self.current_search != pattern:
            # new search
            self.find_all()
        start_index = '1.0'
        if self.last_found_index:
            # move after last found
            start_index = f'{self.last_found_index}+1c'
        idx = self.text.search(pattern, start_index, nocase=1, stopindex=tk.END)
        if not idx:
            # wrap around
            idx = self.text.search(pattern, '1.0', nocase=1, stopindex=tk.END)
            if not idx:
                self.status.config(text=f'No matches for "{pattern}"')
                return
        end = f'{idx}+{len(pattern)}c'
        self.text.tag_remove(tk.SEL, '1.0', tk.END)
        self.text.tag_add(tk.SEL, idx, end)
        self.text.mark_set(tk.INSERT, end)
        self.text.see(idx)
        self.last_found_index = idx
        self.status.config(text=f'Showing match at {idx}')

    def set_topmost(self):
        v = self.topmost_var.get()
        try:
            self.root.attributes('-topmost', v)
        except Exception:
            pass

    def _get_toplevel_hwnd(self):
        """Get the real Win32 toplevel HWND (not tkinter's inner frame)."""
        self.root.update_idletasks()
        return ctypes.windll.user32.GetParent(int(self.root.winfo_id()))

    def apply_taskbar_setting(self):
        if platform.system() != 'Windows':
            self.status.config(text='Hide-from-taskbar only supported on Windows')
            return

        try:
            hwnd = self._get_toplevel_hwnd()

            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            WS_EX_APPWINDOW = 0x00040000
            SW_HIDE = 0
            SW_SHOW = 5

            if self.hide_taskbar_var.get():
                ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                self._orig_exstyle_taskbar = ex
                # Hide, change style, show — Windows refreshes taskbar on show
                ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                new_ex = (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_ex)
                ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self.status.config(text='Hidden from taskbar')
            else:
                orig = getattr(self, '_orig_exstyle_taskbar', None)
                if orig is not None:
                    ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, orig)
                    ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self.status.config(text='Shown in taskbar')

        except Exception as e:
            self.status.config(text=f'Could not change taskbar visibility: {e}')

    def apply_stronger_hide(self):
        if platform.system() != 'Windows':
            self.status.config(text='Stronger hide only supported on Windows')
            return
        try:
            hwnd = self._get_toplevel_hwnd()
            GWL_EXSTYLE = -20
            WS_EX_TOOLWINDOW = 0x00000080
            WS_EX_APPWINDOW = 0x00040000
            SW_HIDE = 0
            SW_SHOW = 5

            ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            if not getattr(self, '_strong_hide', False):
                self._orig_exstyle = ex
                new_ex = (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
                ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_ex)
                ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self._strong_hide = True
                self.status.config(text='Applied stronger hide')
            else:
                orig = getattr(self, '_orig_exstyle', None)
                if orig is not None:
                    ctypes.windll.user32.ShowWindow(hwnd, SW_HIDE)
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, orig)
                    ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW)
                self._strong_hide = False
                self.status.config(text='Reverted stronger hide')
        except Exception as e:
            self.status.config(text=f'Stronger hide failed: {e}')

    # --- System tray support ---
    def _create_image(self, width=64, height=64, color1=(48,48,48), color2=(200,200,200)):
        img = Image.new('RGB', (width, height), color1)
        d = ImageDraw.Draw(img)
        # simple 'N' letter
        d.text((width*0.2, height*0.15), 'N', fill=color2)
        return img

    def _tray_worker(self, icon):
        try:
            icon.run()
        except Exception:
            pass

    def show_in_tray(self):
        if not TRAY_AVAILABLE:
            self.status.config(text='Install pystray + pillow for tray support')
            return
        if getattr(self, 'tray_icon', None) is not None:
            return
        image = self._create_image()
        menu = pystray.Menu(
            pystray.MenuItem('Restore', lambda: self.root.after(0, self._tray_restore)),
            pystray.MenuItem('Quit', lambda: self.root.after(0, self._tray_quit)),
        )
        # Use a neutral name for the bundled exe
        icon = pystray.Icon('WindowsHelper', image, 'Windows Helper', menu)
        self.tray_icon = icon
        t = _threading.Thread(target=self._tray_worker, args=(icon,), daemon=True)
        t.start()
        self.tray_thread = t
        self.status.config(text='Minimized to tray')

    def remove_tray(self):
        if getattr(self, 'tray_icon', None) is None:
            return
        try:
            self.tray_icon.stop()
        except Exception:
            pass
        self.tray_icon = None
        self.tray_thread = None
        self.status.config(text='Removed tray icon')

    def _tray_restore(self):
        self.remove_tray()
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.visible = True
        except Exception:
            pass

    def _tray_quit(self):
        self.remove_tray()
        self.on_close()

    def toggle_visibility(self):
        try:
            if self.visible:
                # hide: if minimize-to-tray enabled, use tray icon; otherwise withdraw
                if self.minimize_tray_var.get():
                    self.root.withdraw()
                    self.show_in_tray()
                else:
                    self.root.withdraw()
                self.visible = False
            else:
                # restore from tray if present
                if getattr(self, 'tray_icon', None) is not None:
                    self.remove_tray()
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                self.visible = True
        except Exception:
            pass

    def on_close(self):
        # save window state
        try:
            self.save_config()
        except Exception:
            pass
        # Hotkey thread is a daemon — exits automatically when process ends
        self.root.destroy()


def main():
    root = tk.Tk()
    app = NoteAssistantApp(root)

    # If a file path was provided on the command line, load it
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if os.path.exists(arg):
            try:
                app.load_file(arg)
            except Exception:
                pass

    # Run the tkinter mainloop in the main thread (keyboard hook runs in background)
    root.mainloop()


if __name__ == '__main__':
    main()
