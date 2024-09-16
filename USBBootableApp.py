import sys
import os
import shutil
import psutil
import threading
import subprocess
import ctypes
import tkinter as tk
import requests
import tempfile
import time
from tkinter import filedialog, messagebox, ttk

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if is_admin():
        print("El script ya se está ejecutando con privilegios de administrador.")
        return True
    else:
        print("Solicitando privilegios de administrador...")
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            return False
        except Exception as e:
            print(f"Error al solicitar privilegios de administrador: {e}")
            return False

if not is_admin():
    if run_as_admin():
        # Esto hará que el script continúe ejecutándose con privilegios de administrador
        sys.exit(0)
    else:
        sys.exit(1)

def list_usb_storage_devices():
    devices = []
    for part in psutil.disk_partitions():
        if 'removable' in part.opts:
            devices.append(part)
    return devices

class USBBootableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ISM-Creador de USB")
        self.run_process = False
        self.stop_copy = False
        self.start_time = None

        
        # Ruta del icono
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, "ism.ico")
            bg_image_path = os.path.join(sys._MEIPASS, "fondo.png")
        else:
            icon_path = os.path.join(os.path.dirname(__file__), "ism.ico")
            bg_image_path = os.path.join(os.path.dirname(__file__), "fondo.png")

        def centrar_ventana(self,root, width, height):
            # Obtener el tamaño de la pantalla
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
                
            # Calcular la posición x y y
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
            
            # Establecer la geometría de la ventana
            self.root.geometry(f'{width}x{height}+{x}+{y}')

        # Definir el tamaño de la ventana
        self.window_width = 670
        self.window_height = 670

        # Centrar la ventana
        centrar_ventana(self, self.root, self.window_width, self.window_height)

        #Icono de la ventana
        self.root.iconbitmap(icon_path)

        #Tamaño de la ventana
        
        self.root.resizable(False, False)
        
        # Create main frame
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=1, relheight=1)
        
        # Set background image
        bg_image = tk.PhotoImage(file=bg_image_path)
        bg_label = tk.Label(self.main_frame, image=bg_image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = bg_image
        bg_label.lift()
        
        # Crear la barra de menú
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Crear el menú "Ayuda"
        help_menu = tk.Menu(menubar,tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=help_menu)
        help_menu.add_command(label="Instrucciones", command=self.show_help)
        help_menu.add_command(label="Info", command=self.show_info)
        
        # Device list
        self.device_label = tk.Label(self.main_frame, text="Dispositivos USB:",highlightthickness=1,highlightbackground="#EF7D00", width=20, height=2,bg="white",font=("Arial",9,"bold"))
        self.device_label.grid(row=0, column=0, padx=(10,0))
        
        self.device_combo = ttk.Combobox(self.main_frame, state="readonly")
        self.device_combo.grid(row=0, column=1)
        self.device_combo.configure(font=("Arial",9,"bold"),justify="center",foreground="black")
        self.refresh_devices()
        

        # Button to rescan USB devices
        self.rescan_button = tk.Button(self.main_frame, text="Escanear", command=self.refresh_devices,width=20,bg="#EF7D00",fg="white",font=("Arial",8,"bold"))
        self.rescan_button.grid(row=0, column=2)
        
        # Filesystem options
        self.filesystem_label = tk.Label(self.main_frame,text="Seleccione formato:",highlightthickness=1,highlightbackground="#EF7D00", width=20, height=2,bg="white",font=("Arial",9,"bold"))
        self.filesystem_label.grid(row=1, column=0, pady=(8, 0),padx=(10,0))
        
        self.filesystem_combo = ttk.Combobox(self.main_frame, values=["NTFS", "FAT32"], state="readonly")
        self.filesystem_combo.grid(row=1, column=1, pady=(8, 0))
        self.filesystem_combo.configure(font=("Arial",9,"bold"),justify="center",foreground="black")
        
        # Source folder selection
        self.source_label = tk.Label(self.main_frame, text="Seleccione Carpeta:",highlightthickness=1,highlightbackground="#EF7D00", width=20, height=2,bg="white",font=("Arial",9,"bold"))
        self.source_label.grid(row=2, column=0, pady=(8, 0),padx=(10,0))
        
        self.source_entry = tk.Entry(self.main_frame, width=50)
        self.source_entry.grid(row=2, column=1, pady=(8, 0))
        
        self.browse_button = tk.Button(self.main_frame, text="Buscar", command=self.browse_folder,width=20,bg="#EF7D00",fg="white",font=("Arial",8,"bold"))
        self.browse_button.grid(row=2, column=2, pady=(8, 0))

        # ISO options
        self.iso_label = tk.Label(self.main_frame, text="Seleccione archivo ISO:",highlightthickness=1,highlightbackground="#EF7D00", width=20, height=2,bg="white",font=("Arial",9,"bold"))
        self.iso_label.grid(row=3, column=0, pady=(8, 0),padx=(10,0))

        self.iso_entry = tk.Entry(self.main_frame, width=50)
        self.iso_entry.grid(row=3, column=1, pady=(8, 0))

        self.iso_browse_button = tk.Button(self.main_frame, text="Buscar", command=self.browse_iso_file,width=20,bg="#EF7D00",fg="white",font=("Arial",8,"bold"))
        self.iso_browse_button.grid(row=3, column=2, pady=(8, 0))

        # Progress bar
        self.progress_label = tk.Label(self.main_frame, text="Progreso:",highlightthickness=1,highlightbackground="#EF7D00", width=20, height=2,bg="white",font=("Arial",9,"bold"))
        self.progress_label.grid(row=4, column=0,pady=(8, 0), padx=(10,0))
        
        self.progress_bar = ttk.Progressbar(self.main_frame, orient="horizontal", length=300, mode="determinate",style="Custom.Horizontal.TProgressbar")
        self.progress_bar.grid(row=4, column=1,pady=(8, 0))
        
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Custom.Horizontal.TProgressbar",troughcolor='white',background='#EF7D00') #thickness=0 

        # Percentage label
        self.percentage_label = tk.Label(self.main_frame,borderwidth=1, relief="solid", text="0%", bg="white",font=("Arial",10,"bold"))
        self.percentage_label.grid(row=4, column=2,pady=(8, 0),padx=(0,0))

        # Log text box
        self.log_text = tk.Text(self.main_frame, height=10, state="disabled",bg="black", fg="white")
        self.log_text.grid(row=5, column=0, columnspan=3,pady=(50,0),padx=(10,0))

        # Start button
        self.start_button = tk.Button(self.main_frame, text="Iniciar", command=self.start_process_thread,width=20,bg="#EF7D00",fg="white",font=("Arial",8,"bold"))
        self.start_button.grid(row=6, column=0,pady=(50,0),padx=(10,0))
        
        #Stop button
        self.stop_button = tk.Button(self.main_frame, text="Detener", command=self.stop_process, width=20,bg="#EF7D00",fg="white",font=("Arial",8,"bold"))
        self.stop_button.grid(row=6, column=1,pady=(50,0))

        # Botón de salir
        self.quit_button = tk.Button(self.main_frame, text="Salir", command=self.close_app, width=20,bg="#EF7D00",fg="white",font=("Arial",8,"bold"))
        self.quit_button.grid(row=6, column=2,pady=(50,0))

        # Timer label
        self.timer_label = tk.Label(self.main_frame, text="Tiempo transcurrido: 00:00:00")
        self.timer_label.grid(row=7, column=1,pady=(100,0))

        # Redirect stdout and stderr
        self.stdout_backup = sys.stdout
        self.stderr_backup = sys.stderr
        sys.stdout = self
        sys.stderr = self
        

    def install_7zip(self):
    # Verificar si 7-Zip ya está instalado
        if self.is_7zip_installed():
            print("7-Zip ya está instalado en el sistema.")
            return

        # URL de descarga del instalador de 7-Zip
        url = "https://www.7-zip.org/a/7z2406-x64.exe"
        
        # Carpeta temporal para almacenar el instalador descargado
        temp_dir = tempfile.gettempdir()
        installer_path = os.path.join(temp_dir, "7z_installer.exe")
        
        try:
            # Descargar el instalador de 7-Zip
            print("Descargando instalador de 7-Zip...")
            with open(installer_path, "wb") as f:
                response = requests.get(url)
                f.write(response.content)
            
            # Ejecutar el instalador
            print("Instalando 7-Zip...")
            subprocess.run([installer_path, "/S"], shell=True, check=True)
            
            print("7-Zip instalado correctamente.")
        except Exception as e:
            print(f"Error al instalar 7-Zip: {e}")
        finally:
            # Eliminar el instalador después de la instalación
            if os.path.exists(installer_path):
                os.remove(installer_path)

    def is_7zip_installed(self):
        # Verificar si 7-Zip está instalado buscando su ejecutable en las rutas del sistema
        # Puedes personalizar este método según la forma en que 7-Zip esté instalado en tu sistema
        return any(
            os.path.exists(os.path.join(path, "7z.exe")) 
            for path in os.environ["PATH"].split(os.pathsep)
        )
    
        
    def write(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def flush(self):
        pass

    def show_help(self):
        help_text = (
            "Instrucciones de uso:\n\n"
            "1. Seleccione el dispositivo USB desde la lista.\n"
            "2. Seleccione el formato de archivo (NTFS o FAT32).\n"
            "3. Seleccione la carpeta de origen con los archivos a copiar.\n"
            "4. Seleccione el archivo ISO de arranque (opcional).\n"
            "5. Haga clic en 'Iniciar' para comenzar el proceso de copiado.\n"
            "6. El progreso se mostrará en la barra de progreso y en el área de registro."
        )
        messagebox.showinfo("Ayuda", help_text)

    def show_info(self):
        info_text = (
            "Información del programa:\n\n"
            "Nombre: ISM USB CREATOR\n"
            "Versión: 0.0.1\n"
            "Autor: Jose Ramon Carvajal Gonzalez\n"
            "Descripción: Aplicación para crear USB de arranque con formato NTFS o FAT32."
        )
        messagebox.showinfo("Información", info_text)    
    
    def refresh_devices(self):
        devices = list_usb_storage_devices()
        self.device_combo['values'] = [f"{device.device} ({device.fstype})" for device in devices]
        if devices:
            self.device_combo.current(0)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder_selected)
    
    def browse_iso_file(self):
        iso_file_selected = filedialog.askopenfilename(filetypes=[("ISO files", "*.iso")])
        if iso_file_selected:
            self.iso_entry.delete(0, tk.END)
            self.iso_entry.insert(0, iso_file_selected)

    def make_usb_bootable(self, device_info, filesystem):
        try:
            device_path = device_info.split()[0]
            drive_letter = device_path[0] + ":"
            if filesystem.upper() == "FAT32":
                # Obtener la ruta completa de fat32format.exe en la carpeta actual
                fat32format_path = os.path.join(os.path.dirname(__file__), "fat32format.exe")
                # Construir el comando con la entrada predeterminada como "Y"
                command = f'echo Y | {fat32format_path} {drive_letter}'
                subprocess.run(command, shell=True, check=True)
                print(f'USB formateada a FAT32 en {drive_letter}')
            else:
                # Usa PowerShell para otros sistemas de archivos
                format_command = f"powershell -Command \"Get-Volume -DriveLetter {drive_letter[0]} | Format-Volume -FileSystem {filesystem} -Confirm:$false\""
                result = subprocess.run(format_command, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                self.write(f"Dispositivo formateado a {filesystem}.\nExtrayendo archivos de la ISO.")
                self.write(f"Salida estándar: {result.stdout}\n")
                self.write(f"Error estándar: {result.stderr}\n")
        except subprocess.CalledProcessError as e:
            self.write(f"Error al formatear el dispositivo: {e}\n")
            self.write(f"Salida estándar: {e.stdout}\n")
            self.write(f"Error estándar: {e.stderr}\n")
        except FileNotFoundError:
            self.write('El programa fat32format no se encuentra en la misma carpeta que el script.\n')
        except OSError as e:
            self.write(f"Error del sistema operativo: {e}\n")


    def copy_files_to_usb(self, source, destination):
        try:
            total_files = sum([len(files) for _, _, files in os.walk(source)])
            file_counter = 0
            for root, dirs, files in os.walk(source):
                if self.stop_copy:
                    self.write("Proceso de copia de archivos detenido por el usuario.\n")
                    return
                for file in files:
                    source_file = os.path.join(root, file)
                    relative_path = os.path.relpath(source_file, source)
                    destination_file = os.path.join(destination, relative_path)
                    destination_dir = os.path.dirname(destination_file)
                    if not os.path.exists(destination_dir):
                        os.makedirs(destination_dir)
                    shutil.copy2(source_file, destination_file)
                    file_counter += 1
                    progress = file_counter / total_files * 100
                    self.progress_bar["value"] = progress
                    self.percentage_label.config(text=f"{progress:.2f}%")
                    self.root.update_idletasks()
                    self.write(f"Copiando archivo: {source_file} -> {destination_file}\n")
            self.write("Archivos copiados correctamente.\n")
        except Exception as e:
            self.write(f"Error al copiar archivos: {e}\n")

    def start_process_thread(self):
        if not self.run_process:
            self.copy_thread = threading.Thread(target=self.start_process)
            self.copy_thread.start()

    def start_process(self):
        device_info = self.device_combo.get()
        if not device_info:
            messagebox.showerror("Error", "No se ha seleccionado ningún dispositivo USB.")
            return

        device_path = device_info.split()[0]
        filesystem = self.filesystem_combo.get()
        if not filesystem:
            messagebox.showerror("Error", "No se ha seleccionado un sistema de archivos.")
            return

        source_folder = self.source_entry.get()
        iso_file = self.iso_entry.get()

        if not source_folder and not iso_file:
            messagebox.showerror("Error", "Debe seleccionar una carpeta o un archivo ISO.")
            return

        # Desactivar el botón de inicio y establecer la bandera de proceso en verdadero
        self.start_button.config(state="disabled")
        self.run_process = True

        try:
            # Primero, formatear el dispositivo USB
            self.make_usb_bootable(device_path, filesystem)
            self.start_time = time.time()  # Start the timer
            threading.Thread(target=self.update_timer).start()  # Start the timer thread

            if source_folder and not iso_file:
                # Si solo se selecciona una carpeta, copiar los archivos de la carpeta al USB
                if not os.path.isdir(source_folder):
                    messagebox.showerror("Error", "La ruta seleccionada no es una carpeta válida.")
                    return
                self.copy_files_to_usb(source_folder, device_path)

            elif iso_file and not source_folder:
                # Si solo se selecciona un archivo ISO, extraer y copiar los archivos al USB
                if not os.path.isfile(iso_file):
                    messagebox.showerror("Error", "La ruta proporcionada no es un archivo ISO válido.")
                    return
                temp_dir = os.path.join(os.path.dirname(__file__), "temp")
                os.makedirs(temp_dir, exist_ok=True)
                
                # Suprimir la ventana del proceso en Windows
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

                # Ajusta la ruta a 7z.exe 
                seven_zip_path = r"C:\Program Files\7-Zip\7z.exe"

                try:
                    result = subprocess.run([seven_zip_path, "x", iso_file, "-o" + temp_dir, "-y"],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,startupinfo=startupinfo)
                
                # Capturar y escribir la salida del proceso
                    self.write(f"Extracción completada.\nSalida estándar: {result.stdout}\n")
                    if result.stderr:
                        self.write(f"Error estándar: {result.stderr}\n")

                    self.copy_files_to_usb(temp_dir, device_path)
                except subprocess.CalledProcessError as e:
                    self.write(f"Error al extraer el archivo ISO: {e}\n")
                    self.write(f"Salida estándar: {e.stdout}\n")
                    self.write(f"Error estándar: {e.stderr}\n")

                # Log de archivos temporales eliminados
                self.write("Borrando archivos temporales de la extracción ISO...\n")
                for root, dirs, files in os.walk(temp_dir, topdown=False):
                    for name in files:
                        temp_file_path = os.path.join(root, name)
                        self.write(f"Eliminando archivo temporal: {temp_file_path}\n")
                        os.remove(temp_file_path)
                    for name in dirs:
                        temp_dir_path = os.path.join(root, name)
                        self.write(f"Eliminando carpeta temporal: {temp_dir_path}\n")
                        os.rmdir(temp_dir_path)
                os.rmdir(temp_dir)
                self.write("Archivos temporales eliminados correctamente.\n")

                # Si se seleccionan ambos, dar un error o manejar de acuerdo a la lógica deseada
            elif source_folder and iso_file:
                messagebox.showerror("Error", "Seleccione solo una carpeta o un archivo ISO, no ambos.")
                return
            
            self.write("Todos los procesos han finalizado.\n")
            self.stop_timer()
        
        finally:
            # Al finalizar el proceso, reactivar el botón de inicio
            self.start_button.config(state="normal")
            self.run_process = False

    def stop_copy_files(self):
        self.stop_copy = True

    def stop_process(self):
        if self.run_process:
            self.run_process = False
            self.stop_copy_files()  # Detener la copia de archivos
            self.write("Proceso detenido por el usuario.\n")
            self.progress_bar["value"] = 0

            # Detener cualquier proceso relacionado con 7z.exe si está en ejecución
            for proc in psutil.process_iter():
                try:
                    if proc.name() == "7z.exe":
                        proc.terminate()
                        proc.wait()
                        self.write("Proceso relacionado con 7z.exe terminado.\n")
                except psutil.AccessDenied:
                    self.write("No se pudo detener el proceso relacionado con 7z.exe debido a un error de permisos.\n")

            # Detener cualquier proceso relacionado con shutil.copy2
            for proc in psutil.process_iter(['pid', 'cmdline']):
                try:
                    if proc.info['cmdline'] and any("copy2" in cmd for cmd in proc.info['cmdline']):
                        proc.terminate()
                        proc.wait()
                        self.write("Proceso relacionado con shutil.copy2 terminado.\n")
                except psutil.AccessDenied:
                    self.write("No se pudo detener el proceso relacionado con shutil.copy2 debido a un error de permisos.\n")
                except Exception as e:
                    self.write(f"Error al detener el proceso relacionado con shutil.copy2: {e}\n")

            # Asegurarse de que la carpeta temp no está en uso antes de eliminarla
            temp_dir = os.path.join(os.path.dirname(__file__), "temp")
            if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
                try:
                    # Intentar liberar archivos en uso antes de eliminarlos
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            try:
                                os.remove(os.path.join(root, file))
                            except Exception as e:
                                self.write(f"No se pudo eliminar el archivo {file}: {e}\n")
                    shutil.rmtree(temp_dir)
                    self.write(f"Carpeta {temp_dir} eliminada.\n")
                except Exception as e:
                    self.write(f"No se pudo eliminar la carpeta {temp_dir}: {e}\n")

        else:
            self.write("No hay ningún proceso en ejecución.\n")    

    def update_timer(self):
        while self.run_process:
            elapsed_time = int(time.time() - self.start_time)
            hours, remainder = divmod(elapsed_time, 3600)
            minutes, seconds = divmod(remainder, 60)
            self.timer_label.config(text=f"Tiempo transcurrido: {hours:02}:{minutes:02}:{seconds:02}")
            time.sleep(1)
            if not self.run_process:
                    break
            
    def stop_timer(self):
        self.run_process = False

    def close_app(self):
        if messagebox.askokcancel("Salir", "¿Está seguro que desea salir?"):
            self.root.destroy()

    def flush(self):
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = USBBootableApp(root)
    app.install_7zip()
    root.mainloop()
