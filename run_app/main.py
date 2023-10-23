import customtkinter
import os
from PIL import Image
from app import ConsultTrainees
import sys, shutil
from tkinter import Toplevel
from tkinter import filedialog
from tkinter import messagebox
import subprocess
from tkinter import StringVar

class Redirector(object):
    def __init__(self, text_widget):
        """Constructor"""
        self.output = text_widget

    def write(self, string):
        """Add text to the end and scroll to the end"""
        self.output.insert('end', string)
        self.output.see('end')
        self.output.update_idletasks()

    def flush(self):
        pass


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Automatización RPA en el Centro de Comercios SENA")
        self.geometry("550x500")
        self.current_root_path = os.getcwd()
        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.attributes("-topmost", True)
        
        self.file_path_entrada = None
        self.filename = ""
        self.path_consolidated = ""
        self.empty_process_path = True


        # load images with light and dark mode image
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "test_images")
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "logo-rpa.png")), size=(90, 90))
        self.background_image_1 = customtkinter.CTkImage(Image.open(os.path.join(image_path, "bg_gradient.jpg")), size=(90, 90))
        self.background_image = customtkinter.CTkImage(Image.open(os.path.join(image_path,  "bg.jpeg")), size=(700, 450))
        self.large_test_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "large_test_image.png")), size=(500, 150))
        self.image_icon_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "image_icon_light.png")), size=(20, 20))
        self.home_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "home_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "home_light.png")), size=(20, 20))
        self.historical_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "historical.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "home_light.png")), size=(20, 20))
        self.chat_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "chat_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "chat_light.png")), size=(20, 20))
        self.add_user_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "add_user_dark.png")),
                                                     dark_image=Image.open(os.path.join(image_path, "add_user_light.png")), size=(20, 20))

        # create navigation frame
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="", image=self.logo_image,
                                                             compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Iniciar Robot",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.home_image, anchor="w", command=self.home_button_event)
        self.home_button.grid(row=1, column=0, sticky="ew")


        self.historical_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Ver Historico",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.historical_image, anchor="w", command=self.historical_button_event)
        self.historical_button.grid(row=2, column=0, sticky="ew")


        
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)
        
        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Inicia Sesion en SofiaPlus", font=("Arial", 20, "bold"))
        self.home_frame_large_image_label.grid(row=0, column=0, padx=20, pady=10)

        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Ingrese Usuario:", font=("Arial", 15))
        self.home_frame_large_image_label.grid(row=1, column=0, padx=15, pady=5)

        self.entry_user = customtkinter.CTkEntry(self.home_frame, placeholder_text="Usuario...")
        self.entry_user.grid(row=2, column=0, padx=15, pady=5)

        self.password_entry = customtkinter.CTkLabel(self.home_frame, text="Ingrese Contrasena:", font=("Arial", 15))
        self.password_entry.grid(row=3, column=0, padx=15, pady=5)

        self.entry_pass = customtkinter.CTkEntry(self.home_frame, show="*", placeholder_text="Contrasena...")
        self.entry_pass.grid(row=4, column=0, padx=15, pady=5)

        self.home_frame_button_4 = customtkinter.CTkButton(self.home_frame, text="Iniciar", compound="bottom", anchor="w",\
                                                           command=lambda: self.next_button(self.entry_user.get(), self.entry_pass.get()))
        self.home_frame_button_4.grid(row=5, column=0, padx=20, pady=30)

        #create space
        self.home_frame_text = customtkinter.CTkTextbox(self.home_frame)
        self.home_frame_text.grid(row=6, column=0, padx=15, pady=5, sticky="ew")


        #frame upload

        self.upload_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.upload_frame.grid_columnconfigure(0, weight=1)

        self.upload_frame_large_process_label = customtkinter.CTkLabel(self.upload_frame, text="Cargue el archivo de entrada", font=("Arial", 20, "bold"))
        self.upload_frame_large_process_label.grid(row=0, column=0, padx=20, pady=10)

        self.upload_button = customtkinter.CTkButton(self.upload_frame, text="Cargar Archivo", command=self.open_file)
        self.upload_button.grid(row=1, column=0, padx=20, pady=30, sticky="ew")

        self.filename_var = StringVar()
        self.filename_var.trace('w', self.update_label)
        self.upload_frame_large_up_label = customtkinter.CTkLabel(self.upload_frame, textvariable=self.filename_var, font=("Arial", 15, "bold"))
        self.upload_frame_large_up_label.grid(row=2, column=0, padx=15, pady=5)


        self.upload_frame_button_4 = customtkinter.CTkButton(self.upload_frame, text="Iniciar proceso - Reporte de Aprendices", compound="bottom", anchor="w",\
                                                           command=lambda: self.report_trainees(self.entry_user.get(), self.entry_pass.get()))
        self.upload_frame_button_4.grid(row=3, column=0, padx=20, pady=30)

        self.upload_frame_text = customtkinter.CTkTextbox(self.upload_frame)
        self.upload_frame_text.grid(row=4, column=0, padx=15, pady=5, sticky="ew")

        #frame upload - juicios

        self.upload_frame_juicios = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.upload_frame_juicios.grid_columnconfigure(0, weight=1)

        self.upload_frame_large_process_label_juicios = customtkinter.CTkLabel(self.upload_frame_juicios, text="Cargue el archivo de entrada", font=("Arial", 20, "bold"))
        self.upload_frame_large_process_label_juicios.grid(row=0, column=0, padx=20, pady=10)

        self.upload_button_juicios = customtkinter.CTkButton(self.upload_frame_juicios, text="Cargar Archivo", command=self.open_file)
        self.upload_button_juicios.grid(row=1, column=0, padx=20, pady=30, sticky="ew")

        self.filename_var_juicios = StringVar()
        self.filename_var_juicios.trace('w', self.update_label)
        self.upload_frame_large_up_label_juicios = customtkinter.CTkLabel(self.upload_frame_juicios, textvariable=self.filename_var, font=("Arial", 15, "bold"))
        self.upload_frame_large_up_label_juicios.grid(row=2, column=0, padx=15, pady=5)


        self.upload_frame_button_4_juicios = customtkinter.CTkButton(self.upload_frame_juicios, text="Iniciar proceso - Reporte de Juicios", compound="bottom", anchor="w",\
                                                           command=lambda: self.report_juicios(self.entry_user.get(), self.entry_pass.get()))
        self.upload_frame_button_4_juicios.grid(row=3, column=0, padx=20, pady=30)

        self.upload_frame_text_juicios = customtkinter.CTkTextbox(self.upload_frame_juicios)
        self.upload_frame_text_juicios.grid(row=4, column=0, padx=15, pady=5, sticky="ew")



                #login

        self.login_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.login_frame.grid_columnconfigure(0, weight=1)

        
        self.login_frame_large_proces_label = customtkinter.CTkLabel(self.login_frame, text="Seleccione proceso", font=("Arial", 20, "bold"))
        self.login_frame_large_proces_label.grid(row=0, column=0, padx=20, pady=10)
        # Create buttons

        self.login_frame_button_4 = customtkinter.CTkButton(self.login_frame, text="Reporte de Aprendices", compound="bottom", anchor="w",\
                                                           command=self.upload_button_event)
        self.login_frame_button_4.grid(row=1, column=0, padx=20, pady=30)

        self.login_frame_button_4 = customtkinter.CTkButton(self.login_frame, text="Reporte de Juicios", compound="bottom", anchor="w",\
                                                           command=self.upload_juicio_button_event)
        self.login_frame_button_4.grid(row=2, column=0, padx=20, pady=30)

        self.login_frame_text = customtkinter.CTkTextbox(self.login_frame)
        self.login_frame_text.grid(row=3, column=0, padx=20, pady=5, sticky="ew")


        #Ver archivo

        self.file_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.file_frame.grid_columnconfigure(0, weight=1)

        self.file_frame_large_process_label = customtkinter.CTkLabel(self.file_frame, text="Consolidado Generado", font=("Arial", 20, "bold"))
        self.file_frame_large_process_label.grid(row=0, column=0, padx=20, pady=10)

        self.file_button = customtkinter.CTkButton(self.file_frame, text="Ver archivo", command=lambda: self.open_entregable(self.path_consolidated))
        self.file_button.grid(row=1, column=0, padx=20, pady=30, sticky="ew")    


         # frame Ver historico

        self.historical_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.historical_frame.grid_columnconfigure(0, weight=1)

        self.historical_frame_large_process_label = customtkinter.CTkLabel(self.historical_frame, text="Historicos de Consolidado", font=("Arial", 20, "bold"))
        self.historical_frame_large_process_label.grid(row=0, column=0, padx=20, pady=10)

        self.historical_button = customtkinter.CTkButton(self.historical_frame, text="Ver historico", command=lambda: self.open_entregable(f"{self.current_root_path}/reporte"))
        self.historical_button.grid(row=1, column=0, padx=20, pady=30, sticky="ew")        

        # select default frame
        self.select_frame_by_name("home")


    def open_file(self):
        self.file_path_entrada = filedialog.askopenfilename()
        self.filename = os.path.basename(self.file_path_entrada)
        if self.file_path_entrada:
            messagebox.showinfo('Archivo cargado', 'Archivo subido exitosamente')
            destination = f"{self.current_root_path}/entrada"
            shutil.copy(self.file_path_entrada, destination)
            self.update_idletasks()
        else:
            messagebox.showinfo('No archivo seleccionado', 'Ningun archivo ha sido seleccionado')
        self.filename_var.set(f"Archivo cargado: {self.filename}")

    def open_entregable(self, path_to_start):
        if os.name == 'nt':
            os.startfile(f'{self.current_root_path}/reporte')
        else:  # Unix
            subprocess.call(['open', f'{self.current_root_path}/reporte'])
        

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.home_button.configure(fg_color=("gray75", "gray25") if name == "home" else "transparent")
        # self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "frame_2" else "transparent")
        
        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "login":
            self.login_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.login_frame.grid_forget()
        if name == "upload":
            self.upload_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.upload_frame.grid_forget()
        if name == "file":
            self.file_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.file_frame.grid_forget()
        if name == "historical":
            self.historical_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.historical_frame.grid_forget()
        if name == "upload_juicio":
            self.upload_frame_juicios.grid(row=0, column=1, sticky="nsew")
        else:
            self.upload_frame_juicios.grid_forget()


    def report_trainees(self, usuario, contrasena):
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = Redirector(self.upload_frame_text)
        sys.stderr = Redirector(self.upload_frame_text)
        try:
            if os.listdir(f'{os.getcwd()}/procesar/'):
                # Directory is not empty
                continue_process = messagebox.askyesno("La carpeta procesar no esta vacía",
                                                   f"La carpeta '{self.current_root_path}/procesar/' contiene archivos descargados. Desea continuar?")
            
                if not continue_process:
                    self.empty_process_path = False  # Exit from the function or loop
                else:
                    self.empty_process_path = True
            else:
                self.empty_process_path = True

            if self.empty_process_path == True:
                BotConsulta = ConsultTrainees()
                BotConsulta.open_webdriver()
                response_login = BotConsulta.login_process(user=usuario, password=contrasena)
                BotConsulta.select_role()
                list_fichas = BotConsulta.obtain_fichas_a_descargar(document_name=self.filename)
                if list_fichas ==False:
                    pass
                else:
                    final_download_files = BotConsulta.depurate_from_existing_files(list_fichas=list_fichas)
                    BotConsulta.download_files(final_download_files)
                    BotConsulta.modified_files()
                    self.path_consolidated = BotConsulta.generate_consolidated_trainees()
                    if self.path_consolidated != False:
                        BotConsulta.delete_all_files_in_procesar_and_entrada(self.path_consolidated)
                        self.select_frame_by_name("file")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


    def report_juicios(self, usuario, contrasena):
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = Redirector(self.upload_frame_text)
        sys.stderr = Redirector(self.upload_frame_text)
        try:
            if os.listdir(f'{os.getcwd()}/procesar/'):
                # Directory is not empty
                continue_process = messagebox.askyesno("La carpeta procesar no esta vacía",
                                                   f"La carpeta '{self.current_root_path}/procesar/' contiene archivos descargados. Desea continuar?")
            
                if not continue_process:
                    self.empty_process_path = False  # Exit from the function or loop
                else:
                    self.empty_process_path = True
            else:
                self.empty_process_path = True

            if self.empty_process_path == True:
                # Directory is empty, continue with the process
                BotConsulta = ConsultTrainees()
                BotConsulta.open_webdriver()
                response_login = BotConsulta.login_process(user=usuario, password=contrasena)
                BotConsulta.select_role()
                list_fichas = BotConsulta.obtain_fichas_a_descargar(document_name=self.filename)
                if list_fichas ==False:
                    pass
                else:
                    final_download_files = BotConsulta.depurate_from_existing_files(list_fichas=list_fichas)
                    BotConsulta.download_juicios_process(final_download_files)
                    BotConsulta.restructure_for_consolidated_file()
                    self.path_consolidated = BotConsulta.generate_consolidated_juicios()
                    if self.path_consolidated != False:
                        BotConsulta.delete_all_files_in_procesar_and_entrada(self.path_consolidated)
                        self.select_frame_by_name("file")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr



    def next_button(self, usuario, contrasena):
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = Redirector(self.home_frame_text)
        sys.stderr = Redirector(self.home_frame_text)
        BotConsulta = ConsultTrainees()
        BotConsulta.open_webdriver()
        response_login = BotConsulta.login_process(user=usuario, password=contrasena)
        if response_login:
            self.select_frame_by_name("login")
        # BotConsulta.select_role()
        # list_fichas = BotConsulta.obtain_fichas_a_descargar()
        # final_download_files = BotConsulta.depurate_from_existing_files(list_fichas=list_fichas)
        # BotConsulta.download_files(final_download_files)
        # BotConsulta.modified_files()
        # BotConsulta.generate_consolidated_trainees()

        sys.stdout = old_stdout
        sys.stderr = old_stderr

    def select_trainees(self):
        pass

    def home_button_event(self):
        self.select_frame_by_name("home")

    def historical_button_event(self):
        self.select_frame_by_name("historical")

    def upload_button_event(self):
        self.select_frame_by_name("upload")

    def upload_juicio_button_event(self):
        self.select_frame_by_name("upload_juicio")

    def frame_3_button_event(self):
        self.select_frame_by_name("frame_3")

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def update_label(self, *args):
        self.upload_frame_large_up_label.config(text=self.filename_var.get())




        


if __name__ == "__main__":
    app = App()
    app.mainloop()

