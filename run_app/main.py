import customtkinter
import os
from PIL import Image
from app import ConsultTrainees

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Automatización RPA en el Centro de Comercios SENA")
        self.geometry("700x450")

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images with light and dark mode image
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "test_images")
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "logo-rpa.png")), size=(90, 90))
        self.background_image_1 = customtkinter.CTkImage(Image.open(os.path.join(image_path, "bg_gradient.jpg")), size=(90, 90))
        self.background_image = customtkinter.CTkImage(Image.open(os.path.join(image_path,  "bg.jpeg")), size=(700, 450))
        self.large_test_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "large_test_image.png")), size=(500, 150))
        self.image_icon_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "image_icon_light.png")), size=(20, 20))
        self.home_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "home_dark.png")),
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

        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Proceso Fichas",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.home_image, anchor="w", command=self.home_button_event)
        self.home_button.grid(row=1, column=0, sticky="ew")

        # self.frame_2_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Opcion 2 FRame",
        #                                               fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
        #                                               image=self.home_image, anchor="w", command=self.frame_2_button_event)
        # self.frame_2_button.grid(row=2, column=0, sticky="ew")



        # self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"],
        #                                                         command=self.change_appearance_mode_event)
        # self.appearance_mode_menu.grid(row=6, column=0, padx=20, pady=20, sticky="s")

        # create home frame
        
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)

        background_label = customtkinter.CTkLabel(self.home_frame, text='', image=self.background_image)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)


        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Inicia Sesion en SofiaPlus", font=("Arial", 20, "bold"))
        self.home_frame_large_image_label.grid(row=0, column=0, padx=20, pady=10)

        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Ingrese Usuario:", font=("Arial", 15))
        self.home_frame_large_image_label.grid(row=1, column=0, padx=15, pady=5)

        self.entry_user = customtkinter.CTkEntry(self.home_frame, placeholder_text="Usuario...")
        self.entry_user.grid(row=2, column=0, padx=15, pady=5)

        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Ingrese Contrasena:", font=("Arial", 15))
        self.home_frame_large_image_label.grid(row=3, column=0, padx=15, pady=5)

        self.entry_pass = customtkinter.CTkEntry(self.home_frame, placeholder_text="Contrasena...")
        self.entry_pass.grid(row=4, column=0, padx=15, pady=5)

        self.home_frame_button_4 = customtkinter.CTkButton(self.home_frame, text="Iniciar", compound="bottom", anchor="w",\
                                                           command=lambda: self.next_button(self.entry_user.get(), self.entry_pass.get()))
        self.home_frame_button_4.grid(row=5, column=0, padx=20, pady=30)

        # create second frame
        self.second_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")

        # create third frame
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")

        # select default frame
        self.select_frame_by_name("home")

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.home_button.configure(fg_color=("gray75", "gray25") if name == "home" else "transparent")
        # self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "frame_2" else "transparent")
        
        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.second_frame.grid_forget()
        if name == "frame_3":
            self.third_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.third_frame.grid_forget()


    def next_button(self, usuario, contrasena):
        BotConsulta = ConsultTrainees()
        BotConsulta.open_webdriver()
        BotConsulta.login_process(user=usuario, password=contrasena)
        BotConsulta.select_role()
        list_fichas = BotConsulta.obtain_fichas_a_descargar()
        final_download_files = BotConsulta.depurate_from_existing_files(list_fichas=list_fichas)
        BotConsulta.download_files(final_download_files)
        BotConsulta.modified_files()
        BotConsulta.generate_consolidated_trainees()


    def home_button_event(self):
        self.select_frame_by_name("home")

    def frame_2_button_event(self):
        self.select_frame_by_name("frame_2")

    def frame_3_button_event(self):
        self.select_frame_by_name("frame_3")

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

        


if __name__ == "__main__":
    app = App()
    app.mainloop()

