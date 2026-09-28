import os # for file paths
import tkinter as tk  # for UI elements and event handling
from tkinter import messagebox # for message boxes and error handling
from PIL import Image, ImageTk # for image handling and display

# "Python, try to run this code."
try:
    # it is used to play audio
    # pywin32 is a Python package that provides a Python interface to the Windows COM API.
    import win32com.client  
    # "Yes, the library is available." 
    HAS_WIN32COM = True
except ImportError: # If the library is not available
    # "No, the library is not available."
    HAS_WIN32COM = False


# --- FILE PATHS --- replace with your own paths
IMAGE_DIR = r"C:\Users\HP G8\Desktop\Balochi\images\renamed_images"
AUDIO_DIR = r"C:\Users\HP G8\Desktop\Balochi\audio"



# --- BASE CLASS --- Every Balochi letter should have some information.
class AlphabetItem:
    """Base class for Balochi letters"""
    def __init__(self, name, sound, example, image_index):
        self.name = name
        self.sound = sound
        self.example = example
        # This creates the complete image filename.
        self.image_path = os.path.join(
            IMAGE_DIR, f"image_{image_index:02d}.jpg"
        )
        # This creates the complete audio filename.
        self.audio_path = os.path.join(
            AUDIO_DIR, f"sound_{image_index:02d}.wav.mp4"
        )


# --- DATASET & CATEGORIES ---
# This list contains many AlphabetItem objects.
# The dataset is a list of Balochi letters 
# Each letter has a name, a sound, an example, and an image index

ALL_LETTERS = [
    AlphabetItem("A", "Short A sound", "Anár (Pomegranate)", 1),
    AlphabetItem("Á", "Long A sound", "Ázmán (sky)", 2),
    AlphabetItem("B", "B sound", "Báli (Airplane)", 3),
    AlphabetItem("Ch", "Ch sound", "Chamm (Eye)", 4),
    AlphabetItem("D", "D sound", "Dast (Hand)", 5),
    AlphabetItem("Dh", "D sound", "Dhál (beans)", 6),
    AlphabetItem("E", "E sound", "Estár (Star)", 7),
    AlphabetItem("É", "Long E sound", "Ésherk (Rhazya stricta)", 8),
    AlphabetItem("G", "G sound", "Gósh (Ear)", 9),
    AlphabetItem("H", "H sound", "Hayk (Egg)", 10),
    AlphabetItem("I", "Short I sound", "Inch (ruler)", 11),
    AlphabetItem("J", "J sound", "Jig (Front Panel)", 12),
    AlphabetItem("K", "K sound", "Kóndh (Knee)", 13),
    AlphabetItem("L", "L sound", "Lonth (Laps)", 14),
    AlphabetItem("M", "M sound", "Móz (Banana)", 15),
    AlphabetItem("N", "N sound", "Nibag (Fruit)", 16),
    AlphabetItem("O", "O sound", "Ostád (Teacher)", 17),
    AlphabetItem("Ó", "Long O sound", "Ózhnág (Swimming)", 18),
    AlphabetItem("P", "P sound", "Poll (Flower)", 19),
    AlphabetItem("R", "R sound", "Róch (Sun)", 20),
    AlphabetItem("S", "S sound", "Seng (Stone)", 21),
    AlphabetItem("Sh", "Sh sound", "Shir (Milk)", 22),
    AlphabetItem("T", "T sound", "Ták (Leave)", 23),
    AlphabetItem("Th", "Retroflex T sound", "Thásin (safety pin)", 24),
    AlphabetItem("U", "U sound", "Ud (Oud wood chips)", 25),
    AlphabetItem("W", "W sound", "Wád (Salt)", 26),
    AlphabetItem("Y", "Y sound", "Yakk (One)", 27),
    AlphabetItem("Z", "Z sound", "Zer (Beach)", 28),
    AlphabetItem("Zh", "Zh sound", "Zháng (ghungroos)", 29),
]


# create a function.
# The function returns a list of AlphabetItem objects.
# Go through every letter in ALL_LETTERS and keep only the letters whose name exists in names_list.
# This function helps to create different categories.
def filter_by_names(names_list):
    return [item for item in ALL_LETTERS if item.name in names_list]


# This is a dictionary.
# It stores different groups of letters.
CATEGORIES = {
    "1. Balóchiay Áb": ("Balóchiay Áb", ALL_LETTERS),
    "2. Balóchiay Jwánáb": (
        "Balóchiay Jwánáb",
        filter_by_names(
            [
                "B",
                "Ch",
                "D",
                "Dh",
                "G",
                "H",
                "J",
                "K",
                "L",
                "M",
                "N",
                "P",
                "R",
                "S",
                "Sh",
                "T",
                "Th",
                "W",
                "Y",
                "Z",
                "Zh",
            ]
        ),
    ),
    "3. Balóchiay Jotkáb": (
        "Balóchiay Jotkáb",
        filter_by_names(["Ch", "Dh", "Sh", "Th", "Zh"]),
    ),
    "4. Drájén Kassháb": (
        "Drájén Kassháb",
        filter_by_names(["Á", "É", "I", "Ó", "U"]),
    ),
    "5. Gwandhén Kassháb": (
        "Gwandhén Kassháb",
        filter_by_names(["A", "E", "O"]),
    ),
}


# --- PREMIUM GUI SUITE ---

# This is the main class.
# The BalochiApp object controls the whole GUI.
class BalochiApp:

    # This is the constructor. 
    # It takes the root window as an argument.
    # The root window is the main window of the GUI.
    def __init__(self, root):
        self.root = root # The root window
        self.root.title("Balochi Bongéji - Studio Edition") # The title of the root window
        self.root.geometry("1000x720") # The size of the root window
        self.root.minsize(950, 680) # The minimum size of the root window
        self.root.config(bg="#0f172a") # The background color of the root window

        # Current Category
        self.current_cat_key = "1. Balóchiay Áb" # The key of the current category
        self.current_list = CATEGORIES[self.current_cat_key][1] # The list of the current category
        self.current_index = 0 # The index of the current item in the current list

        # Background Audio Engine
        self.player = None # The audio player
        if HAS_WIN32COM: # If the audio player is available 
            try: # Try to create the audio player

                # "Python, try to run this code."
                # Now program can use it to play audio.
                self.player = win32com.client.Dispatch("WMPlayer.OCX")
            except Exception: # If the audio player is not available
                # "No, the library is not available."
                self.player = None

        # UI Elements
        # This dictionary will store your category buttons.
        self.sidebar_buttons = {}
        # This is a list. 
        # This list will store letter buttons.
        self.grid_buttons = []

        # Setup UI
        # This function sets up the UI.
        # It creates the UI elements.
        self.setup_ui()
        # Load Item
        self.load_item()

    # This method creates the complete interface.
    # It creates the UI elements.
    def setup_ui(self):
        # Master Layout: Left Sidebar + Main Studio
        self.sidebar = tk.Frame(self.root, bg="#1e293b", width=240)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        self.studio = tk.Frame(self.root, bg="#0f172a")
        self.studio.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- SIDEBAR UI ---
        # App Brand Header
        brand_frame = tk.Frame(self.sidebar, bg="#1e293b", pady=20, padx=15)
        brand_frame.pack(fill=tk.X)

        tk.Label(
            brand_frame,
            text="بلوچی بنگیجی",
            font=("Segoe UI", 22, "bold"),
            fg="#14b8a6",
            bg="#1e293b",
        ).pack(anchor="w")

        tk.Label(
            brand_frame,
            text="BALOCHI LEARNING SUITE",
            font=("Segoe UI", 8, "bold"),
            fg="#64748b",
            bg="#1e293b",
        ).pack(anchor="w", pady=(2, 0))

        tk.Frame(self.sidebar, bg="#334155", height=1).pack(
            fill=tk.X, padx=15, pady=5
        )

        # Navigation Category Links
        nav_frame = tk.Frame(self.sidebar, bg="#1e293b", pady=10)
        nav_frame.pack(fill=tk.X)

        for key, (label, _) in CATEGORIES.items():
            btn = tk.Button(
                nav_frame,
                text=f"  {label}",
                font=("Segoe UI", 10, "bold"),
                anchor="w",
                bg="#1e293b",
                fg="#94a3b8",
                activebackground="#0f766e",
                activeforeground="#ffffff",
                relief=tk.FLAT,
                bd=0,
                padx=15,
                pady=10,
                cursor="hand2",
                command=lambda k=key: self.switch_category(k),
            )
            btn.pack(fill=tk.X, pady=2)
            self.sidebar_buttons[key] = btn

        # --- STUDIO CONTENT UI ---
        # Top Header
        top_bar = tk.Frame(self.studio, bg="#0f172a", pady=15, padx=25)
        top_bar.pack(fill=tk.X)

        self.cat_title = tk.Label(
            top_bar,
            font=("Segoe UI", 14, "bold"),
            fg="#f8fafc",
            bg="#0f172a",
        )
        self.cat_title.pack(side=tk.LEFT)

        self.counter_badge = tk.Label(
            top_bar,
            font=("Segoe UI", 10, "bold"),
            fg="#14b8a6",
            bg="#134e4a",
            padx=12,
            pady=4,
        )
        self.counter_badge.pack(side=tk.RIGHT)

        # Center Showcase Box (Split Image & Letter Card)
        showcase = tk.Frame(self.studio, bg="#0f172a", padx=25)
        showcase.pack(fill=tk.X)

        # 1. Left Card: Image Container
        img_card = tk.Frame(
            showcase,
            bg="#1e293b",
            bd=1,
            relief=tk.SOLID,
            highlightbackground="#334155",
            highlightthickness=1,
            width=330,
            height=330,
        )
        img_card.pack(side=tk.LEFT, padx=(0, 20))
        img_card.pack_propagate(False)

        self.img_label = tk.Label(img_card, bg="#1e293b")
        self.img_label.pack(expand=True, fill=tk.BOTH, padx=8, pady=8)

        # 2. Right Card: Details Box
        info_card = tk.Frame(showcase, bg="#1e293b", padx=20, pady=20)
        info_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(
            info_card,
            text="BALOCHI LETTER",
            font=("Segoe UI", 9, "bold"),
            fg="#64748b",
            bg="#1e293b",
        ).pack(anchor="w")

        self.letter_lbl = tk.Label(
            info_card,
            font=("Segoe UI", 38, "bold"),
            fg="#38bdf8",
            bg="#1e293b",
        )
        self.letter_lbl.pack(anchor="w", pady=(0, 5))

        self.sound_lbl = tk.Label(
            info_card,
            font=("Segoe UI", 11),
            fg="#94a3b8",
            bg="#1e293b",
        )
        self.sound_lbl.pack(anchor="w", pady=(0, 5))

        self.example_lbl = tk.Label(
            info_card,
            font=("Segoe UI", 13, "bold"),
            fg="#34d399",
            bg="#1e293b",
        )
        self.example_lbl.pack(anchor="w", pady=(0, 15))

        # Audio Button
        self.tawar_btn = tk.Button(
            info_card,
            text="🔊  PLAY TAWAR (VOICE)",
            font=("Segoe UI", 11, "bold"),
            bg="#0d9488",
            fg="#ffffff",
            activebackground="#0f766e",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            pady=10,
            cursor="hand2",
            command=self.play_tawar,
        )
        self.tawar_btn.pack(fill=tk.X, pady=(0, 15))

        # Next / Prev Navigation
        nav_row = tk.Frame(info_card, bg="#1e293b")
        nav_row.pack(fill=tk.X)

        self.prev_btn = tk.Button(
            nav_row,
            text="◄ Prev",
            font=("Segoe UI", 10, "bold"),
            bg="#334155",
            fg="#f8fafc",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            pady=6,
            cursor="hand2",
            command=self.prev_item,
        )
        self.prev_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))

        self.next_btn = tk.Button(
            nav_row,
            text="Next ►",
            font=("Segoe UI", 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            pady=6,
            cursor="hand2",
            command=self.next_item,
        )
        self.next_btn.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(4, 0))

        # --- BOTTOM INTERACTIVE KEYBOARD GRID ---
        grid_section = tk.Frame(self.studio, bg="#0f172a", padx=25, pady=20)
        grid_section.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            grid_section,
            text="QUICK SELECT LETTER:",
            font=("Segoe UI", 9, "bold"),
            fg="#64748b",
            bg="#0f172a",
        ).pack(anchor="w", pady=(0, 8))

        self.grid_container = tk.Frame(grid_section, bg="#0f172a")
        self.grid_container.pack(fill=tk.BOTH, expand=True)

    # --- BOTTOM INTERACTIVE KEYBOARD GRID ---
    def rebuild_letter_grid(self):
        """Creates quick clickable letter chips at the bottom"""
        for btn in self.grid_buttons:
            btn.destroy()
        self.grid_buttons.clear()

        for idx, item in enumerate(self.current_list):
            is_active = idx == self.current_index
            bg_color = "#0d9488" if is_active else "#1e293b"
            fg_color = "#ffffff" if is_active else "#94a3b8"

            btn = tk.Button(
                self.grid_container,
                text=item.name,
                font=("Segoe UI", 10, "bold"),
                bg=bg_color,
                fg=fg_color,
                activebackground="#0f766e",
                activeforeground="#ffffff",
                relief=tk.FLAT,
                bd=0,
                width=4,
                pady=4,
                cursor="hand2",
                command=lambda i=idx: self.jump_to_index(i),
            )
            # Arrange in rows of 10
            row = idx // 10
            col = idx % 10
            btn.grid(row=row, column=col, padx=3, pady=3)
            self.grid_buttons.append(btn)

    # --- SIDEBAR BUTTONS ---
    def update_sidebar_styles(self):
        for key, btn in self.sidebar_buttons.items():
            if key == self.current_cat_key:
                btn.config(bg="#0d9488", fg="#ffffff")
            else:
                btn.config(bg="#1e293b", fg="#94a3b8")


# --- SIDEBAR BUTTONS ---
    def switch_category(self, category_key):
        self.current_cat_key = category_key
        self.current_list = CATEGORIES[category_key][1]
        self.current_index = 0
        self.load_item()

    #  --- SIDEBAR BUTTONS ---
    def jump_to_index(self, index):
        self.current_index = index
        self.load_item()

    # --- SIDEBAR BUTTONS ---
    def load_item(self):
        if not self.current_list:
            return

        item = self.current_list[self.current_index]

        # Update Header & Text Details
        cat_label = CATEGORIES[self.current_cat_key][0]
        self.cat_title.config(text=cat_label.upper())
        self.counter_badge.config(
            text=f"{self.current_index + 1} / {len(self.current_list)}"
        )

        self.letter_lbl.config(text=item.name)
        self.sound_lbl.config(text=f"Phonetic: {item.sound}")
        self.example_lbl.config(text=f"Example: {item.example}")

        self.update_sidebar_styles()
        self.rebuild_letter_grid()

        # Load Image
        if os.path.exists(item.image_path):
            img = Image.open(item.image_path)
            img = img.resize((310, 310), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.img_label.config(image=photo, text="")
            self.img_label.image = photo
        else:
            self.img_label.config(
                image="",
                text=f"Image Missing:\n{os.path.basename(item.image_path)}",
                fg="#f87171",
                font=("Segoe UI", 10, "bold"),
            )

    # --- Audio BUTTONS ---
    def play_tawar(self):
        item = self.current_list[self.current_index]

        if not os.path.exists(item.audio_path):
            messagebox.showinfo(
                "Tawar Audio",
                f"Audio file missing:\n{os.path.basename(item.audio_path)}",
            )
            return

        if self.player is not None:
            try:
                self.player.URL = os.path.abspath(item.audio_path)
                self.player.controls.play()
            except Exception as e:
                messagebox.showerror(
                    "Audio Error",
                    f"Unable to play file:\n{os.path.basename(item.audio_path)}\n\nDetails: {e}",
                )
        else:
            messagebox.showerror(
                "Module Missing",
                "pywin32 package is not installed.\nRun: pip install pywin32",
            )

    # --- next BUTTONS ---
    def next_item(self):
        if self.current_index < len(self.current_list) - 1:
            self.current_index += 1
            self.load_item()

    # --- prev BUTTONS ---
    def prev_item(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.load_item()


# Main Application
# __name__ is a special Python variable.
#  Python automatically creates it for every .py file.
if __name__ == "__main__":
    root = tk.Tk() # Create main window 
    app = BalochiApp(root) # Create application object
    # It starts the GUI event loop.
    # It will run until the window is closed. 
    root.mainloop() # Start the main loop
