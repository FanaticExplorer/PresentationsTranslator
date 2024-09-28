import customtkinter as ctk
from tkinter import filedialog, messagebox
from utils import translate_powerpoint
import threading
import os


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("700x300")
        self.title("PowerPoint Translator")
        ctk.set_default_color_theme("ctk_theme.json")

        # Configure grid layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # Input Section
        self.input_label = ctk.CTkLabel(self, text="Input PowerPoint File:")
        self.input_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.input_entry = ctk.CTkEntry(self)
        self.input_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        self.input_button = ctk.CTkButton(
            self, text="Browse", command=self.browse_input
        )
        self.input_button.grid(row=0, column=2, padx=5, pady=10, columnspan=2)

        # Output Section
        self.output_label = ctk.CTkLabel(self, text="Output PowerPoint File:")
        self.output_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.output_entry = ctk.CTkEntry(self)
        self.output_entry.grid(row=1, column=1, padx=5, pady=10, sticky="ew")

        self.output_button = ctk.CTkButton(
            self, text="Browse", command=self.browse_output
        )
        self.output_button.grid(row=1, column=2, padx=5, pady=10, columnspan=2)

        # Target Language Section
        self.language_label = ctk.CTkLabel(self, text="Target Language:")
        self.language_label.grid(
            row=2, column=0, padx=10, pady=10, sticky="w", columnspan=2
        )

        self.language_var = ctk.StringVar(value="uk")  # Default value
        self.language_dropdown = ctk.CTkOptionMenu(
            self, values=["uk", "ru", "en"], variable=self.language_var
        )
        self.language_dropdown.grid(
            row=2, column=1, padx=5, pady=10, columnspan=3, sticky="ew"
        )

        # Slider for Threads
        self.thread_label = ctk.CTkLabel(self, text="Number of Threads:")
        self.thread_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")

        self.thread_slider = ctk.CTkSlider(
            self, from_=1, to=30, command=self.update_thread_label
        )
        self.thread_slider.set(5)  # Default value
        self.thread_slider.grid(
            row=3, column=1, padx=5, pady=10, columnspan=2, sticky="ew"
        )

        self.thread_value_label = ctk.CTkLabel(self, text="5")
        self.thread_value_label.grid(row=3, column=3)

        # Translate Button
        self.translate_button = ctk.CTkButton(
            self, text="Translate", command=self.translate
        )
        self.translate_button.grid(
            row=4, column=0, columnspan=4, pady=10, sticky="ew", padx=10
        )

        self.credits = ctk.CTkLabel(
            self,
            text="Made by: @FanaticExplorer. Bugs are expected.",
            cursor="hand2",
            font=ctk.CTkFont("Courier New"),
        )
        self.credits.grid(row=5, column=1, padx=10, pady=10, sticky="ew")
        self.credits.bind(
            "<Button-1>", lambda e: os.startfile("https://t.me/FanaticExplorer")
        )

        self.warning = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont("Courier New"),
        )
        # self.warning.grid(row=5, column=1, padx=10, pady=10, sticky="ew")

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if file_path:
            self.input_entry.delete(0, ctk.END)
            self.input_entry.insert(0, file_path)

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pptx", filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if file_path:
            self.output_entry.delete(0, ctk.END)
            self.output_entry.insert(0, file_path)

    def update_thread_label(self, value):
        self.thread_value_label.configure(text=str(int(value)))

    def translate(self):
        input_file = self.input_entry.get()
        output_file = self.output_entry.get()
        dest_language = self.language_var.get()
        max_workers = int(self.thread_slider.get())

        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select both input and output files.")
            return

        # Disable the button and update the text
        self.translate_button.configure(state="disabled", text="In progress...")
        self.translate_button.update()

        def run_translation():
            try:
                translate_powerpoint(
                    input_file, output_file, dest_language, max_workers
                )
                messagebox.showinfo(
                    "Success", f"Translated PowerPoint saved as: {output_file}"
                )
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")
            finally:
                # Re-enable the button and reset the text once the translation is done
                self.translate_button.configure(state="normal", text="Translate")
                self.translate_button.update()

        # Start the translation in a new thread
        translation_thread = threading.Thread(target=run_translation)
        translation_thread.start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
