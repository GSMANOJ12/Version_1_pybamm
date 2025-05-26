import customtkinter as CTk

def on_tab_changed():
    selected_tab = tabview.get()
    if selected_tab == "Tab1":
        update_dynamic_widgets()
    else:
        hide_dynamic_widgets()

def update_dynamic_widgets():
    if checkbox_var1.get() == 1:
        extra_label1.grid(row=2, column=0, pady=5, sticky="w")
        extra_entry1.grid(row=3, column=0, pady=5)

def hide_dynamic_widgets():
    extra_label1.grid_forget()
    extra_entry1.grid_forget()

def on_checkbox_click():
    update_dynamic_widgets()

root = CTk.CTk()
root.geometry("500x500")

tabview = CTk.CTkTabview(root, command=on_tab_changed)
tabview.pack(padx=20, pady=20)

tab1 = tabview.add("Tab1")
tab2 = tabview.add("Tab2")

checkbox_var1 = CTk.IntVar()
checkbox1 = CTk.CTkCheckBox(tab1, text="Enable Option A", variable=checkbox_var1, command=on_checkbox_click)
checkbox1.grid(row=0, column=0, pady=10, sticky="w")

extra_label1 = CTk.CTkLabel(tab1, text="Option A - Enter value:")
extra_entry1 = CTk.CTkEntry(tab1)

root.mainloop()
