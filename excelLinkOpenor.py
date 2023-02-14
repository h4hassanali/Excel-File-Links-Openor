import tkinter as tk
from tkinter import messagebox
import pandas as pd
import webbrowser
class UrlOpenor:
    def __init__(self, master):
        self.master = master
        self.master.title('Excel Link Opening Software')
        
        # Create labels and text fields
        self.input_label_0 = tk.Label(self.master,bg='#BCBCEE',text='Excel File Link Openor Software',font=('Times New Roman', 26))
        self.input_label_0.pack(padx=10, pady=10)
        self.input_label_1 = tk.Label(self.master,bg='#BCBCEE',text='\nEnter Excel File Location :',font=('Times New Roman', 18))
        self.input_label_1.pack(padx=10, pady=10)
        self.input_field_1 = tk.Entry(self.master,width=50,font=('Calibri', 15))
        self.input_field_1.pack(padx=10, pady=10)
        
        self.input_label_2 = tk.Label(self.master,bg='#BCBCEE',text='Enter Links Column Name :',font=('Times New Roman', 18))
        self.input_label_2.pack(padx=10, pady=10)
        self.input_field_2 = tk.Entry(self.master,width=50,font=('Times New Roman', 15))
        self.input_field_2.pack(padx=10, pady=10)
        
        # Create a button
        self.submit_button = tk.Button(self.master,bg='#FFFFFA', text='Open Links', command=self.display_text,width=22,height=3)
        self.submit_button.pack(padx=10, pady=10)
        #Acknowledgements
        self.input_label_3 = tk.Label(self.master,bg='#BCBCEE',text='App Developed by Hassan Ali\n All rights reserved.',font=('Times New Roman', 8))
        self.input_label_3.pack(padx=10, pady=10)
        
        # Create variables to store the entered data
        self.data_1 = tk.StringVar()
        self.data_2 = tk.StringVar()
        
    def display_text(self):
        # Get the entered data and store it in the variables
        self.data_1.set(self.input_field_1.get())
        self.data_2.set(self.input_field_2.get())
        excel_file_location = str(self.data_1.get())
        df = pd.read_excel(excel_file_location)
        links_column = str(self.data_2.get())
        column_data = df[links_column].tolist()
        for i in column_data:
             webbrowser.open(i)
        tk.messagebox.showinfo('Opening Excel Links', f'Location You Provided: {self.data_1.get()}\nColumn you provided : {self.data_2.get()}')

        
# Create the main window and start the event loop
root = tk.Tk()
root.geometry("500x500")
root.configure(bg='#BCBCEE')
# # to make the GUI dimensions not fixed
root.resizable(True, True) 
app = UrlOpenor(root)
root.mainloop()
