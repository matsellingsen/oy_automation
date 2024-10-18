import tkinter as tk
root = tk.Tk() 
root.title("My Desktop Application")
label = tk.Label(root, text="Hello, World!") 
label.pack()  
button = tk.Button(root, text="Click Me", command=lambda: print("Button clicked!")) 
button.pack()
root.mainloop()