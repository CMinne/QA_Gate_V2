#from tkinter import *
#import tkinter as tk
#from tkinter import font

#class Passwordchecker(tk.Frame):
#    def __init__(self, parent):
#        tk.Frame.__init__(self, parent)
#        self.parent = parent
#        self.initialize_user_interface()

#    def initialize_user_interface(self):
#        ScreenHeight = int(self.parent.winfo_screenheight()) - 63
#        ScreenWidth = int(self.parent.winfo_screenwidth())

#        self.parent.geometry(str(ScreenWidth) + "x" + str(ScreenHeight)+'+-10+0')
#        self.parent.configure(background='#ffffff')
#        self.parent.state('zoomed')

#        self.label_font = font.Font(self.parent, family='{Courier New}', size=48, weight='normal')

#        font11 = "-family {Segoe UI} -size 48 -weight normal -slant "  \
#            "roman -underline 0 -overstrike 0"
#        font12 = "-family {Segoe UI} -size 24 -weight normal -slant "  \
#            "roman -underline 0 -overstrike 0"
#        font13 = "-family {Segoe UI} -size 36 -weight normal -slant "  \
#            "roman -underline 0 -overstrike 0"

#        self.frame = Frame(self.parent, background = 'blue', width = int(self.parent.winfo_screenwidth()), height = (int(self.parent.winfo_screenheight()) - 63))
#        self.frame.pack(fill = "both", expand=True)
#        self.frame.update()
#        self.UPframe = Frame(self.frame, background = '#ffffff', width = int(self.frame.winfo_screenwidth()), height = int((self.frame.winfo_height())/3))
#        self.UPframe.pack(fill = "both", expand=True)
#        self.UPframe.update()
#        self.Midframe = Frame(self.frame, background = '#454545', width = int(self.frame.winfo_screenwidth()), height = int((self.frame.winfo_height())/3))
#        self.Midframe.pack(fill = "both", expand=True)
#        self.Bottomframe = Frame(self.frame, background = '#000000', width = int(self.frame.winfo_screenwidth()), height = int((self.frame.winfo_height())/3))
#        self.Bottomframe.pack(fill = "both", expand=True)


#        self.LabelText = tk.Label(self.UPframe)
#        self.LabelText.place(relx=0.5, rely=0.05, height = int(0.18*(self.UPframe['height'])), width=int(0.5*(self.Midframe['width'])), anchor = 'center')
#        self.LabelText.configure(background="#f4f4f4")
#        self.LabelText.configure(font=self.label_font)
#        self.LabelText.configure(text='''Extraction données QA Gate 4.0''')

#        self.LabelReference = tk.Label(self.UPframe)
#        self.LabelReference.place(relx=0.5, rely=0.8, height=int(0.14*(self.Midframe['height'])), width=234, anchor = 'center')
#        self.LabelReference.configure(background="#f4f4f4")
#        self.LabelReference.configure(font="-family {Segoe UI} -size 36")
#        self.LabelReference.configure(text='''Reference''')

#        #self.LabelType = tk.Label(self.Midframe)
#        #self.LabelType.place(relx=0.251, rely=0.336, height=51, width=224, anchor = 'center')
#        #self.LabelType.configure(anchor='w')
#        #self.LabelType.configure(background="#f4f4f4")
#        #self.LabelType.configure(font="-family {Segoe UI} -size 24")
#        #self.LabelType.configure(text='''Type de pièce :''')

#        #self.varPG = tk.BooleanVar() 
#        #self.varPG.set(True)
#        #self.ButtonPG = tk.Checkbutton(self.Midframe, variable=self.varPG, indicatoron=0)
#        #self.ButtonPG.configure(bd = 2,font = fontButton, foreground = 'black', background = 'slate gray', command = self.handlerPG)
#        #self.ButtonPG.place(relx=0.37, rely=0.313, height=45, width=166)
#        #self.ButtonPG.configure(text='''PG''')

#        #self.varSG = tk.BooleanVar() 
#        #self.ButtonSG = tk.Checkbutton(self.Midframe, variable=self.varSG, indicatoron=0)
#        #self.ButtonSG.configure(bd = 2,font = fontButton, foreground = 'black', background = 'slate gray', command = self.handlerSG)
#        #self.ButtonSG.place(relx=0.542, rely=0.313, height=45, width=166)
#        #self.ButtonSG.configure(text='''SG''')
#        #self.ButtonSG['state']= tk.DISABLED
        

#        def d(event):
#            self.frame.configure(width = int(self.parent.winfo_width()), height = (int(self.parent.winfo_height()) - 63))
#            self.UPframe.configure(width = int(self.frame.winfo_width()), height = (int(self.frame.winfo_height()))/3)
#            self.Midframe.configure(width = int(self.frame.winfo_width()), height = (int(self.frame.winfo_height()))/3)
#            self.Bottomframe.configure(width = int(self.frame.winfo_width()), height = (int(self.frame.winfo_height()))/3)


#        def e(event):
#            self.LabelText.place(relx=0.5, rely=0.08, height = int(0.18*(self.UPframe['height'])),width=int(0.5*(self.Midframe['width'])), anchor = 'center')
#            self.LabelReference.place(relx=0.5, rely=0.8, height=int(0.14*(self.Midframe['height'])), width=234, anchor = 'center')
#            heightIni = self.LabelText.winfo_height()
#            widthIni = self.LabelText.winfo_width()

#            height = heightIni // 2
#            width = widthIni // 2

#            if height < 10 or width < 30:
#                height = 10
#            elif height < 20 or width < 60:
#                height = 20
#            elif height < 30 or width < 90:
#                height = 30
#            else:
#                height = 40
#            print(str(width))
#            self.label_font['size'] = height


#            height = heightIni // 2

#            if height < 10:
#                height = 8
#            elif height < 20:
#                height = 18
#            elif height < 30:
#                height = 24
#            else:
#                height = 36

#            self.LabelReference.configure(font="-family {Segoe UI} -size " + str(height))
            
#        self.UPframe.bind('<Configure>',e)
#        self.parent.bind('<Configure>',d)





#if __name__ == '__main__':

#    root = tk.Tk()
#    run = Passwordchecker(root)
#    root.mainloop()

from datetime import datetime
# datetime object containing current date and time
now = datetime.now()
 
print("now =", now)
# dd/mm/YY H:M:S
dt_string = now.strftime("%Y%m%d_%H%M%S")
print("date and time =", dt_string)
