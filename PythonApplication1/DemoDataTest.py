import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import time
import calendar
import datetime
from math import pi
from bokeh.core.properties import value
from bokeh.io import show, output_file
from bokeh.plotting import figure
from bokeh.models import PrintfTickFormatter


class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        '''Main Window AKA ROOT'''
        ROOT.geometry("650x250")
        ROOT.title("Excel Import tester")
        tk.Label(self, text="Import data", padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
        tk.Label(self, text="Import data", font=("Helvetica", 12)).pack(side="top")
        self.read_spread_sheet = tk.Button(self, font=("Helvetica", 16))
        self.read_spread_sheet["text"] = "Load a csv spreadsheet"
        self.read_spread_sheet["command"] = self.spread_sheet_window
        self.read_spread_sheet.pack()
        self.quit = tk.Button(self, text="QUIT", fg="red", command=ROOT.destroy, font=("Helvetica", 16))
        self.quit.pack(side="bottom")

    def spread_sheet_window(self):
        '''Load Spreadsheet Window '''
        ROOT.withdraw()
        spread_sheet_window = tk.Toplevel(self)
        data_state = tk.Label(spread_sheet_window, text="Loading  Data", padx=10, pady=10, font=("Helvetica", 16))
        data_state.pack()
        num_of_trys = 4
        for iteration in range(0, num_of_trys):
            try:
                file_path = askopenfilename()
                data = pd.read_csv(file_path, header=8)
                break
            except:
                data_state.configure(text="Sorry that did not work please load vaild data")
                data_state.update()
                if iteration == (num_of_trys-1):
                    data_state.configure(text="Sorry that failed " + str(iteration) + " times, closing")
                    data_state.update()
                    time.sleep(5)
                    ROOT.destroy()
                    quit(0)
        data_state.configure(text="Data Loaded")                
        data_state.update()
        self.choose_date_window(data)
        spread_sheet_window.destroy()
      
    def choose_date_window(self,data):
          choose_date_window = tk.Toplevel(self)
          tk.Label(choose_date_window, text="Data imported was", padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
          tk.Label(choose_date_window, text=str(data.iat[0,0])+" to the "+str(data.iat[-1,0]), padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
          tk.Label(choose_date_window, text="Which month  would like to search", padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
          month = tk.StringVar()
          day = tk.IntVar()
          day.set(1)
          months_choices = []
          for i in range(1,13):
              months_choices.append(calendar.month_name[i])
          month.set(months_choices[0])
          tk.OptionMenu(choose_date_window,month,*months_choices).pack(side="top")
          first_date = datetime.datetime.strptime(data.iat[0,0], '%d/%m/%Y')
          tk.Label(choose_date_window, text="Which day would like to search", padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
          tk.OptionMenu(choose_date_window,day,*range(1,32)).pack(side="top")
          tk.Button(choose_date_window, text="Submit", fg="red", command=lambda:[choose_date_window.destroy(), self.date_vaildiation(day,month,first_date.year,data)], font=("Helvetica", 16)).pack(side="bottom")

    def date_vaildiation(self,day,month,year,data):
        date_vaildiation = tk.Toplevel(self)
        tk.Label(date_vaildiation, text="Data Selectd was", padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
        
        tk.Label(date_vaildiation, text=[day.get(),"-",month.get(),"-",year]).pack(side="top")
        
        try:
            date = datetime.datetime.strptime(str(day.get())+str(month.get())+str(year),'%d%B%Y')
            
        except:
            tk.Label(date_vaildiation, text="Data is invaild", padx=10, pady=10, font=("Helvetica", 16)).pack(side="top")
            tk.Button(date_vaildiation, text="Return", fg="red", command=lambda:[date_vaildiation.destroy(), self.choose_date_window(data)], font=("Helvetica", 16)).pack(side="bottom")
        
        self.ploting_time(data,date)
        date_vaildiation.destroy()
    
    def ploting_time(self,data,date):
        sub_set = data.loc[data['Day'] == date.strftime('%d/%m/%Y').lstrip("0")]
        output_file("stacked.html")
        shows = sub_set['Program Name'].tolist()
        number_of=1
        for x in shows:
            if x in shows: 
                shows[shows.index(x)] = str(shows[shows.index(x)]) + str(number_of)
                number_of+= 1
        demos = list(sub_set.columns.values[6:11])     
        colors = ["#c9d9d3", "#718dbf", "#e84d60","#A9d9d3", "#218dbf"]
        p = figure(x_range=shows, title="Shows by Demo for "+str(date.strftime('%d/%m/%Y').lstrip("0")),
                   toolbar_location=None, tools="hover", tooltips="$name @shows: @$name")        
        y = {'shows':shows}
        for x in demos:
            temp = sub_set[x].tolist()
            temp2 = list()
            for z in range(len(temp)):
                temp2.append(int((temp[z].replace(',', ''))) )
            y[x] = temp2


        p.vbar_stack(demos, x='shows', width=0.9, color=colors, source=y,legend=[value(x) for x in demos])
        p.y_range.start = 0
        p.yaxis[0].formatter = PrintfTickFormatter(format="%5d")
        p.x_range.range_padding = 0.1
        p.xaxis.major_label_orientation = -pi/8
        p.xgrid.grid_line_color = None
        p.axis.minor_tick_line_color = None
        p.outline_line_color = None
        p.sizing_mode = 'stretch_both'
        p.legend.location = "top_left"
        p.legend.orientation = "horizontal"
        show(p)
        ROOT.destroy


ROOT = tk.Tk()
APP = Application(master=ROOT)
APP.mainloop()