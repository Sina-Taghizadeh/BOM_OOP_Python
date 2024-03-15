# Graphical BOM Processor using Python OOP
# by Sina Taghizadeh & Mohammad Zolfaghari
# June 2019
authors = "Sina Taghizadeh, Mohammad Zolfaghari"
license = "GNU GPL Version 3.0"

import xlrd #A library for reading data and formatting information from Excel files
from tkinter import *  #For GUI
from tkinter import filedialog,ttk,messagebox 

#\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\        

class Goods:    #A Class for all Goods
    global goods
    goods={}
    global Goods_in_row
    Goods_in_row=[]
    global _BOM
    _BOM={}
    global G_list
    G_list=[]
    global G_dict
    G_dict=[]    
    
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\
    
    def __init__(self,name,ID,unit,LC,MC,OC,PC):
        self.name=name
        self.id=ID
        self.unit=unit
        self.LC=LC    #Labor Cost
        self.MC=MC    #Machine Cost
        self.OC=OC    #Other Cost
        self.PC=PC    #Purchase Cost
        
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\

    def chosse_rows_from_exel (first_row=1,last_row=20,sheet_number='Sheet1'):  #Creating a method to receive an Excel file and fill in the created lists and dictionaries

        try:
            first_row=int(entry1.get())  #Get first row of the Excel file from user
            last_row=int(entry2.get())   #Get last row of the Excel file from user
            
            if first_row>=last_row:  #An error occurred when the first input is larger than the last input
                messagebox.showerror('IndexError','please enter a corect number')
                Reset1()
            
            book=xlrd.open_workbook(file_name)
            sheet=book.sheet_by_name(entry0.get())  #Select a specific sheet in an Excel file based on user input
            
            for i in range(first_row,last_row):
                s=sheet.row_values(i)  #A list of rows in Excel stored in a variable named s
                Goods_in_row.append(s)  
                
                G_list.append(s[0])   #Append the name item to the G_list list.
                G_list.append(s[1])   #Append the level to the G_list list.
                G_dict.append(s[0])   #Append the name item to the G_dict list.
                G_dict.append(s[4])   #Append the Quantity of item to the G_dict list. 
                
            for i in range(len(Goods_in_row)):  # loop for Goods classe 
                Goods_in_row[i]=Goods(Goods_in_row[i][0],Goods_in_row[i][2],Goods_in_row[i][3],Goods_in_row[i][5],Goods_in_row[i][6],Goods_in_row[i][7],Goods_in_row[i][8]) #Create a list of goods that are present in the list Goods_in_row               
                goods[Goods_in_row[i].name]=Goods_in_row[i] #Create a list of instances of the Goods class in a dictionary called goods with the key being the name of the product and the value being a list of instances of the Goods class.

        except(IndexError,AttributeError):   #Error management with message display and program reset
            messagebox.showerror('IndexError','please enter a correct number')
            Reset1()
            return

        except xlrd.biffh.XLRDError:    #Error management with message display and program reset
            messagebox.showerror('FileEror','Can not find this sheet in this file')
            Reset1()
            return

        except ValueError:         #Error management with message display and program reset
            messagebox.showerror('ValueError','you enterd text,please enter number')
            Reset1()
            return
        
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\        

    def BOM_function():     #Method creation for the BOM class to extract goods with the good name itself
        for i in range(len(Goods_in_row)): #Function to create a BOM (Bill of Materials) from a list of items in the Goods_in_row list.
            Goods_in_row[i]=BOM(Goods_in_row[i].name,Goods_in_row[i].id,Goods_in_row[i].unit,Goods_in_row[i].LC,Goods_in_row[i].MC,Goods_in_row[i].OC,Goods_in_row[i].PC)
        for i in range(len(Goods_in_row)): #To create a dictionary of BOM class members in the _BOM dictionary with the key of the product name and the value of the list of Goods class members
            _BOM[Goods_in_row[i].name]=Goods_in_row[i]
        for keys in _BOM :     
            globals()[_BOM[keys].name]=_BOM[keys]

    
#\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\

class BOM(Goods):    #Class BOM : A class that inherits from the Goods class

    def children(self):                     #A method for finding children and displaying them 
        b=G_list.index(self.name)           #b=index of the name of the product
        x=G_list[b+1]                       #x=level of the good
        a=b
        Children={}
        g_dict=G_dict.copy()                #Create a copy of the list G_dict with the name g_dict
        del(g_dict[0:b])                    #Remove the first b characters from the index of the given good name
        for i in G_list[b+2:]:              #loop for finding children
            a+=1
            if i==x+1 :                     #if i(level) is one level greater than level of the good (i.e. it is a child)
                h=g_dict.index(G_list[a])   
                d=g_dict[h+1]               #child Quantity                  
                Children[G_list[a]]=float(d)  #Adding a child to the children dictionary          
            if i==x:                        #if level b reached to its own level then stop loop
                break            
        return(Children)             
    
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\

    def super_children(self):         #A method for finding all submatrices of a good and displaying them 
        b=G_list.index(self.name)     #index of the good
        x=G_list[b+1]                 #level of the good
        a=b
        h=[]
        g_dict=G_dict.copy()         
        k=[self.name,1]               
        h.append(k)                  
        del(g_dict[0:b])             
        for i in G_list[b+2:]:
            a+=1
            if type(i) is float:
                if i>x:               #if the level is greater than the good level, then it is the good subset.          
                    n=[G_list[a],g_dict[a-b+1]]  
                    h.append(n)   
                if i==x:               #if level b reached to its own level then stop loop
                    break        
        return(h)
    
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\        

    def Cost(self):                       #Method for all costs of goods
        j=self.super_children()           #Finding good subsets 
        lc=[]
        mc=[]
        oc=[]
        pc=[]
        cost={}
        for keys in j:                    #all of subsets with Quantity
            mm=globals()[keys[0]].LC      
            MM=mm*keys[1]                 #LC of all       
            lc.append(MM)       

        for keys in j:                    #all of subsets with Quantity
            mm=globals()[keys[0]].MC
            MM=mm*keys[1]                 #MC of all
            mc.append(MM)           

        for keys in j:                    #all of subsets with Quantity
            mm=globals()[keys[0]].OC
            MM=mm*keys[1]                 #OC of all 
            oc.append(MM)             
            
        for keys in j:                    #all of subsets with Quantity
            mm=globals()[keys[0]].PC
            MM=mm*keys[1]                 #PC of all
            pc.append(MM)            

        cost['LC cost']=sum(lc)          
        cost['MC cost']=sum(mc)          
        cost['OC cost']=sum(oc)          
        cost['PC cost']=sum(pc)           
        return(cost)   
        
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\   

    def Sum_Cost(self):                 #Method for summing all calculated lc,oc,pc,mc in cost
        v=self.Cost()
        b=[]
        for keys in v:
            b.append(v[keys])
        c=sum(b)                        #summing all calculated lc,oc,pc,mc in cost
        return(c)
        
#\\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\ \\

    def Purchase_Cost(self): #All PCs of the good
        return self.Cost()['PC cost']
        
    def Machine_Cost(self):        #All MCs of the good
        return self.Cost()['MC cost']+self.Cost()['LC cost']
        
    def Other_Cost(self):       #All OCs of the good
        return self.Cost()['OC cost']
        
#\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\//\\
        
def Object_in_class(Class):                #Function to show how many object and type of them in a class
        global cxyz
        cxyz=[]
        for keys in _BOM:    
            cxyz.append(keys)
        return cxyz
        
#[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[[start GUI]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]]

def main():      #function main for GUI (for reset button functionality built into the main function)
    global entry0  # for sheetname
    global entry1  # for first row
    global entry2  # for last row
    global Reset1
    
    def Reset1():   #reset in window 1
        win1.destroy()
        main()     
                
    def Reset2():   #reset in window 2
        win1.destroy() 
        
        goods.clear()   
        Goods_in_row.clear()
        _BOM.clear()
        G_list.clear()
        G_dict.clear()
        
        main()    
        return
        
    def choosing_excel():     #for select an excel file
        root=Tk()
        root.withdraw()
        global file_name
        file_name=filedialog.askopenfilename()
        
    
    win1=Tk()                   #Create first window
    win1.title('Welcome')     
    win1.geometry('550x450+450+200') #Size of first window
                                
    line1=Label(win1,text = 'Welcome!\n\nPlease choose your excel file location',font='BTitr') 
    line1.pack()
    
    btn1=Button(win1,bg='pink',text = 'Choose',command= choosing_excel)          #Create button
    btn1.pack(pady=20,ipadx=3,ipady=4)
    
    line2=Label(win1,text = 'Choose rows in excel',font='BTitr') 
    line2.pack()
    
    line0=Label(win1,text = 'Enter the sheetname')  #(entry0)
    line0.pack()
    
    entry0=Entry(win1)        #Get input for sheetname
    entry0.config(width=12) 
    entry0.pack()

    line3=Label(win1,text = 'Enter the first row') #(entry1)
    line3.pack()
    
    entry1=Entry(win1)        #Get input for firt row
    entry1.config(width=3) 
    entry1.pack()

    line4=Label(win1,text = 'Enter the last row') #(entry2)
    line4.pack()
    
    entry2=Entry(win1)        #Get input for last row
    entry2.config(width=3)
    entry2.pack()
    
    btn2=Button(win1,bg='yellow' ,text = 'Confirm',command=lambda :[Goods.chosse_rows_from_exel() , Goods.BOM_function(),sstart()]) 
    btn2.pack(pady=10) #Create a confirmation button to receive Excel document and Create BOM class and explode parts 

    def start():              #for second window
        win1.withdraw()       #Hide window 1
        win2=Toplevel()       #Create window 2 
        win2.title('BOM')      
        win2.geometry('800x400+450+200') 
        
        line1=Label(win2,text='Choose your goods',font='BTitr') 
        line1.pack()
        
        combo_box1=ttk.Combobox(win2,values= list(set(Object_in_class(BOM))))   #Create combobox1 for good selection
        combo_box1.pack(pady=10,ipadx=10)
        
        line2=Label(win2,text='Choose your wanted from combobox',font='BTitr') 
        line2.pack()
        
        combo_box2=ttk.Combobox(win2,values= ('Children(Just first children)', 'Cost', 'Purchase Costs(sum)', 'Machine Costs(sum)', 'Other Costs(sum)','Sum Costs','LC','MC','OC','PC','id','Unit'))   
        combo_box2.pack(pady=10,ipadx=40)  
        
        def show():   #Display output based on input in combobox1, combobox2
            
            if combo_box2.get()=='Children(Just first children)':
                line3.config(text=globals()[combo_box1.get()].children(),font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4
                

                
            if combo_box2.get()=='Cost':
                line3.config(text='LC cost='+str(globals()[combo_box1.get()].Cost()['LC cost']),font='BTitr') #display in line 1
                linemc.config(text='MC cost='+str(globals()[combo_box1.get()].Cost()['MC cost']),font='BTitr') #display in line 2
                lineoc.config(text='OC cost='+str(globals()[combo_box1.get()].Cost()['OC cost']),font='BTitr') #display in line 3
                linepc.config(text='PC cost='+str(globals()[combo_box1.get()].Cost()['PC cost']),font='BTitr') #display in line 4
                

            if combo_box2.get()=='Purchase Costs(sum)':
                line3.config(text=globals()[combo_box1.get()].Purchase_Cost(),font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='Machine Costs(sum)':
                line3.config(text=globals()[combo_box1.get()].Machine_Cost(),font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='Other Costs(sum)':
                line3.config(text=globals()[combo_box1.get()].Other_Cost(),font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='Sum Costs':
                line3.config(text=globals()[combo_box1.get()].Sum_Cost(),font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='LC':
                line3.config(text=globals()[combo_box1.get()].LC,font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='MC':
                line3.config(text=globals()[combo_box1.get()].MC,font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='OC':
                line3.config(text=globals()[combo_box1.get()].OC,font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='PC':
                line3.config(text=globals()[combo_box1.get()].PC,font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='id':
                line3.config(text=globals()[combo_box1.get()].id,font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

            if combo_box2.get()=='Unit':
                line3.config(text=globals()[combo_box1.get()].unit,font='BTitr')
                linemc.config(text='',font='BTitr') #clear line 2
                lineoc.config(text='',font='BTitr') #clear line 3
                linepc.config(text='',font='BTitr') #clear line 4

        btn1=Button(win2,text='R U N',bg='green',command=show) #create the RUN button in window 2 to execute the user request with the show function
        btn1.pack(ipadx=30,ipady=5,pady=10)
        
        line3=Label(win2) #Build lines for display in window 2
        line3.pack()
        linemc=Label(win2)
        linemc.pack()
        lineoc=Label(win2)
        lineoc.pack()
        linepc=Label(win2)
        linepc.pack()
        
        reset2=Button(win2,text='R E S E T',bg='red',command=Reset2)   
        reset2.pack(ipadx=3,ipady=4,side=LEFT)
        reset2.place(x=0,y=0)
        
    def sstart():               #Remove the 'Confirm' button and add a 'START' button in Window 1
        btn3=Button(win1,text='S T A R T',bg='green',command=start)    
        btn3.pack(ipadx=30,ipady=5,pady=10)
        btn2.destroy()
        
    reset1=Button(win1,text='R E S E T',bg='red',command=Reset1) 
    reset1.pack(ipadx=3,ipady=4,side=LEFT)
    reset1.place(x=0,y=0)
    
    mainloop() 
    
main() 
