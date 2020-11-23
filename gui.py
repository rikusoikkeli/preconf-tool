# -*- coding: utf-8 -*-
"""
Created on Sun Sep 13 10:58:54 2020

@author: rikus

"""

# IMPORTS
##############################################################################

from tkinter import *
from tkinter.font import Font
from tkinter import filedialog
from luokat import *
from väripaletit import *
import winsound

##############################################################################


# HELPER FUNCTIONS
##############################################################################

def set_proteus_date(r):
    """
    Proteus framen Radiobutton widgetit käyttävät tätä. Muuttuja r sisältää tiedon
    siitä, mikä nappi on valittuna. Funktio ottaa r arvon ja kopioi sen globaaliin
    muuttujaan proteus_date.
    """
    global proteus_date
    proteus_date = r


def set_proteus_save_path():
    """
    Asettaa GUI :n napilla valitun tallennuspolun Proteus-raportille.
    Polku tallennetaan ohjelman kotikansion tiedostoon nimeltä "tallennuspolku.txt"
    """
    global proteus_save_path
    proteus_save_path = filedialog.askdirectory()
    filehandle = open("tallennuspolku.txt", "w", encoding="utf-8-sig")
    filehandle.write(proteus_save_path + "\n")
    filehandle.close()
    proteus_frame() # päivittää proteus framen, että uusi path tulee näkyviin
    
    
def get_proteus_save_path():
    """
    Yrittää palauttaa ohjelman kotikansion tiedoston "tallennuspolku.txt" ensimmäisen rivin (str).
    Jos se ei onnistu, palauttaa None.
    """
    try:
        filename = "tallennuspolku.txt"
        filehandle = open(filename, "r", encoding="utf-8-sig")
        path = str(filehandle.readline()).replace("\n","").replace("\\", "/")
        if path[1] == ":":
            path = path[2:]
        if path[-1] != "/":
            path = path + "/"
        return path
    except:
        return None
    

def run_proteus_algorithm():
    """
    Ajaa algoritmin, joka generoi Proteus-raportin GUI :n napilla valitusta tiedostosta
    ja tallentaa sen levylle excelinä.
    """
    file_path = filedialog.askopenfilename() # haetaan proteus-raportin lokaatio levyllä
    proteus = Proteus(file_path, proteus_date) # luodaan siitä proteus-objekti
    
    # haetaan proteus-luokan perkaamat df-objektit
    file_PT_tilaukset = proteus.get_file_PT_tilaukset()
    file_PT_virheet = proteus.get_file_PT_virheet()
    file_edit = proteus.get_file_edit()
    
    # lisätään df-objektit samaan listaan
    list_of_df = []
    list_of_df.append(file_PT_tilaukset)
    list_of_df.append(file_PT_virheet)
    list_of_df.append(file_edit)
    
    proteus_save_path = get_proteus_save_path() # noudetaan tallennuslokaatio perattua raporttia varten
    date = proteus.get_date(proteus_date)
    time = proteus.get_time()

    # tallennetaan perattu raportti levylle
    filename = "työjono " + date + " " + time + ".xlsx"
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    if proteus_save_path != None:
        writer = pd.ExcelWriter(proteus_save_path+filename, engine='xlsxwriter')
        n = 1
        for df in list_of_df:
            try:
                # Convert the dataframe to an XlsxWriter Excel object.
                df.to_excel(writer, sheet_name="Sheet"+str(n))
            except:
                pass
            n += 1
        # Close the Pandas Excel writer and output the Excel file.
        try:
            writer.save()
            # Jos tallentaminen onnistuu, toistaa animaation ja soittaa äänen.
            save_success_label = ImageLabel(file_proteus_frame, bg=colour6)
            save_success_label.place(x=300, y=150)
            save_success_label.load("success_animation.gif")
            winsound.PlaySound("save_sound.wav", winsound.SND_ASYNC)
        except:
            pass


def run_sap_algorithm():
    """
    Ajaa algoritmin, joka generoi SAP-raportin GUI :n napilla valitusta tiedostosta
    ja kirjoittaa sen tekstikenttään. Ilmoittaa työrivien pohjalta ennustetun työajan.
    """
    file_path = filedialog.askopenfilename() # noudetaan file_path
    sap = Sap(file_path) # tehdään sap-objekti, joka tekee kaikki tarvittavat muotoilut
    times = sap.get_times() # pyydetään sap-objektilta halutut tiedot
    frequencies = sap.get_frequencies()
    text_widget.delete("1.0", END) # tyhjennetään tekstikenttä kirjoittamista varten
    for line in times: # kirjoitetaan haluttu data tekstikenttään
        text_widget.insert(END, line+"\n")


def run_teraterm_algorithm():
    """
    Ajaa algoritmin, joka ottaa konffit tekstikentästä teraterm_input_box,
    palauttaa sen teratermin lukemassa muodossa muuttujaan teraterm_output_text
    (str) ja kopioi muuttujan sisällön tekstikenttään teraterm_output_box.
    """
    input_text = teraterm_input_box.get("1.0", END)
    teraterm = Teraterm(input_text)
    teraterm_output_text = teraterm.get_konffit_teraterm()
    teraterm_output_box.delete("1.0", END)
    teraterm_output_box.insert(END, teraterm_output_text)
    
  
def hide_all_frames():
    """
    Jotta uuden freimin voi avata, pitää jo olemassa olevat freimit pakata pois.
    Tämä funktio hoitaa sen ja sitä käytetään uuden freimin avaamisen yhteydessä.
    """
    slaves = window.pack_slaves()
    for s in slaves:
        s.pack_forget()
    

def proteus_frame():
    """
    Avaa freimin, jossa on Proteus-raporttiin liittyvät työkalut.
    """
    hide_all_frames()
    file_proteus_frame.pack(fill="both", expand=1)
    proteus_save_button.place( # valitse tallennuskansio
            x=140, 
            y=20)
    proteus_run_button.place( # lataa
            x=140, 
            y=50)
    proteus_text_frame.place( # polku
            height=25, 
            width=395, 
            x=280,
            y=20)
    proteus_text_box.grid()
    
    proteus_text_box.delete("1.0", END)
    proteus_save_path = get_proteus_save_path()
    if proteus_save_path != None:
        proteus_text_box.insert(END, proteus_save_path)
    else:
        proteus_text_box.insert(END, "ei kansiota valittuna")
    

def sap_frame():
    """
    Avaa freimin, jossa on SAP-raporttiin littyvät työkalut.
    """
    hide_all_frames()
    file_sap_frame.pack(fill="both", expand=1)
    text_widget.grid(column=0, 
                     row=0, 
                     padx=180, 
                     pady=20)
    text_widget.delete("1.0", END)
    text_widget.insert(END, "Käytä painiketta ladataksesi ohjelmaan SAP :in\ntyöjonoraportti. Arvioitu työjonon kesto\nilmestyy tähän.")
    lataa_sap_button.grid(column=0, 
                          row=1)
    
    
def teraterm_frame():
    """
    Avaa freimin, jossa on konffien kääntämiseen liittyvät työkalut.
    """
    hide_all_frames()
    file_teraterm_frame.pack(fill="both", expand=1)
    teraterm_input_box.grid(row=0, 
                            column=0,
                            pady=10,
                            padx=10)
    teraterm_output_box.grid(row=0, 
                             column=2,
                             pady=10,
                             padx=10)
    teraterm_convert_button.grid(row=0, 
                                 column=1,
                                 padx=2)
    teraterm_input_box.delete("1.0", END)
    teraterm_input_box.insert(END, "Tähän voit kopioida yhteistyökumppanin\nlähettämät konffit, jotka\nohjelma muuttaa Teratermin\nlukemaan muotoon.")
    teraterm_output_box.delete("1.0", END)
    teraterm_output_box.insert(END, "Korjatut konffit ilmestyvät tähän.")

##############################################################################


# MAIN WINDOW
##############################################################################

window = Tk()
window.title("Preconf Tool 2.1")
window.geometry("700x800")

top_menu = Menu(window)
# different from other widgets, need to use config
# tells tkinter to use top_menu as the menu
window.config(menu=top_menu)

# create a menu item
file_menu = Menu(top_menu, tearoff=0) # tearoff poistaa oudon viivan pudotusvalikon alusta
file_menu.add_command(label="Tietoa työjonosta (virhetilaukset yms.)", command=proteus_frame)
file_menu.add_command(label="Työjonon pituus per esiasentaja", command=sap_frame)
file_menu.add_command(label="Konffityökalu", command=teraterm_frame)

top_menu.add_cascade(label="Valitse toiminto", menu=file_menu) # menu has to be cascaded instead of packed

my_font = Font(family="Comic Sans MS", size=9) # ohjelman käyttämä fontti

##############################################################################


# PROTEUS FRAME
##############################################################################

r = IntVar() # tkinterin metodi, joka seuraa reaaliajassa Radiobutton widgetiin valittua arvoa
proteus_date = r.get() # tkinter metodi, jolla noudetaan r :n viimeisin arvo

file_proteus_frame = Frame(window, bg=colour6)
Radiobutton(file_proteus_frame, 
            text="Tänään", 
            variable=r, 
            value=0,
            bg=colour6,
            font=my_font,
            command=lambda: set_proteus_date(r.get())).place(x=20, y=20)

Radiobutton(file_proteus_frame, 
            text="Huomenna", 
            variable=r, 
            value=1,
            bg=colour6,
            font=my_font,
            command=lambda: set_proteus_date(r.get())).place(x=20, y=40)

Radiobutton(file_proteus_frame, 
            text="Ylihuomenna", 
            variable=r, 
            value=2,
            bg=colour6,
            font=my_font,
            command=lambda: set_proteus_date(r.get())).place(x=20, y=60)

proteus_save_button = Button(file_proteus_frame, 
                             text="Valitse tallennuskansio", 
                             command=set_proteus_save_path,
                             bg=colour3,
                             font=my_font)

proteus_run_button = Button(file_proteus_frame, 
                            text="Lataa tiedosto", 
                            command=run_proteus_algorithm, 
                            width=75, 
                            height=2,
                            bg=colour4,
                            font=my_font)

# Koska Text widgetin muotoilu on itsessään hankalaa, pitää sille tehdä oma freimi, jonka
# muotoilu luonnollisesti muotoilee myös widgetiä.
proteus_text_frame = Frame(file_proteus_frame) # freimi
proteus_text_box = Text(proteus_text_frame, # freimin sisälle tuleva tekstilaatikko
                        bg=colour2, 
                        font=my_font, 
                        wrap=NONE)

##############################################################################


# SAP FRAME
##############################################################################

file_sap_frame = Frame(window, 
                       width=400, 
                       height=400, 
                       bg=colour6)

lataa_sap_button = Button(file_sap_frame, 
                          text="Lataa tiedosto", 
                          command=run_sap_algorithm, 
                          width=25, 
                          height=2,
                          bg=colour4,
                          font=my_font)

text_widget = Text(file_sap_frame, 
                   width=50, 
                   height=25,
                   bg=colour1,
                   font=my_font)

##############################################################################


# TERATERM FRAME
##############################################################################

file_teraterm_frame = Frame(window, 
                            width=400, 
                            height=400, 
                            bg=colour6) # taustan väri
teraterm_input_box = Text(file_teraterm_frame, 
                          height=45, 
                          width=42, 
                          bg=colour1, # vasemman laatikon väri
                          wrap=NONE, # asetetaan, jotta yksittäinen rivi ei jakaannu rumasti useammalle riville
                          font=my_font)
teraterm_output_box = Text(file_teraterm_frame, 
                           height=45, 
                           width=42, 
                           wrap=NONE, # asetetaan, jotta yksittäinen rivi ei jakaannu rumasti useammalle riville
                           bg=colour3, # oikean laatikon väri
                           font=my_font)
teraterm_convert_button = Button(file_teraterm_frame, 
                                 text="Käännä", 
                                 command=run_teraterm_algorithm,
                                 width=7,
                                 height=2,
                                 bg=colour4, # painikkeen väri
                                 font=my_font)

##############################################################################


# MAINLOOP
##############################################################################

proteus_frame() # aloitetaan loop tästä freimistä
window.mainloop()

##############################################################################


