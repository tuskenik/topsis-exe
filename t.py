import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess

def topsis(macierz, wagi, kierunki):
    n_wariantow,n_kryteriow=macierz.shape
#1
    macierz_norm=np.zeros((n_wariantow,n_kryteriow))
    for j in range(n_kryteriow):
        suma_kw=0
        for i in range(n_wariantow):
            suma_kw+=macierz[i,j]**2
        dzielnik=np.sqrt(suma_kw)
        for i in range(n_wariantow):
            macierz_norm[i,j]=macierz[i,j]/dzielnik if dzielnik !=0 else 1

#2
    macierz_wazona=np.zeros((n_wariantow, n_kryteriow))
    for j in range(n_kryteriow):
        waga_j=wagi[j]
        for i in range(n_wariantow):
            macierz_wazona[i,j]=macierz_norm[i,j]*waga_j

#3
    pis=[]
    nis=[]
    for j in range(n_kryteriow):
        kol=macierz_wazona[:,j]
        if kierunki[j]=="MAX":
            pis.append(max(kol))
            nis.append(min(kol))
        else:
            pis.append(min(kol))
            nis.append(max(kol))

#4
    d_plus=[]
    d_minus=[]
    for i in range(n_wariantow):
        d_plus.append(np.sqrt(sum((macierz_wazona[i, j]-pis[j])**2 for j in range(n_kryteriow))))
        d_minus.append(np.sqrt(sum((macierz_wazona[i, j]-nis[j])**2 for j in range(n_kryteriow))))

#5
    wyniki_ci=[]
    for i in range(n_wariantow):
        suma=d_plus[i]+d_minus[i]
        wyniki_ci.append(d_minus[i]/suma if suma !=0 else 1)

    return np.array(wyniki_ci)

class TopsisApilkacja:
    def __init__(self, root):
        self.root=root
        self.root.title("Analiza wrażliwości wag kryteriów w klasycznej metodzie TOPSIS")
        self.root.geometry("900x900")
        self.dane=None
        self.wagi=None
        self.zmienne_gui={}

        main_header=tk.Frame(root)
        main_header.pack(fill="x", pady=10)
        tk.Label(main_header, text="METODA TOPSIS", font=("Arial", 18, "bold"), fg="#2c3e50").pack()
        tk.Label(main_header, text="Analiza wrażliwości wag kryteriów", font=("Arial", 10, "italic")).pack()

        self.btn_load=tk.Button(root, text="WCZYTAJ PLIK EXCEL", command=self.wczytaj_plik,
                                  width=40, height=2, bg="#3498db", fg="black", font=("Arial", 10, "bold"))
        self.btn_load.pack(pady=(10, 0))
        self.file_label=tk.Label(root, text="", font=("Arial", 8), fg="#7f8c8d")
        self.file_label.pack(pady=(2, 5))

        self.btn_calc=tk.Button(root, text="OBLICZ RANKING BAZOWY", command=self.ranking_bazowy,
                                  bg="#27ae60", fg="black", height=2, width=40, font=("Arial", 10, "bold"))
        self.btn_calc.pack(pady=10)

        self.frame_step=tk.LabelFrame(root, text=" PARAMETRY ANALIZY WRAŻLIWOŚCI ", padx=10, pady=10)
        self.frame_step.pack(fill="x", padx=20, pady=5)
        self.step_var=tk.StringVar(value="0.05")
        tk.Label(self.frame_step, text="Krok zmiany wagi (Δw):").pack(side="left", padx=5)
        tk.Entry(self.frame_step, textvariable=self.step_var, width=8).pack(side="left")
        tk.Label(self.frame_step, text=" wpisz liczbę z przedziału (0;1) ", font=("Arial", 8, "italic"), fg="#555555").pack(side="left", padx=10)

        self.canvas_frame=tk.LabelFrame(root, text=" KRYTERIA I KIERUNEK OPTYMALIZACJI ", padx=10, pady=10)
        self.canvas_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.canvas=tk.Canvas(self.canvas_frame, highlightthickness=0)
        self.scrollbar=ttk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scroll_content=tk.Frame(self.canvas)
        self.scroll_content.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_content, anchor="nw", width=840)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

    def wczytaj_plik(self):
        sciezka=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not sciezka: return
        try:
            df_full=pd.read_excel(sciezka, index_col=0)
            if 'Wagi' not in df_full.index:
                messagebox.showerror("Błąd", "Brak wiersza 'Wagi'!")
                return
            pobrane_wagi=df_full.loc['Wagi'].values.astype(float)
            suma_wag=sum(pobrane_wagi)

            if any(w>1.0 for w in pobrane_wagi) or any(w<0 for w in pobrane_wagi):
                messagebox.showerror("Błąd danych", "Wagi w Excelu muszą być z zakresu [0, 1].")
                return

            if not np.isclose(suma_wag, 1.0, atol=0.001):
                messagebox.showerror("Błąd sumy", f"Suma wag wynosi {suma_wag:.4f}.\n"
                                                  "Musi wynosić dokładnie 1.0. Popraw plik.")
                return

            self.wagi=pobrane_wagi
            self.dane=df_full.drop('Wagi').select_dtypes(include=[np.number])
            self.file_label.config(text=f"Wczytano plik: {os.path.basename(sciezka)}")
            for element in self.scroll_content.winfo_children(): element.destroy()
            self.zmienne_gui={}

            tk.Label(self.scroll_content, text="NAZWA KRYTERIUM", font=("Arial", 9, "bold")).grid(row=0, column=0, pady=10)
            tk.Label(self.scroll_content, text="KIERUNEK OPTYMALIZACJI", font=("Arial", 9, "bold")).grid(row=0, column=1, pady=10)
            tk.Label(self.scroll_content, text="ANALIZA", font=("Arial", 9, "bold")).grid(row=0, column=2, pady=10)

            for i, kol in enumerate(self.dane.columns):
                idx=i+1
                tk.Label(self.scroll_content, text=kol, anchor="w").grid(row=idx, column=0, padx=20, sticky="w")
                var=tk.StringVar(value="MAX")
                tk.OptionMenu(self.scroll_content, var, "MAX", "MIN").grid(row=idx, column=1)
                self.zmienne_gui[kol] = var
                tk.Button(self.scroll_content, text="Analiza wrażliwości", width=20, bg="#ecf0f1",
                          command=lambda c=kol: self.analiza_wrazliwosci(c)).grid(row=idx, column=2, pady=5)
            messagebox.showinfo("OK", "Pomyślnie wczytano plik.")
        except Exception as e: messagebox.showerror("Błąd", str(e))

    def ranking_bazowy(self):
        if self.dane is None:
            messagebox.showwarning("Uwaga", "Najpierw wczytaj plik Excel.")
            return
        macierz=self.dane.values.astype(float)
        kierunki=[self.zmienne_gui[c].get() for c in self.dane.columns]
        wyniki_ci=topsis(macierz, self.wagi, kierunki)
        ranking=pd.DataFrame({"Wariant": self.dane.index, "Ci": wyniki_ci}).sort_values(by="Ci", ascending=False)

        res_window=tk.Toplevel(self.root)
        res_window.title("Wyniki Rankingu Bazowego")
        res_window.geometry("550x450")
        tk.Label(res_window, text="WYNIKI RANKINGU BAZOWEGO", font=("Arial", 12, "bold"), pady=10).pack()

        frame=tk.Frame(res_window); frame.pack(padx=20, pady=10, fill="both", expand=True)
        tree=ttk.Treeview(frame, columns=("pos", "name", "score"), show="headings", height=12)
        tree.heading("pos", text="Miejsce"); tree.heading("name", text="Wariant"); tree.heading("score", text="Współczynnik bliskości")
        tree.column("pos", width=80, anchor="center"); tree.column("name", width=250, anchor="center"); tree.column("score", width=150, anchor="center")
        for i, r in enumerate(ranking.itertuples(), 1):
            tree.insert("", "end", values=(i, r.Wariant, f"{r.Ci:.4f}"))
        tree.pack(side="left", fill="both", expand=True)
        ttk.Scrollbar(frame, orient="vertical", command=tree.yview).pack(side="right", fill="y")
        tk.Button(res_window, text="ZAMKNIJ", command=res_window.destroy, width=20, bg="#e74c3c", fg="black").pack(pady=10)

    def analiza_wrazliwosci(self, kryterium):
        if self.dane is None: return
        try:
            krok=float(self.step_var.get().replace(',', '.'))
            if not (0<krok<1): raise ValueError
        except:
            messagebox.showerror("Błąd", "Krok zmiany wagi musi być liczbą z przedziału (0;1)!")
            return

        kolumny=list(self.dane.columns)
        j=kolumny.index(kryterium)
        macierz=self.dane.values.astype(float)
        kierunki=[self.zmienne_gui[c].get() for c in kolumny]

        wagi_bazowe=self.wagi/self.wagi.sum()
        waga_bazowa_j=wagi_bazowe[j]

        ci_bazowe=topsis(macierz, wagi_bazowe, kierunki)
        df_bazowy=pd.DataFrame({"W": self.dane.index, "C": ci_bazowe}).sort_values(by="C", ascending=False)
        rank_bazowy_lista=list(df_bazowy["W"])
        lider_bazowy=rank_bazowy_lista[0]

        raport=[]
        stabilny_ranking=[waga_bazowa_j]; stabilny_lider=[waga_bazowa_j]

        wiersz_bazowy={"Waga (wj)": f"BAZOWA: {round(waga_bazowa_j, 4)}"}
        for i, v in enumerate(rank_bazowy_lista, 1): wiersz_bazowy[f"Miejsce {i}"]=v
        wiersz_bazowy["Stabilność Rankingu"]="-"; wiersz_bazowy["Stabilność Lidera"]="-"
        raport.append(wiersz_bazowy)

        for wj_nowa in np.arange(0, 1.0001,krok):
            wj_nowa=round(float(wj_nowa),4)
            wj_nowa=min(1.0, max(0.0, wj_nowa))

            if abs(wj_nowa-waga_bazowa_j)<(krok/2): continue
            w_nowe=np.zeros(len(wagi_bazowe)); w_nowe[j]=wj_nowa
            suma_reszty=sum(wagi_bazowe[k] for k in range(len(wagi_bazowe)) if k!=j)
            for k in range(len(wagi_bazowe)):
                if k!=j: w_nowe[k]=(wagi_bazowe[k]/suma_reszty)*(1-wj_nowa) if suma_reszty>0 else (1-wj_nowa)/(len(wagi_bazowe)-1)

            ci_akt=topsis(macierz, w_nowe, kierunki)
            rank_akt=list(pd.DataFrame({"W": self.dane.index, "C": ci_akt}).sort_values(by="C", ascending=False)["W"])
            if rank_akt==rank_bazowy_lista: stabilny_ranking.append(wj_nowa)
            if rank_akt[0]==lider_bazowy: stabilny_lider.append(wj_nowa)

            wiersz={"Waga (wj)": wj_nowa}
            for pos, wariant in enumerate(rank_akt, 1): wiersz[f"Miejsce {pos}"]=wariant
            wiersz["Stabilność Rankingu"]="STABILNY" if rank_akt==rank_bazowy_lista else "ZMIANA"
            wiersz["Stabilność Lidera"]="STABILNY" if rank_akt[0]==lider_bazowy else "ZMIANA"
            raport.append(wiersz)

        sciezka_excel=os.path.join(os.path.expanduser("~"), "Desktop", f"Analiza_{kryterium}.xlsx")
        pd.DataFrame(raport).to_excel(sciezka_excel, index=False)

        okno=tk.Toplevel(self.root)
        okno.title("Zakresy Stabilności"); okno.geometry("450x350")
        tk.Label(okno, text=f"KRYTERIUM: {kryterium}", font=("Arial", 11, "bold"), pady=15).pack()
        tk.Label(okno, text="Zakres, w którym RANKING nie ulega zmianie:").pack()
        tk.Label(okno, text=f"[{min(stabilny_ranking):.2f} - {max(stabilny_ranking):.2f}]", font=("Arial", 12, "bold"), fg="blue").pack(pady=5)
        tk.Label(okno, text="Zakres, w którym WARIANT NAJLEPSZY nie ulega zmianie:").pack()
        tk.Label(okno, text=f"[{min(stabilny_lider):.2f} - {max(stabilny_lider):.2f}]", font=("Arial", 12, "bold"), fg="green").pack(pady=5)

        def otworz():
            if os.name=='nt': os.startfile(sciezka_excel)
            else: subprocess.run(["open", sciezka_excel])

        tk.Button(okno, text="ZOBACZ CAŁĄ ANALIZĘ (EXCEL)", command=otworz, width=30, bg="#ecf0f1").pack(pady=20)

if __name__=="__main__":
    okno_programu=tk.Tk()
    app=TopsisApilkacja(okno_programu)
    okno_programu.mainloop()