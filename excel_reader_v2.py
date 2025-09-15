import datetime
import pandas as pd

class Hora:
    def __init__(self, hours, mins):
        self.hours = int(hours)
        self.mins = int(mins)

    def hora_to_float(self):
        return self.hours + self.mins / 60.0

    def hora_to_string(self):
        return f"{self.hours}.{self.mins:02d}"


class ExcelReader:
    def __init__(self):
        self.path = r"C:\Users\ZZ Pizza\Desktop\excel_reader\RELATORIO.xlsx"
        self.faixa_1 = [(9.00, 9.30), (9.30, 10.00), (10.00, 10.30), (10.30, 11.00),
                        (11.00, 11.30), (11.30, 12.00)]
        self.faixa_2 = [(12.00, 13.00), (13.00, 14.00), (14.00, 15.00), (15.00, 16.00),
                        (16.00, 17.00), (17.00, 18.00), (18.00, 19.00), (19.00, 20.00),
                        (20.00, 21.00)]
        self.faixa_3 = [(21.00, 21.30), (21.30, 22.00), (22.00, 22.30), (22.30, 23.00), (23.00, 23.30), (23.30, 23.59)]
        self.traduzir = {'Sunday': 'Domingo', 'Monday': 'Segunda-feira', 'Tuesday': 'Terça-feira',
                         'Wednesday': 'Quarta-feira', 'Thursday': 'Quinta-feira',
                         'Friday': 'Sexta-feira', 'Saturday': 'Sábado'}

    def set_faixa(self, hora: Hora):
        hora_float = hora.hora_to_float()
        for inicio, fim in (self.faixa_1 + self.faixa_2 + self.faixa_3):
            if inicio <= hora_float < fim:

                inicio_h = int(inicio)
                inicio_m = int(round((inicio - inicio_h) * 60))
                fim_h = int(fim)
                fim_m = int(round((fim - fim_h) * 60))

                faixa_str = f"De {inicio_h:02d}:{inicio_m:02d} -> {fim_h:02d}:{fim_m:02d}"
                return faixa_str
        return None

    def traduzir_dia(self, day_of_week):

        day_br = self.traduzir.get(day_of_week)
        if day_br:
            return day_br
        else:
            print("Dia não encontrado!")
            return None

    def integrar_excel(self, df_new, dias, faixas):
        print("integrando...")
        df_new["D"] = dias
        df_new["Faixa"] = faixas
        df_new.to_excel(self.path, index=False)
        print("Integrado!")

    def reader(self):
        try:
            print(f"Lendo arquivo: {self.path}")
            df = pd.read_excel(self.path)
            df_new = df.dropna().copy()

            dias = []
            faixas = []

            column_a = df_new.iloc[:, 0]
            for data in column_a:
                if pd.isna(data):
                    dias.append(None)
                    continue
                new_date = pd.to_datetime(data, dayfirst=True)
                day_of_week = new_date.strftime("%A")
                day_br = self.traduzir_dia(day_of_week)
                dias.append(day_br)

            column_b = df_new.iloc[:, 1]
            for hora in column_b:
                if pd.isna(hora):
                    faixas.append(None)
                    continue
                h, m = str(hora).split(":")
                new_hora = Hora(h, m)
                faixa = self.set_faixa(new_hora)
                faixas.append(faixa)
                print(f"Hora {hora} -> {new_hora.hora_to_string()} caiu na faixa {faixa}")

        except Exception as e:
            print("ERRO!", e)


        self.integrar_excel(df_new, dias, faixas)




excel = ExcelReader()
excel.reader()
