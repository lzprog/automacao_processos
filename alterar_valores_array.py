import fdb
import pandas as pd
import os


class AlterarValores:
    def __init__(self):
        self.firebird_path = r"C:\Program Files\Firebird\Firebird_3_0"
        if self.firebird_path not in os.environ['PATH']:
            os.environ['PATH'] += f";{self.firebird_path}"

        self.enderecos = {
            "dw": r"192.168.1.20:\grdj\Rest_15\Base_ib\VAMO_ZERADO.FDB",
            "mt": r"192.168.7.10:\grdj\Rest_15\Base_ib\VAMO_ABELARDO_BUENO.FDB",
            "sdu": r"192.168.6.20:\grdj\Rest_15\Base_ib\VAMO_SANTOS_DUMONT.FDB",
            "lagoa": r"192.168.3.10:\grdj\Rest_15\Base_ib\VAMO_NOVO.FDB"
        }
        self.excel_path = r"C:\Users\ZZ Pizza\Desktop\excel_reader\alterar.xlsx" 
        self.produtos = []

    def reader(self):
        try:
            file_path = self.excel_path
            print(f"Lendo arquivo: {file_path}")
            
            df = pd.read_excel(file_path, dtype={'Código PDV': str})
            df.dropna(how='all', inplace=True) 
            
            column_a = (df.iloc[:, 0])
            column_c = (df.iloc[:, 2])     
            self.produtos = list(zip(column_a, column_c))
        except Exception as e:
            print("Erro ao ler arquivo !")

    def conectar_banco(self, loja, endereco):
        try:
            con = fdb.connect(
                dsn=endereco,
                user='SYSDBA',
                password='475869',
                charset='ISO8859_2'
            )
            return con
        except fdb.fbcore.DatabaseError as e:
            print(f"\n*****Erro ao conectar ao banco {loja.upper()}*****")
            return None
        except Exception as e:
            print(f"\n*****Erro inesperado ao conectar ao banco {loja.upper()}*****")
            return None

    def procurar_item(self):
        for codigo, preco_padrao in self.produtos:
            for loja, endereco in self.enderecos.items():
                con = self.conectar_banco(loja, endereco)
                if not con:
                    continue
                try:
                    cur = con.cursor()
                    cur.execute("SELECT * FROM RST002 WHERE COD_RED = ?", [codigo])
                    results = cur.fetchall()
                    
                    if results:
                        columns = [desc[0] for desc in cur.description]
                        df = pd.DataFrame(results, columns=columns)
                        descricao = df.loc[df['COD_RED'] == codigo, 'DESC_RES'].values[0]
                        preco = df.loc[df['COD_RED'] == codigo, 'PRECO'].values[0]
                        print("==========================================================")
                        print(f"{loja.upper()} - {codigo} - {descricao}: R$ {preco}")
                    else:
                        print(f"{loja.upper()} - Produto {codigo} não encontrado.")
                    
                    cur.close()
                    con.close()
                except Exception as e:
                    print(f"*****Erro ao consultar {loja.upper()}*****")

        try:
            option = int(input("\nDeseja mudar preco em massa? \n1. sim\n0. nao: "))
            if option not in [0,1]:
                option = 0
        except ValueError:
            option = 0

        self.mudar_preco(option)

    def mudar_preco(self, option):
        if option != 0:
            for codigo, preco_novo in self.produtos:
                for loja, endereco in self.enderecos.items():
                    con = self.conectar_banco(loja, endereco)
                    if not con:
                        continue
                    try:
                        cur = con.cursor()
                        cur.execute("SELECT * FROM RST002 WHERE COD_RED = ?", [codigo])
                        results = cur.fetchall()
                        
                        if not results:
                            print(f"{loja.upper()} - Produto {codigo} não encontrado para atualizar.")
                            continue

                        columns = [desc[0] for desc in cur.description]
                        df = pd.DataFrame(results, columns=columns)
                        descricao = df.loc[df['COD_RED'] == codigo, 'DESC_RES'].values[0]
                        preco_antigo = df.loc[df['COD_RED'] == codigo, 'PRECO'].values[0]
                        
                        cur.execute("UPDATE RST002 SET PRECO = ? WHERE COD_RED = ?", [preco_novo, codigo])
                        con.commit()
                        cur.close()
                        con.close()
                        
                        print("==========================================================")
                        print(f"{loja.upper()} - Preço do produto {codigo} - {descricao} - R${preco_antigo} atualizado para ***R$ {preco_novo}***")
                    except Exception as e:
                        print(f"*****Erro ao atualizar {loja.upper()}*****")


if __name__ == "__main__":
    alteraval = AlterarValores()
    alteraval.reader()
    alteraval.procurar_item()  
