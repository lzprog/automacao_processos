from pathlib import Path
import shutil
from datetime import datetime
import time
from ftplib import FTP


FTP_HOST = "integra01.acom.net.br"
FTP_USER = "sisintegravamo30"
FTP_PASS = "Hya23ae#21wqQ"


class Fiscal:
    def __init__(self, file_paths, ftp_paths, log_path="log_fiscal.txt"):
        self.file_paths = file_paths
        self.ftp_paths = ftp_paths
        self.log_file = Path(log_path)

        self.log("==== INÍCIO DA EXECUÇÃO ====")

        self.sys_date = datetime.now()
        self.days = self.set_days()

        self.ftp_host = FTP_HOST
        self.ftp_user = FTP_USER
        self.ftp_pass = FTP_PASS
        self.ftp = None

    def log(self, mensagem):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.log_file.open("a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {mensagem}\n")

    def ftp_con(self):
        try:
            self.ftp = FTP(timeout=30)
            self.ftp.connect(self.ftp_host, 21)  
            self.ftp.set_pasv(True)
            self.ftp.login(user=self.ftp_user, passwd=self.ftp_pass)
            msg = f"Conectado ao FTP {self.ftp_host}"
            self.log(msg)
            print(msg)
        except Exception as e:
            msg = f"Erro ao conectar ao FTP: {e}"
            self.log(msg)
            print(msg)
            raise

    def set_days(self):
        sys_date = self.sys_date
        days = [(sys_date.day - i) for i in range(1, 5)]
        days = [d for d in days if d > 0]
        print(f"Dias analisados: {days}")
        return days

    def copiar_xmls(self):
        for nome_loja, caminho in self.file_paths.items():
            nome_loja_upper = nome_loja.upper()
            novos_copiados = 0
            ja_existia = 0

            try:
                origem_1 = Path(caminho["input_1"])
                origem_2 = Path(caminho["input_2"])
                destino = Path(caminho["output"])
                destino.mkdir(parents=True, exist_ok=True)

                arquivos_xml = [
                    arq for arq in list(origem_1.glob("*.xml")) + list(origem_2.glob("*.xml"))
                    if datetime.fromtimestamp(arq.stat().st_mtime).day in self.days
                ]
                total = len(arquivos_xml)
                if total == 0:
                    msg = f"[{nome_loja_upper}] Nenhum arquivo XML encontrado na origem para os dias {self.days}."
                    print(msg)
                    self.log(msg)
                    continue

                for arquivo in arquivos_xml:
                    destino_arquivo = destino / arquivo.name

                    if destino_arquivo.exists():
                        ja_existia += 1
                        msg = f"[{nome_loja_upper}] Já existe: {arquivo.name}"
                        print(msg)
                        self.log(msg)
                        continue

                    try:
                        shutil.copy2(arquivo, destino_arquivo)
                        novos_copiados += 1
                        msg = f"[{nome_loja_upper}] Copiado: {arquivo.name}"
                        print(msg)
                        self.log(msg)
                    except Exception as e:
                        msg = f"[{nome_loja_upper}] ERRO ao copiar {arquivo.name}: {e}"
                        print(msg)
                        self.log(msg)

                total_destino = novos_copiados + ja_existia
                if total_destino < total:
                    faltando = total - total_destino
                    msg = f"[{nome_loja_upper}] ATENÇÃO: {faltando} arquivos não foram copiados. Tentando novamente..."
                    print(msg)
                    self.log(msg)

                    time.sleep(3)

                    segunda_tentativa = 0
                    for arquivo in arquivos_xml:
                        destino_arquivo = destino / arquivo.name
                        if destino_arquivo.exists():
                            continue

                        try:
                            shutil.copy2(arquivo, destino_arquivo)
                            segunda_tentativa += 1
                            msg = f"[{nome_loja_upper}] (2ª tentativa) Copiado: {arquivo.name}"
                            print(msg)
                            self.log(msg)
                        except Exception as e:
                            msg = f"[{nome_loja_upper}] ERRO na 2ª tentativa ao copiar {arquivo.name}: {e}"
                            print(msg)
                            self.log(msg)

                    total_pos_tentativa = novos_copiados + ja_existia + segunda_tentativa
                    if total_pos_tentativa < total:
                        msg = f"[{nome_loja_upper}] ERRO: Ainda faltam {total - total_pos_tentativa} arquivos após 2 tentativas."
                    else:
                        msg = f"[{nome_loja_upper}] Sucesso: Todos os {total} arquivos copiados após 2 tentativas."
                else:
                    msg = f"[{nome_loja_upper}] Todos os {total} arquivos copiados (ou já estavam no destino)."
                print(msg)
                self.log(msg)

            except Exception as e:
                msg = f"[{nome_loja_upper}] ERRO ao acessar a pasta de origem: {caminho['input_1']} - {e}"
                print(msg)
                self.log(msg)
                continue

        self.log("==== FIM DA EXECUÇÃO ====\n")

    def enviar_ftp(self, arquivo_local, arquivo_remoto, nome_loja_upper):
        try:
            with open(arquivo_local, "rb") as f:
                self.ftp.storbinary(f"STOR {arquivo_remoto}", f)
            self.log(f"Enviado via FTP[{nome_loja_upper}]: {arquivo_remoto}")
            print(f"Enviado via FTP[{nome_loja_upper}]: {arquivo_remoto}")
            return True
        except Exception as e:
            self.log(f"Erro ao enviar[{nome_loja_upper}] -> {arquivo_local}: {e}")
            print(f"Erro ao enviar[{nome_loja_upper}] -> {arquivo_local}: {e}")
            return False

    def ftp_upload(self):
        self.ftp_con()

        for nome_loja, caminho in self.ftp_paths.items():
            nome_loja_upper = nome_loja.upper()
            enviados = 0

            try:
                origem_1 = Path(caminho["input_1"])
                destino = caminho["output"]

                try:
                    self.ftp.cwd(destino)
                except:
                    msg = f"[{nome_loja_upper}] - Criando pasta_destino in FTP..."
                    self.log(msg)
                    print(msg)
                    try:
                        self.ftp.mkd(destino)
                        self.ftp.cwd(destino)
                    except Exception as e:
                        msg = f"[{nome_loja_upper}] ERRO ao criar diretório {destino} no FTP: {e}"
                        print(msg)
                        self.log(msg)
                        continue

                arquivos = list(origem_1.glob("*"))
                total_origem_1 = len(arquivos)

                if total_origem_1 == 0:
                    msg = f"[{nome_loja_upper}] Nenhum arquivo na pasta local."
                    print(msg)
                    self.log(msg)
                    continue

                for arquivo in arquivos:
                    if self.enviar_ftp(arquivo, arquivo.name, nome_loja_upper):
                        enviados += 1

                if enviados == total_origem_1:
                    msg = f"[{nome_loja_upper}] Todos os arquivos enviados com sucesso!"
                else:
                    msg = f"[{nome_loja_upper}] ATENÇÃO: Apenas {enviados}/{total_origem_1} arquivos enviados."
                print(msg)
                self.log(msg)

            except Exception as e:
                msg = f"[{nome_loja_upper}] ERRO AO ENVIAR: {e}"
                print(msg)
                self.log(msg)


file_paths = {
    "zz": {
        "input_1": r"G:\Rest_15\NFCe\DANFE_XML\202509",
        "input_2": r"G:\Rest_15\NFe\DANFE_XML\34559709000143\202509",
        "output": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\ZZ"
    },
    "sdu": {
        "input_1": r"P:\Rest_15\NFCe\DANFE_XML\202509",
        "input_2": r"P:\Rest_15\NFe\DANFE_XML\23565552000142\202509",
        "output": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\SDU"
    },
    "lagoa": {
        "input_1": r"T:\Rest_15\NFCe\DANFE_XML\202509",
        "input_2": r"T:\Rest_15\NFe\DANFE_XML\47366757000167\202509",
        "output": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\LAGOA"
    },
    "mt": {
        "input_1": r"U:\Rest_15\NFCe\DANFE_XML\202509",
        "input_2": r"U:\Rest_15\NFe\DANFE_XML\34707607000128\202509",
        "output": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\METROPOLITANO"
    },
    "dw": {
        "input_1": r"W:\Rest_15\NFCe\DANFE_XML\202509",
        "input_2": r"W:\Rest_15\NFe\DANFE_XML\29496975000134\202509",
        "output": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\DOWNTOWN"
    }
}

ftp_paths = {
    "zz": {
        "input_1": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\ZZ",
        "output": r"/LSOARES/ZZ"
    },
    "sdu": {
        "input_1": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\SDU",
        "output": r"/LSOARES/SDU"
    },
    "lagoa": {
        "input_1": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\LAGOA",
        "output": r"/LSOARES/LAGOA"
    },
    "mt": {
        "input_1": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\METROPOLITANO",
        "output": r"/LSOARES/METROPOLITANO"
    },
    "dw": {
        "input_1": r"C:\Users\ZZ Pizza\Desktop\NEWPASTE\DOWNTOWN",
        "output": r"/LSOARES/DOWNTOWN"
    }
}

fiscal = Fiscal(file_paths, ftp_paths)
fiscal.copiar_xmls()
fiscal.ftp_upload()
