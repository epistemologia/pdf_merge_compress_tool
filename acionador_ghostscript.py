import yaml
import subprocess
import os
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime


# ----------------------------------------------------------
# Logging utilities
# ----------------------------------------------------------
def init_logging():
    """Ensure /logs exists and create a timestamped log file."""
    log_dir = "logs"
    os.makedirs(log_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = os.path.join(log_dir, f"log_{timestamp}.txt")

    return log_file


def log_write(log_file, text):
    """Append text to the log file."""
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(text + "\n")


# ----------------------------------------------------------
# Helper – run a Ghostscript command and capture result
# ----------------------------------------------------------
def run_gs(cmd, log_file):
    log_write(log_file, "\n--- Running Ghostscript Command ---")
    log_write(log_file, "Command:")
    for c in cmd:
        log_write(log_file, f"  {c}")

    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        log_write(log_file, "\nSTDOUT:")
        log_write(log_file, result.stdout.strip())

        log_write(log_file, "\nSTDERR:")
        log_write(log_file, result.stderr.strip())

        if result.returncode != 0:
            log_write(log_file, f"\nERROR: Exit code {result.returncode}")
            return False, result.stderr

        log_write(log_file, "\nSUCCESS (exit code 0)\n")
        return True, ""

    except Exception as e:
        log_write(log_file, f"\nEXCEPTION: {str(e)}")
        return False, str(e)


# ----------------------------------------------------------
# Load YAML config
# ----------------------------------------------------------
def load_config():
    yaml_path = filedialog.askopenfilename(
        title="Selecione o arquivo config.yaml",
        filetypes=[("YAML Files", "*.yaml *.yml")]
    )
    if not yaml_path:
        messagebox.showerror("Erro", "Nenhum arquivo YAML selecionado.")
        return None

    try:
        with open(yaml_path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    except Exception as e:
        messagebox.showerror("Erro ao ler YAML", str(e))
        return None


# ----------------------------------------------------------
# MAIN PROCESS: merge → compress → pdfa
# ----------------------------------------------------------
def process():
    config = load_config()
    if not config:
        return

    # Start logging
    log_file = init_logging()
    log_write(log_file, "=== PDF Processing Log Started ===")
    log_write(log_file, f"Timestamp: {datetime.now()}")
    log_write(log_file, "----------------------------------\n")

    gs = config["caminhos"]["ghostscript"]
    log_write(log_file, f"Ghostscript path: {gs}")

    # Find PDFA_def.ps
    pdfa_def = os.path.abspath(os.path.join(
        os.path.dirname(gs), "..", "..", "lib", "PDFA_def.ps"
    ))
    log_write(log_file, f"PDF/A definition file: {pdfa_def}")

    if not os.path.exists(pdfa_def):
        msg = f"PDFA_def.ps não encontrado em:\n{pdfa_def}"
        log_write(log_file, msg)
        messagebox.showerror("Erro", msg)
        return

    messages = []

    # ------------------------------------------------------
    # 1. MERGE
    # ------------------------------------------------------
    if config.get("juntar", {}).get("ativado", False):
        merge_files = config["juntar"].get("arquivos", [])
        log_write(log_file, "Estágio de junção ativado.")

        if not merge_files:
            merge_files = filedialog.askopenfilenames(
                title="Selecione os PDFs para juntar"
            )
            merge_files = list(merge_files)

        log_write(log_file, f"Arquivos de entrada para junção: {merge_files}")

        if not merge_files:
            msg = "Nenhum arquivo para juntar."
            log_write(log_file, msg)
            messagebox.showerror("Erro", msg)
            return

        saida_juncao = config["juntar"]["saida"]

    cmd = [
        gs,
        "-dBATCH",
        "-dNOPAUSE",
        "-sDEVICE=pdfwrite",
        f"-sOutputFile={saida_juncao}",
    ]

    # Accept both simple list of strings and list of dicts with page ranges
    for item in merge_files:
        if isinstance(item, str):
            # Old behaviour: whole file
            cmd.append(item)

        elif isinstance(item, dict):
            caminho = item.get("arquivo")
            if not caminho:
                continue

            pagina_inicial = item.get("pagina_inicial")
            pagina_final = item.get("pagina_final")

            # Page range settings apply only immediately before the file
            if pagina_inicial:
                cmd.append(f"-dFirstPage={pagina_inicial}")
            if pagina_final:
                cmd.append(f"-dLastPage={pagina_final}")

            cmd.append(caminho)
     
        else:
            log_write(log_file, f"Entrada não reconhecida em 'arquivos': {item}")

        ok, err = run_gs(cmd, log_file)
        if not ok:
            messagebox.showerror("Erro na junção ", err)
            return

        messages.append(f"Junção OK → {saida_juncao}")
        input_for_next_step = saida_juncao

    # ------------------------------------------------------
    # 2. COMPRESS
    # ------------------------------------------------------
    if config.get("compactar", {}).get("ativado", False):
        log_write(log_file, "Estágio de compressão ativado.")

        if input_for_next_step is None:
            messagebox.showerror("Erro", "Nenhum PDF para compactação.")
            return

        output_compress = config["compactar"]["saida"]
        level = config["parametros"]["compactacao"]

        log_write(log_file, f"Nível de compactação: {level}")

        cmd = [
            gs,
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            f"-dPDFSETTINGS=/{level}",
            f"-sOutputFile={output_compress}",
            input_for_next_step
        ]

        ok, err = run_gs(cmd, log_file)
        if not ok:
            messagebox.showerror("Erro na Compactação", err)
            return

        messages.append(f"Compressão OK → {output_compress}")
        input_for_next_step = output_compress

    # ------------------------------------------------------
    # 3. PDF/A
    # ------------------------------------------------------
    if config.get("pdfa", {}).get("ativado", False):
        log_write(log_file, "Estágio de criação de PDF/A ativado.")

        if input_for_next_step is None:
            messagebox.showerror("Erro", "Nenhum PDF para PDF/A.")
            return

        output_pdfa = config["pdfa"]["saida"]
        profile = config["pdfa"].get("perfil", "PDF/A-2b")

        log_write(log_file, f"PDF/A profile: {profile}")

        cmd = [
            gs,
            "-dPDFA=2",
            "-sDEVICE=pdfwrite",
            "-dNOPAUSE", "-dBATCH", "-dNOOUTERSAVE",
            "-sColorConversionStrategy=RGB",
            "-sProcessColorModel=DeviceRGB",
            f"-sOutputFile={output_pdfa}",
            "-dPDFACompatibilityPolicy=1",
            pdfa_def,
            input_for_next_step
        ]

        ok, err = run_gs(cmd, log_file)
        if not ok:
            messagebox.showerror("Erro PDF/A", err)
            return

        messages.append(f"PDF/A OK → {output_pdfa}")

    # ------------------------------------------------------
    # DONE
    # ------------------------------------------------------
    log_write(log_file, "\n=== Process Finished ===\n")
    messagebox.showinfo("Concluído", "\n".join(messages))


# ----------------------------------------------------------
# GUI window
# ----------------------------------------------------------
root = tk.Tk()
root.title("Ferramenta PDF / Junção / Compactação / PDF-A")
root.geometry("480x160")

tk.Label(root, text="Selecione um arquivo YAML e execute o processo.", pady=10).pack()

btn = tk.Button(root, text="Executar", command=process, width=20)
btn.pack(pady=20)

root.mainloop()
