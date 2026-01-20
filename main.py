import argparse
import os
import time
from pathlib import Path
from typing import Optional

import pandas as pd
import matplotlib.pyplot as plt

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# =========================
# Utilidades
# =========================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera imagens por assessor com clientes em saldo negativo e envia via WhatsApp Web (Selenium)."
    )
    parser.add_argument("--saldos", required=True, help="Caminho do Excel com a base de saldos (ex: data/saldos.xlsx)")
    parser.add_argument("--contatos", required=True, help="Caminho do Excel com a base de contatos (ex: data/contatos.xlsx)")
    parser.add_argument("--out-dir", default="output", help="Pasta para salvar as imagens geradas (default: output)")
    parser.add_argument("--headless", action="store_true", help="Rodar Chrome em modo headless (nem sempre funciona bem com WhatsApp)")
    parser.add_argument("--dry-run", action="store_true", help="Não envia nada; só gera as imagens e imprime o que faria")
    parser.add_argument("--sleep-between", type=int, default=3, help="Pausa entre envios em segundos (default: 3)")
    return parser.parse_args()


def format_brl(value) -> str:
    """Formatação simples estilo BR (sem depender de locale do sistema)."""
    try:
        v = float(value)
    except Exception:
        return ""
    s = f"{v:,.2f}"          # 1,234.56
    s = s.replace(",", "X")  # 1X234.56
    s = s.replace(".", ",")  # 1X234,56
    s = s.replace("X", ".")  # 1.234,56
    return f"R$ {s}"


def ensure_out_dir(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)


# =========================
# Leitura e preparação de dados
# =========================

def load_saldos(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Ajuste aqui se as colunas do seu Excel tiverem nomes diferentes
    required = {"Assessor"}
    if not required.issubset(df.columns):
        raise ValueError(f"Base de saldos precisa ter a coluna {required}. Colunas encontradas: {list(df.columns)}")

    return df


def load_contatos(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    # Ajuste aqui se as colunas do seu Excel tiverem nomes diferentes
    required = {"codigo", "numero"}
    if not required.issubset(df.columns):
        raise ValueError(f"Base de contatos precisa ter colunas {required}. Colunas encontradas: {list(df.columns)}")

    # normaliza codigo e numero
    df["codigo"] = df["codigo"].astype(str).str.strip()
    df["numero"] = df["numero"].astype(str).str.strip()
    return df


def find_phone(contatos_df: pd.DataFrame, assessor_code) -> Optional[str]:
    code = str(assessor_code).strip()
    m = contatos_df[contatos_df["codigo"] == code]
    if m.empty:
        return None
    phone = m.iloc[0]["numero"]
    if not phone or phone.lower() == "nan":
        return None
    return phone


# =========================
# Geração da imagem (tabela)
# =========================

def generate_table_image(df_group: pd.DataFrame, out_path: Path) -> None:
    # Copia para não alterar df original
    data = df_group.copy()

    # Formata colunas monetárias, se existirem
    money_cols = ["D0", "D+1", "Total"]
    for c in money_cols:
        if c in data.columns:
            data[c] = data[c].apply(format_brl)

    # Configura altura proporcional ao tamanho
    largura_total = 2000
    altura = max(1.8, len(data) * 0.35)

    plt.figure(figsize=(largura_total / 100, altura))
    table = plt.table(
        cellText=data.values,
        colLabels=data.columns,
        cellLoc="center",
        loc="center",
        bbox=[0, 0, 1, 1],
    )

    # Cabeçalho em negrito
    for i in range(len(data.columns)):
        table[0, i].set_text_props(weight="bold")

    table.auto_set_font_size(False)
    table.set_fontsize(12)

    plt.axis("off")
    plt.subplots_adjust(left=0.02, right=0.98, top=0.98, bottom=0.02)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(out_path, bbox_inches="tight", pad_inches=0.3)
    plt.close()


# =========================
# WhatsApp Web (Selenium)
# =========================

def build_driver(headless: bool) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")

    # Dica: você pode fixar um perfil para evitar logar sempre:
    # options.add_argument(r"--user-data-dir=C:\temp\chrome-whatsapp-profile")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_window_size(1280, 900)
    return driver


def wait_whatsapp_ready(driver: webdriver.Chrome, timeout: int = 120) -> None:
    driver.get("https://web.whatsapp.com/")
    # Espera aparecer a lateral (indica que logou)
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, "side"))
    )


def open_chat_with_prefill(driver: webdriver.Chrome, phone: str, text: str) -> None:
    import urllib.parse
    encoded = urllib.parse.quote(text)
    driver.get(f"https://web.whatsapp.com/send?phone={phone}&text={encoded}")

    # Espera campo de digitação aparecer
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="main"]//footer//p'))
    )


def send_enter(driver: webdriver.Chrome) -> None:
    # Pressiona Enter no campo
    field = driver.find_element(By.XPATH, '//*[@id="main"]//footer//p')
    field.send_keys(Keys.ENTER)


def attach_image(driver: webdriver.Chrome, image_path: Path) -> None:
    """
    Seletores do WhatsApp mudam com frequência.
    Mantive um approach comum: procurar input file de anexos.
    Se quebrar, o README explica como ajustar.
    """
    # Tenta clicar no botão de anexo (clipe)
    try:
        clip_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='plus'], span[data-icon='clip']"))
        )
        clip_btn.click()
    except Exception:
        # Alguns layouts não precisam clicar no clipe se o input file já existe
        pass

    # Procura pelo input file
    file_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
    )
    file_input.send_keys(str(image_path.resolve()))

    # Espera o botão de enviar (aviãozinho) e clica
    send_btn = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send'], span[data-icon='send-light']"))
    )
    send_btn.click()


# =========================
# Execução
# =========================

def main() -> int:
    args = parse_args()

    saldos_df = load_saldos(args.saldos)
    contatos_df = load_contatos(args.contatos)

    out_dir = Path(args.out_dir)
    ensure_out_dir(out_dir)

    # Agrupa por assessor
    groups = saldos_df.groupby("Assessor")

    # Driver (somente se não for dry-run)
    driver = None
    if not args.dry_run:
        driver = build_driver(args.headless)
        print("Abrindo WhatsApp Web... faça login/QR Code se necessário.")
        wait_whatsapp_ready(driver)

    try:
        for assessor_code, group in groups:
            phone = find_phone(contatos_df, assessor_code)
            if not phone:
                print(f"[SKIP] Assessor {assessor_code}: sem telefone na base.")
                continue

            image_path = out_dir / f"imagem_assessor_{assessor_code}.png"
            generate_table_image(group, image_path)

            message = (
                f"Segue posição de clientes com saldo negativo para o assessor {assessor_code}.\n"
                f"(Mensagem automática)\n"
            )

            if args.dry_run:
                print(f"[DRY-RUN] Enviaria para {assessor_code} ({phone}) | imagem: {image_path}")
                continue

            # Abre chat com mensagem pré-preenchida e envia
            open_chat_with_prefill(driver, phone, message)
            send_enter(driver)
            time.sleep(1)

            # Anexa imagem e envia
            attach_image(driver, image_path)

            print(f"[OK] Enviado para assessor {assessor_code} ({phone}).")
            time.sleep(args.sleep_between)

        return 0

    finally:
        if driver:
            driver.quit()


if __name__ == "__main__":
    raise SystemExit(main())
