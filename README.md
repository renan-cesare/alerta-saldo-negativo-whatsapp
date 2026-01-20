# Alerta de Saldo Negativo via WhatsApp (Selenium + Pandas)

Automação em Python que **gera uma imagem por assessor** com a lista de clientes em **saldo negativo** (a partir de uma planilha) e **envia via WhatsApp Web** usando Selenium.

> Projeto profissional **sanitizado**: sem dados reais, sem telefones reais, sem nomes internos e sem base de clientes.

---

## Contexto

Em rotinas operacionais, clientes com saldo negativo podem sofrer encargos/multas. Para acelerar a regularização, era necessário avisar diariamente os assessores com a lista de clientes que precisavam de atenção.

Este script automatiza esse fluxo:

* lê base de saldos
* separa por assessor
* gera uma imagem (tabela)
* envia mensagem + imagem no WhatsApp Web

---

## O que o projeto faz

1. Lê um Excel de **saldos** (com coluna `Assessor`)
2. Lê um Excel de **contatos** (`codigo`, `numero`)
3. Agrupa por assessor
4. Gera uma imagem PNG com a tabela de cada assessor
5. Abre o WhatsApp Web e envia:

   * mensagem padrão
   * anexo com a imagem do assessor

---

## Tecnologias

* Python
* Pandas
* Matplotlib (geração de imagem/tabela)
* Selenium + WebDriver Manager (WhatsApp Web)

---

## Instalação

```bash
pip install -r requirements.txt
```

---

## Como usar

### 1) Prepare os arquivos

Você precisa de dois Excel:

**Base de saldos (`--saldos`)**

* obrigatória a coluna: `Assessor`
* pode ter colunas monetárias como: `D0`, `D+1`, `Total` (opcional)

**Base de contatos (`--contatos`)**

* colunas obrigatórias:

  * `codigo` (código do assessor)
  * `numero` (telefone no formato internacional, ex.: `5511999999999`)

> Importante: não suba seus arquivos reais no GitHub. Use exemplos fake se quiser manter modelos.

---

### 2) Rodar em modo de teste (recomendado)

Gera imagens e imprime o que faria, mas não envia:

```bash
python main.py --saldos data/saldos_exemplo.xlsx --contatos data/contatos_exemplo.xlsx --dry-run
```

---

### 3) Rodar enviando de verdade

```bash
python main.py --saldos data/saldos.xlsx --contatos data/contatos.xlsx
```

Na primeira vez, o Chrome vai abrir e você precisa **logar no WhatsApp Web** (QR Code). Depois disso, ele segue o envio automático.

---

## Observações importantes (WhatsApp Web)

* O WhatsApp Web muda o HTML com frequência. Se algum seletor quebrar, ajuste a função `attach_image()` e/ou os seletores de espera.
* Automação via WhatsApp Web pode estar sujeita a regras/limitações do próprio WhatsApp. Use com responsabilidade.

---

## Parâmetros úteis

* `--out-dir` pasta onde as imagens serão geradas (default: `output`)
* `--sleep-between` pausa entre envios (default: 3s)
* `--dry-run` não envia nada (apenas gera e simula)
* `--headless` roda sem interface (não recomendado para WhatsApp, mas disponível)

---

## Limitações

* Depende do WhatsApp Web estar acessível e estável
* Seletores podem mudar (exigindo manutenção)
* Não é “API oficial”; é automação de navegador

---

## Licença

MIT.
