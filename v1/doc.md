# Guia de Execução da Automação RPA-CCEE

Para que a automação funcione corretamente sem o erro *"target window already closed"*, é obrigatório seguir esta ordem lógica estrita. A automação depende de encontrar um navegador **já aberto** e **logado**.

## 1. Configuração (Antes de abrir qualquer coisa)
Edite o arquivo `config.py` para definir o alvo da extração.
* **MES**: O mês do relatório (ex: `"12"`).
* **ANO**: O ano do relatório (ex: `"2025"`).
* **NOME_RELATORIO**: O código do relatório (ex: `"RCAP002"`).
* **empresas_alvo**: Verifique se a lista de empresas está atualizada.

---

## 2. Inicialização do Ambiente
Este passo prepara o navegador para ser controlado pelo robô.

1.  **Feche todos os Chromes**: Garanta que não há nenhuma janela do Google Chrome aberta no computador.
2.  **Execute o Arquivo BAT**:
    * Dê duplo clique no arquivo `iniciar_robo.bat`.
    * *O que vai acontecer:* Uma janela branca/nova do Chrome irá abrir. Não a feche!

---

## 3. Preparação Manual (Login)
Com a janela do Chrome que o `.bat` abriu:

1.  Acesse ao site da CCEE (ex: `ccee.org.br` ou o link direto do BI).
2.  **Faça o Login**: Insira usuário, senha e resolva o Captcha manualmente.
3.  **Navegue até ao Relatório**:
    * Vá até ao Painel de Controle / Dashboard onde está o relatório desejado.
    * Deixe a tela pronta, exibindo os filtros (Ano/Mês/Agente), exatamente como na imagem que enviou.
    * **NÃO FECHE ESTA JANELA.**

---

## 4. Execução do Robô
Agora que o Chrome está pronto e esperando na porta 9222:

1.  Abra o seu terminal ou VS Code.
2.  Execute o script Python:
    ```bash
    python app.py
    ```
    *(Ou clique no botão "Play" do seu editor)*.

---

## 5. Aguardar (Crucial)
Assim que o script exibir a mensagem *"Iniciando extração..."*:

* 🛑 **SOLTE O RATO E O TECLADO.**
* Não clique em outras janelas.
* Não minimize o Chrome.
* *Motivo:* O robô usa o `PyAutoGUI` para salvar o PDF (janela de "Salvar Como"). Se você mexer no mouse, ele vai clicar no lugar errado e o processo vai falhar.

---

## Resumo Gráfico do Fluxo

`config.py` (Editar) ➔ `iniciar_robo.bat` (Rodar) ➔ **Chrome** (Logar Manualmente) ➔ `app.py` (Executar) ➔ 🛑 (Não mexer)